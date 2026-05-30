"""Transparent overlay window for PPT Touch Controller.

Creates a frameless, always-on-top, semi-transparent window that floats
above the PowerPoint slideshow. The window is click-through except on
button areas, where touch/click events are captured and translated to
slide navigation commands.

Key Windows API tricks:
- WS_EX_LAYERED + WS_EX_TRANSPARENT: Makes window transparent and
  click-through by default
- WM_NCHITTEST subclassing: Returns HTTRANSPARENT for transparent areas
  (pass through to PowerPoint) and HTCLIENT for button areas (capture)
- Long-press detection: 500ms hold → drag mode to reposition buttons
"""

import ctypes
from ctypes import wintypes
from PySide6.QtCore import (
    Qt, QTimer, QPoint, QSize, Signal, QEasingCurve,
    QPropertyAnimation,
)
from PySide6.QtGui import (
    QPainter, QColor, QFont, QPen, QBrush, QPainterPath, QFontMetrics,
    QMouseEvent, QTouchEvent,
)
from PySide6.QtWidgets import (
    QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QLabel,
    QGraphicsOpacityEffect, QSizePolicy, QSlider, QDialog,
    QDialogButtonBox, QFormLayout, QComboBox, QSpinBox,
)

from settings_manager import SettingsManager

# Windows API constants
WS_EX_LAYERED = 0x00080000
WS_EX_TRANSPARENT = 0x00000020
WS_EX_TOOLWINDOW = 0x00000080
GWL_EXSTYLE = -20

WM_NCHITTEST = 0x0084
HTCLIENT = 1
HTTRANSPARENT = -1
HTCAPTION = 2

# Window style helpers
SetWindowLongPtrW = ctypes.windll.user32.SetWindowLongPtrW
GetWindowLongPtrW = ctypes.windll.user32.GetWindowLongPtrW

LONG_PTR = ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long
SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, LONG_PTR]
SetWindowLongPtrW.restype = LONG_PTR
GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
GetWindowLongPtrW.restype = LONG_PTR


# MSG structure for nativeEvent
class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt_x", ctypes.c_long),
        ("pt_y", ctypes.c_long),
    ]


class TouchButton(QPushButton):
    """Large touch-friendly circular button.

    Features:
    - Semi-transparent dark background with white arrow
    - Long-press detection (500ms) for drag mode
    - Visual feedback: scale up on press, highlight on long-press
    """

    long_pressed = Signal()
    drag_moved = Signal(QPoint)  # Emitted during drag (delta)

    LONG_PRESS_MS = 500

    def __init__(self, text: str, size: int = 80, parent=None):
        super().__init__(text, parent)
        self._btn_size = size
        self.setFixedSize(size, size)
        self.setCursor(Qt.PointingHandCursor)

        # Touch/long-press state
        self._press_timer = QTimer(self)
        self._press_timer.setSingleShot(True)
        self._press_timer.timeout.connect(self._on_long_press)
        self._long_press_active = False
        self._is_dragging = False
        self._press_start_pos = QPoint()
        self._opacity = 0.75

    def set_opacity(self, value: float) -> None:
        """Set button opacity (0.0 - 1.0)."""
        self._opacity = max(0.3, min(1.0, value))
        self.update()

    def paintEvent(self, event) -> None:
        """Draw semi-transparent circular button with text."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        s = self._btn_size
        margin = 3
        r = (s - 2 * margin) / 2
        cx, cy = s / 2, s / 2

        # Background circle
        bg_alpha = int(self._opacity * 200)
        if self._long_press_active:
            bg_alpha = int(self._opacity * 255)  # Full opacity during drag
            painter.setPen(QPen(QColor(255, 255, 255, 200), 4))
        else:
            painter.setPen(QPen(QColor(255, 255, 255, 150), 3))

        painter.setBrush(QColor(0, 0, 0, bg_alpha))
        painter.drawEllipse(QPoint(int(cx), int(cy)), int(r), int(r))

        # Arrow text
        font = QFont("Segoe UI", int(s * 0.35), QFont.Bold)
        painter.setFont(font)
        painter.setPen(QColor(255, 255, 255, int(self._opacity * 255)))
        painter.drawText(self.rect(), Qt.AlignCenter, self.text())

    def mousePressEvent(self, event: QMouseEvent) -> None:
        """Record press position, start long-press timer."""
        if event.button() == Qt.LeftButton:
            self._press_start_pos = event.globalPosition().toPoint()
            self._is_dragging = False
            self._long_press_active = False
            self._press_timer.start(self.LONG_PRESS_MS)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent) -> None:
        """Handle drag movement."""
        if self._long_press_active:
            current_pos = event.globalPosition().toPoint()
            delta = current_pos - self._press_start_pos
            self._press_start_pos = current_pos
            self._is_dragging = True
            self.drag_moved.emit(delta)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        """Cancel long-press timer, emit click if not dragging."""
        self._press_timer.stop()
        was_dragging = self._is_dragging
        was_long_press = self._long_press_active
        self._is_dragging = False
        self._long_press_active = False

        if not was_dragging and not was_long_press:
            # Normal click
            super().mouseReleaseEvent(event)
        else:
            # Was a drag or long press, don't emit clicked
            event.accept()
        self.update()

    def _on_long_press(self) -> None:
        """Long press detected: enter drag mode."""
        self._long_press_active = True
        self.long_pressed.emit()
        self.update()


class SettingsDialog(QDialog):
    """Settings panel for customizing button appearance."""

    def __init__(self, settings: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("触控按钮设置")
        self.setWindowFlags(
            Qt.Dialog | Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint
        )
        self.setFixedSize(320, 280)
        self._settings = settings
        self._setup_ui()

    def _setup_ui(self) -> None:
        layout = QFormLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        # Button size
        self._size_slider = QSlider(Qt.Horizontal)
        self._size_slider.setRange(60, 120)
        self._size_slider.setValue(self._settings.get("button_size", 80))
        self._size_slider.setTickPosition(QSlider.TicksBelow)
        self._size_slider.setTickInterval(10)
        layout.addRow("按钮大小:", self._size_slider)

        # Opacity
        self._opacity_slider = QSlider(Qt.Horizontal)
        self._opacity_slider.setRange(30, 100)
        self._opacity_slider.setValue(self._settings.get("button_opacity", 75))
        self._opacity_slider.setTickPosition(QSlider.TicksBelow)
        self._opacity_slider.setTickInterval(10)
        layout.addRow("透明度:", self._opacity_slider)

        # Color
        self._color_combo = QComboBox()
        self._color_combo.addItems(["蓝", "绿", "灰", "白"])
        default_color = self._settings.get("button_color", "#0078D4")
        color_map = {
            "#0078D4": 0, "#107C10": 1, "#555555": 2, "#FFFFFF": 3,
        }
        self._color_combo.setCurrentIndex(color_map.get(default_color, 0))
        layout.addRow("颜色主题:", self._color_combo)

        # Hand mode
        self._hand_combo = QComboBox()
        self._hand_combo.addItems(["右手 (下一页在右)", "左手 (下一页在左)"])
        hand_idx = 0 if self._settings.get("hand_mode", "right") == "right" else 1
        self._hand_combo.setCurrentIndex(hand_idx)
        layout.addRow("手模式:", self._hand_combo)

        # Buttons
        btn_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Reset
        )
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        btn_box.button(QDialogButtonBox.Reset).clicked.connect(self._reset)
        layout.addRow(btn_box)

        self.setStyleSheet("""
            QDialog {
                background: #2D2D2D;
                border: 2px solid #555;
                border-radius: 10px;
                color: white;
                font-size: 14px;
            }
            QLabel { color: white; }
            QSlider::groove:horizontal {
                height: 8px; background: #555; border-radius: 4px;
            }
            QSlider::handle:horizontal {
                width: 24px; height: 24px; margin: -8px 0;
                background: #0078D4; border-radius: 12px;
            }
            QComboBox {
                background: #444; color: white; border: 1px solid #666;
                padding: 4px 8px; border-radius: 4px;
            }
            QPushButton {
                background: #444; color: white; border: 1px solid #666;
                padding: 6px 16px; border-radius: 4px;
            }
            QPushButton:hover { background: #555; }
        """)

    def _reset(self) -> None:
        """Reset all values to defaults."""
        defaults = SettingsManager.DEFAULTS
        self._size_slider.setValue(defaults["button_size"])
        self._opacity_slider.setValue(defaults["button_opacity"])
        self._color_combo.setCurrentIndex(0)
        self._hand_combo.setCurrentIndex(0)

    def get_values(self) -> dict:
        """Get current dialog values as settings dict."""
        color_map = {0: "#0078D4", 1: "#107C10", 2: "#555555", 3: "#FFFFFF"}
        hand_map = {0: "right", 1: "left"}
        return {
            "button_size": self._size_slider.value(),
            "button_opacity": self._opacity_slider.value(),
            "button_color": color_map[self._color_combo.currentIndex()],
            "hand_mode": hand_map[self._hand_combo.currentIndex()],
        }


class OverlayWindow(QWidget):
    """Transparent overlay floating above PowerPoint slideshow.

    Provides large touch-friendly prev/next buttons. Background is
    invisible and click-through (via WM_NCHITTEST handled in nativeEvent).
    Buttons capture touches. Long-press a button to drag the entire
    overlay to a new position.
    """

    next_requested = Signal()
    prev_requested = Signal()
    exit_requested = Signal()

    def __init__(self, settings: dict = None):
        super().__init__()
        self._settings = settings or SettingsManager.load()
        self._drag_active = False

        self._setup_window()
        self._create_buttons()
        self._apply_position()
        self._apply_style()

    def _setup_window(self) -> None:
        """Configure frameless, always-on-top, tool window."""
        self.setWindowFlags(
            Qt.FramelessWindowHint
            | Qt.WindowStaysOnTopHint
            | Qt.Tool
            | Qt.NoDropShadowWindowHint
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WA_AcceptTouchEvents)

        # Overall window size depends on button size and layout
        btn_size = self._settings.get("button_size", 80)
        self.setFixedSize(btn_size * 2 + 40, btn_size + 20)

    def showEvent(self, event):
        """Apply WS_EX_LAYERED after the native window is created."""
        super().showEvent(event)
        hwnd = int(self.winId())
        ex_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        ex_style |= WS_EX_LAYERED | WS_EX_TOOLWINDOW
        # Ensure WS_EX_TRANSPARENT is NOT set so buttons can receive clicks
        ex_style &= ~WS_EX_TRANSPARENT
        SetWindowLongPtrW(hwnd, GWL_EXSTYLE, ex_style)

    def nativeEvent(self, eventType, message):
        """Handle WM_NCHITTEST to implement click-through.

        Qt's native event handler — no ctypes WNDPROC subclassing needed.
        Returns HTTRANSPARENT for non-button areas (pass-through to
        PowerPoint), HTCLIENT for button areas (capture clicks).
        """
        msg = ctypes.cast(
            int(message), ctypes.POINTER(MSG)
        ).contents

        if msg.message == WM_NCHITTEST:
            # Extract screen coordinates from lParam
            x = msg.lParam & 0xFFFF
            y = (msg.lParam >> 16) & 0xFFFF

            # Map to widget-local coordinates
            local_pos = self.mapFromGlobal(QPoint(x, y))

            # Check if position is over a button
            if hasattr(self, '_prev_btn') and hasattr(self, '_next_btn'):
                prev_geo = self._prev_btn.geometry()
                next_geo = self._next_btn.geometry()

                if prev_geo.contains(local_pos) or next_geo.contains(local_pos):
                    return False, HTCLIENT  # Capture click on buttons
                elif self._drag_active:
                    return False, HTCAPTION  # Allow window drag

            # Pass through to window underneath (PowerPoint)
            return True, HTTRANSPARENT

        return False, 0  # Let Qt handle all other messages

    def _create_buttons(self) -> None:
        """Create prev/next buttons based on hand mode."""
        btn_size = self._settings.get("button_size", 80)
        hand_mode = self._settings.get("hand_mode", "right")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(12)

        self._prev_btn = TouchButton("<", btn_size, self)
        self._next_btn = TouchButton(">", btn_size, self)

        if hand_mode == "left":
            # Left-handed: next on left, prev on right
            layout.addWidget(self._next_btn)
            layout.addWidget(self._prev_btn)
        else:
            # Right-handed (default): prev on left, next on right
            layout.addWidget(self._prev_btn)
            layout.addWidget(self._next_btn)

        # Connect signals
        self._prev_btn.clicked.connect(self.prev_requested.emit)
        self._next_btn.clicked.connect(self.next_requested.emit)

        # Drag support
        self._prev_btn.drag_moved.connect(self._on_button_drag)
        self._prev_btn.long_pressed.connect(self._on_long_press_start)
        self._next_btn.drag_moved.connect(self._on_button_drag)
        self._next_btn.long_pressed.connect(self._on_long_press_start)

        self._apply_button_opacity()

    def _apply_button_opacity(self) -> None:
        """Apply opacity setting to both buttons."""
        opacity = self._settings.get("button_opacity", 75) / 100.0
        self._prev_btn.set_opacity(opacity)
        self._next_btn.set_opacity(opacity)

    def _apply_position(self) -> None:
        """Position the overlay window based on saved settings or auto-detect."""
        pos = self._settings.get("overlay_position", {})
        saved_x = pos.get("x")
        saved_y = pos.get("y")

        if saved_x is not None and saved_y is not None:
            self.move(saved_x, saved_y)
        else:
            # Default: bottom-center of screen
            screen = self.screen()
            if screen:
                screen_geo = screen.availableGeometry()
                x = (screen_geo.width() - self.width()) // 2 + screen_geo.x()
                y = screen_geo.height() - self.height() - 30 + screen_geo.y()
                self.move(x, y)

    def _apply_style(self) -> None:
        """Apply stylesheet — transparent for click-through."""
        self.setStyleSheet("background: transparent;")

    def _on_button_drag(self, delta: QPoint) -> None:
        """Move the entire overlay window when a button is dragged."""
        self._drag_active = True
        new_pos = self.pos() + delta
        self.move(new_pos)

    def _on_long_press_start(self) -> None:
        """Visual feedback when long press detected."""
        pass  # Buttons handle their own visual feedback

    def mouseReleaseEvent(self, event: QMouseEvent) -> None:
        """End drag mode and save position."""
        if self._drag_active:
            self._drag_active = False
            # Save position
            pos = self.pos()
            SettingsManager.update(
                "overlay_position", {"x": pos.x(), "y": pos.y()}
            )
        super().mouseReleaseEvent(event)

    def save_position(self) -> None:
        """Persist current window position to settings."""
        pos = self.pos()
        SettingsManager.update(
            "overlay_position", {"x": pos.x(), "y": pos.y()}
        )

    def show_settings_dialog(self) -> dict:
        """Open settings dialog and return updated settings if accepted.

        Returns:
            Updated settings dict, or None if cancelled.
        """
        dialog = SettingsDialog(self._settings, self)
        if dialog.exec() == QDialog.Accepted:
            new_values = dialog.get_values()
            self._settings.update(new_values)
            SettingsManager.save(self._settings)
            self._recreate_buttons()
            return new_values
        return None

    def _recreate_buttons(self) -> None:
        """Rebuild buttons after settings change."""
        # Remove old layout
        old_layout = self.layout()
        if old_layout:
            while old_layout.count():
                item = old_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            old_layout.deleteLater()

        btn_size = self._settings.get("button_size", 80)
        hand_mode = self._settings.get("hand_mode", "right")
        btn_spacing = max(8, int(btn_size * 0.15))
        # Window padding
        pad = 10
        self.setFixedSize(btn_size * 2 + 2 * pad + btn_spacing, btn_size + 2 * pad)

        self._create_buttons()
        self._apply_position()

    def set_button_size(self, size: int) -> None:
        """Update button size at runtime."""
        self._settings["button_size"] = size
        self._recreate_buttons()
        SettingsManager.update("button_size", size)

    def set_hand_mode(self, mode: str) -> None:
        """Switch hand mode (left/right)."""
        self._settings["hand_mode"] = mode
        self._recreate_buttons()
        SettingsManager.update("hand_mode", mode)
