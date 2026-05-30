"""PPT Touch Controller - Main Entry Point.

A touch-screen-friendly PowerPoint slideshow controller for teaching.
Creates a transparent overlay with customizable prev/next buttons that
float above a PowerPoint fullscreen slideshow.

Usage:
    python main.py <path_to_pptx_file>
    python main.py                    (opens file dialog)

The application:
1. Detects PowerPoint on the system
2. Opens the PPTX file and starts a fullscreen slideshow via COM
3. Creates a transparent overlay with large touch-friendly buttons
4. Monitors slideshow status and exits gracefully when slideshow ends
"""

import sys
import os
import logging
from pathlib import Path

from PySide6.QtCore import (
    Qt, QTimer, QThread, QSharedMemory, QDataStream,
)
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import (
    QApplication, QMessageBox, QFileDialog, QMenu,
)
from PySide6.QtGui import QIcon, QAction

from settings_manager import SettingsManager
from ppt_controller import PPTController
from overlay_window import OverlayWindow

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("PPTouchController")

# Single-instance identifiers
SHARED_MEMORY_KEY = "PPTTouchController_Instance_v1"
IPC_SERVER_NAME = "PPTTouchController_IPC_v1"


def detect_powerpoint() -> bool:
    """Check if PowerPoint is installed via Windows registry."""
    import winreg
    progids = [
        "PowerPoint.Application",
        "PowerPoint.Application.16",
        "PowerPoint.Application.15",
        "PPTViewer.Application",
    ]
    for progid in progids:
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID")
            winreg.CloseKey(key)
            logger.info(f"Found PowerPoint via {progid}")
            return True
        except OSError:
            continue
    return False


class PPTTouchApp(QApplication):
    """Main application with single-instance enforcement."""

    def __init__(self, argv):
        super().__init__(argv)
        self.setApplicationName("PPT Touch Controller")
        self.setApplicationVersion("1.0.0")
        self.setOrganizationName("PPTTouchController")

        # Single-instance check
        self._memory = QSharedMemory(SHARED_MEMORY_KEY)
        self._is_primary = False
        self._ipc_server = None

        # Core components
        self._com_thread = None
        self._com_worker = None
        self._overlay = None
        self._status_timer = None
        self._file_path = None
        self._slideshow_active = False

        # Cleanup on exit
        self.aboutToQuit.connect(self._cleanup_slideshow)

    def try_become_primary(self) -> bool:
        """Attempt to become the primary instance."""
        if self._memory.create(1):
            self._is_primary = True
            self._start_ipc_server()
            return True
        else:
            self._is_primary = False
            return False

    def send_to_primary(self, file_path: str) -> bool:
        """Send a file path to the already-running primary instance."""
        socket = QLocalSocket(self)
        socket.connectToServer(IPC_SERVER_NAME)
        if socket.waitForConnected(2000):
            stream = QDataStream(socket)
            stream.setVersion(QDataStream.Qt_6_0)
            data = file_path.encode("utf-8")
            stream.writeInt32(len(data))
            stream.writeRawData(data)
            socket.waitForBytesWritten(2000)
            socket.disconnectFromServer()
            logger.info(f"Sent file path to primary instance: {file_path}")
            return True
        else:
            logger.warning("Failed to connect to primary instance")
            return False

    def _start_ipc_server(self) -> None:
        """Start local server to receive file paths from secondary instances."""
        self._ipc_server = QLocalServer(self)
        QLocalServer.removeServer(IPC_SERVER_NAME)
        if self._ipc_server.listen(IPC_SERVER_NAME):
            self._ipc_server.newConnection.connect(self._on_ipc_connection)
            logger.info("IPC server started")

    def _on_ipc_connection(self) -> None:
        """Handle incoming connection from a secondary instance."""
        socket = self._ipc_server.nextPendingConnection()
        if socket and socket.waitForReadyRead(2000):
            stream = QDataStream(socket)
            stream.setVersion(QDataStream.Qt_6_0)
            length = stream.readInt32()
            data = stream.readRawData(length)
            file_path = data.decode("utf-8")
            socket.disconnectFromServer()
            logger.info(f"Received file path via IPC: {file_path}")
            self._open_file(file_path)

    def run(self, file_path: str = None) -> int:
        """Main entry point after initialization."""
        if not self._is_primary:
            if file_path:
                self.send_to_primary(file_path)
            return 0

        if file_path:
            self._open_file(file_path)
        else:
            self._show_welcome()

        return self.exec()

    def _open_file(self, file_path: str) -> None:
        """Open a PPTX file and start the slideshow with overlay."""
        path = Path(file_path)
        if not path.exists():
            QMessageBox.critical(
                None, "文件未找到",
                f"找不到文件:\n{file_path}"
            )
            return

        if path.suffix.lower() not in (".pptx", ".ppt"):
            QMessageBox.warning(
                None, "不支持的文件格式",
                f"请选择 .pptx 或 .ppt 文件。\n当前文件: {path.name}"
            )
            return

        SettingsManager.update("last_directory", str(path.parent))

        if not detect_powerpoint():
            self._handle_no_powerpoint(file_path)
            return

        self._file_path = file_path

        # Shut down previous slideshow if any
        self._cleanup_slideshow()

        # Create overlay window
        settings = SettingsManager.load()
        self._overlay = OverlayWindow(settings)
        self._overlay.show()

        # Create COM worker and thread
        self._com_thread = QThread(self)
        self._com_worker = PPTController()
        self._com_worker.moveToThread(self._com_thread)

        # Connect signals (cross-thread via Qt.QueuedConnection)
        self._overlay.next_requested.connect(
            self._com_worker.next_step, Qt.QueuedConnection
        )
        self._overlay.prev_requested.connect(
            self._com_worker.prev_step, Qt.QueuedConnection
        )
        self._com_worker.slideshow_started.connect(
            self._on_slideshow_started, Qt.QueuedConnection
        )
        self._com_worker.slideshow_ended.connect(
            self._on_slideshow_ended, Qt.QueuedConnection
        )
        self._com_worker.error_occurred.connect(
            self._on_error, Qt.QueuedConnection
        )

        # Start COM thread
        self._com_thread.started.connect(
            lambda fp=file_path: self._com_worker.open_and_start(fp)
        )
        self._com_thread.start()
        self._slideshow_active = True

        # Start polling slideshow status
        self._status_timer = QTimer(self)
        self._status_timer.timeout.connect(self._poll_slideshow_status)
        self._status_timer.start(500)

        logger.info(f"Opening file: {file_path}")

    def _on_slideshow_started(self) -> None:
        """Called when PowerPoint slideshow has started."""
        logger.info("Slideshow started successfully")
        if self._overlay:
            self._overlay.raise_()
            self._overlay.show()

    def _on_slideshow_ended(self) -> None:
        """Called when PowerPoint slideshow ends (from COM worker signal)."""
        logger.info("Slideshow ended signal received")
        self._slideshow_active = False
        if self._overlay:
            self._overlay.hide()
        # Don't do full cleanup here - let polling handle it

    def _on_error(self, error_msg: str) -> None:
        """Handle errors from COM worker."""
        logger.error(f"COM error: {error_msg}")
        QMessageBox.critical(
            self._overlay, "PowerPoint 错误",
            f"无法控制 PowerPoint:\n\n{error_msg}\n\n"
            "请确保 PowerPoint 已安装并能正常打开此文件。"
        )
        self._slideshow_active = False
        self._cleanup_slideshow()

    def _poll_slideshow_status(self) -> None:
        """Periodically check if slideshow thread is still running."""
        if not self._slideshow_active:
            # Already ended, clean up
            self._cleanup_slideshow()
            return

        if self._com_thread and not self._com_thread.isRunning():
            # Thread stopped (slideshow ended or crashed)
            logger.info("COM thread stopped, cleaning up")
            self._slideshow_active = False
            if self._overlay:
                self._overlay.hide()
            self._cleanup_slideshow()

    def _cleanup_slideshow(self) -> None:
        """Stop slideshow monitoring, clean up overlay and COM thread."""
        self._slideshow_active = False

        # Stop timer
        if self._status_timer:
            self._status_timer.stop()
            self._status_timer.deleteLater()
            self._status_timer = None

        # Hide and clean overlay
        if self._overlay:
            try:
                self._overlay.save_position()
            except Exception:
                pass
            self._overlay.hide()
            self._overlay.deleteLater()
            self._overlay = None

        # Clean COM worker
        if self._com_worker:
            try:
                self._com_worker.shutdown()
            except Exception:
                pass
            self._com_worker.deleteLater()
            self._com_worker = None

        # Stop COM thread
        if self._com_thread:
            try:
                self._com_thread.quit()
                self._com_thread.wait(3000)
            except Exception:
                pass
            self._com_thread.deleteLater()
            self._com_thread = None

    def _handle_no_powerpoint(self, file_path: str) -> None:
        """Fallback when PowerPoint is not installed."""
        reply = QMessageBox.question(
            None,
            "未检测到 PowerPoint",
            "系统中未找到 Microsoft PowerPoint。\n\n"
            "是否使用简易查看器打开？\n"
            "(简易查看器不支持动画效果)\n\n"
            "你也可以从 Microsoft Store 下载免费的 "
            "PowerPoint 移动版。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply == QMessageBox.Yes:
            self._open_fallback_viewer(file_path)

    def _open_fallback_viewer(self, file_path: str) -> None:
        """Open a simple static slide viewer using python-pptx."""
        try:
            from fallback_viewer import FallbackViewer
            viewer = FallbackViewer(file_path)
            viewer.show()
            self._fallback_viewer = viewer
        except Exception as e:
            QMessageBox.critical(
                None, "查看器错误",
                f"无法打开简易查看器:\n{e}"
            )

    def _show_welcome(self) -> None:
        """Show welcome dialog when no file path provided."""
        msg = QMessageBox()
        msg.setWindowTitle("PPT Touch Controller")
        msg.setText(
            "<h2>PPT 触控控制器</h2>"
            "<p>专为教学触控屏设计的 PPT 查看工具。</p>"
            "<p><b>使用方法:</b></p>"
            "<ul>"
            "<li>将 .pptx 文件拖放到此程序图标上</li>"
            "<li>或将 .pptx 文件关联到此程序后双击打开</li>"
            "<li>或点击下方按钮选择文件</li>"
            "</ul>"
            "<p>全屏播放时，屏幕底部会显示触控按钮。</p>"
            "<p>长按按钮可拖拽到任意位置。</p>"
        )
        open_btn = msg.addButton("打开 PPTX 文件...", QMessageBox.ActionRole)
        msg.addButton(QMessageBox.Close)
        msg.exec()

        if msg.clickedButton() == open_btn:
            file_path, _ = QFileDialog.getOpenFileName(
                None,
                "选择 PowerPoint 文件",
                SettingsManager.load().get("last_directory", ""),
                "PowerPoint 文件 (*.pptx *.ppt);;所有文件 (*.*)",
            )
            if file_path:
                self._open_file(file_path)


def main():
    """Application entry point."""
    # Parse command line
    file_path = None
    if len(sys.argv) > 1:
        file_path = " ".join(sys.argv[1:])

    # High-DPI support for touch screens
    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = PPTTouchApp(sys.argv)

    if not app.try_become_primary():
        if file_path:
            app.send_to_primary(file_path)
        return 0

    return app.run(file_path)


if __name__ == "__main__":
    sys.exit(main())
