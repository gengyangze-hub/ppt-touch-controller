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
import atexit
import subprocess
import ctypes
from ctypes import wintypes
import logging
from pathlib import Path

from PySide6.QtCore import (
    Qt, QTimer, QDataStream,
)
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import (
    QApplication, QMessageBox, QFileDialog, QMenu,
)
from PySide6.QtGui import QIcon, QAction

from settings_manager import SettingsManager
from ppt_controller import PPTController
from overlay_window import OverlayWindow

# ── Windows single-instance mutex ──────────────────────────────────
# A named kernel mutex is the most reliable single-instance lock on
# Windows: the OS auto-releases it when the owning process terminates,
# even on crash.  No stale-segment issues like QSharedMemory.
kernel32 = ctypes.windll.kernel32
CreateMutexW = kernel32.CreateMutexW
CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
CreateMutexW.restype = wintypes.HANDLE
GetLastError = kernel32.GetLastError
CloseHandle = kernel32.CloseHandle
CloseHandle.argtypes = [wintypes.HANDLE]
CloseHandle.restype = wintypes.BOOL
ERROR_ALREADY_EXISTS = 183

MUTEX_NAME = "Local\\PPTTouchController_SingleInstance_v1"
PID_FILE = Path(os.environ.get("TEMP", "")) / "PPTTouchController.pid"

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("PPTouchController")

# Single-instance identifier (QLocalServer-based, no stale-segment issues)
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
    """Main application with single-instance enforcement.

    Uses a Windows named mutex for single-instance detection — the OS
    auto-releases the mutex when the process exits, even on crash, so
    there are no stale-lock issues.
    """

    def __init__(self, argv):
        super().__init__(argv)
        self.setApplicationName("PPT Touch Controller")
        self.setApplicationVersion("1.0.0")
        self.setOrganizationName("PPTTouchController")

        # Single-instance via Windows named mutex
        self._mutex_handle = None
        self._is_primary = False
        self._ipc_server = None

        # Core components
        self._com_worker = None
        self._overlay = None
        self._status_timer = None
        self._file_path = None
        self._slideshow_active = False
        self._fallback_viewer = None

        # Cleanup on exit (direct cleanup — no deleteLater on shutdown)
        self.aboutToQuit.connect(self._cleanup_on_quit)
        atexit.register(PPTTouchApp._remove_pid_file)

    def try_become_primary(self) -> bool:
        """Attempt to become the primary instance.

        Creates a Windows named mutex.  If the mutex already exists,
        another instance is running and we become a secondary (forwarding
        the file path to the primary).

        The OS automatically destroys the mutex when the process exits,
        so stale locks from crashes are impossible.
        """
        handle = CreateMutexW(None, False, MUTEX_NAME)
        if handle == 0:
            logger.error("CreateMutexW failed — cannot determine instance status")
            self._is_primary = False
            return False

        self._mutex_handle = handle
        if GetLastError() == ERROR_ALREADY_EXISTS:
            # Mutex exists — another instance may be running.
            # Try IPC to check if it's alive.
            if self._check_primary_alive():
                # Genuine second instance — forward and exit
                self._is_primary = False
                logger.info("Another instance is already running (IPC confirmed)")
                return False
            else:
                # Zombie: mutex held but primary is unresponsive
                logger.warning("Existing mutex but IPC failed — zombie detected")
                if self._ask_force_restart():
                    self._kill_zombie()
                    # Retry: close the existing mutex handle, then create fresh
                    CloseHandle(handle)
                    handle = CreateMutexW(None, False, MUTEX_NAME)
                    self._mutex_handle = handle
                    self._is_primary = True
                    self._write_pid_file()
                    self._start_ipc_server()
                    logger.info("Primary instance started (forced after zombie cleanup)")
                    return True
                else:
                    self._is_primary = False
                    return False

        # We created the mutex — we are the primary
        self._is_primary = True
        self._write_pid_file()
        self._start_ipc_server()
        logger.info("Primary instance started")
        return True

    def _check_primary_alive(self) -> bool:
        """Check if the existing primary instance is responsive via IPC."""
        socket = QLocalSocket()
        socket.connectToServer(IPC_SERVER_NAME)
        alive = socket.waitForConnected(1500)
        if alive:
            socket.disconnectFromServer()
        return alive

    def _ask_force_restart(self) -> bool:
        """Show dialog asking if user wants to kill the unresponsive instance."""
        # Must be called before QApplication.exec(), so use a raw QMessageBox
        reply = QMessageBox.question(
            None,
            "程序已在运行",
            "检测到程序已有实例在运行，但该实例无响应。\n\n"
            "是否终止旧实例并重新启动？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        return reply == QMessageBox.Yes

    def _kill_zombie(self) -> None:
        """Kill the unresponsive primary process by its saved PID."""
        try:
            if PID_FILE.exists():
                pid_text = PID_FILE.read_text().strip()
                pid = int(pid_text)
                subprocess.run(
                    ['taskkill', '/F', '/PID', str(pid)],
                    capture_output=True,
                )
                PID_FILE.unlink(missing_ok=True)
                logger.info(f"Killed zombie process PID {pid}")
        except Exception as e:
            logger.warning(f"Failed to kill zombie: {e}")

    def _write_pid_file(self) -> None:
        """Save current process PID for zombie detection by future instances."""
        try:
            PID_FILE.write_text(str(os.getpid()))
        except OSError:
            pass

    @staticmethod
    def _remove_pid_file() -> None:
        """Clean up PID file on normal exit."""
        try:
            PID_FILE.unlink(missing_ok=True)
        except OSError:
            pass

    def _start_ipc_server(self) -> None:
        """Start local server to receive file paths from secondary instances."""
        QLocalServer.removeServer(IPC_SERVER_NAME)
        self._ipc_server = QLocalServer(self)
        if self._ipc_server.listen(IPC_SERVER_NAME):
            self._ipc_server.newConnection.connect(self._on_ipc_connection)
            logger.info("IPC server listening")
        else:
            logger.warning(f"IPC server failed to listen: {self._ipc_server.errorString()}")

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

        # ── Single-threaded architecture ──
        # Initialize COM on the main thread (STA required by PowerPoint).
        # PPTController handles CoInitializeEx internally.
        self._com_worker = PPTController()

        # Connect signals (same thread, no queuing needed)
        self._com_worker.slideshow_started.connect(
            self._on_slideshow_started
        )
        self._com_worker.slideshow_ended.connect(
            self._on_slideshow_ended
        )
        self._com_worker.error_occurred.connect(
            self._on_error
        )

        # Create overlay window
        settings = SettingsManager.load()
        self._overlay = OverlayWindow(settings)

        # Connect overlay buttons → COM worker (same thread)
        self._overlay.next_requested.connect(
            self._com_worker.next_step
        )
        self._overlay.prev_requested.connect(
            self._com_worker.prev_step
        )

        self._overlay.show()

        # Open PowerPoint and start slideshow (blocking, but fast)
        self._com_worker.open_and_start(file_path)
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
        """Periodically check if slideshow is still running."""
        if not self._slideshow_active:
            self._cleanup_slideshow()
            return

        # Check COM worker status
        if self._com_worker and not self._com_worker.check_status():
            logger.info("Slideshow no longer running, cleaning up")
            self._slideshow_active = False
            if self._overlay:
                self._overlay.hide()
            self._cleanup_slideshow()

    def _cleanup_slideshow(self) -> None:
        """Stop slideshow monitoring, clean up overlay and COM worker."""
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

    def _cleanup_on_quit(self) -> None:
        """Direct cleanup for application shutdown (aboutToQuit).

        Avoids deleteLater() since the event loop is winding down and
        deferred deletions will never be processed.
        """
        self._slideshow_active = False

        if self._status_timer:
            self._status_timer.stop()
            self._status_timer = None

        if self._overlay:
            try:
                self._overlay.save_position()
            except Exception:
                pass
            self._overlay.hide()
            self._overlay = None

        if self._com_worker:
            try:
                self._com_worker.shutdown()
            except Exception:
                pass
            self._com_worker = None

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
    # AA_EnableHighDpiScaling / AA_UseHighDpiPixmaps are deprecated in Qt 6
    # — high-DPI is enabled by default. Keeping for Qt 5 compatibility.
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except AttributeError:
        pass

    app = PPTTouchApp(sys.argv)

    if not app.try_become_primary():
        if file_path:
            app.send_to_primary(file_path)
        return 0

    return app.run(file_path)


if __name__ == "__main__":
    sys.exit(main())
