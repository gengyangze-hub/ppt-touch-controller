"""PowerPoint COM automation for PPT Touch Controller.

Manages the PowerPoint application lifecycle on a dedicated STA thread,
since PowerPoint COM requires Single-Threaded Apartment mode while Qt
runs in Multi-Threaded Apartment.
"""

import time
import logging
from PySide6.QtCore import QObject, Signal, QMutex, QMutexLocker, Slot

logger = logging.getLogger(__name__)


class PPTController(QObject):
    """Controls PowerPoint via COM, running on a dedicated thread.

    Signals are thread-safe via Qt::QueuedConnection (automatic for
    cross-thread signal-slot connections).

    Public slots (callable from any thread):
        open_and_start(file_path: str)
        next_step()
        prev_step()
        exit_slideshow()
        shutdown()
        check_status() -> bool
    """

    # Signals (emitted to main thread)
    slideshow_started = Signal()
    slideshow_ended = Signal()
    error_occurred = Signal(str)
    status_changed = Signal(bool)  # True = running, False = ended

    def __init__(self, parent=None):
        super().__init__(parent)
        self._app = None
        self._presentation = None
        self._mutex = QMutex()
        self._running = False

    def _get_slideshow_view(self):
        """Get the active slideshow view dynamically.

        Accessing the view fresh each time avoids stale COM dispatch
        references. Returns None if no slideshow is running.
        """
        try:
            return self._app.SlideShowWindows(1).View
        except Exception:
            return None

    @Slot(str)
    def open_and_start(self, file_path: str) -> None:
        """Open a PPTX file and start slideshow in fullscreen.

        Must be called from the STA thread.
        """
        import pythoncom
        import win32com.client

        with QMutexLocker(self._mutex):
            try:
                # Initialize COM for this thread (STA mode)
                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

                # Use EnsureDispatch for early-bound COM proxy (static
                # dispatch — resolves all methods/properties from the
                # PowerPoint type library).
                try:
                    self._app = win32com.client.GetActiveObject(
                        "PowerPoint.Application"
                    )
                    # Wrap as EnsureDispatch to get static proxy
                    self._app = win32com.client.gencache.EnsureDispatch(
                        self._app
                    )
                    logger.info("Attached to existing PowerPoint instance")
                except Exception:
                    self._app = win32com.client.gencache.EnsureDispatch(
                        "PowerPoint.Application"
                    )
                    logger.info("Created new PowerPoint instance")

                self._app.Visible = True  # Required for slideshow

                # Open the presentation
                self._presentation = self._app.Presentations.Open(
                    FileName=file_path,
                    WithWindow=True,
                )
                logger.info(f"Opened presentation: {file_path}")

                # Start slideshow (fullscreen)
                slideshow_settings = self._presentation.SlideShowSettings
                slideshow = slideshow_settings.Run()
                # Don't cache .View — access it fresh each time
                self._running = True

                logger.info("Slideshow started")
                self.slideshow_started.emit()

            except Exception as e:
                msg = f"Failed to open presentation: {e}"
                logger.error(msg)
                self._running = False
                self.error_occurred.emit(msg)

    @Slot()
    def next_step(self) -> None:
        """Advance to next animation step or slide."""
        with QMutexLocker(self._mutex):
            if not self._running:
                logger.warning("next_step called but slideshow not running")
                return
            try:
                view = self._get_slideshow_view()
                if view:
                    logger.info("Next: advancing")
                    view.Next()
                    logger.info("Next: done")
            except Exception as e:
                logger.error(f"Next failed: {e}", exc_info=True)
                self._check_if_ended()

    @Slot()
    def prev_step(self) -> None:
        """Go back to previous animation step or slide."""
        with QMutexLocker(self._mutex):
            if not self._running:
                logger.warning("prev_step called but slideshow not running")
                return
            try:
                view = self._get_slideshow_view()
                if view:
                    logger.info("Previous: going back")
                    view.Previous()
                    logger.info("Previous: done")
            except Exception as e:
                logger.error(f"Previous failed: {e}", exc_info=True)
                self._check_if_ended()

    @Slot(int)
    def goto_slide(self, index: int) -> None:
        """Jump to a specific slide (1-indexed, skips animations)."""
        with QMutexLocker(self._mutex):
            if not self._running:
                return
            try:
                view = self._get_slideshow_view()
                if view:
                    view.GotoSlide(index)
            except Exception as e:
                logger.debug(f"GotoSlide failed: {e}")

    def check_status(self) -> bool:
        """Check if the slideshow is still running.

        Returns:
            True if slideshow is active, False otherwise.
        """
        with QMutexLocker(self._mutex):
            return self._check_running()

    def _check_running(self) -> bool:
        """Internal: check slideshow status without locking."""
        if not self._app or not self._running:
            return False
        try:
            # SlideShowWindows(1) throws if slideshow has ended
            _ = self._app.SlideShowWindows(1)
            return True
        except Exception:
            self._running = False
            return False

    def _check_if_ended(self) -> None:
        """Check if slideshow ended and emit signal if so."""
        if not self._check_running():
            self._running = False
            self.slideshow_ended.emit()
            self.status_changed.emit(False)

    @Slot()
    def exit_slideshow(self) -> None:
        """Exit slideshow and close presentation."""
        with QMutexLocker(self._mutex):
            try:
                view = self._get_slideshow_view()
                if view:
                    view.Exit()
            except Exception:
                pass
            self._running = False
            self.slideshow_ended.emit()
            self.status_changed.emit(False)

    @Slot()
    def shutdown(self) -> None:
        """Exit slideshow, close presentation, quit PowerPoint."""
        with QMutexLocker(self._mutex):
            self._do_cleanup()
            self._running = False

    def _do_cleanup(self) -> None:
        """Internal cleanup without mutex (caller must hold lock)."""
        try:
            view = self._get_slideshow_view()
            if view:
                view.Exit()
        except Exception:
            pass
        try:
            if self._presentation:
                self._presentation.Close()
                self._presentation = None
        except Exception:
            pass
        try:
            if self._app:
                self._app.Quit()
                self._app = None
        except Exception:
            pass

    def get_slide_info(self) -> dict:
        """Get current slide number and total count."""
        with QMutexLocker(self._mutex):
            try:
                if self._running:
                    view = self._get_slideshow_view()
                    if view:
                        current = view.CurrentShowPosition
                        total = self._presentation.Slides.Count
                        return {"current": current, "total": total}
            except Exception:
                pass
        return {"current": 0, "total": 0}
