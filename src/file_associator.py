"""Windows file association helpers for PPT Touch Controller.

Registers .pptx and .ppt file associations so that double-clicking
a PowerPoint file opens it directly in PPT Touch Controller.
"""

import sys
import os
import winreg
import ctypes
import logging

logger = logging.getLogger(__name__)

PROG_ID = "PPTTouchController.pptx"


def get_exe_path() -> str:
    """Get the path to the running executable."""
    if getattr(sys, "frozen", False):
        return sys.executable
    else:
        return os.path.abspath(sys.argv[0])


def register_file_associations(exe_path: str = None) -> bool:
    """Register .pptx and .ppt file associations for current user.

    Args:
        exe_path: Path to the executable. Auto-detected if None.

    Returns:
        True on success, False on failure.
    """
    if exe_path is None:
        exe_path = get_exe_path()

    if not os.path.exists(exe_path):
        logger.error(f"Executable not found: {exe_path}")
        return False

    try:
        # Register ProgID
        _set_registry(
            winreg.HKEY_CURRENT_USER,
            f"Software\\Classes\\{PROG_ID}",
            "",
            "PPT Touch Controller Presentation",
        )

        # Default icon
        _set_registry(
            winreg.HKEY_CURRENT_USER,
            f"Software\\Classes\\{PROG_ID}\\DefaultIcon",
            "",
            f'"{exe_path}",0',
        )

        # Open command
        _set_registry(
            winreg.HKEY_CURRENT_USER,
            f"Software\\Classes\\{PROG_ID}\\shell\\open\\command",
            "",
            f'"{exe_path}" "%1"',
        )

        # Associate .pptx
        _set_registry(
            winreg.HKEY_CURRENT_USER,
            "Software\\Classes\\.pptx\\OpenWithProgids",
            PROG_ID,
            "",
        )

        # Associate .ppt
        _set_registry(
            winreg.HKEY_CURRENT_USER,
            "Software\\Classes\\.ppt\\OpenWithProgids",
            PROG_ID,
            "",
        )

        # Notify Windows shell
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)

        logger.info("File associations registered successfully")
        return True

    except Exception as e:
        logger.error(f"Failed to register file associations: {e}")
        return False


def unregister_file_associations() -> bool:
    """Remove file associations for current user.

    Returns:
        True on success, False on failure.
    """
    try:
        # Remove ProgID
        _delete_key(winreg.HKEY_CURRENT_USER, f"Software\\Classes\\{PROG_ID}")

        # Remove .pptx association
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                "Software\\Classes\\.pptx\\OpenWithProgids",
                0, winreg.KEY_SET_VALUE,
            )
            winreg.DeleteValue(key, PROG_ID)
            winreg.CloseKey(key)
        except OSError:
            pass

        # Remove .ppt association
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                "Software\\Classes\\.ppt\\OpenWithProgids",
                0, winreg.KEY_SET_VALUE,
            )
            winreg.DeleteValue(key, PROG_ID)
            winreg.CloseKey(key)
        except OSError:
            pass

        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)

        logger.info("File associations unregistered")
        return True

    except Exception as e:
        logger.error(f"Failed to unregister file associations: {e}")
        return False


def is_registered() -> bool:
    """Check if file associations are registered."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            f"Software\\Classes\\{PROG_ID}",
        )
        winreg.CloseKey(key)
        return True
    except OSError:
        return False


def _set_registry(hkey, subkey, name, value) -> None:
    """Set a registry value, creating keys as needed."""
    key = winreg.CreateKey(hkey, subkey)
    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    winreg.CloseKey(key)


def _delete_key(hkey, subkey) -> None:
    """Recursively delete a registry key."""
    try:
        winreg.DeleteKey(hkey, subkey)
    except OSError:
        pass


if __name__ == "__main__":
    # CLI for manual registration
    import argparse
    parser = argparse.ArgumentParser(description="PPT Touch Controller file associations")
    parser.add_argument("action", choices=["register", "unregister", "status"],
                        help="Action to perform")
    args = parser.parse_args()

    if args.action == "register":
        if register_file_associations():
            print("File associations registered. Double-click .pptx files to open.")
        else:
            print("Failed to register file associations.")
    elif args.action == "unregister":
        if unregister_file_associations():
            print("File associations removed.")
        else:
            print("Failed to remove file associations.")
    elif args.action == "status":
        if is_registered():
            print("File associations are registered.")
        else:
            print("File associations are NOT registered.")
