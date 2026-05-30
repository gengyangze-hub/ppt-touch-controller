"""Settings persistence for PPT Touch Controller.

Stores overlay button positions, sizes, opacity, etc.
Uses JSON file in %APPDATA%/PPTTouchController/.
"""

import json
import os
from pathlib import Path


class SettingsManager:
    """Simple JSON-based configuration persistence."""

    CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / "PPTTouchController"
    CONFIG_FILE = CONFIG_DIR / "config.json"

    DEFAULTS = {
        "overlay_position": {"x": None, "y": None},  # None = auto-center bottom
        "button_size": 80,          # px, range 60-120
        "button_opacity": 75,       # percent, range 30-100
        "button_color": "#0078D4",  # default blue
        "last_directory": "",
        "confirm_exit": False,
        "hand_mode": "right",       # "left" or "right" (which side the next button is on)
    }

    @classmethod
    def load(cls) -> dict:
        """Load settings from disk, merging with defaults."""
        settings = dict(cls.DEFAULTS)
        try:
            if cls.CONFIG_FILE.exists():
                data = json.loads(cls.CONFIG_FILE.read_text(encoding="utf-8"))
                settings.update(data)
        except (json.JSONDecodeError, OSError):
            # Corrupted config, use defaults
            pass
        return settings

    @classmethod
    def save(cls, settings: dict) -> None:
        """Save settings to disk."""
        try:
            cls.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            cls.CONFIG_FILE.write_text(
                json.dumps(settings, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except OSError:
            pass  # Non-critical failure

    @classmethod
    def update(cls, key: str, value) -> None:
        """Update a single setting key."""
        settings = cls.load()
        settings[key] = value
        cls.save(settings)
