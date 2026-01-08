"""Application configuration management."""

import json
from pathlib import Path
from typing import Any


class AppConfig:
    """Application settings and recent files management."""

    DEFAULT_CONFIG = {
        "recent_excel_files": [],
        "recent_template_files": [],
        "recent_output_dirs": [],
        "last_mapping_file": "",
        "max_recent_files": 10,
        "window_geometry": None,
    }

    def __init__(self, config_dir: Path | None = None):
        """Initialize application config.

        Args:
            config_dir: Directory to store config. Defaults to user's config dir.
        """
        if config_dir is None:
            config_dir = Path.home() / ".doc-generator"
        self.config_dir = Path(config_dir)
        self.config_file = self.config_dir / "config.json"
        self._config: dict[str, Any] = dict(self.DEFAULT_CONFIG)
        self._load()

    def _load(self) -> None:
        """Load config from file."""
        if self.config_file.exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    self._config.update(loaded)
            except Exception:
                pass  # Use defaults on error

    def save(self) -> None:
        """Save config to file."""
        self.config_dir.mkdir(parents=True, exist_ok=True)
        with open(self.config_file, "w", encoding="utf-8") as f:
            json.dump(self._config, f, ensure_ascii=False, indent=2)

    def get(self, key: str, default: Any = None) -> Any:
        """Get a config value."""
        return self._config.get(key, default)

    def set(self, key: str, value: Any) -> None:
        """Set a config value."""
        self._config[key] = value

    def add_recent_file(self, category: str, path: str) -> None:
        """Add a file to recent files list.

        Args:
            category: One of 'excel', 'template', 'output_dir'.
            path: Path to add.
        """
        key = f"recent_{category}_files"
        if category == "output_dir":
            key = "recent_output_dirs"

        recent = self._config.get(key, [])
        if path in recent:
            recent.remove(path)
        recent.insert(0, path)

        max_files = self._config.get("max_recent_files", 10)
        self._config[key] = recent[:max_files]

    def get_recent_files(self, category: str) -> list[str]:
        """Get recent files list."""
        key = f"recent_{category}_files"
        if category == "output_dir":
            key = "recent_output_dirs"
        return self._config.get(key, [])


# Global app config instance
_app_config: AppConfig | None = None


def get_app_config() -> AppConfig:
    """Get the global app config instance."""
    global _app_config
    if _app_config is None:
        _app_config = AppConfig()
    return _app_config
