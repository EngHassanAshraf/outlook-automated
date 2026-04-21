import yaml
import logging
from pathlib import Path
from typing import Dict, Any, Optional


# Minimal hardcoded fallback — only used when config.yaml is completely missing.
# All real defaults live in config.yaml, which is the single source of truth.
_BARE_MINIMUM_DEFAULTS: Dict[str, Any] = {
    'output': {'base_folder': 'MV', 'year_format': 'MV-{year}'},
    'processing': {'process_unread': False, 'archive_processed': True, 'mark_as_read': True},
    'categories': {'others': {'name': 'Others - أخرى'}},
    'attachments': {'accepted_types': ['docx', 'pdf', 'xlsx', 'pptx'], 'ignored_files': []},
    'logging': {
        'level': 'INFO',
        'format': '%(asctime)s - %(levelname)s - %(message)s',
        'file': 'outlook_auto.log',
        'console': True,
    },
    'outlook': {
        'application': 'Outlook.Application',
        'namespace': 'MAPI',
        'inbox_folder_number': 6,
        'archive_root_folder': 'Archives',
        'archive_folder_name': 'Archive',
    },
    'error_handling': {'retry_attempts': 3, 'retry_delay': 5, 'continue_on_error': False},
}


class ConfigManager:
    """Manages application configuration from YAML file."""

    def __init__(self, config_path: str = "config.yaml"):
        """
        Initialize configuration manager.

        Args:
            config_path: Path to the YAML configuration file.
        """
        self.config_path = Path(config_path)
        self.config = self._load_config()

    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from YAML file, falling back to bare-minimum defaults."""
        try:
            if not self.config_path.exists():
                logging.warning(
                    f"Config file '{self.config_path}' not found — using built-in defaults. "
                    "Create config.yaml to customise behaviour."
                )
                return _BARE_MINIMUM_DEFAULTS.copy()

            with open(self.config_path, 'r', encoding='utf-8') as fh:
                loaded = yaml.safe_load(fh)
                logging.info(f"Configuration loaded from {self.config_path}")
                return loaded or _BARE_MINIMUM_DEFAULTS.copy()

        except Exception as exc:
            logging.error(f"Error loading config: {exc}")
            return _BARE_MINIMUM_DEFAULTS.copy()

    # ------------------------------------------------------------------
    # Generic accessor
    # ------------------------------------------------------------------

    def get(self, key_path: str, default: Any = None) -> Any:
        """
        Get a configuration value using dot-notation.

        Args:
            key_path: Dot-separated path, e.g. 'output.base_folder'.
            default:  Returned when the key is absent.

        Returns:
            The configuration value, or *default*.
        """
        keys = key_path.split('.')
        value = self.config
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            logging.warning(f"Config key '{key_path}' not found, using default: {default}")
            return default

    # ------------------------------------------------------------------
    # Typed convenience getters
    # ------------------------------------------------------------------

    def get_output_base_folder(self) -> str:
        return self.get('output.base_folder', 'MV')

    def get_year_format(self) -> str:
        return self.get('output.year_format', 'MV-{year}')

    def should_process_unread(self) -> bool:
        return self.get('processing.process_unread', False)

    def should_archive_processed(self) -> bool:
        return self.get('processing.archive_processed', True)

    def should_mark_as_read(self) -> bool:
        return self.get('processing.mark_as_read', True)

    def get_accepted_types(self) -> list:
        return self.get('attachments.accepted_types', ['docx', 'pdf', 'xlsx', 'pptx'])

    def get_ignored_files(self) -> list:
        return self.get('attachments.ignored_files', [])

    def get_logging_config(self) -> Dict[str, Any]:
        return self.get('logging', _BARE_MINIMUM_DEFAULTS['logging'])

    def get_outlook_config(self) -> Dict[str, Any]:
        return self.get('outlook', _BARE_MINIMUM_DEFAULTS['outlook'])

    def get_error_handling_config(self) -> Dict[str, Any]:
        return self.get('error_handling', _BARE_MINIMUM_DEFAULTS['error_handling'])

    def get_category_config(self) -> Dict[str, Any]:
        return self.get('categories', {})

    # ------------------------------------------------------------------
    # Mutation (runtime overrides)
    # ------------------------------------------------------------------

    def update_config(self, key_path: str, value: Any) -> bool:
        """
        Update a configuration value in memory and persist to disk.

        Args:
            key_path: Dot-separated path to the target key.
            value:    New value to set.

        Returns:
            True on success, False on failure.
        """
        try:
            keys = key_path.split('.')
            node = self.config
            for key in keys[:-1]:
                node = node.setdefault(key, {})
            node[keys[-1]] = value
            self._save_config()
            logging.info(f"Configuration updated: {key_path} = {value}")
            return True
        except Exception as exc:
            logging.error(f"Error updating config: {exc}")
            return False

    def _save_config(self) -> None:
        """Persist current configuration to the YAML file."""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as fh:
                yaml.dump(self.config, fh, default_flow_style=False, allow_unicode=True)
        except Exception as exc:
            logging.error(f"Error saving config: {exc}")
            raise


# ---------------------------------------------------------------------------
# Module-level singleton — kept for backward compatibility.
# Prefer passing a ConfigManager instance explicitly in new code.
# ---------------------------------------------------------------------------
config = ConfigManager()
