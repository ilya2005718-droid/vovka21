import json
import os
import threading
import tempfile
from typing import Dict, Any

# Config file location:
# - If env VOVKA21_CONFIG set => use it
# - Else use project-root/config.json (project root is two levels up from this file)
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DEFAULT_CONFIG_PATH = os.environ.get("VOVKA21_CONFIG", os.path.join(PROJECT_ROOT, "config.json"))
EXAMPLE_CONFIG_PATH = os.path.join(PROJECT_ROOT, "config.example.json")

_lock = threading.RLock()
_config_cache: Dict[str, Any] = None
_config_path = DEFAULT_CONFIG_PATH

def _read_json_file(path: str) -> Dict[str, Any]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except Exception:
        raise

def load_config() -> Dict[str, Any]:
    """
    Load configuration from disk. If config.json doesn't exist, fall back to example config.
    Returns a dict.
    """
    global _config_cache
    with _lock:
        if _config_cache is not None:
            return _config_cache

        cfg = _read_json_file(_config_path)
        if not cfg:
            # Try example config as fallback
            cfg = _read_json_file(EXAMPLE_CONFIG_PATH)
        if not cfg:
            cfg = {}
        _config_cache = cfg
        return cfg

def save_config(new_cfg: Dict[str, Any]) -> None:
    """
    Atomically save configuration to disk (overwrite).
    """
    global _config_cache
    with _lock:
        os.makedirs(os.path.dirname(_config_path), exist_ok=True)
        # atomic write to temp file then replace
        fd, tmp_path = tempfile.mkstemp(dir=os.path.dirname(_config_path))
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as tmp:
                json.dump(new_cfg, tmp, ensure_ascii=False, indent=2)
            os.replace(tmp_path, _config_path)
            _config_cache = new_cfg
        finally:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass

def get_db_settings() -> Dict[str, Any]:
    """
    Returns DB related settings (subset of config).
    """
    cfg = load_config()
    return cfg.get("databases", {})

def update_db_settings(updates: Dict[str, Any]) -> None:
    """
    Merge updates into existing databases config and save.
    """
    cfg = load_config()
    databases = cfg.get("databases", {})
    # deep merge for mysql dict
    for k, v in updates.items():
        if isinstance(v, dict) and isinstance(databases.get(k), dict):
            databases[k].update(v)
        else:
            databases[k] = v
    cfg["databases"] = databases
    save_config(cfg)

def get_config_path() -> str:
    return _config_path
