"""Central config loading.

Every path in config.ini is resolved against the project root (the folder that
holds config.ini) rather than the current working directory, so the bot behaves
the same whether you launch it from the repo root or from src/.
"""

import configparser
import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(PROJECT_ROOT, "config.ini")

_cfg = None


def exists() -> bool:
    return os.path.exists(CONFIG_PATH)


def load() -> configparser.ConfigParser:
    global _cfg
    if _cfg is None:
        if not exists():
            raise FileNotFoundError(
                f"config.ini not found at {CONFIG_PATH}\n"
                "Copy config.ini.template to config.ini and fill in your values."
            )
        cfg = configparser.ConfigParser()
        cfg.read(CONFIG_PATH, encoding="utf-8")
        _cfg = cfg
    return _cfg


def get(section: str, key: str, default: str = "") -> str:
    return load().get(section, key, fallback=default).strip()


def get_int(section: str, key: str, default: int) -> int:
    raw = get("settings", key) if section == "settings" else get(section, key)
    try:
        return int(raw)
    except (TypeError, ValueError):
        return default


def path(key: str, default: str) -> str:
    """Return an absolute path for a [settings] file path."""
    raw = get("settings", key, default) or default
    return raw if os.path.isabs(raw) else os.path.join(PROJECT_ROOT, raw)
