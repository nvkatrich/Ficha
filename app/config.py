"""Application paths and small persistence helpers."""
from __future__ import annotations

import json
import os
from pathlib import Path
from threading import Lock
from typing import Any

APP_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ.get("CP_DATA_DIR", APP_DIR / "data")).resolve()
TEMPLATE_DIR = DATA_DIR / "templates"
DOCUMENT_DIR = DATA_DIR / "documents"
SETTINGS_FILE = DATA_DIR / "settings.json"
CATALOG_FILE = DATA_DIR / "templates.json"
DOCUMENT_HISTORY_FILE = DATA_DIR / "document_history.json"
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
ALLOWED_TEMPLATE_EXTENSIONS = {".docx", ".pptx"}
LOCK = Lock()

DEFAULT_SETTINGS: dict[str, Any] = {
    "bitrix_auth_type": "webhook",
    "bitrix_webhook_url": "",
    "bitrix_domain": "",
    "bitrix_access_token": "",
    "automation_enabled": "false",
    "automation_template_id": "",
    "automation_stage_id": "",
    "automation_token": "",
    "company_name": "",
    "company_inn": "",
    "company_phone": "",
    "company_email": "",
    "company_address": "",
    "document_prefix": "KP",
    "timezone": "Europe/Moscow",
}


def ensure_storage() -> None:
    """Create private local storage on first boot."""
    for path in (DATA_DIR, TEMPLATE_DIR, DOCUMENT_DIR):
        path.mkdir(parents=True, exist_ok=True)
        try:
            path.chmod(0o700)
        except OSError:
            pass
    if not SETTINGS_FILE.exists():
        save_json(SETTINGS_FILE, DEFAULT_SETTINGS)
    if not CATALOG_FILE.exists():
        save_json(CATALOG_FILE, [])
    if not DOCUMENT_HISTORY_FILE.exists():
        save_json(DOCUMENT_HISTORY_FILE, [])


def load_json(path: Path, default: Any) -> Any:
    try:
        with LOCK, path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except (OSError, json.JSONDecodeError):
        return default


def save_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    with LOCK, temporary.open("w", encoding="utf-8") as file:
        json.dump(payload, file, ensure_ascii=False, indent=2)
    temporary.replace(path)
    try:
        path.chmod(0o600)
    except OSError:
        pass


def get_settings() -> dict[str, Any]:
    stored = load_json(SETTINGS_FILE, {})
    return {**DEFAULT_SETTINGS, **stored}


def save_settings(values: dict[str, Any]) -> dict[str, Any]:
    current = get_settings()
    clean = {key: str(values.get(key, "")).strip() for key in DEFAULT_SETTINGS}
    # Empty password-style fields mean “keep the local value”; secrets never need
    # to be rendered back to the browser just to retain a connection.
    for key in ("bitrix_webhook_url", "bitrix_access_token", "automation_token"):
        if not clean[key] and current.get(key):
            clean[key] = current[key]
    current.update(clean)
    save_json(SETTINGS_FILE, current)
    return current


def get_templates() -> list[dict[str, Any]]:
    return load_json(CATALOG_FILE, [])


def save_templates(templates: list[dict[str, Any]]) -> None:
    save_json(CATALOG_FILE, templates)


def get_document_history() -> list[dict[str, Any]]:
    return load_json(DOCUMENT_HISTORY_FILE, [])


def add_document_history(item: dict[str, Any]) -> None:
    history = get_document_history()
    history.insert(0, item)
    # Keep just enough metadata for a useful local activity stream.
    save_json(DOCUMENT_HISTORY_FILE, history[:200])
