"""DOCX/PPTX template catalog and placeholder inspection."""
from __future__ import annotations

import re
import shutil
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from werkzeug.utils import secure_filename

from .config import ALLOWED_TEMPLATE_EXTENSIONS, TEMPLATE_DIR, get_templates, save_templates

PLACEHOLDER_PATTERN = re.compile(r"{{\s*([A-Za-zА-Яа-яЁё0-9_.]+)\s*}}")


def extract_placeholders(file_path: Path) -> list[str]:
    """Extract placeholders from Word or PowerPoint XML parts."""
    try:
        with zipfile.ZipFile(file_path) as archive:
            fragments: list[str] = []
            for entry in archive.namelist():
                is_document_xml = entry.startswith("word/") and entry.endswith(".xml")
                is_slide_xml = entry.startswith("ppt/slides/") and entry.endswith(".xml")
                if is_document_xml or is_slide_xml:
                    xml = archive.read(entry).decode("utf-8", errors="ignore")
                    fragments.append(re.sub(r"<[^>]+>", "", xml))
        return sorted(set(PLACEHOLDER_PATTERN.findall("".join(fragments))))
    except (OSError, zipfile.BadZipFile) as error:
        raise ValueError("Файл не является корректным документом .docx или .pptx.") from error


def add_template(upload: Any, display_name: str) -> dict[str, Any]:
    raw_name = secure_filename(upload.filename or "")
    suffix = Path(raw_name).suffix.lower()
    if suffix not in ALLOWED_TEMPLATE_EXTENSIONS:
        raise ValueError("Загрузите шаблон в формате .docx или .pptx.")
    template_id = uuid.uuid4().hex
    stored_name = f"{template_id}{suffix}"
    target = TEMPLATE_DIR / stored_name
    upload.save(target)
    try:
        placeholders = extract_placeholders(target)
    except ValueError:
        target.unlink(missing_ok=True)
        raise
    item = {
        "id": template_id,
        "name": display_name.strip() or Path(raw_name).stem,
        "filename": stored_name,
        "kind": suffix[1:],
        "uploaded_at": datetime.now(timezone.utc).isoformat(),
        "placeholders": placeholders,
    }
    templates = get_templates()
    templates.append(item)
    save_templates(templates)
    return item


def remove_template(template_id: str) -> bool:
    templates = get_templates()
    selected = next((item for item in templates if item["id"] == template_id), None)
    if not selected:
        return False
    (TEMPLATE_DIR / selected["filename"]).unlink(missing_ok=True)
    save_templates([item for item in templates if item["id"] != template_id])
    return True


def get_template(template_id: str) -> dict[str, Any] | None:
    return next((item for item in get_templates() if item["id"] == template_id), None)


def template_path(template: dict[str, Any]) -> Path:
    path = (TEMPLATE_DIR / template["filename"]).resolve()
    if TEMPLATE_DIR.resolve() not in path.parents or not path.is_file():
        raise FileNotFoundError("Файл шаблона не найден.")
    return path


def copy_template_for_edit(template: dict[str, Any], output_path: Path) -> None:
    shutil.copy2(template_path(template), output_path)
