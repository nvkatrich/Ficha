"""Local Bitrix24 commercial-proposal generator."""
from __future__ import annotations

import hmac
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from flask import Flask, abort, flash, jsonify, redirect, render_template, request, send_from_directory, url_for

from .bitrix import BitrixClient, BitrixError
from .config import DOCUMENT_DIR, MAX_UPLOAD_BYTES, add_document_history, ensure_storage, get_document_history, get_settings, get_templates, save_settings
from .document_service import DocumentError, build_context, convert_to_pdf, generate_docx, lookup_dotted
from .pptx_service import PowerPointError, generate_pptx
from .specification import SpecificationError, parse_specification
from .template_store import add_template, get_template, get_templates as read_templates, remove_template, template_path


def create_app(test_config: dict[str, Any] | None = None) -> Flask:
    ensure_storage()
    app = Flask(__name__)
    app.config.update(
        SECRET_KEY=os.environ.get("CP_SECRET_KEY", os.urandom(32)),
        MAX_CONTENT_LENGTH=MAX_UPLOAD_BYTES,
        TEMPLATES_AUTO_RELOAD=True,
    )
    if test_config:
        app.config.update(test_config)

    @app.get("/")
    def index():
        settings = get_settings()
        view_settings = {key: value for key, value in settings.items() if key not in {"bitrix_webhook_url", "bitrix_access_token", "automation_token"}}
        view_settings["has_webhook"] = bool(settings.get("bitrix_webhook_url"))
        view_settings["has_oauth_token"] = bool(settings.get("bitrix_access_token"))
        view_settings["has_automation_token"] = bool(settings.get("automation_token"))
        return render_template("index.html", templates=read_templates(), settings=view_settings, documents=get_document_history()[:8])

    @app.post("/settings")
    def update_settings():
        values = {key: request.form.get(key, "") for key in (
            "bitrix_auth_type", "bitrix_webhook_url", "bitrix_domain", "bitrix_access_token",
            "company_name", "company_inn", "company_phone", "company_email", "company_address",
            "document_prefix", "timezone", "automation_template_id", "automation_stage_id", "automation_token",
        )}
        values["automation_enabled"] = "true" if request.form.get("automation_enabled") else "false"
        save_settings(values)
        flash("Настройки сохранены локально. URL вебхука и токены не передаются в браузер.", "success")
        return redirect(url_for("index") + "#settings")

    @app.post("/settings/check")
    def check_connection():
        try:
            profile = client_from_settings().check()
            return jsonify({"ok": True, "message": f"Подключено: {profile.get('NAME', '')} {profile.get('LAST_NAME', '')}".strip()})
        except BitrixError as error:
            return jsonify({"ok": False, "message": str(error)}), 400

    @app.post("/templates")
    def upload_template():
        upload = request.files.get("template")
        if not upload or not upload.filename:
            flash("Выберите файл шаблона .docx или .pptx.", "error")
            return redirect(url_for("index") + "#templates")
        try:
            item = add_template(upload, request.form.get("display_name", ""))
            found = ", ".join(f"{{{{{key}}}}}" for key in item["placeholders"]) or "не найдены (для PPTX используются также имена объектов)"
            flash(f"Шаблон добавлен. Найденные метки: {found}.", "success")
        except ValueError as error:
            flash(str(error), "error")
        return redirect(url_for("index") + "#templates")

    @app.post("/templates/<template_id>/delete")
    def delete_template(template_id: str):
        if remove_template(template_id):
            flash("Шаблон удалён вместе с локальной копией.", "success")
        else:
            flash("Шаблон уже отсутствует.", "error")
        return redirect(url_for("index") + "#templates")

    @app.get("/templates/<template_id>/source")
    def download_template(template_id: str):
        template = get_template(template_id)
        if not template:
            abort(404)
        path = template_path(template)
        return send_from_directory(path.parent, path.name, as_attachment=True, download_name=f"{template['name']}{path.suffix}")

    @app.get("/api/deals/<int:deal_id>/specifications")
    def list_specifications(deal_id: int):
        try:
            files = client_from_settings().get_deal_comment_files(deal_id)
            return jsonify({"ok": True, "files": [
                {key: file.get(key, "") for key in ("key", "name", "created", "comment")}
                for file in files
            ]})
        except BitrixError as error:
            return jsonify({"ok": False, "message": str(error)}), 400

    @app.get("/api/deals/<int:deal_id>/preview")
    def preview_deal(deal_id: int):
        template_id = request.args.get("template_id", "")
        specification_key = request.args.get("specification_key", "")
        template = get_template(template_id)
        if not template:
            return jsonify({"ok": False, "message": "Сначала выберите шаблон."}), 400
        try:
            context = build_context(bundle_with_specification(deal_id, specification_key), get_settings())
            fields = [{"key": key, "value": lookup_dotted(context, key)} for key in template.get("placeholders", [])]
            return jsonify({"ok": True, "fields": fields, "deal_title": context["deal"]["title"], "products": context["products"], "specification": context.get("specification", {})})
        except (BitrixError, SpecificationError) as error:
            return jsonify({"ok": False, "message": str(error)}), 400

    @app.post("/generate")
    def generate():
        template = get_template(request.form.get("template_id", ""))
        if not template:
            flash("Выберите существующий шаблон.", "error")
            return redirect(url_for("index") + "#generator")
        try:
            deal_id = int(request.form.get("deal_id", ""))
            if deal_id < 1:
                raise ValueError
        except ValueError:
            flash("ID сделки должен быть положительным числом.", "error")
            return redirect(url_for("index") + "#generator")
        keys = request.form.getlist("override_key")
        values = request.form.getlist("override_value")
        overrides = {key: value for key, value in zip(keys, values) if key}
        try:
            context = build_context(bundle_with_specification(deal_id, request.form.get("specification_key", "")), get_settings(), overrides)
            output_name = request.form.get("output_name", "") or f"КП_{deal_id}"
            if template.get("kind", "docx") == "pptx":
                source_path = generate_pptx(template, context, output_name)
            else:
                source_path = generate_docx(template, context, output_name)
            output_path = convert_to_pdf(source_path) if request.form.get("output_format") == "pdf" else source_path
            remember_document(output_path, context, template, "manual")
            flash("Коммерческое предложение сформировано.", "success")
            return redirect(url_for("download_document", filename=output_path.name))
        except (BitrixError, DocumentError, PowerPointError, SpecificationError) as error:
            flash(str(error), "error")
            return redirect(url_for("index") + "#generator")

    @app.post("/bitrix/app")
    def receive_bitrix_app_context():
        access_token = request.form.get("AUTH_ID", "") or request.form.get("auth[access_token]", "")
        domain = request.form.get("DOMAIN", "") or request.form.get("auth[domain]", "")
        if not access_token or not domain:
            return "Bitrix24 did not supply authorization data.", 400
        parsed = urlparse(domain if "://" in domain else f"https://{domain}")
        if not parsed.netloc:
            return "Invalid Bitrix24 domain.", 400
        settings = get_settings()
        settings.update({"bitrix_auth_type": "oauth", "bitrix_domain": parsed.netloc, "bitrix_access_token": access_token})
        save_settings(settings)
        return redirect(url_for("index"))

    @app.post("/bitrix/events/deal")
    def receive_deal_event():
        settings = get_settings()
        if settings.get("automation_enabled") != "true":
            return jsonify({"ok": True, "ignored": "automation_disabled"})
        expected_token = settings.get("automation_token", "")
        received_token = request.form.get("auth[application_token]", "")
        if not expected_token or not received_token or not hmac.compare_digest(expected_token, received_token):
            return jsonify({"ok": False, "message": "Invalid application token"}), 403
        event_name = request.form.get("event", "")
        if event_name not in {"ONCRMDEALADD", "ONCRMDEALUPDATE"}:
            return jsonify({"ok": True, "ignored": "unsupported_event"})
        deal_id = request.form.get("data[FIELDS][ID]", "")
        try:
            bundle = bundle_with_specification(int(deal_id), "__latest__")
            stage = settings.get("automation_stage_id", "")
            if stage and bundle["deal"].get("STAGE_ID") != stage:
                return jsonify({"ok": True, "ignored": "stage_mismatch"})
            template = get_template(settings.get("automation_template_id", ""))
            if not template:
                return jsonify({"ok": False, "message": "Automation template unavailable"}), 400
            context = build_context(bundle, settings)
            output_path = generate_pptx(template, context, f"КП_{deal_id}") if template.get("kind") == "pptx" else generate_docx(template, context, f"КП_{deal_id}")
            remember_document(output_path, context, template, "automation")
            return jsonify({"ok": True, "document": output_path.name})
        except (BitrixError, DocumentError, PowerPointError, SpecificationError, ValueError) as error:
            return jsonify({"ok": False, "message": str(error)}), 400

    @app.get("/documents/<filename>")
    def download_document(filename: str):
        if Path(filename).name != filename or Path(filename).suffix.lower() not in {".docx", ".pptx", ".pdf"}:
            abort(404)
        path = DOCUMENT_DIR / filename
        if not path.is_file():
            abort(404)
        return send_from_directory(DOCUMENT_DIR, filename, as_attachment=True, download_name=filename)

    @app.errorhandler(413)
    def too_large(_error: Any):
        flash("Размер шаблона превышает лимит 20 МБ.", "error")
        return redirect(url_for("index") + "#templates"), 413

    return app


def client_from_settings() -> BitrixClient:
    settings = get_settings()
    if settings.get("bitrix_auth_type") == "oauth":
        return BitrixClient(webhook_url="", oauth_domain=settings.get("bitrix_domain", ""), access_token=settings.get("bitrix_access_token", ""))
    return BitrixClient(webhook_url=settings.get("bitrix_webhook_url", ""))


def bundle_with_specification(deal_id: int, specification_key: str = "") -> dict[str, Any]:
    client = client_from_settings()
    bundle = client.get_deal_bundle(deal_id)
    if not specification_key:
        return bundle
    files = client.get_deal_comment_files(deal_id)
    supported = {".xlsx", ".xlsm", ".csv", ".docx", ".pptx", ".pdf"}
    if specification_key == "__latest__":
        selected = next((item for item in files if Path(str(item.get("name", ""))).suffix.lower() in supported), None)
    else:
        selected = next((item for item in files if item.get("key") == specification_key), None)
    if not selected:
        raise SpecificationError("Выбранная спецификация не найдена среди комментариев сделки.")
    suffix = Path(str(selected.get("name", "specification.xlsx"))).suffix.lower()
    if suffix not in {".xlsx", ".xls", ".xlsm", ".csv", ".docx", ".pptx", ".pdf"}:
        raise SpecificationError("Выберите спецификацию в формате XLSX, XLS, CSV, DOCX, PPTX или текстового PDF.")
    with tempfile.NamedTemporaryFile(prefix="cp_spec_", suffix=suffix, delete=False) as temporary:
        temporary_path = Path(temporary.name)
    try:
        client.download_comment_file(selected, temporary_path)
        parsed = parse_specification(temporary_path)
    finally:
        temporary_path.unlink(missing_ok=True)
    currency = bundle.get("deal", {}).get("CURRENCY_ID", "")
    bundle["products"] = [
        {"PRODUCT_NAME": item.get("name", ""), "QUANTITY": item.get("quantity", ""), "MEASURE_NAME": item.get("unit", ""), "PRICE": item.get("price", ""), "PRICE_ACCOUNT": item.get("total", ""), "CURRENCY": currency, "COMMENTS": item.get("comment", "")}
        for item in parsed["products"]
    ]
    bundle["specification"] = {"name": selected.get("name", ""), "comment_id": selected.get("comment_id", ""), "created": selected.get("created", ""), "warnings": parsed.get("warnings", []), "rows": len(parsed["products"])}
    return bundle


def remember_document(path: Path, context: dict[str, Any], template: dict[str, Any], source: str) -> None:
    add_document_history({"filename": path.name, "created_at": datetime.now(timezone.utc).isoformat(), "deal_id": context.get("deal", {}).get("id", ""), "deal_title": context.get("deal", {}).get("title", ""), "template_name": template.get("name", ""), "source": source})
