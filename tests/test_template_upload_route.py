import io
from pathlib import Path

from app import config
from app import template_store
from app.main import create_app


def test_upload_route_saves_docx_and_lists_it(tmp_path, monkeypatch):
    data = tmp_path / "data"
    monkeypatch.setattr(config, "DATA_DIR", data)
    monkeypatch.setattr(config, "TEMPLATE_DIR", data / "templates")
    monkeypatch.setattr(config, "DOCUMENT_DIR", data / "documents")
    monkeypatch.setattr(config, "SETTINGS_FILE", data / "settings.json")
    monkeypatch.setattr(config, "CATALOG_FILE", data / "templates.json")
    monkeypatch.setattr(config, "DOCUMENT_HISTORY_FILE", data / "document_history.json")
    monkeypatch.setattr(template_store, "TEMPLATE_DIR", data / "templates")

    app = create_app({"TESTING": True, "SECRET_KEY": "test"})
    source = Path(__file__).resolve().parent.parent / "examples" / "Шаблон_КП_пример.docx"
    with source.open("rb") as file:
        response = app.test_client().post(
            "/templates",
            data={"display_name": "Проверочный шаблон", "template": (io.BytesIO(file.read()), "template.docx")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
    assert response.status_code == 200
    assert "Проверочный шаблон" in response.text
    assert "Шаблон добавлен" in response.text
    items = config.get_templates()
    assert len(items) == 1
    assert "products.table" in items[0]["placeholders"]
