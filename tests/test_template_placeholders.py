from pathlib import Path

from app.template_store import extract_placeholders


def test_example_template_exposes_expected_placeholders():
    path = Path(__file__).resolve().parent.parent / "examples" / "Шаблон_КП_пример.docx"
    placeholders = extract_placeholders(path)
    assert "document.number" in placeholders
    assert "deal.title" in placeholders
    assert "products.table" in placeholders
