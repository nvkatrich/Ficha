from pathlib import Path

from openpyxl import Workbook

from app.specification import parse_specification


def test_xlsx_specification_normalizes_product_rows(tmp_path):
    path = tmp_path / "spec.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Позиция", "Кол.", "Ед.", "Цена", "Сумма", "Комментарий"])
    sheet.append(["Камера ГРЗ", 2, "шт.", 36500, 73000, "Наружная установка"])
    sheet.append(["Монтаж", 1, "усл.", 50000, 50000, ""])
    workbook.save(path)

    result = parse_specification(path)
    assert len(result["products"]) == 2
    assert result["products"][0]["name"] == "Камера ГРЗ"
    assert result["products"][0]["quantity"] == "2"
    assert result["products"][0]["total"] == "73000"
    assert result["warnings"] == []
