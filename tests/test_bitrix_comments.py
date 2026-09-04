from app.bitrix import BitrixClient


def test_comment_files_are_flattened_with_stable_selection_keys(monkeypatch):
    client = BitrixClient("https://portal.bitrix24.ru/rest/1/secret/")
    calls = []

    def fake_call(method, payload):
        calls.append((method, payload))
        return [{
            "ID": "900",
            "CREATED": "2026-09-03T12:00:00+03:00",
            "COMMENT": "Актуальная спецификация",
            "FILES": {"77": {"id": 77, "name": "spec.xlsx", "urlDownload": "https://portal.bitrix24.ru/disk/downloadFile/77/"}},
        }]

    monkeypatch.setattr(client, "call", fake_call)
    files = client.get_deal_comment_files(410)
    assert files[0]["key"] == "900:77"
    assert files[0]["name"] == "spec.xlsx"
    assert calls[0][1]["filter"] == {"ENTITY_ID": 410, "ENTITY_TYPE": "deal"}
