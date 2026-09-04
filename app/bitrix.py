"""Minimal, auditable Bitrix24 REST client for CRM deal data."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any
from urllib.parse import urlparse

import requests


class BitrixError(RuntimeError):
    """A safe error message suitable for the interface."""


def _safe_endpoint(webhook_url: str) -> str:
    value = webhook_url.strip().rstrip("/")
    if not value:
        raise BitrixError("Укажите URL входящего вебхука Bitrix24 в настройках.")
    parsed = urlparse(value)
    if parsed.scheme != "https" or not parsed.netloc or "/rest/" not in parsed.path:
        raise BitrixError("Нужен полный HTTPS-адрес вебхука вида https://portal.bitrix24.ru/rest/1/код/.")
    # A URL copied from a test method can include method and .json. Store only REST base.
    rest_path = parsed.path.split("/rest/", 1)[1].strip("/").split("/")
    if len(rest_path) < 2:
        raise BitrixError("URL вебхука не содержит идентификатор пользователя и секретный код.")
    return f"https://{parsed.netloc}/rest/{rest_path[0]}/{rest_path[1]}"


@dataclass
class BitrixClient:
    webhook_url: str
    oauth_domain: str = ""
    access_token: str = ""
    timeout_seconds: int = 20

    @property
    def base_url(self) -> str:
        if self.oauth_domain:
            parsed = urlparse(self.oauth_domain if "://" in self.oauth_domain else f"https://{self.oauth_domain}")
            if parsed.scheme != "https" or not parsed.netloc:
                raise BitrixError("Укажите корректный HTTPS-домен Bitrix24.")
            if not self.access_token:
                raise BitrixError("Токен локального приложения отсутствует. Откройте приложение из Bitrix24 повторно.")
            return f"https://{parsed.netloc}/rest"
        return _safe_endpoint(self.webhook_url)

    def call(self, method: str, payload: dict[str, Any] | None = None) -> Any:
        request_payload = dict(payload or {})
        if self.oauth_domain:
            request_payload["auth"] = self.access_token
        try:
            response = requests.post(
                f"{self.base_url}/{method}.json",
                json=request_payload,
                timeout=self.timeout_seconds,
            )
            response.raise_for_status()
            content = response.json()
        except requests.RequestException as error:
            raise BitrixError("Не удалось связаться с Bitrix24. Проверьте URL вебхука и подключение.") from error
        except ValueError as error:
            raise BitrixError("Bitrix24 вернул некорректный ответ.") from error
        if content.get("error"):
            description = content.get("error_description", "Неизвестная ошибка REST API")
            raise BitrixError(f"Bitrix24: {description}")
        return content.get("result")

    def check(self) -> dict[str, Any]:
        result = self.call("profile")
        if not isinstance(result, dict):
            raise BitrixError("Bitrix24 не вернул профиль пользователя.")
        return result

    def get_deal_bundle(self, deal_id: int) -> dict[str, Any]:
        deal = self.call("crm.deal.get", {"id": deal_id})
        if not isinstance(deal, dict):
            raise BitrixError("Сделка не найдена или недоступна.")

        company = self._get_optional("crm.company.get", deal.get("COMPANY_ID"))
        contact = self._get_optional("crm.contact.get", deal.get("CONTACT_ID"))
        manager = self._get_optional("user.get", deal.get("ASSIGNED_BY_ID"), key="ID")
        products = self._get_products(deal_id)
        return {
            "deal": deal,
            "company": company,
            "contact": contact,
            "manager": manager,
            "products": products,
        }

    def get_deal_comment_files(self, deal_id: int) -> list[dict[str, Any]]:
        """Return files attached to deal timeline comments, newest first."""
        comments: list[dict[str, Any]] = []
        start = 0
        while True:
            page = self.call("crm.timeline.comment.list", {
                "filter": {"ENTITY_ID": deal_id, "ENTITY_TYPE": "deal"},
                "select": ["ID", "CREATED", "COMMENT", "FILES"],
                "order": {"ID": "DESC"},
                "start": start,
            })
            if not isinstance(page, list):
                break
            comments.extend(page)
            if len(page) < 50:
                break
            start += 50
        files: list[dict[str, Any]] = []
        for comment in comments:
            attachments = comment.get("FILES") or {}
            if isinstance(attachments, list):
                attachments = {str(item.get("id", index)): item for index, item in enumerate(attachments)}
            if not isinstance(attachments, dict):
                continue
            for file_key, file_info in attachments.items():
                if not isinstance(file_info, dict) or not file_info.get("urlDownload"):
                    continue
                files.append({
                    "key": f"{comment.get('ID')}:{file_key}",
                    "comment_id": str(comment.get("ID", "")),
                    "file_id": str(file_info.get("id", file_key)),
                    "name": file_info.get("name", f"file_{file_key}"),
                    "url": file_info.get("urlDownload", ""),
                    "created": comment.get("CREATED", ""),
                    "comment": comment.get("COMMENT", ""),
                })
        return files

    def download_comment_file(self, file_info: dict[str, Any], target: Any) -> None:
        """Download a signed Bitrix URL after checking it belongs to the portal."""
        url = str(file_info.get("url", ""))
        source = urlparse(url)
        portal = urlparse(self.oauth_domain if self.oauth_domain else self.webhook_url)
        if source.scheme != "https" or not source.netloc or source.netloc != portal.netloc:
            raise BitrixError("Bitrix24 вернул файл с неподтверждённого адреса.")
        try:
            response = requests.get(url, timeout=self.timeout_seconds, stream=True)
            response.raise_for_status()
            with open(target, "wb") as output:
                for chunk in response.iter_content(1024 * 128):
                    if chunk:
                        output.write(chunk)
        except requests.RequestException as error:
            raise BitrixError("Не удалось скачать спецификацию из комментария сделки.") from error

    def _get_optional(self, method: str, entity_id: Any, key: str = "id") -> dict[str, Any]:
        if not entity_id or str(entity_id) in {"0", "None"}:
            return {}
        result = self.call(method, {key: entity_id})
        # user.get returns a list; entity calls return a dictionary.
        if isinstance(result, list):
            return result[0] if result else {}
        return result if isinstance(result, dict) else {}

    def _get_products(self, deal_id: int) -> list[dict[str, Any]]:
        result = self.call("crm.deal.productrows.get", {"id": deal_id})
        return result if isinstance(result, list) else []
