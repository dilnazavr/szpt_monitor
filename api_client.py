"""
Клиент OWS API goszakup.gov.kz (только REST, без HTML-парсинга).
"""
from __future__ import annotations

import time
from typing import Any, Dict, Generator, List, Optional

import requests

from config import (
    API_PAUSE_SEC,
    KATO_PREFIX,
    OWS_BASE_URL,
    OWS_TOKEN,
    REGION_KEYWORDS,
)


class GoszakupAPIError(Exception):
    def __init__(self, message: str, status_code: Optional[int] = None):
        super().__init__(message)
        self.status_code = status_code


class GoszakupClient:
    def __init__(self, token: str | None = None, base_url: str | None = None):
        self.token = token or OWS_TOKEN
        self.base_url = (base_url or OWS_BASE_URL).rstrip("/")
        if not self.token:
            raise GoszakupAPIError(
                "OWS_TOKEN не задан. Укажите в .env или переменной окружения."
            )
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        })

    def _get(self, path: str, params: Optional[Dict] = None) -> Dict[str, Any]:
        url = f"{self.base_url}{path}"
        try:
            resp = self.session.get(url, params=params or {}, timeout=60)
        except requests.RequestException as e:
            raise GoszakupAPIError(f"Сетевая ошибка: {e}") from e

        if resp.status_code == 401:
            raise GoszakupAPIError(
                "401 Unauthorized — проверьте OWS_TOKEN", status_code=401
            )
        if resp.status_code == 403:
            raise GoszakupAPIError(
                "403 Forbidden — нет доступа к ресурсу", status_code=403
            )
        if resp.status_code >= 400:
            raise GoszakupAPIError(
                f"HTTP {resp.status_code}: {resp.text[:300]}",
                status_code=resp.status_code,
            )
        try:
            return resp.json()
        except ValueError:
            raise GoszakupAPIError("Ответ не является JSON")

    def _paginate(
        self,
        path: str,
        params: Optional[Dict] = None,
        max_pages: int = 200,
    ) -> Generator[Dict[str, Any], None, None]:
        params = dict(params or {})
        params.setdefault("limit", 50)
        page = 0
        next_url: Optional[str] = None

        while page < max_pages:
            if next_url:
                if next_url.startswith("http"):
                    resp = self.session.get(next_url, timeout=60)
                    if resp.status_code >= 400:
                        break
                    data = resp.json()
                else:
                    data = self._get(next_url if next_url.startswith("/") else f"/{next_url}")
            else:
                data = self._get(path, params)

            items = data.get("items") or data.get("data") or []
            if not items:
                break
            for item in items:
                yield item

            next_page = data.get("next_page") or data.get("next")
            if next_page:
                next_url = next_page
            else:
                total = data.get("total")
                if total is not None and (page + 1) * params["limit"] >= total:
                    break
                params["offset"] = (page + 1) * params["limit"]
                next_url = None

            page += 1
            time.sleep(API_PAUSE_SEC)

    def search_lots(
        self,
        keywords: List[str] | None = None,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
        customer_bin: Optional[str] = None,
        limit_pages: int = 100,
    ) -> List[Dict[str, Any]]:
        params: Dict[str, Any] = {"limit": 50}
        if customer_bin:
            params["customer_bin"] = customer_bin
        if date_from:
            params["date_from"] = date_from
        if date_to:
            params["date_to"] = date_to

        results = []
        seen_ids = set()
        paths_to_try = ["/v3/lots", "/v2/lots", "/lots"]

        for path in paths_to_try:
            try:
                for item in self._paginate(path, params, max_pages=limit_pages):
                    lid = item.get("id")
                    if lid in seen_ids:
                        continue
                    seen_ids.add(lid)
                    results.append(item)
                if results:
                    break
            except GoszakupAPIError as e:
                if e.status_code in (404, 405):
                    continue
                raise
        return results

    def get_lot(self, lot_id: int | str) -> Dict[str, Any]:
        for path in (f"/v3/lots/{lot_id}", f"/v2/lots/{lot_id}", f"/lots/{lot_id}"):
            try:
                return self._get(path)
            except GoszakupAPIError as e:
                if e.status_code in (404, 405):
                    continue
                raise
        raise GoszakupAPIError(f"Лот {lot_id} не найден")


def is_aktobe_region(item: Dict[str, Any]) -> bool:
    kato_fields = [
        item.get("kato"),
        item.get("kato_code"),
        item.get("ref_kato_code"),
        item.get("customer_kato"),
    ]
    for k in kato_fields:
        if k and str(k).startswith(KATO_PREFIX):
            return True

    text_fields = [
        item.get("name_ru") or "",
        item.get("name_kz") or "",
        item.get("customer_name_ru") or "",
        item.get("customer_name_kz") or "",
        item.get("description_ru") or "",
        item.get("description_kz") or "",
        str(item.get("customer_bin") or ""),
    ]
    for key in ("delivery_place", "address", "place", "kato_list"):
        val = item.get(key)
        if isinstance(val, list):
            text_fields.extend(str(x) for x in val)
        elif val:
            text_fields.append(str(val))

    combined = " ".join(text_fields).lower()
    return any(kw in combined for kw in REGION_KEYWORDS)


def normalize_lot(raw: Dict[str, Any]) -> Dict[str, Any]:
    amount = raw.get("amount") or raw.get("total_sum") or raw.get("price") or 0
    qty = raw.get("count") or raw.get("quantity") or raw.get("qty") or 0
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        amount = 0.0
    try:
        qty = float(qty)
    except (TypeError, ValueError):
        qty = 0.0

    return {
        "id": raw.get("id"),
        "lot_number": raw.get("lot_number") or raw.get("number") or str(raw.get("id", "")),
        "name_ru": raw.get("name_ru") or raw.get("name") or "",
        "name_kz": raw.get("name_kz") or "",
        "description_ru": raw.get("description_ru") or "",
        "amount": amount,
        "qty": qty,
        "customer_name": (
            raw.get("customer_name_ru")
            or raw.get("customer_name")
            or raw.get("subject_name_ru")
            or ""
        ),
        "customer_bin": raw.get("customer_bin") or raw.get("subject_biin") or "",
        "status_id": raw.get("ref_lot_status_id") or raw.get("ref_buy_status_id"),
        "status": str(raw.get("ref_lot_status_id") or raw.get("status") or ""),
        "trd_buy_id": raw.get("trd_buy_id") or raw.get("trdBuyId"),
        "number_anno": raw.get("trd_buy_number_anno") or raw.get("number_anno") or "",
        "date": raw.get("last_update_date") or raw.get("crdate") or raw.get("publish_date") or "",
        "contract_number": raw.get("contract_number") or "",
        "supplier_biin": raw.get("supplier_biin") or raw.get("supplier_bin") or "",
        "raw": raw,
    }