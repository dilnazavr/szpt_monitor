"""
Клиент OWS API goszakup.gov.kz — без лимита страниц, с прогрессом и поиском по keywords.
"""
from __future__ import annotations

import sys
import time
from typing import Any, Callable, Dict, Generator, List, Optional, Set

import requests

from config import (
    API_PAUSE_SEC,
    KATO_PREFIX,
    LOT_URL_BY_ID,
    LOT_URL_TEMPLATE,
    OWS_BASE_URL,
    OWS_TOKEN,
    REGION_KEYWORDS,
    SZPT_PRODUCTS,
)


class GoszakupAPIError(Exception):
    def __init__(self, message: str, status_code: Optional[int] = None):
        super().__init__(message)
        self.status_code = status_code


ProgressCb = Optional[Callable[[str, int, int], None]]


def _default_progress(msg: str, current: int = 0, total: int = 0) -> None:
    if total:
        pct = min(100, int(100 * current / max(total, 1)))
        bar = "#" * (pct // 5) + "-" * (20 - pct // 5)
        line = f"\r[{bar}] {pct}% | {msg} | лотов: {current}"
    else:
        line = f"\r>>> {msg} | загружено: {current}"
    sys.stdout.write(line)
    sys.stdout.flush()


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
            resp = self.session.get(url, params=params or {}, timeout=90)
        except requests.RequestException as e:
            raise GoszakupAPIError(f"Сетевая ошибка: {e}") from e

        if resp.status_code == 401:
            raise GoszakupAPIError("401 Unauthorized — проверьте OWS_TOKEN", status_code=401)
        if resp.status_code == 403:
            raise GoszakupAPIError("403 Forbidden", status_code=403)
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
        progress: ProgressCb = None,
        progress_label: str = "",
    ) -> Generator[Dict[str, Any], None, None]:
        """Пагинация БЕЗ лимита страниц — пока API отдаёт данные."""
        params = dict(params or {})
        params.setdefault("limit", 50)
        page = 0
        next_url: Optional[str] = None
        loaded = 0
        total_hint = 0
        cb = progress or _default_progress

        while True:
            try:
                if next_url:
                    if next_url.startswith("http"):
                        resp = self.session.get(next_url, timeout=90)
                        if resp.status_code >= 400:
                            break
                        data = resp.json()
                    else:
                        data = self._get(
                            next_url if next_url.startswith("/") else f"/{next_url}"
                        )
                else:
                    data = self._get(path, params)
            except GoszakupAPIError:
                break

            items = data.get("items") or data.get("data") or []
            if not items:
                break

            total_hint = int(data.get("total") or total_hint or 0)
            for item in items:
                loaded += 1
                yield item

            if page % 2 == 0 or not data.get("next_page"):
                cb(
                    progress_label or path,
                    loaded,
                    total_hint or loaded,
                )

            next_page = data.get("next_page") or data.get("next")
            if next_page:
                next_url = next_page
            else:
                total = data.get("total")
                if total is not None and (page + 1) * params["limit"] >= int(total):
                    break
                if len(items) < params["limit"]:
                    break
                params["offset"] = (page + 1) * params["limit"]
                next_url = None

            page += 1
            time.sleep(API_PAUSE_SEC)

        sys.stdout.write("\n")
        sys.stdout.flush()

    def search_lots(
        self,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
        progress: ProgressCb = None,
    ) -> List[Dict[str, Any]]:
        """
        Полный сбор лотов:
        1) общий список за период (все страницы);
        2) доп. поиск по ключевым словам каждого товара СЗПТ.
        """
        seen: Set[Any] = set()
        results: List[Dict[str, Any]] = []
        cb = progress or _default_progress

        def add_items(items_iter, label: str):
            for item in items_iter:
                lid = item.get("id")
                if lid is None or lid in seen:
                    continue
                seen.add(lid)
                results.append(item)

        params: Dict[str, Any] = {"limit": 50}
        if date_from:
            params["date_from"] = date_from
        if date_to:
            params["date_to"] = date_to

        paths = ["/v3/lots", "/v2/lots", "/lots"]
        base_ok = False
        for path in paths:
            try:
                cb(f"Загрузка {path}…", 0, 0)
                add_items(
                    self._paginate(path, dict(params), progress=cb, progress_label=path),
                    path,
                )
                if results:
                    base_ok = True
                    break
            except GoszakupAPIError as e:
                if e.status_code in (404, 405):
                    continue
                raise

        # Поиск по keywords — чтобы не пропустить лоты с товарами СЗПТ
        search_terms: List[str] = []
        for kws in SZPT_PRODUCTS.values():
            # берём 1–2 самых характерных слова
            for kw in kws[:3]:
                kw = kw.strip()
                if len(kw) >= 3 and kw not in search_terms:
                    search_terms.append(kw)

        # уникальные короткие запросы
        search_terms = search_terms[:40]

        for i, term in enumerate(search_terms):
            cb(f"Поиск «{term}» ({i + 1}/{len(search_terms)})", len(results), 0)
            for path in paths:
                for param_name in ("name_ru", "nameRu", "search", "q"):
                    p = dict(params)
                    p[param_name] = term
                    try:
                        for item in self._paginate(
                            path, p, progress=None, progress_label=term
                        ):
                            lid = item.get("id")
                            if lid is None or lid in seen:
                                continue
                            seen.add(lid)
                            results.append(item)
                        break
                    except GoszakupAPIError:
                        continue
                else:
                    continue
                break
            time.sleep(API_PAUSE_SEC)

        cb(f"Итого уникальных лотов: {len(results)}", len(results), len(results))
        print(f"\n[api] Всего уникальных лотов: {len(results)}")
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
        item.get("kato"), item.get("kato_code"), item.get("ref_kato_code"),
        item.get("customer_kato"), item.get("refKatoCode"),
    ]
    for k in kato_fields:
        if k and str(k).startswith(KATO_PREFIX):
            return True
    if isinstance(item.get("kato_list"), list):
        for k in item["kato_list"]:
            if str(k).startswith(KATO_PREFIX):
                return True

    text_fields = [
        item.get("name_ru") or "",
        item.get("name_kz") or "",
        item.get("customer_name_ru") or "",
        item.get("customer_name_kz") or "",
        item.get("description_ru") or "",
        item.get("description_kz") or "",
        item.get("supplier_name_ru") or "",
    ]
    for key in ("delivery_place", "address", "place", "point_list"):
        val = item.get(key)
        if isinstance(val, list):
            text_fields.extend(str(x) for x in val)
        elif val:
            text_fields.append(str(val))

    combined = " ".join(text_fields).lower()
    return any(kw in combined for kw in REGION_KEYWORDS)


def build_lot_url(number_anno: str = "", trd_buy_id: Any = None, lot_id: Any = None) -> str:
    if number_anno:
        return LOT_URL_TEMPLATE.format(number_anno=number_anno)
    if trd_buy_id:
        return LOT_URL_BY_ID.format(trd_buy_id=trd_buy_id)
    if lot_id:
        return f"https://goszakup.gov.kz/ru/search/lots?filter%5Blot_number%5D={lot_id}"
    return ""


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

    number_anno = (
        raw.get("trd_buy_number_anno")
        or raw.get("number_anno")
        or raw.get("trdBuyNumberAnno")
        or ""
    )
    trd_buy_id = raw.get("trd_buy_id") or raw.get("trdBuyId")
    lot_id = raw.get("id")

    return {
        "id": lot_id,
        "lot_number": raw.get("lot_number") or raw.get("number") or str(lot_id or ""),
        "name_ru": raw.get("name_ru") or raw.get("name") or raw.get("nameRu") or "",
        "name_kz": raw.get("name_kz") or raw.get("nameKz") or "",
        "description_ru": (
            raw.get("description_ru") or raw.get("descriptionRu") or raw.get("description") or ""
        ),
        "amount": amount,
        "qty": qty,
        "customer_name": (
            raw.get("customer_name_ru")
            or raw.get("customer_name")
            or raw.get("customerNameRu")
            or raw.get("subject_name_ru")
            or ""
        ),
        "customer_bin": (
            raw.get("customer_bin")
            or raw.get("customerBin")
            or raw.get("subject_biin")
            or ""
        ),
        "status_id": raw.get("ref_lot_status_id") or raw.get("ref_buy_status_id"),
        "status": str(raw.get("ref_lot_status_id") or raw.get("status") or ""),
        "trd_buy_id": trd_buy_id,
        "number_anno": number_anno,
        "date": (
            raw.get("last_update_date")
            or raw.get("crdate")
            or raw.get("publish_date")
            or raw.get("index_date")
            or ""
        ),
        "contract_number": raw.get("contract_number") or raw.get("contractNumber") or "",
        "supplier_biin": (
            raw.get("supplier_biin")
            or raw.get("supplier_bin")
            or raw.get("supplierBiin")
            or raw.get("supplier_iin")
            or ""
        ),
        "supplier_name": (
            raw.get("supplier_name_ru")
            or raw.get("supplier_name")
            or raw.get("supplierNameRu")
            or raw.get("supplier_name_kz")
            or raw.get("supplierNameKz")
            or ""
        ),
        "lot_url": build_lot_url(str(number_anno or ""), trd_buy_id, lot_id),
        "raw": raw,
    }
