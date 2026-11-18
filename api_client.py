"""
Клиент OWS API goszakup.gov.kz — поиск по Актобе / Актюбинской области.

GraphQL: строки (не массивы) для nameDescriptionRu / customerNameRu — так требует живой API.
REST: запасной путь + локальный фильтр региона.
"""
from __future__ import annotations

import sys
import time
from typing import Any, Callable, Dict, List, Optional, Set

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

try:
    from config import SZPT_SEARCH_TERMS
except ImportError:
    SZPT_SEARCH_TERMS = []


try:
    from config import LOT_STATUS_IDS
except ImportError:
    LOT_STATUS_IDS = [210, 220, 230, 240, 250, 260, 270, 280, 350, 360]

try:
    from config import AKTOBE_CUSTOMER_TERMS
except ImportError:
    AKTOBE_CUSTOMER_TERMS = ["*Актюбин", "*Актобе", "*Ақтөбе", "*Актөбе"]

try:
    from config import PAGES_PER_KEYWORD
except ImportError:
    PAGES_PER_KEYWORD = 6


class GoszakupAPIError(Exception):
    def __init__(self, message: str, status_code: Optional[int] = None):
        super().__init__(message)
        self.status_code = status_code


ProgressCb = Optional[Callable[[str, int, int], None]]


def _default_progress(msg: str, current: int = 0, total: int = 0) -> None:
    if total:
        pct = min(100, int(100 * current / max(total, 1)))
        bar = "#" * (pct // 5) + "-" * (20 - pct // 5)
        line = f"\r[{bar}] {pct}% | {msg} | {current}/{total}"
    else:
        line = f"\r>>> {msg} | {current}"
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
        if resp.status_code >= 400:
            raise GoszakupAPIError(
                f"HTTP {resp.status_code}: {resp.text[:300]}",
                status_code=resp.status_code,
            )
        try:
            return resp.json()
        except ValueError:
            raise GoszakupAPIError("Ответ не является JSON")

    def _graphql(self, query: str, variables: Optional[Dict] = None) -> Dict[str, Any]:
        url = f"{self.base_url}/v3/graphql"
        payload = {"query": query, "variables": variables or {}}
        try:
            resp = self.session.post(url, json=payload, timeout=90)
        except requests.RequestException as e:
            raise GoszakupAPIError(f"GraphQL сеть: {e}") from e
        if resp.status_code == 401:
            raise GoszakupAPIError("401 Unauthorized — проверьте OWS_TOKEN", status_code=401)
        if resp.status_code >= 400:
            raise GoszakupAPIError(
                f"GraphQL HTTP {resp.status_code}: {resp.text[:400]}",
                status_code=resp.status_code,
            )
        try:
            data = resp.json()
        except ValueError:
            raise GoszakupAPIError("GraphQL ответ не JSON")
        if data.get("errors"):
            msg = str(data["errors"])[:500]
            raise GoszakupAPIError(f"GraphQL errors: {msg}")
        return data.get("data") or {}

    _LOTS_QUERY = """
    query($limit: Int, $after: Int, $filter: LotsFiltersInput) {
      Lots(limit: $limit, after: $after, filter: $filter) {
        id
        lotNumber
        refLotStatusId
        lastUpdateDate
        count
        amount
        nameRu
        nameKz
        descriptionRu
        descriptionKz
        customerId
        customerBin
        customerNameRu
        customerNameKz
        trdBuyNumberAnno
        trdBuyId
        plnPointKatoList
        indexDate
      }
    }
    """

    def _graphql_lots_page(
        self,
        filt: Dict[str, Any],
        limit: int = 200,
        after: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        variables: Dict[str, Any] = {"limit": limit, "filter": filt}
        if after is not None:
            variables["after"] = after
        data = self._graphql(self._LOTS_QUERY, variables)
        items = data.get("Lots") or data.get("lots") or []
        return items if isinstance(items, list) else []

    def _graphql_collect(
        self,
        filt: Dict[str, Any],
        max_pages: int,
        progress: ProgressCb,
        label: str,
        seen: Set[Any],
        results: List[Dict[str, Any]],
        require_aktobe: bool = True,
    ) -> int:
        cb = progress or _default_progress
        added = 0
        after: Optional[int] = None
        for page in range(max_pages):
            try:
                items = self._graphql_lots_page(filt, limit=200, after=after)
            except GoszakupAPIError as e:
                if page == 0:
                    # не роняем весь скан — пробуем другие фильтры
                    err = str(e)
                    if "invalid value" in err or "Expected type" in err:
                        print(f"\n[api] фильтр отклонён ({label}): типы полей — пробуем другой вариант")
                    else:
                        print(f"\n[api] GraphQL ({label}): {err[:200]}")
                break
            if not items:
                break
            for raw in items:
                lid = raw.get("id")
                if lid is None or lid in seen:
                    continue
                item = _gql_to_rest_shape(raw)
                if require_aktobe and not is_aktobe_region(item):
                    continue
                seen.add(lid)
                results.append(item)
                added += 1
            cb(label, len(results), max_pages * 200)
            last_id = items[-1].get("id")
            if last_id is None:
                break
            after = int(last_id)
            if len(items) < 200:
                break
            time.sleep(API_PAUSE_SEC)
        return added

    def _build_filters(
        self,
        term: str,
        customer: Optional[str] = None,
        kato: Optional[str] = None,
        date_filter: Optional[List[str]] = None,
        with_status: bool = True,
    ) -> List[Dict[str, Any]]:
        """
        Несколько вариантов фильтра: живой API ждёт String, не [String],
        для части полей. Пробуем по очереди.
        """
        variants: List[Dict[str, Any]] = []

        def base() -> Dict[str, Any]:
            f: Dict[str, Any] = {}
            if with_status and LOT_STATUS_IDS:
                # массив Int для статусов обычно ок
                f["refLotStatusId"] = list(LOT_STATUS_IDS)
            if date_filter:
                f["lastUpdateDate"] = date_filter
            return f

        # 1) nameDescriptionRu + customer (строки)
        if customer:
            f = base()
            f["nameDescriptionRu"] = term
            f["customerNameRu"] = customer
            variants.append(f)
            f2 = base()
            f2["nameRu"] = f"?{term}"  # частичное совпадение по доке
            f2["customerNameRu"] = customer
            variants.append(f2)

        # 2) только товар + КАТО
        if kato:
            f = base()
            f["nameDescriptionRu"] = term
            f["plnPointKatoList"] = kato
            variants.append(f)
            f2 = base()
            f2["nameRu"] = f"?{term}"
            f2["plnPointKatoList"] = kato
            variants.append(f2)

        # 3) только заказчик Актобе (без товара) — один раз снаружи лучше
        return variants

    def search_lots(
        self,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
        progress: ProgressCb = None,
        mode: str = "full",
        max_pages: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        seen: Set[Any] = set()
        results: List[Dict[str, Any]] = []
        cb = progress or _default_progress

        if max_pages is None or max_pages == 0:
            pages = PAGES_PER_KEYWORD if mode == "full" else max(3, PAGES_PER_KEYWORD // 2)
        else:
            pages = max(1, min(int(max_pages), 30))

        date_filter: Optional[List[str]] = None
        if date_from and date_to:
            date_filter = [date_from[:10], date_to[:10]]
        elif date_from:
            date_filter = [date_from[:10]]

        # Сильные поисковые слова СЗПТ
        terms: List[str] = []
        for kw in (SZPT_SEARCH_TERMS or []):
            kw = kw.strip()
            if len(kw) >= 3 and kw not in terms:
                terms.append(kw)
        if not terms:
            for kws in SZPT_PRODUCTS.values():
                for kw in kws[:2]:
                    kw = kw.strip()
                    if len(kw) >= 4 and kw not in terms:
                        terms.append(kw)
        terms = terms[:16] if mode == "fast" else terms[:28]

        print(
            f"\n[api] Режим={mode}. Актобе/область. "
            f"Слов СЗПТ: {len(terms)}, до {pages} стр./запрос"
        )

        # ---- A) каждое слово СЗПТ + заказчик Актобе ----
        for i, term in enumerate(terms):
            label = f"СЗПТ «{term}» + Актобе ({i + 1}/{len(terms)})"
            cb(label, i, len(terms))
            got = 0
            for cust in AKTOBE_CUSTOMER_TERMS[:2]:
                for filt in self._build_filters(term, customer=cust, date_filter=date_filter):
                    n = self._graphql_collect(filt, pages, cb, label, seen, results)
                    got += n
                    if n:
                        print(f"\n[api] {label}: +{n} (всего {len(results)})")
                        break
                if got:
                    break
                time.sleep(API_PAUSE_SEC)
            time.sleep(API_PAUSE_SEC)

        # ---- B) КАТО 15 + каждое слово СЗПТ ----
        for i, term in enumerate(terms):
            label = f"КАТО15 «{term}» ({i + 1}/{len(terms)})"
            cb(label, len(results), 0)
            for filt in self._build_filters(
                term, kato=KATO_PREFIX, date_filter=date_filter
            ):
                n = self._graphql_collect(
                    filt, max(3, pages // 2), cb, label, seen, results
                )
                if n:
                    print(f"\n[api] {label}: +{n} (всего {len(results)})")
                    break
            time.sleep(API_PAUSE_SEC)

        # ---- C) товар по названию + локальный фильтр Актобе ----
        print("\n[api] СЗПТ по названию + локальный фильтр Актобе…")
        for i, term in enumerate(terms[:20]):
            label = f"Товар «{term}» ({i + 1})"
            cb(label, len(results), 0)
            for name_field, name_val in (
                ("nameDescriptionRu", term),
                ("nameRu", f"?{term}"),
            ):
                filt = {name_field: name_val}
                if date_filter:
                    filt["lastUpdateDate"] = date_filter
                n = self._graphql_collect(
                    filt, max(2, pages // 2), cb, label, seen, results, require_aktobe=True
                )
                if n:
                    print(f"\n[api] {label}: +{n} Актобе (всего {len(results)})")
                    break
            time.sleep(API_PAUSE_SEC)

        # ---- D) широкий заказчик — только если мало ----
        if len(results) < 30:
            print("\n[api] Добор: все лоты заказчиков Актобе…")
            for cust in AKTOBE_CUSTOMER_TERMS[:2]:
                label = f"Заказчик {cust}"
                filt = {"customerNameRu": cust}
                if date_filter:
                    filt["lastUpdateDate"] = date_filter
                n = self._graphql_collect(filt, pages, cb, label, seen, results)
                if n:
                    print(f"\n[api] {label}: +{n} (всего {len(results)})")

        # ---- D) REST fallback ----
        if len(results) < 5:
            print("\n[api] REST fallback…")
            results.extend(
                self._rest_search_aktobe(
                    terms, date_from, date_to, pages, seen, cb
                )
            )

        cb(f"Итого Актобе: {len(results)}", len(results), len(results))
        print(f"\n[api] Готово. Уникальных лотов Актобе/области: {len(results)}")
        return results

    def _rest_search_aktobe(
        self,
        terms: List[str],
        date_from: Optional[str],
        date_to: Optional[str],
        pages: int,
        seen: Set[Any],
        cb: ProgressCb,
    ) -> List[Dict[str, Any]]:
        out: List[Dict[str, Any]] = []
        paths = ["/v3/lots", "/v2/lots"]

        # 1) по словам СЗПТ
        for i, term in enumerate(terms[:15]):
            label = f"REST «{term}» ({i + 1})"
            cb(label if cb else label, len(out) + len(seen), 0)
            for path in paths:
                offset = 0
                for page in range(pages):
                    params: Dict[str, Any] = {"limit": 100, "offset": offset}
                    # разные имена параметров REST
                    for key in ("name_ru", "nameRu", "search"):
                        params[key] = term
                    if date_from:
                        params["date_from"] = date_from[:10]
                        params["last_update_date_from"] = date_from[:10]
                    if date_to:
                        params["date_to"] = date_to[:10]
                        params["last_update_date_to"] = date_to[:10]
                    try:
                        data = self._get(path, params)
                    except GoszakupAPIError:
                        break
                    items = data.get("items") or data.get("data") or []
                    if not items:
                        break
                    for raw in items:
                        lid = raw.get("id")
                        if lid is None or lid in seen:
                            continue
                        if not is_aktobe_region(raw):
                            continue
                        seen.add(lid)
                        out.append(raw)
                    offset += 100
                    if len(items) < 100:
                        break
                    time.sleep(API_PAUSE_SEC)
                if out:
                    break
            print(f"\n[api] {label}: Актобе в выборке, всего найдено {len(seen)}")

        # 2) без name — только даты, фильтр Актобе (если API отдаёт customer_name)
        if len(out) < 3 and date_from:
            print("\n[api] REST: период без name, фильтр Актобе…")
            for path in paths:
                offset = 0
                for page in range(min(pages * 2, 20)):
                    params = {
                        "limit": 100,
                        "offset": offset,
                        "date_from": date_from[:10],
                    }
                    if date_to:
                        params["date_to"] = date_to[:10]
                    try:
                        data = self._get(path, params)
                    except GoszakupAPIError:
                        break
                    items = data.get("items") or data.get("data") or []
                    if not items:
                        break
                    for raw in items:
                        lid = raw.get("id")
                        if lid is None or lid in seen:
                            continue
                        if not is_aktobe_region(raw):
                            continue
                        seen.add(lid)
                        out.append(raw)
                    offset += 100
                    if cb:
                        cb("REST период+Актобе", len(seen), 0)
                    if len(items) < 100:
                        break
                    time.sleep(API_PAUSE_SEC)
                if out:
                    break
        return out


    _CONTRACT_QUERY = """
    query($limit: Int, $filter: ContractFiltersInput) {
      Contract(limit: $limit, filter: $filter) {
        id
        trdBuyId
        trdBuyNumberAnno
        contractNumber
        contractNumberSys
        supplierId
        supplierBiin
        customerBin
        Supplier {
          id
          bin
          nameRu
          nameKz
        }
      }
    }
    """

    def fetch_supplier_by_trd_buy(
        self,
        trd_buy_id: Any = None,
        number_anno: str = "",
    ) -> Dict[str, str]:
        """БИН и название поставщика из договора (реестр Contract)."""
        out = {"supplier_name": "", "supplier_biin": "", "contract_number": ""}

        if trd_buy_id:
            try:
                tid = int(trd_buy_id)
                data = self._graphql(
                    self._CONTRACT_QUERY,
                    {"limit": 5, "filter": {"trdBuyId": [tid]}},
                )
                items = data.get("Contract") or data.get("contract") or []
                if items:
                    c = items[0]
                    biin = c.get("supplierBiin") or ""
                    name = ""
                    sup = c.get("Supplier") or {}
                    if isinstance(sup, dict):
                        name = sup.get("nameRu") or sup.get("nameKz") or ""
                        biin = biin or sup.get("bin") or ""
                    out["supplier_biin"] = str(biin or "")
                    out["supplier_name"] = str(name or "")
                    out["contract_number"] = str(
                        c.get("contractNumberSys")
                        or c.get("contractNumber")
                        or ""
                    )
                    if out["supplier_biin"] or out["supplier_name"]:
                        return out
            except Exception:
                pass

        if trd_buy_id:
            for path in ("/v3/contract", "/v2/contract", "/contract"):
                try:
                    data = self._get(path, {
                        "limit": 10,
                        "trd_buy_id": trd_buy_id,
                    })
                    items = data.get("items") or data.get("data") or []
                    if not items and isinstance(data, dict) and data.get("id"):
                        items = [data]
                    for c in items:
                        biin = (
                            c.get("supplier_biin")
                            or c.get("supplierBiin")
                            or ""
                        )
                        if not biin and not c.get("supplier_id"):
                            continue
                        out["supplier_biin"] = str(biin or "")
                        out["contract_number"] = str(
                            c.get("contract_number_sys")
                            or c.get("contractNumberSys")
                            or c.get("contract_number")
                            or ""
                        )
                        if out["supplier_biin"] and not out["supplier_name"]:
                            out["supplier_name"] = self._subject_name(
                                out["supplier_biin"]
                            )
                        if out["supplier_biin"] or out["supplier_name"]:
                            return out
                except GoszakupAPIError:
                    continue

        if number_anno:
            for path in ("/v3/contract", "/v2/contract"):
                try:
                    data = self._get(path, {
                        "limit": 10,
                        "trd_buy_number_anno": number_anno,
                    })
                    items = data.get("items") or data.get("data") or []
                    for c in items:
                        biin = c.get("supplier_biin") or c.get("supplierBiin") or ""
                        if not biin:
                            continue
                        out["supplier_biin"] = str(biin)
                        out["contract_number"] = str(
                            c.get("contract_number_sys")
                            or c.get("contract_number")
                            or ""
                        )
                        out["supplier_name"] = self._subject_name(out["supplier_biin"])
                        if out["supplier_biin"]:
                            return out
                except GoszakupAPIError:
                    continue
        return out

    def _subject_name(self, biin: str) -> str:
        if not biin:
            return ""
        for path in (
            f"/v3/subject/biin/{biin}",
            f"/v2/subject/biin/{biin}",
            f"/subject/biin/{biin}",
        ):
            try:
                data = self._get(path)
                if isinstance(data, dict):
                    items = data.get("items") or data.get("data")
                    if items and isinstance(items, list):
                        data = items[0]
                    name = (
                        data.get("name_ru")
                        or data.get("nameRu")
                        or data.get("full_name_ru")
                        or data.get("name")
                        or ""
                    )
                    if name:
                        return str(name)
            except GoszakupAPIError:
                continue
        return ""

    def enrich_with_suppliers(
        self,
        rows: List[Dict[str, Any]],
        progress: ProgressCb = None,
    ) -> List[Dict[str, Any]]:
        """Дополняет строки БИН/названием поставщика из договоров."""
        if not rows:
            return rows
        cb = progress or _default_progress
        cache: Dict[str, Dict[str, str]] = {}
        total = len(rows)
        for i, row in enumerate(rows):
            if row.get("supplier_biin") and row.get("supplier_name"):
                continue
            trd_id = row.get("trd_buy_id")
            if not trd_id:
                raw = row.get("raw") or {}
                trd_id = raw.get("trd_buy_id") or raw.get("trdBuyId")
            anno = str(row.get("number_anno") or "")
            key = f"id:{trd_id}" if trd_id else f"anno:{anno}"
            if key not in cache:
                cache[key] = self.fetch_supplier_by_trd_buy(trd_id, anno)
                time.sleep(API_PAUSE_SEC)
            info = cache[key]
            if info.get("supplier_biin") and not row.get("supplier_biin"):
                row["supplier_biin"] = info["supplier_biin"]
            if info.get("supplier_name") and not row.get("supplier_name"):
                row["supplier_name"] = info["supplier_name"]
            if info.get("contract_number") and not row.get("contract_number"):
                row["contract_number"] = info["contract_number"]
            if i % 3 == 0 or i + 1 == total:
                cb("Поставщик (договор)", i + 1, total)
        return rows

    def get_lot(self, lot_id: int | str) -> Dict[str, Any]:
        for path in (f"/v3/lots/{lot_id}", f"/v2/lots/{lot_id}", f"/lots/{lot_id}"):
            try:
                return self._get(path)
            except GoszakupAPIError as e:
                if e.status_code in (404, 405):
                    continue
                raise
        raise GoszakupAPIError(f"Лот {lot_id} не найден")


def _gql_to_rest_shape(raw: Dict[str, Any]) -> Dict[str, Any]:
    mapping = {
        "id": "id",
        "lotNumber": "lot_number",
        "refLotStatusId": "ref_lot_status_id",
        "lastUpdateDate": "last_update_date",
        "count": "count",
        "amount": "amount",
        "nameRu": "name_ru",
        "nameKz": "name_kz",
        "descriptionRu": "description_ru",
        "descriptionKz": "description_kz",
        "customerId": "customer_id",
        "customerBin": "customer_bin",
        "customerNameRu": "customer_name_ru",
        "customerNameKz": "customer_name_kz",
        "trdBuyNumberAnno": "trd_buy_number_anno",
        "trdBuyId": "trd_buy_id",
        "plnPointKatoList": "pln_point_kato_list",
        "indexDate": "index_date",
    }
    out: Dict[str, Any] = dict(raw)
    for camel, snake in mapping.items():
        if camel in raw and snake not in out:
            out[snake] = raw[camel]
    return out


def is_aktobe_region(item: Dict[str, Any]) -> bool:
    kato_fields = [
        item.get("kato"),
        item.get("kato_code"),
        item.get("ref_kato_code"),
        item.get("customer_kato"),
        item.get("refKatoCode"),
        item.get("pln_point_kato_list"),
        item.get("plnPointKatoList"),
    ]
    for k in kato_fields:
        if k is None:
            continue
        if isinstance(k, (list, tuple)):
            for x in k:
                if str(x).startswith(KATO_PREFIX):
                    return True
        else:
            s = str(k).replace(" ", "")
            if s.startswith(KATO_PREFIX) or f",{KATO_PREFIX}" in s:
                return True

    text_parts = [
        item.get("customer_name_ru") or "",
        item.get("customer_name_kz") or "",
        item.get("customerNameRu") or "",
        item.get("customerNameKz") or "",
        item.get("name_ru") or "",
        item.get("description_ru") or "",
        item.get("nameRu") or "",
        item.get("descriptionRu") or "",
    ]
    blob = " ".join(str(x) for x in text_parts).lower()
    return any(kw in blob for kw in REGION_KEYWORDS)


def build_lot_url(number_anno: str, trd_buy_id: Any, lot_id: Any) -> str:
    if number_anno:
        return LOT_URL_TEMPLATE.format(number_anno=number_anno)
    if lot_id:
        return LOT_URL_BY_ID.format(lot_id=lot_id)
    return ""


def normalize_lot(raw: Dict[str, Any]) -> Dict[str, Any]:
    lot_id = raw.get("id")
    number_anno = (
        raw.get("trd_buy_number_anno")
        or raw.get("number_anno")
        or raw.get("trdBuyNumberAnno")
        or ""
    )
    trd_buy_id = raw.get("trd_buy_id") or raw.get("trdBuyId")
    status = raw.get("ref_lot_status_id") or raw.get("refLotStatusId") or ""
    return {
        "id": lot_id,
        "name_ru": raw.get("name_ru") or raw.get("nameRu") or "",
        "name_kz": raw.get("name_kz") or raw.get("nameKz") or "",
        "description_ru": raw.get("description_ru") or raw.get("descriptionRu") or "",
        "amount": raw.get("amount") or 0,
        "qty": raw.get("count") or raw.get("qty") or 0,
        "date": (
            raw.get("last_update_date")
            or raw.get("lastUpdateDate")
            or raw.get("index_date")
            or raw.get("indexDate")
            or ""
        ),
        "customer_name": (
            raw.get("customer_name_ru")
            or raw.get("customerNameRu")
            or raw.get("customer_name")
            or ""
        ),
        "customer_bin": raw.get("customer_bin") or raw.get("customerBin") or "",
        "supplier_name": (
            raw.get("supplier_name")
            or raw.get("supplierNameRu")
            or raw.get("supplier_name_ru")
            or ""
        ),
        "supplier_biin": (
            raw.get("supplier_biin")
            or raw.get("supplierBiin")
            or raw.get("supplierBin")
            or raw.get("supplier_bin")
            or ""
        ),
        "lot_number": raw.get("lot_number") or raw.get("lotNumber") or "",
        "number_anno": number_anno,
        "trd_buy_id": trd_buy_id,
        "contract_number": raw.get("contract_number") or "",
        "status": status,
        "lot_url": build_lot_url(str(number_anno or ""), trd_buy_id, lot_id),
        "raw": raw,
    }
