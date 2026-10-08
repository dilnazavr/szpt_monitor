"""
Сопоставление лотов с эталоном и расчёт процента завышения.
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from config import BLACKLIST, SZPT_PRODUCTS


def _normalize(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = text.lower().strip()
    text = re.sub(r"\s+", " ", text)
    return text


def is_blacklisted(name: str) -> bool:
    n = _normalize(name)
    return any(bl in n for bl in BLACKLIST)


def match_product(name: str, description: str = "") -> Optional[str]:
    if is_blacklisted(name) or is_blacklisted(description):
        return None
    combined = _normalize(name + " " + description)
    candidates: List[Tuple[int, str]] = []
    for key, keywords in SZPT_PRODUCTS.items():
        for kw in keywords:
            if kw in combined:
                candidates.append((len(kw), key))
    if not candidates:
        return None
    candidates.sort(reverse=True)
    return candidates[0][1]


def calc_unit_price(amount: float, qty: float) -> Optional[float]:
    if qty is None or qty <= 0 or amount is None or amount <= 0:
        return None
    return amount / qty


def calc_percent(unit_price: float, etalon: float) -> Optional[float]:
    if etalon is None or etalon <= 0:
        return None
    return (unit_price / etalon - 1.0) * 100.0


def process_lots(
    lots: List[Dict[str, Any]],
    etalon: Dict[str, float],
    threshold: float = 25.0,
):
    results = []
    stats = {
        "total": len(lots),
        "matched_product": 0,
        "has_etalon": 0,
        "above_threshold": 0,
        "skipped_no_qty": 0,
    }

    for lot in lots:
        name = lot.get("name_ru") or lot.get("name_kz") or ""
        desc = lot.get("description_ru") or ""
        product = match_product(name, desc)
        if not product:
            continue
        stats["matched_product"] += 1

        etalon_price = etalon.get(product)
        if etalon_price is None:
            continue
        stats["has_etalon"] += 1

        unit = calc_unit_price(lot.get("amount", 0), lot.get("qty", 0))
        if unit is None:
            stats["skipped_no_qty"] += 1
            continue

        percent = calc_percent(unit, etalon_price)
        if percent is None or percent < threshold:
            continue

        stats["above_threshold"] += 1
        year = ""
        date_str = lot.get("date") or ""
        if date_str:
            m = re.search(r"(20\d{2})", str(date_str))
            if m:
                year = m.group(1)

        results.append({
            "year": year,
            "date": date_str,
            "product": product,
            "lot_name": name,
            "unit_price": round(unit, 2),
            "etalon": round(etalon_price, 2),
            "percent": round(percent, 1),
            "amount": lot.get("amount"),
            "qty": lot.get("qty"),
            "customer": lot.get("customer_name") or "",
            "customer_bin": lot.get("customer_bin") or "",
            "lot_number": lot.get("lot_number") or "",
            "number_anno": lot.get("number_anno") or "",
            "contract_number": lot.get("contract_number") or "",
            "status": lot.get("status") or "",
            "supplier_biin": lot.get("supplier_biin") or "",
            "lot_id": lot.get("id"),
        })

    results.sort(key=lambda x: x["percent"], reverse=True)
    return results, stats