"""
Сопоставление лотов с эталоном СЗПТ.
Целые слова / основы; непищевые лоты отсекаются.
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from config import BLACKLIST, SZPT_PRODUCTS
from etalon_loader import get_etalon_for_date

NON_FOOD = [
    "автомобил", "машина", "легков", "грузов", "автобус", "трактор",
    "насос", "компрессор", "двигател", "генератор",
    "шуруп", "болт", "гайк", "гвозд", "саморез", "метиз",
    "выключател", "розетк", "кабель", "провод", "ламп", "светильн",
    "приставк", "монитор", "компьютер", "принтер", "ноутбук", "сервер",
    "анализатор", "микроскоп", "мебель", "стул", "стол ", "шкаф",
    "услуг", "работ ", "строитель", "ремонт", "проектн",
    "медикамент", "лекарств", "издел медицин", "шприц", "бинт",
    "канцтовар", "бумаг", "картридж", "тонер",
    "одежд", "обув", "спецодежд",
    "бензин", "дизел", "топлив", "масло моторн",
]


def _normalize(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = text.lower().replace("ё", "е")
    text = re.sub(r"[«»\"'()\[\]{},;:+/\\|_.\-]", " ", text)
    text = re.sub(r"\s+", " ", text.strip())
    return text


def _is_non_food(text: str) -> bool:
    n = _normalize(text)
    return any(x in n for x in NON_FOOD)


def is_blacklisted(name: str) -> bool:
    n = _normalize(name)
    if any(bl in n for bl in BLACKLIST):
        return True
    return _is_non_food(name)


def _kw_in_text(text: str, kw: str) -> bool:
    """
    Keyword совпадает с целым словом или слово НАЧИНАЕТСЯ с keyword
    (говядин → говядина), но НЕ ищет кусок в середине (лук ∉ полукруглой).
    """
    k = _normalize(kw).strip()
    if len(k) < 3:
        return False
    words = _normalize(text).split()
    for w in words:
        if w == k:
            return True
        # основа: слово начинается с keyword, keyword достаточно длинный
        if len(k) >= 4 and w.startswith(k):
            return True
        # keyword из 2+ слов: «масло подсолнеч»
        if " " in k and k in text:
            return True
    return False


def match_product(name: str, description: str = "") -> Optional[str]:
    if not name or not str(name).strip():
        return None
    if is_blacklisted(name) or _is_non_food(name):
        return None

    name_n = _normalize(name)
    desc_n = _normalize(description or "")

    scores: Dict[str, int] = {}
    for key, keywords in SZPT_PRODUCTS.items():
        best = 0
        for kw in keywords:
            k = _normalize(kw).strip()
            if len(k) < 3:
                continue
            if _kw_in_text(name_n, k):
                best = max(best, len(k) * 10)
            elif desc_n and len(k) >= 5 and _kw_in_text(desc_n, k):
                best = max(best, len(k) * 2)
        if best > 0:
            scores[key] = best

    if not scores:
        return None
    winner_key, winner_score = max(scores.items(), key=lambda x: x[1])
    if winner_score < 30:
        return None
    return winner_key


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
    etalon_by_period: Dict[str, Dict[str, float]],
    threshold: float = 25.0,
):
    results = []
    stats = {
        "total": len(lots),
        "matched_product": 0,
        "has_etalon": 0,
        "above_threshold": 0,
        "skipped_no_qty": 0,
        "by_product": {},
    }

    for lot in lots:
        name = lot.get("name_ru") or lot.get("name_kz") or ""
        desc = lot.get("description_ru") or ""
        product = match_product(name, desc)
        if not product:
            continue
        stats["matched_product"] += 1
        stats["by_product"][product] = stats["by_product"].get(product, 0) + 1

        date_str = lot.get("date") or ""
        etalon_price, etalon_period = get_etalon_for_date(
            etalon_by_period, date_str, product
        )
        if etalon_price is None:
            continue
        stats["has_etalon"] += 1

        amount = float(lot.get("amount") or 0)
        qty = float(lot.get("qty") or 0)
        unit = calc_unit_price(amount, qty)
        if unit is None:
            stats["skipped_no_qty"] += 1
            continue

        percent = calc_percent(unit, etalon_price)
        if percent is None or percent < threshold:
            continue

        stats["above_threshold"] += 1
        amount_szpt = round(etalon_price * qty, 2)
        difference = round(amount - amount_szpt, 2)

        year = ""
        if date_str:
            m = re.search(r"(20\d{2})", str(date_str))
            if m:
                year = m.group(1)

        results.append({
            "year": year,
            "date": date_str,
            "etalon_period": etalon_period,
            "product": product,
            "lot_name": name,
            "unit_price": round(unit, 2),
            "etalon": round(etalon_price, 2),
            "percent": round(percent, 1),
            "amount": amount,
            "amount_szpt": amount_szpt,
            "difference": difference,
            "qty": qty,
            "customer": lot.get("customer_name") or "",
            "customer_bin": lot.get("customer_bin") or "",
            "supplier_name": lot.get("supplier_name") or "",
            "supplier_biin": lot.get("supplier_biin") or "",
            "lot_number": lot.get("lot_number") or "",
            "number_anno": lot.get("number_anno") or "",
            "contract_number": lot.get("contract_number") or "",
            "status": lot.get("status") or "",
            "lot_id": lot.get("id"),
            "lot_url": lot.get("lot_url") or "",
        })

    results.sort(key=lambda x: x["percent"], reverse=True)
    return results, stats
