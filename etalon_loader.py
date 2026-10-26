"""
Эталон СЗПТ по Актобе — цены СТРОГО по периоду (месяц/квартал).
Лот за январь 2026 сравнивается с эталоном января 2026 (не с июнем!).
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from config import ETALON_DIR, ETALON_PERIOD_MODE, SZPT_PRODUCTS

MONTHS_RU = {
    "январ": 1, "феврал": 2, "март": 3, "апрел": 4,
    "май": 5, "мае": 5, "мая": 5, "июн": 6, "июл": 7,
    "август": 8, "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
}


def _normalize(text: str) -> str:
    if not isinstance(text, str):
        return ""
    return re.sub(r"\s+", " ", text.lower().strip())


def _match_product(name: str) -> Optional[str]:
    n = _normalize(name)
    best, best_len = None, 0
    for key, keywords in SZPT_PRODUCTS.items():
        for kw in keywords:
            k = _normalize(kw)
            if k and k in n and len(k) > best_len:
                best, best_len = key, len(k)
    return best


def _period_from_ym(year: int, month: int) -> str:
    month = max(1, min(12, int(month)))
    if ETALON_PERIOD_MODE == "quarter":
        return f"{year}-Q{(month - 1) // 3 + 1}"
    return f"{year}-{month:02d}"


def _parse_period_from_text(text: str) -> Optional[str]:
    """Вытащить период из имени файла или заголовка листа."""
    n = _normalize(text)
    if not n:
        return None

    # 2025-09-30, 2025_09_30, 2025.09.30
    m = re.search(r"(20\d{2})[-_./](\d{1,2})[-_./](\d{1,2})", n)
    if m:
        return _period_from_ym(int(m.group(1)), int(m.group(2)))

    # 30.09.2025
    m = re.search(r"(\d{1,2})[.](\d{1,2})[.](20\d{2})", n)
    if m:
        return _period_from_ym(int(m.group(3)), int(m.group(2)))

    # на 30 сентября 2025 / сентябрь 2025
    year_m = re.search(r"(20\d{2})", n)
    if year_m:
        y = int(year_m.group(1))
        for key, mo in MONTHS_RU.items():
            if key in n:
                return _period_from_ym(y, mo)

    return None


def _period_to_sort_key(period: str) -> Tuple[int, int]:
    """Для сравнения близости периодов: (year, month)."""
    m = re.match(r"(20\d{2})-(\d{2})$", period)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r"(20\d{2})-Q(\d)", period)
    if m:
        return int(m.group(1)), int(m.group(2)) * 3
    m = re.match(r"(20\d{2})$", period)
    if m:
        return int(m.group(1)), 6
    return 0, 0


def _months_diff(p1: str, p2: str) -> int:
    y1, m1 = _period_to_sort_key(p1)
    y2, m2 = _period_to_sort_key(p2)
    return abs((y1 * 12 + m1) - (y2 * 12 + m2))


def parse_lot_period(date_str: str) -> Optional[str]:
    """Период лота из даты API (ISO, RU и т.п.)."""
    if not date_str:
        return None
    s = str(date_str).strip()
    # 2025-12-15T10:00:00
    m = re.search(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        return _period_from_ym(int(m.group(1)), int(m.group(2)))
    m = re.search(r"(\d{1,2})[.](\d{1,2})[.](20\d{2})", s)
    if m:
        return _period_from_ym(int(m.group(3)), int(m.group(2)))
    try:
        ts = pd.Timestamp(s)
        if not pd.isna(ts):
            return _period_from_ym(int(ts.year), int(ts.month))
    except Exception:
        pass
    return None


def _find_aktobe_column(df: pd.DataFrame) -> Optional[int]:
    for col_idx, col in enumerate(df.columns):
        val = _normalize(str(col))
        if "актобе" in val or "актөбе" in val or "ақтөбе" in val:
            return col_idx
    for col_idx in range(min(40, len(df.columns))):
        for row_idx in range(min(20, len(df))):
            val = _normalize(str(df.iloc[row_idx, col_idx]))
            if "актобе" in val or "актөбе" in val or "ақтөбе" in val:
                return col_idx
    return None


def _find_product_column(df: pd.DataFrame) -> int:
    for col_idx in range(min(6, len(df.columns))):
        sample = " ".join(str(df.iloc[i, col_idx]) for i in range(min(30, len(df))))
        if any(kw in _normalize(sample) for prod in SZPT_PRODUCTS.values() for kw in prod):
            return col_idx
    return 0


def _extract_period_from_sheet(df: pd.DataFrame, filename: str) -> Optional[str]:
    """Период: сначала из имени файла, иначе из первых ячеек листа."""
    p = _parse_period_from_text(filename)
    if p:
        return p
    # первые 15x10 ячеек — часто «на 30 сентября 2025 г.»
    chunks: List[str] = [filename]
    for r in range(min(15, df.shape[0])):
        for c in range(min(10, df.shape[1])):
            chunks.append(str(df.iloc[r, c]))
    return _parse_period_from_text(" ".join(chunks))


def load_single_file(path: Path) -> Tuple[Optional[str], Dict[str, float]]:
    prices: Dict[str, float] = {}
    period: Optional[str] = None
    try:
        xl = pd.ExcelFile(path)
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(path, sheet_name=sheet_name, header=None)
            if df.empty or df.shape[1] < 2:
                continue
            if period is None:
                period = _extract_period_from_sheet(df, path.name + " " + str(sheet_name))

            aktobe_col = _find_aktobe_column(df)
            if aktobe_col is None:
                continue
            prod_col = _find_product_column(df)
            for r in range(df.shape[0]):
                key = _match_product(str(df.iloc[r, prod_col]))
                if not key:
                    continue
                try:
                    val = df.iloc[r, aktobe_col]
                    if pd.isna(val):
                        continue
                    price = float(str(val).replace(",", ".").replace(" ", "").replace("\xa0", ""))
                    if price > 0:
                        prices[key] = price
                except (ValueError, TypeError):
                    continue
            if prices:
                break
    except Exception as e:
        print(f"[etalon] Ошибка чтения {path.name}: {e}")
    return period, prices


def load_etalon_by_period(directory: Path | None = None) -> Dict[str, Dict[str, float]]:
    """
    { '2025-10': {'молоко': 362.0, 'помидоры': 890.0, ...}, '2026-01': {...} }

    Важно: кладите несколько файлов за разные месяцы!
    """
    directory = directory or ETALON_DIR
    files = sorted(directory.glob("*.xlsx")) + sorted(directory.glob("*.xls"))
    if not files:
        print(f"[etalon] Нет файлов в {directory}")
        return {}

    by_period: Dict[str, Dict[str, float]] = {}
    for f in files:
        period, part = load_single_file(f)
        if not part:
            print(f"[etalon] {f.name}: товары не распознаны")
            continue
        key = period or "unknown"
        by_period.setdefault(key, {}).update(part)
        print(f"[etalon] {f.name} -> период {key}: {len(part)} товаров — {list(part.keys())}")

    print(f"[etalon] Всего периодов: {len(by_period)} → {sorted(by_period.keys())}")
    if len(by_period) <= 1:
        print(
            "[etalon] ВНИМАНИЕ: только один период. "
            "Для сезонности (помидоры зимой/летом) добавьте xlsx за разные месяцы."
        )
    return by_period


def load_etalon(directory: Path | None = None) -> Dict[str, float]:
    by_p = load_etalon_by_period(directory)
    merged: Dict[str, float] = {}
    for p in sorted(by_p.keys()):
        merged.update(by_p[p])
    return merged


def get_etalon_for_date(
    by_period: Dict[str, Dict[str, float]],
    date_str: str,
    product: str,
    max_month_gap: int = 2,
) -> Tuple[Optional[float], str]:
    """
    Цена эталона для товара на дату лота.

    Правила:
    1) Точное совпадение месяца/квартала
    2) Ближайший период в пределах max_month_gap месяцев (по умолчанию ±2)
    3) Если ничего рядом нет — None (не берём летнюю цену для зимнего лота!)

    Возвращает (price, period_key_used или пояснение).
    """
    if not by_period or not product:
        return None, ""

    lot_period = parse_lot_period(date_str)

    # 1. Точное совпадение
    if lot_period and lot_period in by_period and product in by_period[lot_period]:
        return by_period[lot_period][product], lot_period

    # 2. Ближайший период с этим товаром, но не дальше max_month_gap
    candidates: List[Tuple[int, str]] = []
    for k, prices in by_period.items():
        if product not in prices:
            continue
        if lot_period:
            gap = _months_diff(lot_period, k)
            if gap <= max_month_gap:
                candidates.append((gap, k))
        else:
            # даты лота нет — берём самый свежий период
            y, m = _period_to_sort_key(k)
            candidates.append((-(y * 12 + m), k))

    if candidates:
        candidates.sort(key=lambda x: (x[0], x[1]))
        gap, k = candidates[0]
        return by_period[k][product], k if gap == 0 else f"{k}(~{gap}мес)"

    # 3. Не подставляем далёкий сезон
    return None, f"нет эталона рядом с {lot_period or date_str}"


def get_etalon_period_hint(directory: Path | None = None) -> str:
    by_p = load_etalon_by_period(directory)
    if not by_p:
        return "период не определён"
    keys = sorted(by_p.keys())
    return f"{keys[0]} … {keys[-1]}" if len(keys) > 1 else keys[0]
