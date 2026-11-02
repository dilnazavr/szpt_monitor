"""
Эталон СЗПТ по Актобе из файлов Т-15-05-Н (ДД.ММ.ГГГГ) рус.xls

Нужный лист: «Средние цены на социально-значимые продовольственные товары»
Колонка: Актобе
"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from config import ETALON_DIR, ETALON_PERIOD_MODE, SZPT_PRODUCTS

try:
    from config import ETALON_MAX_DAY_GAP
except ImportError:
    ETALON_MAX_DAY_GAP = 10

MONTHS_RU = {
    "январ": 1, "феврал": 2, "март": 3, "апрел": 4,
    "май": 5, "мае": 5, "мая": 5, "июн": 6, "июл": 7,
    "август": 8, "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
}


def _normalize(text: str) -> str:
    if not isinstance(text, str):
        return ""
    return re.sub(r"\s+", " ", text.lower().replace("ё", "е").strip())


def _match_product(name: str) -> Optional[str]:
    n = _normalize(name)
    if not n or n.startswith("социально") or "тенге" in n:
        return None
    # Явные приоритеты по корню названия (лист Stat.gov)
    if "хлеб" in n:
        return "хлеб"
    if n.startswith("мука") or n.startswith("муки"):
        return "мука"
    if "говядина бескост" in n or "бескостная говяд" in n:
        return "говядина_бескостная"
    if "говядина" in n or "говядин" in n:
        return "говядина"
    if "мясо кур" in n or n.startswith("куры"):
        return "мясо_кур"
    best, best_len = None, 0
    for key, keywords in SZPT_PRODUCTS.items():
        for kw in keywords:
            k = _normalize(kw)
            if len(k) < 3:
                continue
            if k in n and len(k) > best_len:
                best, best_len = key, len(k)
    return best


def _iso_week_key(d: date) -> str:
    iso = d.isocalendar()
    return f"{iso[0]}-W{iso[1]:02d}"


def _period_from_date(d: date) -> str:
    mode = (ETALON_PERIOD_MODE or "month").lower()
    if mode == "week":
        return _iso_week_key(d)
    if mode == "quarter":
        return f"{d.year}-Q{(d.month - 1) // 3 + 1}"
    return f"{d.year}-{d.month:02d}"


def _parse_date_from_text(text: str) -> Optional[date]:
    n = _normalize(text)
    if not n:
        return None
    m = re.search(r"(20\d{2})[-_./](\d{1,2})[-_./](\d{1,2})", n)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    m = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](20\d{2})", n)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    year_m = re.search(r"(20\d{2})", n)
    if year_m:
        y = int(year_m.group(1))
        day_m = re.search(
            r"\b(\d{1,2})\s*(?:январ|феврал|март|апрел|ма[йяе]|июн|июл|август|сентябр|октябр|ноябр|декабр)",
            n,
        )
        day = int(day_m.group(1)) if day_m else 15
        for key, mo in MONTHS_RU.items():
            if key in n:
                try:
                    return date(y, mo, max(1, min(28, day)))
                except ValueError:
                    return date(y, mo, 15)
    return None


def _parse_period_from_text(text: str) -> Optional[str]:
    d = _parse_date_from_text(text)
    return _period_from_date(d) if d else None


def _period_to_date(period: str) -> Optional[date]:
    m = re.match(r"(20\d{2})-W(\d{2})$", period)
    if m:
        try:
            return date.fromisocalendar(int(m.group(1)), int(m.group(2)), 3)
        except ValueError:
            return None
    m = re.match(r"(20\d{2})-(\d{2})$", period)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), 15)
        except ValueError:
            return None
    return None


def _days_between_periods(p1: str, p2: str) -> int:
    d1, d2 = _period_to_date(p1), _period_to_date(p2)
    if not d1 or not d2:
        return 9999
    return abs((d1 - d2).days)


def parse_lot_date(date_str: str) -> Optional[date]:
    if not date_str:
        return None
    s = str(date_str).strip()
    m = re.search(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    m = re.search(r"(\d{1,2})[.](\d{1,2})[.](20\d{2})", s)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    try:
        ts = pd.Timestamp(s)
        if not pd.isna(ts):
            return ts.date()
    except Exception:
        pass
    return None


def parse_lot_period(date_str: str) -> Optional[str]:
    d = parse_lot_date(date_str)
    return _period_from_date(d) if d else None


def _find_aktobe_column(df: pd.DataFrame) -> Optional[int]:
    for col_idx in range(min(40, df.shape[1])):
        for row_idx in range(min(8, df.shape[0])):
            val = _normalize(str(df.iloc[row_idx, col_idx]))
            if val in ("актобе", "актөбе", "ақтөбе") or val.startswith("актобе"):
                return col_idx
    return None


def _find_product_column(df: pd.DataFrame) -> int:
    for col_idx in range(min(3, df.shape[1])):
        hits = sum(1 for i in range(min(40, df.shape[0])) if _match_product(str(df.iloc[i, col_idx])))
        if hits >= 3:
            return col_idx
    return 0


def _sheet_is_average_prices(df: pd.DataFrame, sheet_name: str) -> bool:
    blob = _normalize(str(sheet_name))
    for r in range(min(5, df.shape[0])):
        for c in range(min(10, df.shape[1])):
            blob += " " + _normalize(str(df.iloc[r, c]))
    if "средн" in blob and "цен" in blob:
        return True
    if "тенге" in blob and _find_aktobe_column(df) is not None:
        return True
    if _find_aktobe_column(df) is not None:
        prod_col = _find_product_column(df)
        hits = sum(1 for i in range(min(40, df.shape[0])) if _match_product(str(df.iloc[i, prod_col])))
        if hits >= 5:
            return True
    return False


def load_single_file(path: Path) -> Tuple[Optional[str], Dict[str, float]]:
    prices: Dict[str, float] = {}
    period: Optional[str] = _parse_period_from_text(path.name)
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        print(f"[etalon] Не открыть {path.name}: {e} (для .xls: pip install xlrd)")
        return period, prices

    ordered: List[str] = []
    for sn in xl.sheet_names:
        try:
            df0 = pd.read_excel(path, sheet_name=sn, header=None)
        except Exception:
            continue
        if _sheet_is_average_prices(df0, sn):
            ordered.insert(0, sn)
        else:
            ordered.append(sn)

    for sheet_name in ordered:
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, header=None)
        except Exception:
            continue
        if df.empty or df.shape[1] < 3:
            continue
        if not _sheet_is_average_prices(df, sheet_name):
            continue
        aktobe_col = _find_aktobe_column(df)
        if aktobe_col is None:
            continue
        prod_col = _find_product_column(df)
        found: Dict[str, float] = {}
        for r in range(df.shape[0]):
            key = _match_product(str(df.iloc[r, prod_col]))
            if not key:
                continue
            try:
                val = df.iloc[r, aktobe_col]
                if pd.isna(val):
                    continue
                raw = str(val).replace(",", ".").replace(" ", "").replace("\xa0", "")
                raw = re.sub(r"[^\d.]", "", raw)
                if not raw:
                    continue
                price = float(raw)
                if price > 0:
                    found[key] = price
            except (ValueError, TypeError):
                continue
        if len(found) >= 3:
            prices = found
            print(f"[etalon] {path.name} лист «{sheet_name}» Актобе col={aktobe_col}: {len(found)} товаров")
            break
    return period, prices


def load_etalon_by_period(directory: Path | None = None) -> Dict[str, Dict[str, float]]:
    directory = directory or ETALON_DIR
    files = sorted(directory.glob("*.xlsx")) + sorted(directory.glob("*.xls"))
    if not files:
        print(f"[etalon] Нет файлов в {directory}")
        return {}
    by_period: Dict[str, Dict[str, float]] = {}
    for f in files:
        period, part = load_single_file(f)
        if not part:
            print(f"[etalon] {f.name}: средние цены / Актобе не найдены")
            continue
        key = period or "unknown"
        by_period.setdefault(key, {}).update(part)
        print(f"[etalon] {f.name} -> {key}: {list(part.keys())}")
    print(f"[etalon] Режим={ETALON_PERIOD_MODE}. Периодов: {len(by_period)} → {sorted(by_period.keys())}")
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
    max_day_gap: Optional[int] = None,
) -> Tuple[Optional[float], str]:
    if not by_period or not product:
        return None, ""
    lot_period = parse_lot_period(date_str)
    gap_days = max_day_gap if max_day_gap is not None else ETALON_MAX_DAY_GAP
    mode = (ETALON_PERIOD_MODE or "month").lower()
    if lot_period and lot_period in by_period and product in by_period[lot_period]:
        return by_period[lot_period][product], lot_period
    candidates: List[Tuple[int, str]] = []
    for k, prices in by_period.items():
        if product not in prices:
            continue
        if lot_period:
            gap = _days_between_periods(lot_period, k)
            limit = gap_days if mode == "week" else max_month_gap * 31
            if gap <= limit:
                candidates.append((gap, k))
        else:
            d = _period_to_date(k)
            candidates.append((-(d.toordinal() if d else 0), k))
    if candidates:
        candidates.sort(key=lambda x: (x[0], x[1]))
        gap, k = candidates[0]
        if gap == 0:
            return by_period[k][product], k
        return by_period[k][product], f"{k}(~{gap}дн)"
    return None, f"нет эталона рядом с {lot_period or date_str}"


def get_etalon_period_hint(directory: Path | None = None) -> str:
    by_p = load_etalon_by_period(directory)
    if not by_p:
        return "период не определён"
    keys = sorted(by_p.keys())
    return f"{keys[0]} … {keys[-1]}" if len(keys) > 1 else keys[0]