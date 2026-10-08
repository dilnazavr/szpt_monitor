"""
Загрузка эталонных цен СЗПТ по городу Актобе из локальных xlsx файлов stat.gov.kz.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Optional

import pandas as pd

from config import ETALON_DIR, SZPT_PRODUCTS


def _normalize(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = text.lower().strip()
    text = re.sub(r"\s+", " ", text)
    return text


def _match_product(name: str) -> Optional[str]:
    n = _normalize(name)
    for key, keywords in SZPT_PRODUCTS.items():
        for kw in keywords:
            if kw in n:
                return key
    return None


def _find_aktobe_column(df: pd.DataFrame) -> Optional[int]:
    for col_idx, col in enumerate(df.columns):
        val = _normalize(str(col))
        if "актобе" in val or "актөбе" in val or "ақтөбе" in val:
            return col_idx
    for col_idx in range(min(30, len(df.columns))):
        for row_idx in range(min(15, len(df))):
            val = _normalize(str(df.iloc[row_idx, col_idx]))
            if "актобе" in val or "актөбе" in val:
                return col_idx
    return None


def _find_product_column(df: pd.DataFrame) -> int:
    for col_idx in range(min(5, len(df.columns))):
        sample = " ".join(str(df.iloc[i, col_idx]) for i in range(min(20, len(df))))
        if any(kw in _normalize(sample) for prod in SZPT_PRODUCTS.values() for kw in prod):
            return col_idx
    return 0


def load_single_file(path: Path) -> Dict[str, float]:
    prices: Dict[str, float] = {}
    try:
        xl = pd.ExcelFile(path)
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(path, sheet_name=sheet_name, header=None)
            if df.empty or df.shape[1] < 2:
                continue

            aktobe_col = _find_aktobe_column(df)
            if aktobe_col is None:
                for c in range(df.shape[1]):
                    for r in range(min(10, df.shape[0])):
                        if "актобе" in _normalize(str(df.iloc[r, c])):
                            aktobe_col = c
                            break
                    if aktobe_col is not None:
                        break
            if aktobe_col is None:
                continue

            prod_col = _find_product_column(df)

            for r in range(df.shape[0]):
                prod_name = str(df.iloc[r, prod_col])
                key = _match_product(prod_name)
                if not key:
                    continue
                try:
                    val = df.iloc[r, aktobe_col]
                    if pd.isna(val):
                        continue
                    price = float(str(val).replace(",", ".").replace(" ", ""))
                    if price > 0:
                        prices[key] = price
                except (ValueError, TypeError):
                    continue
            if prices:
                break
    except Exception as e:
        print(f"[etalon] Ошибка чтения {path.name}: {e}")
    return prices


def load_etalon(directory: Path | None = None) -> Dict[str, float]:
    directory = directory or ETALON_DIR
    files = sorted(directory.glob("*.xlsx")) + sorted(directory.glob("*.xls"))
    if not files:
        print(f"[etalon] Нет файлов в {directory}. Положите xlsx с stat.gov.kz")
        return {}

    merged: Dict[str, float] = {}
    for f in files:
        part = load_single_file(f)
        if part:
            print(f"[etalon] {f.name}: найдено {len(part)} товаров")
            merged.update(part)

    if not merged:
        print("[etalon] Не удалось извлечь цены по Актобе.")
    else:
        print(f"[etalon] Итого эталонных цен: {len(merged)}")
    return merged


def get_etalon_period_hint(directory: Path | None = None) -> str:
    directory = directory or ETALON_DIR
    files = list(directory.glob("*.xlsx")) + list(directory.glob("*.xls"))
    if not files:
        return "период не определён (нет файлов эталона)"
    names = " ".join(f.name for f in files)
    years = re.findall(r"20\d{2}", names)
    if years:
        return f"{min(years)}–{max(years)}" if len(set(years)) > 1 else years[0]
    return "текущий период"