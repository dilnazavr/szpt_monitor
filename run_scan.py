#!/usr/bin/env python3
"""
CLI: полный скан завышений СЗПТ (без лимита страниц, с прогрессом).
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from api_client import GoszakupAPIError, GoszakupClient, is_aktobe_region, normalize_lot
from config import DEFAULT_THRESHOLD_MONITOR, OUTPUT_DIR
from etalon_loader import load_etalon_by_period
from matcher import process_lots

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

PROGRESS_FILE = OUTPUT_DIR / "scan_progress.json"


def write_progress(stage: str, current: int = 0, total: int = 0, extra: str = "") -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    data = {
        "stage": stage,
        "current": current,
        "total": total,
        "extra": extra,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    pct = int(100 * current / total) if total else 0
    bar_len = 30
    filled = int(bar_len * pct / 100)
    bar = "#" * filled + "-" * (bar_len - filled)
    msg = f"\r[{bar}] {pct:3d}% | {stage}"
    if extra:
        msg += f" | {extra}"
    msg += f" | {current}" + (f"/{total}" if total else "")
    sys.stdout.write(msg)
    sys.stdout.flush()


def parse_args():
    p = argparse.ArgumentParser(description="SZPT Monitor — полный скан")
    p.add_argument("--threshold", type=float, default=DEFAULT_THRESHOLD_MONITOR)
    p.add_argument("--days", type=int, default=365)
    p.add_argument("--output", type=str, default=None)
    p.add_argument("--etalon-dir", type=str, default=None)
    # совместимость: игнорируем max-pages, грузим всё
    p.add_argument("--max-pages", type=int, default=0, help="устарело, не используется")
    return p.parse_args()


def main():
    args = parse_args()
    print("=" * 60)
    print("SZPT Monitor — Актобе (полный обход API)")
    print("=" * 60)

    write_progress("Загрузка эталона", 0, 0)
    etalon_by_period = load_etalon_by_period(
        Path(args.etalon_dir) if args.etalon_dir else None
    )
    if not etalon_by_period:
        print("\nОШИБКА: нет эталонных цен. Положите xlsx в data/etalon/")
        sys.exit(1)
    print(f"\nПериодов эталона: {list(etalon_by_period.keys())}")

    try:
        client = GoszakupClient()
    except GoszakupAPIError as e:
        print(f"\nОШИБКА API: {e}")
        sys.exit(1)

    date_to = datetime.now().strftime("%Y-%m-%d")
    date_from = (datetime.now() - timedelta(days=args.days)).strftime("%Y-%m-%d")
    print(f"Период: {date_from} … {date_to}")
    print("Лимит страниц: НЕТ (загружаем все доступные данные)\n")

    def progress_cb(msg: str, current: int = 0, total: int = 0):
        write_progress(msg, current, total)

    write_progress("Запрос лотов API", 0, 0)
    try:
        raw_lots = client.search_lots(
            date_from=date_from,
            date_to=date_to,
            progress=progress_cb,
        )
    except GoszakupAPIError as e:
        print(f"\nОШИБКА при запросе лотов: {e}")
        if e.status_code == 401:
            print("-> Проверьте OWS_TOKEN в .env")
        sys.exit(1)

    print(f"\nПолучено сырых лотов: {len(raw_lots)}")
    write_progress("Фильтр региона Актобе", 0, len(raw_lots))

    lots = []
    for i, raw in enumerate(raw_lots):
        if is_aktobe_region(raw):
            lots.append(normalize_lot(raw))
        if i % 200 == 0:
            write_progress("Фильтр региона Актобе", i, len(raw_lots))

    print(f"После фильтра Актобе/область: {len(lots)}")
    write_progress("Сопоставление с эталоном", 0, len(lots))

    results, stats = process_lots(lots, etalon_by_period, threshold=args.threshold)
    print(f"\nСопоставлено с СЗПТ: {stats['matched_product']}")
    print(f"По товарам: {stats.get('by_product', {})}")
    print(f"Имеют эталон (по дате): {stats['has_etalon']}")
    print(f"Выше порога {args.threshold}%: {stats['above_threshold']}")
    print(f"Пропущено (нет qty): {stats['skipped_no_qty']}")

    out_path = Path(args.output) if args.output else (
        OUTPUT_DIR / f"zavysheniya_{datetime.now():%Y%m%d_%H%M}.xlsx"
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(results)
    if not df.empty:
        cols = [
            "year", "date", "etalon_period", "product", "lot_name",
            "unit_price", "etalon", "percent",
            "amount", "amount_szpt", "difference", "qty",
            "customer", "customer_bin",
            "supplier_name", "supplier_biin",
            "lot_number", "number_anno", "contract_number",
            "status", "lot_id", "lot_url",
        ]
        df = df[[c for c in cols if c in df.columns]]
        df = df.rename(columns={
            "year": "Год",
            "date": "Дата",
            "etalon_period": "Период эталона",
            "product": "Товар",
            "lot_name": "Предмет лота",
            "unit_price": "Цена за ед.",
            "etalon": "Эталон Актобе",
            "percent": "% отклонения",
            "amount": "Сумма по договору",
            "amount_szpt": "Сумма по цене СЗПТ",
            "difference": "Разница",
            "qty": "Количество",
            "customer": "Заказчик",
            "customer_bin": "БИН заказчика",
            "supplier_name": "Поставщик",
            "supplier_biin": "БИН поставщика",
            "lot_number": "Номер лота",
            "number_anno": "Номер объявления",
            "contract_number": "Номер договора",
            "status": "Статус",
            "lot_id": "ID лота",
            "lot_url": "Ссылка",
        })

    df.to_excel(out_path, index=False, engine="openpyxl")
    print(f"\nРезультат: {out_path}")

    json_path = OUTPUT_DIR / "last_results.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump({
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "count": len(results),
            "meta": {
                "threshold": args.threshold,
                "days": args.days,
                "file": out_path.name,
            },
            "rows": results,
        }, f, ensure_ascii=False, indent=2)

    write_progress("Готово", len(results), len(results), f"найдено {len(results)}")
    print(f"JSON: {json_path}")
    print("Готово.")


if __name__ == "__main__":
    main()
