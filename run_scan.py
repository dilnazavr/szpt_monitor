#!/usr/bin/env python3
"""
CLI: сканирование госзакупок СЗПТ по Актобе и выгрузка завышений в Excel.
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from api_client import GoszakupAPIError, GoszakupClient, is_aktobe_region, normalize_lot
from config import DEFAULT_THRESHOLD_MONITOR, OUTPUT_DIR
from etalon_loader import load_etalon
from matcher import process_lots


def parse_args():
    p = argparse.ArgumentParser(description="SZPT Monitor — сканирование завышений (Актобе)")
    p.add_argument("--threshold", type=float, default=DEFAULT_THRESHOLD_MONITOR)
    p.add_argument("--days", type=int, default=365)
    p.add_argument("--max-pages", type=int, default=80)
    p.add_argument("--output", type=str, default=None)
    p.add_argument("--etalon-dir", type=str, default=None)
    return p.parse_args()


def main():
    args = parse_args()
    print("=" * 60)
    print("SZPT Monitor — Актобе")
    print("=" * 60)

    etalon = load_etalon(Path(args.etalon_dir) if args.etalon_dir else None)
    if not etalon:
        print("ОШИБКА: нет эталонных цен. Положите xlsx в data/etalon/")
        sys.exit(1)
    print(f"Эталон загружен: {list(etalon.keys())}")

    try:
        client = GoszakupClient()
    except GoszakupAPIError as e:
        print(f"ОШИБКА API: {e}")
        sys.exit(1)

    date_to = datetime.now().strftime("%Y-%m-%d")
    date_from = (datetime.now() - timedelta(days=args.days)).strftime("%Y-%m-%d")
    print(f"Период: {date_from} … {date_to}")

    print("Запрос лотов через OWS API…")
    try:
        raw_lots = client.search_lots(
            date_from=date_from,
            date_to=date_to,
            limit_pages=args.max_pages,
        )
    except GoszakupAPIError as e:
        print(f"ОШИБКА при запросе лотов: {e}")
        if e.status_code == 401:
            print("→ Проверьте OWS_TOKEN в .env")
        sys.exit(1)

    print(f"Получено сырых лотов: {len(raw_lots)}")

    lots = []
    for raw in raw_lots:
        if not is_aktobe_region(raw):
            continue
        lots.append(normalize_lot(raw))

    print(f"После фильтра Актобе/область: {len(lots)}")

    results, stats = process_lots(lots, etalon, threshold=args.threshold)
    print(f"Сопоставлено с товарами СЗПТ: {stats['matched_product']}")
    print(f"Имеют эталон: {stats['has_etalon']}")
    print(f"Выше порога {args.threshold}%: {stats['above_threshold']}")
    print(f"Пропущено (нет qty): {stats['skipped_no_qty']}")

    out_path = Path(args.output) if args.output else (
        OUTPUT_DIR / f"zavysheniya_{datetime.now():%Y%m%d_%H%M}.xlsx"
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(results)
    if not df.empty:
        cols = [
            "year", "date", "product", "lot_name",
            "unit_price", "etalon", "percent",
            "amount", "qty",
            "customer", "customer_bin",
            "lot_number", "number_anno", "contract_number",
            "status", "supplier_biin", "lot_id",
        ]
        df = df[[c for c in cols if c in df.columns]]
        df = df.rename(columns={
            "year": "Год",
            "date": "Дата",
            "product": "Товар",
            "lot_name": "Предмет лота",
            "unit_price": "Цена за ед.",
            "etalon": "Эталон Актобе",
            "percent": "% отклонения",
            "amount": "Сумма",
            "qty": "Количество",
            "customer": "Заказчик",
            "customer_bin": "БИН заказчика",
            "lot_number": "Номер лота",
            "number_anno": "Номер объявления",
            "contract_number": "Номер договора",
            "status": "Статус",
            "supplier_biin": "БИН поставщика",
            "lot_id": "ID лота",
        })

    df.to_excel(out_path, index=False, engine="openpyxl")
    print(f"\nРезультат сохранён: {out_path}")
    print("Готово.")

    # JSON для веб-интерфейса
    import json
    json_path = OUTPUT_DIR / "last_results.json"
    payload = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "count": len(results),
        "meta": {
            "threshold": args.threshold,
            "days": args.days,
            "file": out_path.name,
        },
        "rows": results,
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"JSON для сайта: {json_path}")
    print("Готово.")
    
if __name__ == "__main__":
    main()