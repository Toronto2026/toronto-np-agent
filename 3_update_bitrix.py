"""
Скрипт 3: ttn_results_*.xlsx → оновлення поля ТТН в угодах Битрікс24.

Запуск:
  python 3_update_bitrix.py
  python 3_update_bitrix.py --dry-run
  python 3_update_bitrix.py --ttn output/ttn_results_20260331.xlsx
"""
import argparse
import sys
import time
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).parent))

from config import Config
from utils.excel import read_ttn_results

OUTPUT_DIR = Path(__file__).parent / "output"


def update_deal(webhook: str, field: str, deal_id: str, ttn: str) -> bool:
    url = webhook.rstrip("/") + "/crm.deal.update"
    resp = requests.post(url, json={"id": int(deal_id), "fields": {field: ttn}}, timeout=10)
    resp.raise_for_status()
    return resp.json().get("result", False)


def main():
    parser = argparse.ArgumentParser(description="Оновлення ТТН в угодах Битрікс24")
    parser.add_argument("--ttn", help="Файл ttn_results_*.xlsx (за замовчуванням — останній з output/)")
    parser.add_argument("--dry-run", action="store_true", help="Не оновлювати, лише показати що буде зроблено")
    args = parser.parse_args()

    cfg = Config()

    if not cfg.BITRIX_WEBHOOK:
        print("❌ BITRIX_WEBHOOK не задано у .env / Secrets")
        sys.exit(1)

    # Знайти файл результатів
    if args.ttn:
        ttn_path = Path(args.ttn)
    else:
        files = sorted(OUTPUT_DIR.glob("ttn_results_*.xlsx"), reverse=True)
        if not files:
            print("❌ Файл ttn_results_*.xlsx не знайдено у output/")
            sys.exit(1)
        ttn_path = files[0]

    print(f"📂 Читання: {ttn_path.name}")
    rows = read_ttn_results(ttn_path)

    # Залишаємо тільки рядки з реальним ТТН (не DRY-RUN, не порожній)
    to_update = [
        r for r in rows
        if r.get("ТТН") and r.get("ТТН") not in ("DRY-RUN", "") and r.get("Статус") == "OK"
    ]

    if not to_update:
        print("⚠️  Немає рядків з ТТН для оновлення")
        sys.exit(0)

    print(f"   Угод для оновлення: {len(to_update)}")
    if args.dry_run:
        print("   (режим dry-run — API не викликається)\n")

    ok = errors = 0
    for row in to_update:
        ttn = row["ТТН"]
        ids_raw = row.get("ID_угод", "")
        deal_ids = [d.strip() for d in ids_raw.split(",") if d.strip()]

        for deal_id in deal_ids:
            if args.dry_run:
                print(f"  [dry-run] Угода {deal_id} → ТТН {ttn}")
                ok += 1
                continue

            try:
                result = update_deal(cfg.BITRIX_WEBHOOK, cfg.BITRIX_TTN_FIELD, deal_id, ttn)
                if result:
                    print(f"  ✅ Угода {deal_id} → ТТН {ttn}")
                    ok += 1
                else:
                    print(f"  ❌ Угода {deal_id}: хибна відповідь API")
                    errors += 1
            except Exception as e:
                print(f"  ❌ Угода {deal_id}: {e}")
                errors += 1

            time.sleep(0.3)  # не перевищувати ліміт Б24 (~3 req/s)

    print()
    print(f"✅ Оновлено угод: {ok}")
    if errors:
        print(f"❌ Помилок: {errors}")


if __name__ == "__main__":
    main()
