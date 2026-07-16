"""
Скрипт 4: Звірка Бітрікс24 ↔ ТТН — напряму через API, без Excel-експорту.

Знаходить угоди, де заповнені НП-дані отримувача (телефон/місто/відділення),
але поле ТТН порожнє — незалежно від того, чи потрапила угода у будь-який
Excel-експорт. Саме так пропускаються угоди, що дозріли ПІСЛЯ того, як
експорт для Кроку 1 вже зняли.

Запуск:
  python 4_reconcile.py
  python 4_reconcile.py --closedate-from 2026-04-15
"""
import argparse
import sys
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).parent))

from config import Config
from utils.excel import write_reconcile_missing

OUTPUT_DIR = Path(__file__).parent / "output"
DEFAULT_CLOSEDATE_FROM = "2026-01-01"


def fetch_orphans(cfg: Config, closedate_from: str) -> tuple[list[dict], dict[str, str]]:
    """Питає Бітрікс24: CATEGORY_ID=ТОРОНТО, телефон НП заповнений,
    ТТН порожній, угода в роботі (не LOSE/WON), дата завершення в межах вікна."""
    webhook = cfg.BITRIX_WEBHOOK.rstrip("/")
    field_map = {
        "phone": cfg.NP_FIELD_PHONE,
        "city": cfg.NP_FIELD_CITY,
        "warehouse": cfg.NP_FIELD_WAREHOUSE,
        "name": cfg.NP_FIELD_NAME,
    }
    filter_ = {
        "CATEGORY_ID": cfg.NP_CATEGORY_ID,
        "!" + field_map["phone"]: "",
        cfg.BITRIX_TTN_FIELD: "",
        "STAGE_SEMANTIC_ID": "P",
        ">=CLOSEDATE": closedate_from,
    }
    select = ["ID", "STAGE_ID", "CLOSEDATE", *field_map.values()]

    results: list[dict] = []
    start = 0
    while True:
        resp = requests.post(
            f"{webhook}/crm.deal.list",
            json={"filter": filter_, "select": select, "start": start},
            timeout=30,
        ).json()
        if "error" in resp:
            raise RuntimeError(f"Bitrix API: {resp}")
        results.extend(resp.get("result", []))
        nxt = resp.get("next")
        if not nxt:
            break
        start = nxt
    return results, field_map


def main():
    parser = argparse.ArgumentParser(description="Звірка Бітрікс24 ↔ ТТН напряму через API")
    parser.add_argument(
        "--closedate-from", default=DEFAULT_CLOSEDATE_FROM,
        help=f"Врахувати угоди з датою завершення від цієї дати, YYYY-MM-DD (за замовчуванням {DEFAULT_CLOSEDATE_FROM})",
    )
    args = parser.parse_args()

    cfg = Config()
    if not cfg.BITRIX_WEBHOOK:
        print("❌ BITRIX_WEBHOOK не задано у .env / Secrets")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    print(f"🔍 Опитую Бітрікс24 (CATEGORY_ID={cfg.NP_CATEGORY_ID}, дата завершення ≥ {args.closedate_from})...")
    orphans, field_map = fetch_orphans(cfg, args.closedate_from)

    print(f"\n📋 Угод з НП-даними отримувача, але БЕЗ ТТН: {len(orphans)}")
    for d in orphans:
        name = (d.get(field_map["name"]) or "").strip()
        city = (d.get(field_map["city"]) or "").strip()
        print(f"  {d.get('ID'):<8} {name:<35} {city}")

    if orphans:
        path = write_reconcile_missing(orphans, field_map, OUTPUT_DIR)
        print(f"\n📁 Файл готовий (можна завантажити у Крок 1): {path.name}")
    else:
        print("\n✅ Розбіжностей не знайдено — всі угоди з НП-даними мають ТТН.")


if __name__ == "__main__":
    main()
