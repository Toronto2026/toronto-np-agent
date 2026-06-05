"""
Відновлення: ТТН з НП → Битрікс24 через маппінг по телефону.

Алгоритм:
  1. Витягує всі ТТН відправника за вказану дату з NP API (RecipientsPhone + IntDocNumber)
  2. Читає оригінальний Excel Бітрікс24 → телефон + ID угоди
  3. Будує маппінг: phone → TTN → deal_id
  4. Оновлює угоди в Битрікс24

Запуск:
  python recover_ttn.py --file export_bitrix.xlsx
  python recover_ttn.py --file export_bitrix.xlsx --dry-run
  python recover_ttn.py --file export_bitrix.xlsx --date 05.06.2026
"""
import argparse
import sys
import time
from datetime import date
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).parent))

from config import Config
from utils.np_api import NovaPoshtaAPI
from utils.excel import read_bitrix_export, COL_ID, COL_PHONE, write_ttn_results

OUTPUT_DIR = Path(__file__).parent / "output"
NP_API_URL = "https://api.novaposhta.ua/v2.0/json/"


def normalize_phone(phone: str) -> str:
    digits = "".join(c for c in phone if c.isdigit())
    if digits.startswith("380") and len(digits) == 12:
        return digits
    if digits.startswith("80") and len(digits) == 11:
        return "3" + digits
    if digits.startswith("0") and len(digits) == 10:
        return "38" + digits
    if len(digits) == 9:
        return "380" + digits
    return digits


def get_ttn_list(api_key: str, date_str: str) -> list[dict]:
    """Отримати список всіх ТТН відправника за дату (дд.мм.рррр)."""
    results = []
    page = 1
    while True:
        payload = {
            "apiKey": api_key,
            "modelName": "InternetDocument",
            "calledMethod": "getDocumentList",
            "methodProperties": {
                "DateTimeFrom": date_str,
                "DateTimeTo":   date_str,
                "Page":  str(page),
                "Count": "100",
                "GetFullList": "1",
            },
        }
        resp = requests.post(NP_API_URL, json=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if not data.get("success"):
            errors = ", ".join(data.get("errors", ["невідома помилка"]))
            raise RuntimeError(f"getDocumentList: {errors}")
        items = data.get("data", [])
        if not items:
            break
        results.extend(items)
        info = data.get("info", {})
        total = int(info.get("totalCount", 0)) if isinstance(info, dict) else 0
        if total and len(results) >= total:
            break
        if len(items) < 100:
            break
        page += 1
    return results


def update_bitrix_deal(webhook: str, field: str, deal_id: str, ttn: str) -> bool:
    url = webhook.rstrip("/") + "/crm.deal.update"
    resp = requests.post(url, json={"id": int(deal_id), "fields": {field: ttn}}, timeout=10)
    resp.raise_for_status()
    return bool(resp.json().get("result", False))


def save_recovery_excel(rows: list[dict]) -> Path:
    from openpyxl import Workbook
    today_str = date.today().strftime("%Y%m%d")
    path = OUTPUT_DIR / f"ttn_recover_{today_str}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Відновлені ТТН"
    ws.append(["ТТН", "ID_угод", "Телефон", "Отримувач", "Статус Б24"])
    for r in rows:
        ws.append([r.get("ttn"), r.get("ids"), r.get("phone"), r.get("name"), r.get("b24_status", "")])
    wb.save(path)
    return path


def main():
    parser = argparse.ArgumentParser(description="Відновлення ТТН → Битрікс24")
    parser.add_argument("--file", required=True, help="Excel-файл Бітрікс24 (оригінал)")
    parser.add_argument("--dry-run",   action="store_true", help="Не оновлювати Б24")
    parser.add_argument("--date", default=None, help="Дата дд.мм.рррр (за замовч. — сьогодні)")
    args = parser.parse_args()

    cfg = Config()

    target_date = args.date.strip() if args.date else date.today().strftime("%d.%m.%Y")
    print(f"Date: {target_date}")
    print(f"NP API: {cfg.NP_API_KEY[:8]}...")

    # === Крок 1: ТТН з НП ===
    print("\nStep 1: get TTN list from Nova Poshta...")
    try:
        docs = get_ttn_list(cfg.NP_API_KEY, target_date)
    except Exception as e:
        print(f"ERROR getDocumentList: {e}")
        sys.exit(1)

    print(f"  Found {len(docs)} documents for {target_date}")

    # phone → TTN (нормалізований телефон)
    phone_to_ttn: dict[str, str] = {}
    phone_to_name: dict[str, str] = {}
    for doc in docs:
        raw_phone = doc.get("RecipientsPhone", "")
        ttn = doc.get("IntDocNumber", "")
        name = doc.get("RecipientDescription", "") or doc.get("RecipientFullName", "")
        if raw_phone and ttn:
            norm = normalize_phone(raw_phone)
            phone_to_ttn[norm] = ttn
            phone_to_name[norm] = name

    print(f"  TTN by phone: {len(phone_to_ttn)}")

    # === Крок 2: Excel Бітрікс24 ===
    print(f"\nStep 2: read Bitrix24 export: {args.file}")
    excel_rows = read_bitrix_export(Path(args.file))
    print(f"  Rows: {len(excel_rows)}")

    # phone → [deal_ids]
    phone_to_ids: dict[str, list[str]] = {}
    for row in excel_rows:
        raw_phone = row.get(COL_PHONE, "").strip()
        deal_id   = str(row.get(COL_ID, "")).strip()
        if not raw_phone or not deal_id:
            continue
        norm = normalize_phone(raw_phone)
        phone_to_ids.setdefault(norm, []).append(deal_id)

    # === Крок 3: маппінг ===
    print("\nStep 3: match phone -> TTN -> deal_ids")
    matched = []
    unmatched_phones = []

    for phone, deal_ids in phone_to_ids.items():
        ttn = phone_to_ttn.get(phone)
        if ttn:
            matched.append({
                "ttn":  ttn,
                "ids":  ",".join(deal_ids),
                "phone": phone,
                "name": phone_to_name.get(phone, ""),
                "b24_status": "",
            })
        else:
            unmatched_phones.append((phone, deal_ids))

    print(f"  Matched:   {len(matched)}")
    print(f"  Unmatched: {len(unmatched_phones)} (не знайдено ТТН для цих телефонів)")

    if unmatched_phones:
        print("  Unmatched phones (deals without TTN):")
        for ph, ids in unmatched_phones:
            print(f"    {ph} -> deals {','.join(ids)}")

    # TTN у НП без відповідного запису в Excel
    excel_phones = set(phone_to_ids.keys())
    np_phones    = set(phone_to_ttn.keys())
    extra_np = np_phones - excel_phones
    if extra_np:
        print(f"\n  NOTE: {len(extra_np)} TTN in NP have no matching phone in Excel (previous batch?):")
        for ph in sorted(extra_np):
            print(f"    {ph} -> TTN {phone_to_ttn[ph]}")

    if not matched:
        print("\nERROR: no matches found. Check date and Excel file.")
        sys.exit(1)

    # === Крок 4: оновлення Б24 ===
    OUTPUT_DIR.mkdir(exist_ok=True)

    print(f"\nStep 4: update Bitrix24 ({len(matched)} TTN)")

    if args.dry_run:
        print("  (dry-run mode — not updating Bitrix24)\n")
        for r in matched:
            ids_list = [i.strip() for i in r["ids"].split(",") if i.strip()]
            for deal_id in ids_list:
                print(f"  [dry-run] Deal {deal_id} -> TTN {r['ttn']} ({r['name']})")
        excel_path = save_recovery_excel(matched)
        print(f"\nSaved: {excel_path.name}")
        return

    if not cfg.BITRIX_WEBHOOK:
        print("ERROR: BITRIX_WEBHOOK not set in .env")
        sys.exit(1)

    ok = errors = 0
    for r in matched:
        ids_list = [i.strip() for i in r["ids"].split(",") if i.strip()]
        for deal_id in ids_list:
            try:
                result = update_bitrix_deal(cfg.BITRIX_WEBHOOK, cfg.BITRIX_TTN_FIELD, deal_id, r["ttn"])
                if result:
                    print(f"  OK  Deal {deal_id} -> TTN {r['ttn']}")
                    ok += 1
                    r["b24_status"] = "OK"
                else:
                    print(f"  ERR Deal {deal_id}: false result from API")
                    errors += 1
                    r["b24_status"] = "ERROR"
            except Exception as e:
                print(f"  ERR Deal {deal_id}: {e}")
                errors += 1
                r["b24_status"] = f"ERROR: {e}"
            time.sleep(0.3)

    excel_path = save_recovery_excel(matched)
    print(f"\nDone: {ok} updated, {errors} errors")
    print(f"Saved: {excel_path.name}")


if __name__ == "__main__":
    main()
