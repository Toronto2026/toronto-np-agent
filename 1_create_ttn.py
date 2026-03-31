"""
Скрипт 1: Excel-експорт Бітрікс24 → ТТН Нова Пошта.

Запуск:
  python 1_create_ttn.py --file export_bitrix.xlsx
  python 1_create_ttn.py --file export_bitrix.xlsx --dry-run
"""
import argparse
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from config import Config
from utils.np_api import NovaPoshtaAPI, NovaPoshtaError
from utils.excel import (
    read_bitrix_export, write_missing, write_ttn_results, write_ttn_per_deal,
    COL_ID, COL_PHONE, COL_CITY, COL_WAREHOUSE, COL_NAME, COL_PRODUCT, COL_QTY,
)

OUTPUT_DIR = Path(__file__).parent / "output"


def normalize_phone(phone: str) -> str:
    """Нормалізувати номер до формату 380XXXXXXXXX (12 цифр)."""
    digits = "".join(c for c in phone if c.isdigit())
    if digits.startswith("380") and len(digits) == 12:
        return digits
    if digits.startswith("380"):
        return digits
    if digits.startswith("80") and len(digits) == 11:
        return "3" + digits
    if digits.startswith("0") and len(digits) == 10:
        return "38" + digits
    # 9 цифр без коду країни (напр. 671234567)
    if len(digits) == 9:
        return "380" + digits
    return digits


def normalize_city(city: str) -> str:
    """Прибрати префікси 'с.', 'м.', 'смт.' та зайву адресу після коми."""
    import re
    city = city.strip()
    # Тільки якщо є крапка — щоб не чіпати "Самар", "Мукачево"
    city = re.sub(r"^(смт\.\s*|м\.\s*|с\.\s*|сел\.\s*)", "", city, flags=re.IGNORECASE)
    # Видаляємо вулицю після " , " (напр. "Мукачево , вулиця Берегівська 66")
    city = re.sub(r"\s*,.*$", "", city)
    return city.strip()


def split_name(full_name: str) -> tuple[str, str, str]:
    """Розбити ПІБ на (прізвище, ім'я, по-батькові)."""
    parts = full_name.strip().split()
    last = parts[0] if len(parts) > 0 else ""
    first = parts[1] if len(parts) > 1 else ""
    middle = parts[2] if len(parts) > 2 else ""
    return last, first, middle


_SKIP_KEYWORDS = ["тільки електронні", "только электронн", "електронна версія", "electronic only"]

def is_electronic_only(row: dict) -> bool:
    """Повертає True якщо замовлено тільки електронні версії — ТТН не потрібен."""
    product = row.get(COL_PRODUCT, "").lower()
    return any(kw in product for kw in _SKIP_KEYWORDS)


def is_complete(row: dict) -> bool:
    """Перевірити, що всі необхідні НП-поля заповнені і потрібна фізична відправка."""
    if is_electronic_only(row):
        return False
    return all(row.get(f, "").strip() for f in [COL_PHONE, COL_CITY, COL_WAREHOUSE, COL_NAME])


def group_by_phone(rows: list[dict]) -> dict[str, list[dict]]:
    groups: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        phone = normalize_phone(row[COL_PHONE])
        groups[phone].append(row)
    return dict(groups)


def build_ttn_params(cfg: Config, group: list[dict], city_ref: str, warehouse_ref: str, recipient: dict) -> dict:
    ids = ",".join(r[COL_ID] for r in group)
    return {
        "PayerType": "Recipient",
        "PaymentMethod": "Cash",
        "CargoType": "Parcel",
        "Weight": cfg.WEIGHT,
        "SeatsAmount": 1,
        "ServiceType": "WarehouseWarehouse",
        "Description": cfg.DESCRIPTION,
        "Cost": cfg.DECLARED_VALUE,
        "OptionsSeat": [{"weight": str(cfg.WEIGHT), "volumetricLength": str(cfg.LENGTH), "volumetricWidth": str(cfg.WIDTH), "volumetricHeight": str(cfg.HEIGHT)}],
        "CitySender": cfg.NP_SENDER_CITY_REF,
        "Sender": cfg.NP_SENDER_REF,
        "SenderAddress": cfg.NP_SENDER_ADDRESS_REF,
        "ContactSender": cfg.NP_SENDER_CONTACT_REF,
        "SendersPhone": cfg.NP_SENDER_PHONE,
        "CityRecipient": city_ref,
        "Recipient": recipient["counterparty_ref"],
        "RecipientAddress": warehouse_ref,
        "ContactRecipient": recipient["contact_ref"],
        "RecipientsPhone": normalize_phone(group[0][COL_PHONE]),
        "InternalNumber": ids,
    }


def process_group(api: NovaPoshtaAPI, cfg: Config, phone: str, group: list[dict], dry_run: bool) -> dict:
    """Обробити одну групу телефону. Повертає рядок результату."""
    first = group[0]
    city_name = first[COL_CITY]
    warehouse_num = first[COL_WAREHOUSE]
    full_name = first[COL_NAME]
    ids = ",".join(r[COL_ID] for r in group)
    products = ";".join(
        f"{r[COL_PRODUCT]} ({r[COL_QTY]})" if r.get(COL_QTY) else r[COL_PRODUCT]
        for r in group
    )

    result_base = {
        "ids": ids,
        "phone": phone,
        "name": full_name,
        "city": city_name,
        "warehouse": warehouse_num,
        "products": products,
    }

    if dry_run:
        print(f"  [dry-run] {ids} | {full_name} | {city_name} №{warehouse_num} | тел.{phone}")
        return {**result_base, "ttn": "DRY-RUN", "status": "dry-run"}

    try:
        # Перевірка: чи вже існує ТТН для цих угод (захист від дублікатів при повторному запуску)
        existing = api.find_ttn_by_internal_number(ids)
        if existing:
            print(f"  ⏭️  Вже існує ТТН {existing} | {full_name} | {city_name} | {ids}")
            return {**result_base, "ttn": existing, "status": "OK"}

        city_ref = api.get_city_ref(normalize_city(city_name))
        warehouse_ref = api.get_warehouse_ref(city_ref, warehouse_num)
        last, first_name, middle = split_name(full_name)
        recipient = api.create_counterparty(first_name, last, middle, normalize_phone(phone))
        params = build_ttn_params(cfg, group, city_ref, warehouse_ref, recipient)
        ttn = api.create_ttn(params)
        print(f"  ✅ ТТН {ttn} | {full_name} | {city_name} | {ids}")
        return {**result_base, "ttn": ttn, "status": "OK"}
    except NovaPoshtaError as e:
        print(f"  ❌ Помилка [{ids}]: {e}")
        return {**result_base, "ttn": "", "status": str(e)}


def main():
    parser = argparse.ArgumentParser(description="Створення ТТН Нова Пошта з Excel-експорту Бітрікс24")
    parser.add_argument("--file", required=True, help="Шлях до Excel-файлу Бітрікс24")
    parser.add_argument("--dry-run", action="store_true", help="Тестовий режим без API-запитів")
    args = parser.parse_args()

    excel_path = Path(args.file)
    if not excel_path.exists():
        print(f"❌ Файл не знайдено: {excel_path}")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)
    cfg = Config()
    api = NovaPoshtaAPI(cfg.NP_API_KEY)

    # Авто-визначення міста відправника з адреси в НП (потрібно для CitySender у ТТН)
    if not args.dry_run and not cfg.NP_SENDER_CITY_REF:
        try:
            addrs = api.get_counterparty_addresses(cfg.NP_SENDER_REF)
            for addr in addrs:
                if addr.get("Ref") == cfg.NP_SENDER_ADDRESS_REF:
                    cfg.NP_SENDER_CITY_REF = addr["CityRef"]
                    print(f"🏙️  CitySender ({addr.get('CityDescription', '?')}): {cfg.NP_SENDER_CITY_REF}")
                    break
            if not cfg.NP_SENDER_CITY_REF:
                print("⚠️  Адресу відправника не знайдено у списку — додайте NP_SENDER_CITY_REF у .env вручну")
        except Exception as e:
            print(f"⚠️  Не вдалося визначити місто відправника: {e}")

    print(f"📂 Читання: {excel_path.name}")
    all_rows = read_bitrix_export(excel_path)
    print(f"   Всього рядків: {len(all_rows)}")

    complete = [r for r in all_rows if is_complete(r)]
    missing = [r for r in all_rows if not is_complete(r)]

    if missing:
        missing_path = write_missing(missing, OUTPUT_DIR)
        print(f"⏭️  Пропущено (немає даних НП): {len(missing)} → {missing_path.name}")

    groups = group_by_phone(complete)
    print(f"\n🔄 Груп за телефоном: {len(groups)}")
    if args.dry_run:
        print("   (режим dry-run — API не викликається)\n")

    results = []
    errors = 0
    for phone, group in groups.items():
        row = process_group(api, cfg, phone, group, dry_run=args.dry_run)
        results.append(row)
        if row["status"] not in ("OK", "dry-run"):
            errors += 1

    # Маппінг: ID угоди → ТТН
    id_to_ttn: dict[str, str] = {}
    for r in results:
        if r.get("ttn") and r["status"] in ("OK", "dry-run"):
            for deal_id in r["ids"].split(","):
                id_to_ttn[deal_id.strip()] = r["ttn"]

    result_path = write_ttn_results(results, OUTPUT_DIR)
    per_deal_path = write_ttn_per_deal(all_rows, id_to_ttn, OUTPUT_DIR)
    today = date.today().strftime("%Y%m%d")

    print()
    ok_count = sum(1 for r in results if r["status"] == "OK")
    print(f"✅ Створено ТТН: {ok_count}")
    if missing:
        print(f"⏭️  Пропущено (немає даних НП): {len(missing)} → missing_{today}.xlsx")
    if errors:
        print(f"❌ Помилок API: {errors}")
    print(f"📁 Результат: {result_path.name}")
    print(f"📋 ТТН по угодах: {per_deal_path.name}")


if __name__ == "__main__":
    main()
