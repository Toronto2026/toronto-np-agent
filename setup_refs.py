"""
Крок 0: Отримати ref відправника, контакту та адреси складу фулфілменту.
Запуск: python setup_refs.py
"""
import os
import sys
from pathlib import Path

# Дозволяємо запускати з будь-якої директорії
sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv
load_dotenv(Path(__file__).parent / ".env")

api_key = os.getenv("NP_API_KEY")
if not api_key or api_key.startswith("В"):
    print("❌ Вставте ваш NP_API_KEY у файл .env")
    sys.exit(1)

from utils.np_api import NovaPoshtaAPI, NovaPoshtaError

api = NovaPoshtaAPI(api_key)

print("=" * 60)
print("КРОК 1: Список відправників (CounterpartyProperty=Sender)")
print("=" * 60)
try:
    senders = api.get_counterparties("Sender")
    if not senders:
        print("  (пусто — можливо, ваш рахунок не є відправником у НП)")
    for s in senders:
        print(f"  Ref:  {s['Ref']}")
        print(f"  Назва: {s.get('Description', '')} / {s.get('FullName', '')}")
        print(f"  ЄДРПОУ: {s.get('EDRPOU', '')}")
        print()
except NovaPoshtaError as e:
    print(f"  ❌ {e}")

print()
sender_ref = input("Введіть NP_SENDER_REF (скопіюйте Ref вашого ФОП з переліку вище): ").strip()
if not sender_ref:
    print("Відмінено.")
    sys.exit(0)

print()
print("=" * 60)
print("КРОК 2: Контактні особи відправника")
print("=" * 60)
try:
    contacts = api.get_counterparty_contact_persons(sender_ref)
    if not contacts:
        print("  (пусто)")
    for c in contacts:
        print(f"  Ref:  {c['Ref']}")
        print(f"  ПІБ:  {c.get('LastName', '')} {c.get('FirstName', '')} {c.get('MiddleName', '')}")
        print()
except NovaPoshtaError as e:
    print(f"  ❌ {e}")
    contacts = []

contact_ref = contacts[0]["Ref"] if contacts else ""
print(f"  → Буде використано перший контакт: {contact_ref}")

print()
print("=" * 60)
print("КРОК 3: Адреси відправника (шукаємо фулфілмент)")
print("=" * 60)
try:
    addresses = api.get_counterparty_addresses(sender_ref)
    if not addresses:
        print("  (пусто)")
    for a in addresses:
        desc = a.get("Description", "")
        print(f"  Ref:  {a['Ref']}")
        print(f"  Адреса: {desc}")
        if "фулфілм" in desc.lower() or "броварськ" in desc.lower() or "склад" in desc.lower():
            print("  ^^^ МОЖЛИВИЙ СКЛАД ФУЛФІЛМЕНТУ")
        print()
except NovaPoshtaError as e:
    print(f"  ❌ {e}")
    addresses = []

address_ref = ""
for a in addresses:
    desc = a.get("Description", "").lower()
    if "фулфілм" in desc or "броварськ" in desc or "склад №3" in desc:
        address_ref = a["Ref"]
        print(f"  → Автоматично вибрано склад фулфілменту: {a.get('Description', '')}")
        break
if not address_ref:
    address_ref = addresses[0]["Ref"] if addresses else ""
    if address_ref:
        print(f"  → Склад фулфілменту не знайдено автоматично, вибрано першу адресу")
    else:
        address_ref = input("Адреса не знайдена. Введіть NP_SENDER_ADDRESS_REF вручну: ").strip()

print()
print("=" * 60)
print("Вставте у .env:")
print("=" * 60)
print(f"NP_SENDER_REF={sender_ref}")
print(f"NP_SENDER_CONTACT_REF={contact_ref}")
print(f"NP_SENDER_ADDRESS_REF={address_ref}")
print()
print("✅ Готово! Тепер запустіть:")
print("   python 1_create_ttn.py --file export_bitrix.xlsx --dry-run")
