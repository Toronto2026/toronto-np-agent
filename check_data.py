import sys, re
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')
from utils.excel import read_bitrix_export, COL_PHONE, COL_CITY

path = r'C:\Users\ozonv\OneDrive\Documents\ТОРОНТО\ЛЮТИЙ 2026\НОВА ПОШТА\!!!Формування ТТН для НП - січень 2026.xls.xlsx'
rows = read_bitrix_export(path)

def normalize_phone(phone):
    digits = ''.join(c for c in phone if c.isdigit())
    if digits.startswith('380') and len(digits) == 12: return digits
    if digits.startswith('80') and len(digits) == 11: return '3' + digits
    if digits.startswith('0') and len(digits) == 10: return '38' + digits
    if len(digits) == 9: return '380' + digits
    return digits

def normalize_city(city):
    city = city.strip()
    city = re.sub(r'^(смт\.\s*|м\.\s*|с\.\s*|сел\.\s*)', '', city, flags=re.IGNORECASE)
    city = re.sub(r'\s*,.*$', '', city)
    return city.strip()

print('=== Телефони після виправлення ===')
seen = set()
for r in rows:
    p = r.get(COL_PHONE,'')
    if not p or p in seen: continue
    seen.add(p)
    norm = normalize_phone(p)
    nd = ''.join(c for c in norm if c.isdigit())
    ok = len(nd) == 12
    changed = (p != norm)
    if not ok:
        print(f'  ERR  {p!r} -> {norm!r}  ({len(nd)} цифр)')
    elif changed:
        print(f'  FIX  {p!r} -> {norm!r}')

print()
print('=== Міста з префіксами ===')
seen2 = set()
for r in rows:
    c = r.get(COL_CITY,'')
    norm = normalize_city(c)
    if c != norm and c not in seen2:
        seen2.add(c)
        print(f'  {c!r} -> {norm!r}')
