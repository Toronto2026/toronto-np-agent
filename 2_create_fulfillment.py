"""
Скрипт 2: ttn_per_deal_*.xlsx → Таблиця замовлень фулфілменту НП.

Запуск:
  python 2_create_fulfillment.py --ttn output/ttn_per_deal_20260303.xlsx

Вихідний файл: output/fulfillment_orders_YYYYMMDD.xlsx
Колонки: ТТН | Номер замовлення | Артикул | Кількість | ПІБ | Місто
"""
import argparse
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from config import Config
from utils.excel import read_ttn_per_deal, write_fulfillment_orders

OUTPUT_DIR = Path(__file__).parent / "output"

# Українська → латиниця (для Номера замовлення)
_UA_TO_LAT = {
    'А': 'A',  'Б': 'B',  'В': 'V',  'Г': 'H',  'Д': 'D',  'Е': 'E',
    'Є': 'YE', 'Ж': 'ZH', 'З': 'Z',  'И': 'Y',  'І': 'I',  'Ї': 'I',
    'Й': 'Y',  'К': 'K',  'Л': 'L',  'М': 'M',  'Н': 'N',  'О': 'O',
    'П': 'P',  'Р': 'R',  'С': 'S',  'Т': 'T',  'У': 'U',  'Ф': 'F',
    'Х': 'KH', 'Ц': 'TS', 'Ч': 'CH', 'Ш': 'SH', 'Щ': 'SHCH',
    'Ю': 'YU', 'Я': 'YA',
}


def city_code(city_name: str) -> str:
    """Перші 3 латинські символи назви міста (верхній регістр).

    Приклади:
      Запоріжжя → ZAP, Дніпро → DNI, Київ → KYI
      Жовква → ZHO, Чернівці → CHE, Шостка → SHO
    """
    result = ""
    for char in city_name.upper():
        if len(result) >= 3:
            break
        lat = _UA_TO_LAT.get(char, char if (char.isascii() and char.isalpha()) else "")
        result += lat
    return result[:3].upper()


# Ключові слова для рядків-оплат які треба ігнорувати (не фізичний товар)
_SKIP_KEYWORDS = ["внесок", "реєстрація", "subscription", "оплата", "послуга"]


def should_skip(product_name: str) -> bool:
    """Повертає True якщо товар є оплатою/послугою — не фізичний для фулфілменту.

    Виняток: якщо назва містить "комплект" — це фізичний набір нагород,
    навіть якщо в назві є слово "внесок" чи "організаційний".
    """
    lower = product_name.lower()
    if "комплект" in lower:
        return False
    return any(kw in lower for kw in _SKIP_KEYWORDS)


def _norm(s: str) -> str:
    """Нормалізація для порівняння: є→е, ї→и, нижній регістр."""
    return s.lower().replace("є", "е").replace("ї", "и")


def resolve_articles(product_name: str, deal_id: str, deal_qty: int, cfg: Config) -> list[tuple[str, int]]:
    """Повертає список (артикул, кількість) для одного рядка угоди.

    Правила:
      - "повний комплект" → диплом + подяка (завжди 1) + медаль + статуєтка
      - "диплом"          → {deal_id},  qty = deal_qty
      - "подяк"           → P-{deal_id}, qty = 1 (один керівник незалежно від кількості)
      - "медаль"          → ARTICLE_MEDAL,     qty = deal_qty
      - "статует/статує"  → ARTICLE_STATUETTE, qty = deal_qty
      - "кубок"           → ARTICLE_CUP,       qty = deal_qty
    """
    n = _norm(product_name)

    dyplom_art   = cfg.ARTICLE_DYPLOM    or deal_id
    podyaka_art  = cfg.ARTICLE_PODYAKA   or f"P-{deal_id}"

    # Повний комплект: 4 позиції, кількість = кількість комплектів
    if "комплект" in n:
        items: list[tuple[str, int]] = [
            (dyplom_art,  deal_qty),
            (podyaka_art, deal_qty),
        ]
        if cfg.ARTICLE_MEDAL:
            items.append((cfg.ARTICLE_MEDAL, deal_qty))
        if cfg.ARTICLE_STATUETTE:
            items.append((cfg.ARTICLE_STATUETTE, deal_qty))
        return items

    if "диплом" in n:
        return [(dyplom_art, deal_qty)]

    if "подяк" in n:
        return [(podyaka_art, 1)]

    # "статует" (з е) або "статує" (з є, після нормалізації → "статуе")
    if "статует" in n or "статуе" in n:
        art = cfg.ARTICLE_STATUETTE
        return [(art, deal_qty)] if art else []

    if "медаль" in n:
        art = cfg.ARTICLE_MEDAL
        return [(art, deal_qty)] if art else []

    if "кубок" in n:
        art = cfg.ARTICLE_CUP
        return [(art, deal_qty)] if art else []

    return []


def main():
    parser = argparse.ArgumentParser(
        description="Формування таблиці замовлень фулфілменту НП"
    )
    parser.add_argument("--ttn", required=True, help="Файл ttn_per_deal_*.xlsx")
    args = parser.parse_args()

    ttn_path = Path(args.ttn)
    if not ttn_path.exists():
        print(f"❌ Файл не знайдено: {ttn_path}")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)
    cfg = Config()

    rows = read_ttn_per_deal(ttn_path)
    print(f"📂 Рядків у файлі: {len(rows)}")

    # Групуємо по ТТН (зберігаємо порядок)
    ttn_groups: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        ttn = row.get("ТТН", "").strip()
        if ttn and ttn != "DRY-RUN":
            ttn_groups[ttn].append(row)

    print(f"📦 Унікальних ТТН: {len(ttn_groups)}")

    fulfillment_rows = []
    unknown_count = 0
    order_counter: dict[str, int] = {}   # base → кількість використань

    for ttn, group in ttn_groups.items():
        first = group[0]
        city  = first.get("Місто", "")
        wh    = str(first.get("Відділення", "")).strip()
        name  = first.get("ПІБ отримувача", "")

        # Номер замовлення: 3 літери міста + номер відділення.
        # Якщо два різних ТТН мають однаковий базовий код → додаємо суфікс -2, -3, …
        base = city_code(city) + wh
        order_counter[base] = order_counter.get(base, 0) + 1
        cnt = order_counter[base]
        order_num = base if cnt == 1 else f"{base}-{cnt}"

        # Збираємо артикули по всіх рядках ТТН (агрегуємо однакові)
        articles: dict[str, int] = {}
        for row in group:
            deal_id = str(row.get("ID угоди", "")).strip()
            product = str(row.get("Товар", "")).strip()
            try:
                qty = int(float(str(row.get("Кількість", 1) or 1)))
            except (ValueError, TypeError):
                qty = 1

            # Ігнорувати оплати / організаційні внески
            if should_skip(product):
                continue

            pairs = resolve_articles(product, deal_id, qty, cfg)
            if not pairs:
                print(f"  ⚠️  Невизначений артикул: ТТН={ttn}, Товар={product!r}")
                unknown_count += 1
                articles[f"???({product[:15]})"] = articles.get(f"???({product[:15]})", 0) + qty
            else:
                for art, art_qty in pairs:
                    articles[art] = articles.get(art, 0) + art_qty

        # Один рядок на артикул
        for article, qty in articles.items():
            fulfillment_rows.append({
                "ttn":          ttn,
                "order_number": order_num,
                "article":      article,
                "qty":          qty,
                "name":         name,
                "city":         city,
            })
            print(f"  📋 {ttn} | {order_num} | {article} × {qty} | {name}")

    out_path = write_fulfillment_orders(fulfillment_rows, OUTPUT_DIR)

    print()
    print(f"✅ Рядків у таблиці: {len(fulfillment_rows)}")
    if unknown_count:
        print(f"⚠️  Невизначених артикулів: {unknown_count} → перевірте вручну")
    print(f"📁 Результат: {out_path.name}")


if __name__ == "__main__":
    main()
