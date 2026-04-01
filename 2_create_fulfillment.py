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

from config import Config, ARTICLE_KEYWORDS
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
    """Повертає True якщо товар є оплатою/послугою — не фізичний для фулфілменту."""
    lower = product_name.lower()
    return any(kw in lower for kw in _SKIP_KEYWORDS)


def resolve_article(product_name: str, deal_id: str, cfg: Config) -> str:
    """Визначити артикул за назвою товару (пошук підрядка без урахування регістру)."""
    lower = product_name.lower()
    for keyword, article_fn in ARTICLE_KEYWORDS:
        if keyword in lower:
            article = article_fn(cfg)
            if not article:
                # Fallback якщо артикул не заданий у .env:
                # диплом → ID угоди, подяка → P-{ID угоди} (формат НП фулфілменту)
                if "диплом" in lower:
                    return deal_id
                if "подяк" in lower:
                    return f"P-{deal_id}"
            return article
    return ""


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

            article = resolve_article(product, deal_id, cfg)
            if not article:
                print(f"  ⚠️  Невизначений артикул: ТТН={ttn}, Товар={product!r}")
                unknown_count += 1
                article = f"???({product[:15]})"

            articles[article] = articles.get(article, 0) + qty

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
