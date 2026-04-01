"""Конфігурація агента ТТН: завантаження .env та константи."""
import os
from pathlib import Path
from dotenv import load_dotenv

# Шукаємо .env у директорії скрипта
_env_path = Path(__file__).parent / ".env"
load_dotenv(_env_path)


def _require(name: str) -> str:
    val = os.getenv(name)
    if not val:
        raise EnvironmentError(
            f"Змінна {name} не задана у .env\n"
            f"Скопіюйте .env.template → .env та заповніть всі поля."
        )
    return val


class Config:
    # API
    NP_API_KEY: str = _require("NP_API_KEY")
    NP_SENDER_REF: str = _require("NP_SENDER_REF")
    NP_SENDER_CONTACT_REF: str = _require("NP_SENDER_CONTACT_REF")
    NP_SENDER_ADDRESS_REF: str = _require("NP_SENDER_ADDRESS_REF")
    NP_SENDER_PHONE: str = _require("NP_SENDER_PHONE")
    NP_SENDER_CITY_REF: str = os.getenv("NP_SENDER_CITY_REF", "")  # авто-визначається при запуску

    # Параметри посилки
    DESCRIPTION: str = os.getenv("DESCRIPTION", "Нагороди фестивалю Торонто 2025")
    DECLARED_VALUE: int = int(os.getenv("DECLARED_VALUE", "200"))
    WEIGHT: float = float(os.getenv("WEIGHT", "0.5"))
    LENGTH: int = int(os.getenv("LENGTH", "21"))
    WIDTH: int = int(os.getenv("WIDTH", "5"))
    HEIGHT: int = int(os.getenv("HEIGHT", "30"))

    # Битрікс24
    BITRIX_WEBHOOK: str = os.getenv("BITRIX_WEBHOOK", "")
    BITRIX_TTN_FIELD: str = os.getenv("BITRIX_TTN_FIELD", "UF_CRM_1704712295456")

    # Артикули
    ARTICLE_MEDAL: str = os.getenv("ARTICLE_MEDAL", "MED-001")
    ARTICLE_STATUETTE: str = os.getenv("ARTICLE_STATUETTE", "STAT-001")
    ARTICLE_CUP: str = os.getenv("ARTICLE_CUP", "CUP-001")
    ARTICLE_DYPLOM: str = os.getenv("ARTICLE_DYPLOM", "")
    ARTICLE_PODYAKA: str = os.getenv("ARTICLE_PODYAKA", "")


# Маппінг ключових слів → артикул (перевіряється у порядку списку)
ARTICLE_KEYWORDS = [
    ("медаль", lambda cfg: cfg.ARTICLE_MEDAL),
    ("статует", lambda cfg: cfg.ARTICLE_STATUETTE),
    ("кубок", lambda cfg: cfg.ARTICLE_CUP),
    ("диплом", lambda cfg: cfg.ARTICLE_DYPLOM),
    ("подяк", lambda cfg: cfg.ARTICLE_PODYAKA),
]
