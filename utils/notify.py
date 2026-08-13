"""Сповіщення в Telegram про результат прогону агента ТТН."""
import requests


def notify_telegram(text: str, bot_token: str, chat_id: str) -> None:
    """Надіслати повідомлення в Telegram. Тихо ігнорує помилки — сповіщення
    не повинно рвати основний прогін, якщо Telegram недоступний."""
    if not bot_token or not chat_id:
        return
    try:
        requests.post(
            f"https://api.telegram.org/bot{bot_token}/sendMessage",
            json={"chat_id": chat_id, "text": text},
            timeout=10,
        )
    except requests.exceptions.RequestException as e:
        print(f"⚠️  Не вдалося надіслати сповіщення в Telegram: {e}")
