"""
Streamlit-додаток: Агент ТТН — Нова Пошта.
Запуск локально:  streamlit run app.py
Деплой:           Streamlit Cloud → https://share.streamlit.io
"""
import io
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import streamlit as st

# ─── Конфіг сторінки ──────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Агент ТТН — Нова Пошта",
    page_icon="🚀",
    layout="wide",
)

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


# ─── Облікові дані: Streamlit Secrets → env vars ──────────────────────────────
def load_credentials() -> dict[str, str]:
    """Читає ключі з st.secrets (Streamlit Cloud) або зі змінних середовища (.env)."""
    keys = [
        "NP_API_KEY", "NP_SENDER_REF", "NP_SENDER_CONTACT_REF",
        "NP_SENDER_ADDRESS_REF", "NP_SENDER_PHONE", "NP_SENDER_CITY_REF",
        "ARTICLE_MEDAL", "ARTICLE_STATUETTE", "ARTICLE_CUP",
        "ARTICLE_DYPLOM", "ARTICLE_PODYAKA",
        "DESCRIPTION", "DECLARED_VALUE",
        "WEIGHT", "LENGTH", "WIDTH", "HEIGHT",
        "BITRIX_WEBHOOK", "BITRIX_TTN_FIELD",
    ]
    creds: dict[str, str] = {}
    for k in keys:
        try:
            val = st.secrets[k]          # Streamlit Cloud secrets
        except Exception:
            val = os.getenv(k, "")       # локальний .env (через load_dotenv у config.py)
        creds[k] = str(val) if val else ""
    return creds


def run_script(script: str, args: list[str], creds: dict) -> tuple[str, int]:
    """Запускає Python-скрипт, передає credentials через env-змінні."""
    env = os.environ.copy()
    env.update(creds)
    env["PYTHONIOENCODING"] = "utf-8"
    res = subprocess.run(
        [sys.executable, str(BASE_DIR / script)] + args,
        capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        cwd=str(BASE_DIR), timeout=300, env=env,
    )
    return (res.stdout or "") + (res.stderr or ""), res.returncode


def colorize(text: str) -> str:
    """Додає кольорове форматування до консольного виводу (HTML)."""
    lines = []
    for line in text.splitlines():
        if any(x in line for x in ["❌", "Помилка", "Error"]):
            lines.append(f'<span style="color:#ff5555">{line}</span>')
        elif any(x in line for x in ["✅", "ТТН 2"]):
            lines.append(f'<span style="color:#50fa7b">{line}</span>')
        elif any(x in line for x in ["⚠️", "⏭️", "dry-run"]):
            lines.append(f'<span style="color:#f1fa8c">{line}</span>')
        elif any(x in line for x in ["📂", "📦", "📁", "📋", "🏙️"]):
            lines.append(f'<span style="color:#8be9fd">{line}</span>')
        else:
            lines.append(line)
    return "<br>".join(lines)


def console_block(text: str):
    """Рендерить текст у стилізованому консольному блоці."""
    st.markdown(
        f'<div style="background:#282a36;color:#f8f8f2;padding:14px 16px;'
        f'border-radius:8px;font-family:\'Courier New\',monospace;font-size:13px;'
        f'line-height:1.6;max-height:420px;overflow-y:auto;white-space:pre-wrap">'
        f'{colorize(text)}</div>',
        unsafe_allow_html=True,
    )


def download_latest(pattern: str, label: str, col=None):
    """Показує кнопку завантаження для останнього файлу за шаблоном."""
    files = sorted(OUTPUT_DIR.glob(pattern), reverse=True)
    if not files:
        return
    target = col or st
    with open(files[0], "rb") as f:
        target.download_button(
            label=f"⬇ {label}",
            data=f.read(),
            file_name=files[0].name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ─── Шапка ────────────────────────────────────────────────────────────────────
st.title("🚀 Агент ТТН — Нова Пошта")
st.caption("Автоматичне створення накладних для фестивалю Торонто")

creds = load_credentials()

# Статус конфігу
required_keys = [
    ("NP_API_KEY",            "API KEY"),
    ("NP_SENDER_REF",         "SENDER REF"),
    ("NP_SENDER_CONTACT_REF", "CONTACT REF"),
    ("NP_SENDER_ADDRESS_REF", "ADDRESS REF"),
    ("NP_SENDER_PHONE",       "PHONE"),
]
cols = st.columns(len(required_keys) + 1)
all_ok = True
for col, (key, label) in zip(cols, required_keys):
    ok = bool(creds.get(key))
    all_ok = all_ok and ok
    col.markdown(f"{'🟢' if ok else '🔴'} **{label}**")

if all_ok:
    cols[-1].success("✓ Готово до роботи")
else:
    cols[-1].error("⚠ Налаштуйте Secrets")

st.divider()

# ─── Вкладки ──────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📦  Крок 1 — Створення ТТН", "🏭  Крок 2 — Фулфілмент", "🔄  Крок 3 — Оновити Битрікс24"])

# ══════════════════════════════════════════════════════════════════════════════
# КРОК 1
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Створення ТТН з Excel-експорту Бітрікс24")

    uploaded = st.file_uploader(
        "Перетягніть або виберіть Excel-файл з Бітрікс24",
        type=["xlsx", "xls"],
        key="bitrix_file",
    )

    col_dry, col_btn = st.columns([3, 1])
    dry_run = col_dry.checkbox(
        "Тестовий режим (dry-run) — не створювати реальні ТТН",
        value=True,
    )
    run1 = col_btn.button(
        "▶ Запустити",
        type="primary",
        disabled=not (uploaded and all_ok),
        use_container_width=True,
    )

    if run1 and uploaded:
        suffix = Path(uploaded.name).suffix or ".xlsx"
        fd, tmp_path = tempfile.mkstemp(suffix=suffix)
        try:
            os.close(fd)
            with open(tmp_path, "wb") as f:
                f.write(uploaded.getvalue())

            args = ["--file", tmp_path]
            if dry_run:
                args.append("--dry-run")

            with st.spinner("Виконується… це може зайняти до 2 хв."):
                output, rc = run_script("1_create_ttn.py", args, creds)
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

        console_block(output)

        if rc == 0 or "Створено ТТН" in output:
            c1, c2, c3 = st.columns(3)
            download_latest("ttn_per_deal_*.xlsx", "ttn_per_deal.xlsx", c1)
            download_latest("ttn_results_*.xlsx",  "ttn_results.xlsx",  c2)
            download_latest("missing_*.xlsx",       "missing.xlsx",      c3)

            # Зберігаємо шлях для Кроку 2
            ttn_files = sorted(OUTPUT_DIR.glob("ttn_per_deal_*.xlsx"), reverse=True)
            if ttn_files:
                st.session_state["ttn_per_deal"] = str(ttn_files[0])
        else:
            st.error("Скрипт завершився з помилкою. Перевірте вивід вище.")

# ══════════════════════════════════════════════════════════════════════════════
# КРОК 2
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Генерація таблиці замовлень фулфілменту НП")
    st.caption("Вхід: **ttn_per_deal_*.xlsx** з Кроку 1. Вихід: **fulfillment_orders_*.xlsx** (2 аркуші: замовлення + звіт)")

    # Файл — або з output/ (поточна сесія), або завантажений вручну
    uploaded2 = st.file_uploader(
        "Завантажте ttn_per_deal_*.xlsx (якщо щойно запустили Крок 1 — файл підтягнеться автоматично)",
        type=["xlsx"],
        key="ttn_per_deal_file",
    )

    if uploaded2:
        new2 = uploaded2.getvalue()
        if st.session_state.get("_b2_bytes") != new2:
            st.session_state["_b2_bytes"] = new2
            st.session_state["_b2_name"]  = uploaded2.name
    else:
        ttn_files = sorted(OUTPUT_DIR.glob("ttn_per_deal_*.xlsx"), reverse=True)
        if ttn_files and "_b2_bytes" not in st.session_state:
            with open(ttn_files[0], "rb") as f:
                st.session_state["_b2_bytes"] = f.read()
            st.session_state["_b2_name"] = ttn_files[0].name

    b2_bytes = st.session_state.get("_b2_bytes")
    b2_name  = st.session_state.get("_b2_name", "ttn_per_deal.xlsx")

    if b2_bytes:
        st.info(f"📄 Файл: **{b2_name}**")

        run2 = st.button(
            "▶ Сформувати замовлення фулфілменту",
            type="primary",
            use_container_width=False,
        )

        if run2:
            # Зберігаємо у тимчасовий файл
            fd2, tmp2 = tempfile.mkstemp(suffix=".xlsx")
            try:
                os.close(fd2)
                with open(tmp2, "wb") as f:
                    f.write(b2_bytes)
                with st.spinner("Формую таблицю..."):
                    output, rc = run_script(
                        "2_create_fulfillment.py",
                        ["--ttn", tmp2],
                        creds,
                    )
            finally:
                try:
                    os.unlink(tmp2)
                except OSError:
                    pass

            console_block(output)

            ful_files = sorted(OUTPUT_DIR.glob("fulfillment_orders_*.xlsx"), reverse=True)
            if ful_files:
                st.success("✅ Таблиця готова!")
                download_latest("fulfillment_orders_*.xlsx", "fulfillment_orders.xlsx")

                # Попередній перегляд: головний аркуш + звіт
                try:
                    import pandas as pd
                    st.markdown("**📋 Замовлення фулфілменту**")
                    df_main = pd.read_excel(ful_files[0], sheet_name="Фулфілмент замовлення")
                    st.dataframe(df_main, use_container_width=True, height=350)

                    st.markdown("**📊 Звіт**")
                    df_report = pd.read_excel(ful_files[0], sheet_name="Звіт", header=None)
                    st.dataframe(df_report, use_container_width=True, height=300)
                except Exception:
                    pass
    else:
        st.warning("⚠ Завантажте файл ttn_per_deal або спочатку запустіть Крок 1.")

# ══════════════════════════════════════════════════════════════════════════════
# КРОК 3
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("Оновлення ТТН в угодах Битрікс24")
    st.caption("Завантажте **ttn_results_*.xlsx** з Кроку 1 — агент запише номери ТТН у всі угоди.")

    webhook_ok = bool(creds.get("BITRIX_WEBHOOK"))

    if not webhook_ok:
        st.error("⚠ BITRIX_WEBHOOK не налаштовано у Secrets")
    else:
        # Спочатку шукаємо файл з поточної сесії (Крок 1 → Крок 3 без виходу)
        result_files = sorted(OUTPUT_DIR.glob("ttn_results_*.xlsx"), reverse=True)

        uploaded3 = st.file_uploader(
            "Завантажте ttn_results_*.xlsx (якщо щойно запустили Крок 1 — файл підтягнеться автоматично)",
            type=["xlsx"],
            key="results_file",
        )

        # Визначаємо джерело файлу — зберігаємо байти в session_state щоб
        # preview і реальний запуск завжди використовували ОДИН і той самий файл
        if uploaded3:
            # Новий файл завантажено — скидаємо старий preview
            new_bytes = uploaded3.getvalue()
            if st.session_state.get("_b3_bytes") != new_bytes:
                st.session_state["_b3_bytes"] = new_bytes
                st.session_state["_b3_name"] = uploaded3.name
                st.session_state.pop("dry_preview_3", None)
        elif result_files and "_b3_bytes" not in st.session_state:
            # Файл з поточної сесії (Крок 1 → Крок 3)
            with open(result_files[0], "rb") as f:
                st.session_state["_b3_bytes"] = f.read()
            st.session_state["_b3_name"] = result_files[0].name

        b3_bytes = st.session_state.get("_b3_bytes")
        b3_name  = st.session_state.get("_b3_name", "ttn_results.xlsx")

        if b3_bytes:
            st.info(f"📄 Файл: **{b3_name}**")

            # Записуємо у тимчасовий файл (один раз за сесію)
            if "ttn_results_tmp" not in st.session_state:
                fd3, tmp3 = tempfile.mkstemp(suffix=".xlsx")
                os.close(fd3)
                with open(tmp3, "wb") as f:
                    f.write(b3_bytes)
                st.session_state["ttn_results_tmp"] = tmp3
            ttn_results_path = st.session_state["ttn_results_tmp"]

            # Автоматичний dry-run — показуємо список одразу
            if "dry_preview_3" not in st.session_state:
                with st.spinner("Перевірка списку угод..."):
                    preview_out, _ = run_script("3_update_bitrix.py",
                                                ["--ttn", ttn_results_path, "--dry-run"], creds)
                st.session_state["dry_preview_3"] = preview_out

            console_block(st.session_state["dry_preview_3"])

            st.divider()
            confirmed = st.checkbox("✅ Список правильний — оновити угоди в Битрікс24", key="confirm3")
            run3 = st.button("▶ Оновити Битрікс24", type="primary",
                             disabled=not confirmed, key="run3", use_container_width=False)

            if run3:
                with st.spinner("Оновлення угод..."):
                    output3, rc3 = run_script("3_update_bitrix.py", ["--ttn", ttn_results_path], creds)
                console_block(output3)
                if rc3 == 0:
                    st.success("✅ Угоди оновлено в Битрікс24!")
                    # Скидаємо кеш для наступного запуску
                    for k in ("dry_preview_3", "ttn_results_tmp", "_b3_bytes", "_b3_name"):
                        st.session_state.pop(k, None)
        else:
            st.warning("⚠ Завантажте файл ttn_results або спочатку запустіть Крок 1.")
