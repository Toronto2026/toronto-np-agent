"""Веб-інтерфейс агента ТТН Нова Пошта."""
import io
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from flask import Flask, render_template, request, jsonify, Response
from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent
load_dotenv(BASE_DIR / ".env")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB


def get_env_status() -> dict:
    keys = ["NP_API_KEY", "NP_SENDER_REF", "NP_SENDER_CONTACT_REF",
            "NP_SENDER_ADDRESS_REF", "NP_SENDER_PHONE"]
    status = {}
    for k in keys:
        val = os.getenv(k, "")
        status[k] = bool(val and not val.startswith("х") and not val.startswith("В") and "xxx" not in val)
    return status


@app.route("/")
def index():
    env_status = get_env_status()
    configured = all(env_status.values())
    return render_template("index.html", env_status=env_status, configured=configured)


@app.route("/run", methods=["POST"])
def run_script():
    script = request.form.get("script", "ttn")  # "ttn" or "fulfillment"
    dry_run = request.form.get("dry_run") == "1"
    file = request.files.get("file")
    ttn_file = request.files.get("ttn_file")

    def save_temp(f, suffix=".xlsx") -> str:
        """Зберегти завантажений файл у тимчасовий файл. Повертає шлях."""
        fd, path = tempfile.mkstemp(suffix=suffix)
        os.close(fd)
        f.save(path)
        return path

    if script == "ttn":
        if not file or not file.filename:
            return jsonify({"error": "Файл не вибрано"}), 400
        suffix = Path(file.filename).suffix or ".xlsx"
        tmp_path = save_temp(file, suffix)
        cmd = [sys.executable, str(BASE_DIR / "1_create_ttn.py"), "--file", tmp_path]
        if dry_run:
            cmd.append("--dry-run")

    elif script == "fulfillment":
        # Автоматично використовуємо останній ttn_per_deal файл з output/
        output_dir = BASE_DIR / "output"
        ttn_per_deal_files = sorted(output_dir.glob("ttn_per_deal_*.xlsx"), reverse=True)
        if not ttn_per_deal_files:
            return jsonify({"error": "Файл ttn_per_deal не знайдено. Спочатку запустіть Крок 1."}), 400
        ttn_path = str(ttn_per_deal_files[0])
        cmd = [sys.executable, str(BASE_DIR / "2_create_fulfillment.py"), "--ttn", ttn_path]
    else:
        return jsonify({"error": "Невідомий скрипт"}), 400

    try:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=str(BASE_DIR),
            timeout=300,
            env=env,
        )
        output = (result.stdout or "") + (result.stderr or "")

        # Знаходимо вихідний файл
        output_files = []
        output_dir = BASE_DIR / "output"
        if script == "ttn":
            for f in sorted(output_dir.glob("ttn_per_deal_*.xlsx"), reverse=True)[:1]:
                output_files.append(f.name)
            for f in sorted(output_dir.glob("ttn_results_*.xlsx"), reverse=True)[:1]:
                output_files.append(f.name)
            for f in sorted(output_dir.glob("missing_*.xlsx"), reverse=True)[:1]:
                output_files.append(f.name)
        elif script == "fulfillment":
            for f in sorted(output_dir.glob("fulfillment_orders_*.xlsx"), reverse=True)[:1]:
                output_files.append(f.name)

        return jsonify({
            "output": output,
            "returncode": result.returncode,
            "output_files": output_files,
        })
    except subprocess.TimeoutExpired:
        return jsonify({"error": "Таймаут (5 хв). Перевірте підключення до інтернету."}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/output/<filename>")
def download_output(filename):
    from flask import send_from_directory
    return send_from_directory(BASE_DIR / "output", filename, as_attachment=True)


if __name__ == "__main__":
    print("=" * 50)
    print("  Агент ТТН — Веб-інтерфейс")
    print("  Відкрийте браузер: http://localhost:5055")
    print("=" * 50)
    app.run(host="127.0.0.1", port=5055, debug=False)
