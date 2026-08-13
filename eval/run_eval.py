"""Evaluation pipeline для класифікації заявок агента ТТН.

Датасет (test_cases.json) зібраний з реальних заявок, які раніше ламали
або зашумлювали звіт агента (missing_20260605.xlsx мав 314 "пропущених"
рядків, з яких лише 3 були реальною проблемою). Грейдер тут — код-базовий:
порівнює категорію, яку визначає 1_create_ttn.py, з очікуваною.

Запуск:
  python eval/run_eval.py
"""
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

AGENT_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(AGENT_DIR))

from importlib import import_module
agent = import_module("1_create_ttn")

TEST_CASES_PATH = Path(__file__).parent / "test_cases.json"


def classify(row: dict) -> str:
    """Та сама логіка розгалуження, що й у 1_create_ttn.py main()."""
    if agent.is_electronic_only(row):
        return "electronic_only"
    if agent.is_foreign_phone(row):
        return "foreign"
    if agent.is_complete(row):
        return "complete"
    return "missing_data"


def main() -> int:
    cases = json.loads(TEST_CASES_PATH.read_text(encoding="utf-8"))

    passed = 0
    failed = []
    for case in cases:
        row = {k: case[k] for k in ("id", "name", "product", "qty", "phone", "city", "warehouse")}
        actual = classify(row)
        expected = case["expected_category"]
        ok = actual == expected

        if ok and expected == "complete" and case["phone"]:
            digits = agent.normalize_phone(case["phone"])
            ok = digits.isdigit() and len(digits) == 12

        if ok:
            passed += 1
        else:
            failed.append((case, actual))

    total = len(cases)
    print(f"{'=' * 60}")
    for case, actual in failed:
        print(f"❌ FAIL  id={case['id']:<10} очікували={case['expected_category']:<15} отримали={actual}")
        print(f"        {case['note']}")
    print(f"{'=' * 60}")
    print(f"Пройдено: {passed}/{total} ({passed / total * 100:.0f}%)")

    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
