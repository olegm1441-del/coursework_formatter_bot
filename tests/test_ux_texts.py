from __future__ import annotations

import os
import sys
from pathlib import Path

os.environ.setdefault("DATABASE_URL", "sqlite:////tmp/test_ux_texts.sqlite")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import services


def test_file_received_format_text() -> tuple[bool, str]:
    expected = (
        "Файл получен.\n\n"
        "Начинаю оформление по методичке.\n"
        "Готовый .docx-файл отправлю через минуту."
    )
    actual = services.build_file_received_text("format")
    if actual != expected:
        return False, f"unexpected format text: {actual!r}"
    return True, "format file-received text matches product copy"


def test_file_received_check_text_unchanged() -> tuple[bool, str]:
    expected = (
        "Файл получен.\n\n"
        "Начинаю проверку оформления по методичке КФУ."
    )
    actual = services.build_file_received_text("check")
    if actual != expected:
        return False, f"unexpected check text: {actual!r}"
    return True, "check file-received text unchanged"


def main() -> int:
    tests = [
        ("file received format text", test_file_received_format_text),
        ("file received check text", test_file_received_check_text_unchanged),
    ]
    failed = 0
    for name, fn in tests:
        ok, msg = fn()
        status = "PASS" if ok else "FAIL"
        print(f"[{status}] {name} — {msg}")
        if not ok:
            failed += 1
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
