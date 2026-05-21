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
        "📄 Начинаю оформление по методичке КФУ.\n\n"
        "Готовый .docx-файл отправлю сюда автоматически.\n"
        "После обработки документ будет удалён из системы"
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


def test_start_text_mentions_file_deletion_and_informal_consent() -> tuple[bool, str]:
    actual = services.build_start_text(
        balance=1,
        is_new=True,
        active_guide_title="КФУ — курсовая 2025",
        referral_progress=0,
        referral_target=3,
    )
    expected_fragments = [
        "Готовый файл вернётся сюда автоматически и удалится из нашей системы.",
        "Продолжая использование бота, ты соглашаешься с",
    ]
    missing = [fragment for fragment in expected_fragments if fragment not in actual]
    if missing:
        return False, f"missing start text fragments: {missing!r}"
    return True, "start text includes deletion notice and informal consent"


def main() -> int:
    tests = [
        ("file received format text", test_file_received_format_text),
        ("file received check text", test_file_received_check_text_unchanged),
        ("start text deletion notice", test_start_text_mentions_file_deletion_and_informal_consent),
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
