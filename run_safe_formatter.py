import sys
from pathlib import Path

from safe_formatter import process_document


def collect_docx_files(input_dir: Path):
    return sorted(
        p for p in input_dir.iterdir()
        if p.is_file()
        and p.suffix.lower() == ".docx"
        and not p.name.startswith("~$")
        and not p.stem.endswith("_updated")
        and not p.stem.endswith("_safe")
    )


def main():
    if len(sys.argv) != 3:
        print("Использование:")
        print("  python3 run_safe_formatter.py input output")
        sys.exit(1)

    input_dir = Path(sys.argv[1]).expanduser().resolve()
    output_dir = Path(sys.argv[2]).expanduser().resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Ошибка: папка input не найдена: {input_dir}")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    files = collect_docx_files(input_dir)
    if not files:
        print(f"В папке {input_dir} не найдено .docx файлов")
        sys.exit(0)

    ok = 0
    err = 0

    for file_path in files:
        out_path = output_dir / f"{file_path.stem}_safe.docx"
        try:
            process_document(file_path, out_path)
            print(f"[OK] {file_path.name} -> {out_path.name}")
            ok += 1
        except Exception as e:
            print(f"[ERR] {file_path.name}: {e}")
            err += 1

    print("-" * 60)
    print(f"Готово. Успешно: {ok}, ошибок: {err}")
    print(f"Результат: {output_dir}")


if __name__ == "__main__":
    main()
