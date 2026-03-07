from pathlib import Path
from safe_formatter import process_document


def format_docx(input_path: str, output_path: str) -> str:
    """
    Форматирует один docx файл.
    Возвращает путь к выходному файлу.
    """

    input_path = Path(input_path)
    output_path = Path(output_path)

    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    if input_path.suffix.lower() != ".docx":
        raise ValueError("Поддерживаются только .docx файлы")

    process_document(input_path, output_path)

    if not output_path.exists():
        raise RuntimeError("Файл не был создан")

    return str(output_path)
