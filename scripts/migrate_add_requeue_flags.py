from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from sqlalchemy import inspect, text

from db import engine


TABLE_NAME = "formatting_requests"


def main() -> None:
    inspector = inspect(engine)
    columns = {column["name"] for column in inspector.get_columns(TABLE_NAME)}

    with engine.begin() as conn:
        if "silent_fail" not in columns:
            conn.execute(
                text(
                    "ALTER TABLE formatting_requests "
                    "ADD COLUMN silent_fail BOOLEAN NOT NULL DEFAULT FALSE"
                )
            )
            print("Added formatting_requests.silent_fail")
        else:
            print("Column formatting_requests.silent_fail already exists")

        if "retry_source" not in columns:
            conn.execute(
                text(
                    "ALTER TABLE formatting_requests "
                    "ADD COLUMN retry_source VARCHAR(100)"
                )
            )
            print("Added formatting_requests.retry_source")
        else:
            print("Column formatting_requests.retry_source already exists")

        indexes = {index["name"] for index in inspector.get_indexes(TABLE_NAME)}
        if "ix_formatting_requests_retry_source" not in indexes:
            conn.execute(
                text(
                    "CREATE INDEX ix_formatting_requests_retry_source "
                    "ON formatting_requests (retry_source)"
                )
            )
            print("Created index ix_formatting_requests_retry_source")
        else:
            print("Index ix_formatting_requests_retry_source already exists")


if __name__ == "__main__":
    main()
