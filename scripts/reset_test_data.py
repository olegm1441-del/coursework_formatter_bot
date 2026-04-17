"""
Reset test data without dropping tables.

Safety:
    TEST_MODE=true python scripts/reset_test_data.py
"""

import os
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from db import SessionLocal
from models import CreditLedger, Document, FormattingRequest, Referral, User


def main() -> None:
    if os.getenv("TEST_MODE", "").strip().lower() != "true":
        raise SystemExit("Refusing to reset data: set TEST_MODE=true")

    db = SessionLocal()
    try:
        for model in (FormattingRequest, Document, CreditLedger, Referral, User):
            deleted = db.query(model).delete(synchronize_session=False)
            print(f"Deleted {deleted} rows from {model.__tablename__}")
        db.commit()
        print("Test data reset complete")
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


if __name__ == "__main__":
    main()
