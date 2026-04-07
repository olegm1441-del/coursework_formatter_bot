"""
Migration: add selected_guide_code column to users table.

Run once:
    python scripts/migrate_add_selected_guide_code.py
"""

from dotenv import load_dotenv
load_dotenv()

from db import engine

with engine.connect() as conn:
    try:
        conn.execute(
            __import__("sqlalchemy").text(
                "ALTER TABLE users ADD COLUMN selected_guide_code VARCHAR(100)"
            )
        )
        conn.commit()
        print("OK: column selected_guide_code added to users")
    except Exception as e:
        if "duplicate column" in str(e).lower() or "already exists" in str(e).lower():
            print("SKIP: column selected_guide_code already exists")
        else:
            raise
