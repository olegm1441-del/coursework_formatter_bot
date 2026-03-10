import argparse
from dataclasses import dataclass
from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from sqlalchemy import and_, exists

from db import SessionLocal
from models import Document, FormattingRequest, User


RETRY_SOURCE = "mass_requeue_after_formatter_fix"
DEFAULT_FAILED_STATUSES = ["failed", "rejected", "formatting_failed", "unprocessable", "invalid", "timeout"]


@dataclass
class Counters:
    found: int = 0
    requeued: int = 0
    skipped: int = 0
    prep_failed: int = 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Mass requeue failed formatting requests")
    parser.add_argument(
        "--statuses",
        default=",".join(DEFAULT_FAILED_STATUSES),
        help="Comma-separated failed statuses to include",
    )
    parser.add_argument(
        "--retry-source",
        default=RETRY_SOURCE,
        help="Value to write into formatting_requests.retry_source",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Actually insert queued retry requests (without this flag runs in dry-run mode)",
    )
    return parser.parse_args()


def get_candidate_documents(db, failed_statuses: list[str], retry_source: str) -> list[Document]:
    has_failed = exists().where(
        and_(
            FormattingRequest.document_id == Document.id,
            FormattingRequest.status.in_(failed_statuses),
        )
    )
    has_success = exists().where(
        and_(
            FormattingRequest.document_id == Document.id,
            FormattingRequest.status == "done",
            FormattingRequest.result_file_path.isnot(None),
        )
    )
    already_requeued = exists().where(
        and_(
            FormattingRequest.document_id == Document.id,
            FormattingRequest.retry_source == retry_source,
        )
    )

    return (
        db.query(Document)
        .join(User, User.id == Document.user_id)
        .filter(Document.storage_path.isnot(None))
        .filter(Document.storage_path != "")
        .filter(has_failed)
        .filter(~has_success)
        .filter(~already_requeued)
        .order_by(Document.id.asc())
        .all()
    )


def build_requeue_request(document: Document, latest_failed_request: FormattingRequest, retry_source: str) -> FormattingRequest:
    return FormattingRequest(
        user_id=document.user_id,
        document_id=document.id,
        service_type=latest_failed_request.service_type,
        university_code=latest_failed_request.university_code,
        document_type=latest_failed_request.document_type,
        guideline_version=latest_failed_request.guideline_version,
        status="queued",
        result_file_path=None,
        error_message=None,
        completed_at=None,
        silent_fail=True,
        retry_source=retry_source,
    )


def main() -> None:
    args = parse_args()
    failed_statuses = [status.strip() for status in args.statuses.split(",") if status.strip()]
    counters = Counters()

    db = SessionLocal()
    try:
        candidates = get_candidate_documents(db, failed_statuses, args.retry_source)
        counters.found = len(candidates)

        for document in candidates:
            latest_failed_request = (
                db.query(FormattingRequest)
                .filter(FormattingRequest.document_id == document.id)
                .filter(FormattingRequest.status.in_(failed_statuses))
                .order_by(FormattingRequest.id.desc())
                .first()
            )

            if not latest_failed_request:
                counters.prep_failed += 1
                continue

            already_requeued = (
                db.query(FormattingRequest)
                .filter(FormattingRequest.document_id == document.id)
                .filter(FormattingRequest.retry_source == args.retry_source)
                .first()
            )
            if already_requeued:
                counters.skipped += 1
                continue

            if not args.execute:
                counters.requeued += 1
                continue

            try:
                db.add(build_requeue_request(document, latest_failed_request, args.retry_source))
                db.commit()
                counters.requeued += 1
            except Exception:
                db.rollback()
                counters.prep_failed += 1

    finally:
        db.close()

    mode = "EXECUTE" if args.execute else "DRY-RUN"
    print(f"Mode: {mode}")
    print(f"retry_source={args.retry_source}")
    print(f"statuses={failed_statuses}")
    print(f"found={counters.found}")
    print(f"requeued={counters.requeued}")
    print(f"skipped={counters.skipped}")
    print(f"prep_failed={counters.prep_failed}")


if __name__ == "__main__":
    main()
