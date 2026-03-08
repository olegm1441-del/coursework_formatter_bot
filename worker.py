import asyncio
import logging
import os
import time
import traceback
from datetime import datetime, timedelta
from multiprocessing import Process, Queue
from pathlib import Path

from dotenv import load_dotenv
from sqlalchemy.orm import Session
from telegram import Bot

from db import Base, SessionLocal, engine
import models  # noqa: F401
from models import Document, FormattingRequest, User
import services


FORMAT_TIMEOUT_SECONDS = 60
STALE_PROCESSING_SECONDS = 5 * 60
POLL_INTERVAL_SECONDS = 2
EMPTY_POLLS_LOG_EVERY = 45

logger = logging.getLogger(__name__)


def setup_logging() -> None:
    logging.basicConfig(
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        level=logging.INFO,
    )


def get_bot() -> Bot:
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("Переменная BOT_TOKEN не задана")
    return Bot(token=token)


async def send_text(bot: Bot, chat_id: int, text: str) -> None:
    await bot.send_message(chat_id=chat_id, text=text)


async def send_document(bot: Bot, chat_id: int, path: Path, filename: str, caption: str) -> None:
    with open(path, "rb") as f:
        await bot.send_document(
            chat_id=chat_id,
            document=f,
            filename=filename,
            caption=caption,
        )


def _formatter_process_target(guide_code: str, input_path: str, output_path: str, queue: Queue) -> None:
    try:
        result = services.format_document_by_guide(
            guide_code=guide_code,
            input_path=input_path,
            output_path=output_path,
        )
        queue.put({
            "ok": True,
            "result": result,
        })
    except Exception:
        queue.put({
            "ok": False,
            "error": traceback.format_exc(),
        })


def run_format_with_timeout(guide_code: str, input_path: str, output_path: str, timeout_seconds: int) -> str:
    queue: Queue = Queue()
    process = Process(
        target=_formatter_process_target,
        args=(guide_code, input_path, output_path, queue),
        daemon=True,
    )

    process.start()
    process.join(timeout_seconds)

    if process.is_alive():
        process.terminate()
        process.join(5)
        raise TimeoutError(f"Formatting timeout after {timeout_seconds} seconds")

    if queue.empty():
        raise RuntimeError("Formatter process finished without returning a result")

    payload = queue.get()

    if payload.get("ok"):
        return payload.get("result") or output_path

    raise RuntimeError(payload.get("error") or "Unknown formatter error")


def build_worker_output_path(request_id: int, original_filename: str) -> Path:
    original_name = Path(original_filename)
    safe_stem = original_name.stem or f"request_{request_id}"
    safe_name = f"{safe_stem}_safe.docx"
    return services.TEMP_DIR / f"{request_id}_{safe_name}"


def fail_request(
    request_id: int,
    user_id: int,
    error_text: str,
    bot: Bot | None = None,
    telegram_id: int | None = None,
) -> None:
    db = SessionLocal()
    try:
        request = (
            db.query(FormattingRequest)
            .filter(FormattingRequest.id == request_id)
            .first()
        )
        if request:
            request.status = "failed"
            request.error_message = (error_text or "")[:1000]
            request.completed_at = datetime.utcnow()
            db.commit()
    finally:
        db.close()

    services.refund_one_credit_in_new_session(user_id, str(request_id))

    if bot is not None and telegram_id is not None:
        try:
            asyncio.run(send_text(
                bot,
                telegram_id,
                (
                    "Не удалось обработать документ из-за технической ошибки.\n"
                    "Кредит возвращён.\n"
                    "Пожалуйста, перешлите этот файл разработчику для багфикса."
                ),
            ))
        except Exception:
            logger.exception("failed_to_notify_user_about_failure request_id=%s", request_id)


def reclaim_stale_processing_requests(bot: Bot) -> int:
    now = datetime.utcnow()
    stale_before = now - timedelta(seconds=STALE_PROCESSING_SECONDS)

    db = SessionLocal()
    reclaimed = 0

    try:
        processing_requests = (
            db.query(FormattingRequest)
            .filter(FormattingRequest.status == "processing")
            .all()
        )

        for request in processing_requests:
            # В MVP используем completed_at как "время входа в processing".
            # Если completed_at пустой (legacy-запись), используем created_at.
            processing_started_at = request.completed_at or request.created_at

            if not processing_started_at:
                continue

            if processing_started_at >= stale_before:
                continue

            user = db.query(User).filter(User.id == request.user_id).first()

            request.status = "failed"
            request.error_message = "stuck in processing timeout"
            request.completed_at = datetime.utcnow()
            db.commit()

            services.refund_one_credit_in_new_session(request.user_id, str(request.id))

            if user:
                try:
                    asyncio.run(send_text(
                        bot,
                        user.telegram_id,
                        (
                            "Не удалось обработать документ из-за технической ошибки.\n"
                            "Кредит возвращён.\n"
                            "Пожалуйста, перешлите этот файл разработчику для багфикса."
                        ),
                    ))
                except Exception:
                    logger.exception("failed_to_notify_stale_request_user request_id=%s", request.id)

            logger.warning(
                "stale_request_failed request_id=%s user_id=%s started_at=%s",
                request.id,
                request.user_id,
                processing_started_at,
            )
            reclaimed += 1

    finally:
        db.close()

    return reclaimed


def pick_next_queued_request_id() -> int | None:
    db: Session = SessionLocal()
    try:
        request = (
            db.query(FormattingRequest)
            .filter(FormattingRequest.status == "queued")
            .order_by(FormattingRequest.id.asc())
            .with_for_update(skip_locked=True)
            .first()
        )

        if not request:
            return None

        request.status = "processing"
        request.error_message = None

        # В MVP временно используем completed_at как started_at для processing.
        request.completed_at = datetime.utcnow()

        db.commit()
        db.refresh(request)

        logger.info(
            "request_picked request_id=%s user_id=%s document_id=%s",
            request.id,
            request.user_id,
            request.document_id,
        )
        return request.id

    except Exception:
        db.rollback()
        logger.exception("pick_next_queued_request_failed")
        return None
    finally:
        db.close()


def process_one_request(request_id: int, bot: Bot) -> bool:
    db = SessionLocal()
    output_path: Path | None = None

    try:
        request = (
            db.query(FormattingRequest)
            .filter(FormattingRequest.id == request_id)
            .first()
        )
        if not request:
            logger.warning("request_not_found request_id=%s", request_id)
            return False

        user = db.query(User).filter(User.id == request.user_id).first()
        document = db.query(Document).filter(Document.id == request.document_id).first()

        if not user:
            fail_request(
                request_id=request_id,
                user_id=request.user_id,
                error_text="User not found",
                bot=None,
                telegram_id=None,
            )
            return False

        if not document:
            fail_request(
                request_id=request_id,
                user_id=request.user_id,
                error_text="Document not found",
                bot=bot,
                telegram_id=user.telegram_id,
            )
            return False

        input_path = Path(document.storage_path)
        if not input_path.exists():
            fail_request(
                request_id=request_id,
                user_id=request.user_id,
                error_text=f"Input file not found: {input_path}",
                bot=bot,
                telegram_id=user.telegram_id,
            )
            return False

        guide_code = services.get_user_selected_guide_code(user)
        output_path = build_worker_output_path(request.id, document.original_filename)

        logger.info(
            "formatting_start request_id=%s user_id=%s input_path=%s output_path=%s guide=%s",
            request.id,
            user.id,
            input_path,
            output_path,
            guide_code,
        )

        run_format_with_timeout(
            guide_code=guide_code,
            input_path=str(input_path),
            output_path=str(output_path),
            timeout_seconds=FORMAT_TIMEOUT_SECONDS,
        )

        request.status = "done"
        request.result_file_path = str(output_path)
        request.error_message = None
        request.completed_at = datetime.utcnow()
        db.commit()

        services.grant_referral_upload_bonus_if_needed(db, user.id)

        logger.info(
            "formatting_done request_id=%s user_id=%s output_path=%s",
            request.id,
            user.id,
            output_path,
        )

        try:
            asyncio.run(send_document(
                bot,
                user.telegram_id,
                output_path,
                output_path.name,
                "Документ оформлен.",
            ))
            logger.info(
                "document_sent request_id=%s user_id=%s filename=%s",
                request.id,
                user.id,
                output_path.name,
            )
        except Exception:
            logger.exception(
                "document_send_failed request_id=%s user_id=%s output_path=%s",
                request.id,
                user.id,
                output_path,
            )

        return True

    except TimeoutError as e:
        logger.warning(
            "formatting_timeout request_id=%s timeout=%s",
            request_id,
            FORMAT_TIMEOUT_SECONDS,
        )

        try:
            request = (
                db.query(FormattingRequest)
                .filter(FormattingRequest.id == request_id)
                .first()
            )
            user_id = request.user_id if request else None
            telegram_id = None

            if request:
                user = db.query(User).filter(User.id == request.user_id).first()
                telegram_id = user.telegram_id if user else None

            if user_id is not None:
                fail_request(
                    request_id=request_id,
                    user_id=user_id,
                    error_text=str(e),
                    bot=bot,
                    telegram_id=telegram_id,
                )
        except Exception:
            logger.exception("timeout_fail_request_failed request_id=%s", request_id)

        if output_path and output_path.exists():
            services.cleanup_temp_files(output_path)

        return False

    except Exception as e:
        logger.exception("process_request_failed request_id=%s", request_id)

        try:
            request = (
                db.query(FormattingRequest)
                .filter(FormattingRequest.id == request_id)
                .first()
            )
            user_id = request.user_id if request else None
            telegram_id = None

            if request:
                user = db.query(User).filter(User.id == request.user_id).first()
                telegram_id = user.telegram_id if user else None

            if user_id is not None:
                fail_request(
                    request_id=request_id,
                    user_id=user_id,
                    error_text=str(e),
                    bot=bot,
                    telegram_id=telegram_id,
                )
        except Exception:
            logger.exception("generic_fail_request_failed request_id=%s", request_id)

        if output_path and output_path.exists():
            services.cleanup_temp_files(output_path)

        return False

    finally:
        db.close()


def main() -> None:
    load_dotenv()
    setup_logging()

    Base.metadata.create_all(bind=engine)
    bot = get_bot()

    logger.info(
        "worker_started timeout=%s stale_timeout=%s poll_interval=%s",
        FORMAT_TIMEOUT_SECONDS,
        STALE_PROCESSING_SECONDS,
        POLL_INTERVAL_SECONDS,
    )

    empty_polls = 0

    while True:
        try:
            reclaimed = reclaim_stale_processing_requests(bot)
            if reclaimed:
                logger.warning("stale_requests_reclaimed=%s", reclaimed)

            request_id = pick_next_queued_request_id()

            if request_id is None:
                empty_polls += 1
                if empty_polls >= EMPTY_POLLS_LOG_EVERY:
                    logger.info("queue_idle empty_polls=%s", empty_polls)
                    empty_polls = 0

                time.sleep(POLL_INTERVAL_SECONDS)
                continue

            empty_polls = 0
            process_one_request(request_id, bot)

        except Exception:
            logger.exception("worker_loop_crashed")
            time.sleep(POLL_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
