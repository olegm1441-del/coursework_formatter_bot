import asyncio

import logging
import os
import time
import traceback
from datetime import UTC, datetime, timedelta
from multiprocessing import Process, Queue
from pathlib import Path

import httpx
import telegram

from dotenv import load_dotenv
from sqlalchemy.orm import Session
from telegram import Bot
from telegram.request import HTTPXRequest

from db import Base, SessionLocal, engine
import models  # noqa: F401
from models import Document, FormattingRequest, User
from keyboards import (
    get_action_inline_keyboard,
    get_check_result_inline_keyboard,
    get_referral_progress_inline_keyboard,
)
import services


FORMAT_TIMEOUT_SECONDS = 180
STALE_PROCESSING_SECONDS = 5 * 60
POLL_INTERVAL_SECONDS = 2
EMPTY_POLLS_LOG_EVERY = 45
SEND_DOCUMENT_WRITE_TIMEOUT_SECONDS = 45
SEND_DOCUMENT_READ_TIMEOUT_SECONDS = 120
SEND_DOCUMENT_CONNECT_TIMEOUT_SECONDS = 30
SEND_DOCUMENT_POOL_TIMEOUT_SECONDS = 30

logger = logging.getLogger(__name__)


def utcnow_naive() -> datetime:
    """UTC timestamp without tzinfo for DB fields declared as naive DateTime."""
    return datetime.now(UTC).replace(tzinfo=None)


def setup_logging() -> None:
    logging.basicConfig(
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        level=logging.INFO,
    )


def get_bot_token() -> str:
    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("Переменная BOT_TOKEN не задана")
    return token


async def send_text(bot_token: str, chat_id: int, text: str, reply_markup=None) -> None:
    bot = Bot(token=bot_token)
    await bot.send_message(chat_id=chat_id, text=text, reply_markup=reply_markup)


async def send_document(
    bot_token: str,
    chat_id: int,
    path: Path,
    filename: str,
    caption: str,
    warnings: list[str] | None = None,
) -> None:
    """
    Upload the formatted DOCX and, if there are formatter warnings, send them
    in a separate text message immediately after.

    Telegram limits document captions to 1024 chars; regular messages allow up
    to 4096 chars.  Warnings are therefore sent as a follow-up message to avoid
    BadRequest: Message caption is too long.
    """
    request = HTTPXRequest(
        write_timeout=SEND_DOCUMENT_WRITE_TIMEOUT_SECONDS,
        connect_timeout=SEND_DOCUMENT_CONNECT_TIMEOUT_SECONDS,
        read_timeout=SEND_DOCUMENT_READ_TIMEOUT_SECONDS,
        pool_timeout=SEND_DOCUMENT_POOL_TIMEOUT_SECONDS,
    )
    bot = Bot(token=bot_token, request=request)
    logger.info(
        "send_document_config ptb_version=%s write_timeout=%s read_timeout=%s connect_timeout=%s pool_timeout=%s",
        telegram.__version__,
        SEND_DOCUMENT_WRITE_TIMEOUT_SECONDS,
        SEND_DOCUMENT_READ_TIMEOUT_SECONDS,
        SEND_DOCUMENT_CONNECT_TIMEOUT_SECONDS,
        SEND_DOCUMENT_POOL_TIMEOUT_SECONDS,
    )
    with open(path, "rb") as f:
        await bot.send_document(
            chat_id=chat_id,
            document=f,
            filename=filename,
            caption=caption,
            write_timeout=SEND_DOCUMENT_WRITE_TIMEOUT_SECONDS,
            read_timeout=SEND_DOCUMENT_READ_TIMEOUT_SECONDS,
            connect_timeout=SEND_DOCUMENT_CONNECT_TIMEOUT_SECONDS,
            pool_timeout=SEND_DOCUMENT_POOL_TIMEOUT_SECONDS,
        )

    # Send warnings as a separate message so caption length is never an issue.
    if warnings:
        warning_lines = "\n".join(f"⚠️ {w}" for w in warnings)
        await bot.send_message(chat_id=chat_id, text=warning_lines)

async def send_referral_bonus_notification(
    bot_token: str,
    inviter_telegram_id: int,
    text: str,
) -> None:
    bot = Bot(token=bot_token)
    await bot.send_message(chat_id=inviter_telegram_id, text=text)
    
def _formatter_process_target(
    guide_code: str,
    input_path: str,
    output_path: str,
    queue: Queue,
) -> None:
    try:
        path, warnings = services.format_document_by_guide(
            guide_code=guide_code,
            input_path=input_path,
            output_path=output_path,
        )
        queue.put({"ok": True, "result": path, "warnings": warnings})
    except Exception:
        queue.put({"ok": False, "error": traceback.format_exc()})


def run_format_with_timeout(
    guide_code: str,
    input_path: str,
    output_path: str,
    timeout_seconds: int,
) -> tuple[str, list[str]]:
    """
    Run the formatter in a subprocess with a hard timeout.

    Returns (output_path, warnings) on success.
    Raises TimeoutError or RuntimeError on failure.
    """
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
        path = payload.get("result") or output_path
        warnings = payload.get("warnings") or []
        return path, warnings

    raise RuntimeError(payload.get("error") or "Unknown formatter error")


def build_worker_output_path(request_id: int, original_filename: str) -> Path:
    original_name = Path(original_filename)
    safe_stem = original_name.stem or f"request_{request_id}"
    safe_name = f"{safe_stem}.docx"
    return services.TEMP_DIR / f"{request_id}_{safe_name}"


def notify_failure(bot_token: str, telegram_id: int) -> None:
    try:
        asyncio.run(
            send_text(
                bot_token,
                telegram_id,
                (
                    "Не удалось обработать документ.\n"
                    "Кредит возвращён.\n\n"
                    "Проверьте, что файл соответствует базовым требованиям:\n"
                    "• файл в формате .docx\n"
                    "• это сама курсовая работа, а не методичка или шаблон\n"
                    "• в документе есть раздел «Введение» отдельным абзацем\n"
                    "• есть структура: введение → главы → заключение → источники\n"
                    "• заголовки разделов написаны обычным текстом\n\n"
                    "Если файл соответствует этим требованиям, "
                    "перешлите его разработчику для проверки: @aelart"
                ),
            )
        )
    except Exception:
        logger.exception("failed_to_notify_user_about_failure telegram_id=%s", telegram_id)


def fail_request(
    request_id: int,
    user_id: int,
    error_text: str,
    bot_token: str | None = None,
    telegram_id: int | None = None,
    silent_fail: bool = False,
) -> None:
    db = SessionLocal()
    effective_silent_fail = bool(silent_fail)

    try:
        request = (
            db.query(FormattingRequest)
            .filter(FormattingRequest.id == request_id)
            .first()
        )
        if request:
            request.status = "failed"
            request.error_message = (error_text or "")[:1000]
            request.completed_at = utcnow_naive()
            effective_silent_fail = effective_silent_fail or bool(request.silent_fail)
            db.commit()
    finally:
        db.close()

    services.refund_one_credit_in_new_session(user_id, str(request_id))

    if bot_token and telegram_id and not effective_silent_fail:
        notify_failure(bot_token, telegram_id)


def reclaim_stale_processing_requests(bot_token: str) -> int:
    now = utcnow_naive()
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
            processing_started_at = request.completed_at or request.created_at

            if not processing_started_at:
                continue

            if processing_started_at >= stale_before:
                continue

            user = db.query(User).filter(User.id == request.user_id).first()

            request.status = "failed"
            request.error_message = "stuck in processing timeout"
            request.completed_at = utcnow_naive()
            db.commit()

            services.refund_one_credit_in_new_session(request.user_id, str(request.id))

            if user and not request.silent_fail:
                notify_failure(bot_token, user.telegram_id)

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
        request.completed_at = utcnow_naive()
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



def prepare_worker_input_path(
    storage_path: str,
    request_id: int,
    original_filename: str,
    bot_token: str,
) -> Path:
    suffix = Path(original_filename).suffix or ".docx"

    if storage_path.startswith(("http://", "https://")):
        source_url = storage_path
    elif storage_path.startswith("documents/") or storage_path.startswith("photos/"):
        source_url = f"https://api.telegram.org/file/bot{bot_token}/{storage_path.lstrip('/')}"
    else:
        return Path(storage_path)

    local_input_path = services.TEMP_DIR / f"{request_id}_downloaded{suffix}"
    response = httpx.get(source_url, timeout=60.0)
    response.raise_for_status()
    local_input_path.write_bytes(response.content)
    return local_input_path


def process_one_request(request_id: int, bot_token: str) -> bool:
    db = SessionLocal()
    output_path: Path | None = None
    input_path: Path | None = None

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
                silent_fail=bool(request.silent_fail),
            )
            return False

        if not document:
            fail_request(
                request_id=request_id,
                user_id=request.user_id,
                error_text="Document not found",
                bot_token=bot_token,
                telegram_id=user.telegram_id,
                silent_fail=bool(request.silent_fail),
            )
            return False

        input_path = prepare_worker_input_path(
            storage_path=document.storage_path,
            request_id=request.id,
            original_filename=document.original_filename,
            bot_token=bot_token,
        )
        if not input_path.exists():
            fail_request(
                request_id=request_id,
                user_id=request.user_id,
                error_text=f"Input file not found: {input_path}",
                bot_token=bot_token,
                telegram_id=user.telegram_id,
                silent_fail=bool(request.silent_fail),
            )
            return False

        if request.service_type == "check":
            logger.info(
                "check_stub_start request_id=%s user_id=%s input_path=%s",
                request.id,
                user.id,
                input_path,
            )
            logger.info(
                "referral_check_upload_seen invited_user_id=%s invited_telegram_id=%s",
                user.id,
                user.telegram_id,
            )

            problems: list[str] = []
            request.status = "done"
            request.result_file_path = None
            request.error_message = None
            request.completed_at = utcnow_naive()
            db.commit()

            services.track_event(
                db,
                event_name="check_completed",
                user_id=user.id,
                source="worker",
                payload_json=f'{{"request_id": {request.id}, "document_id": {document.id}}}',
            )

            inviter_user_id = services.grant_referral_upload_bonus_if_needed(db, user.id)
            if inviter_user_id:
                inviter = (
                    db.query(User)
                    .filter(User.id == inviter_user_id)
                    .first()
                )
                if inviter:
                    bonus_text = services.build_referral_upload_bonus_awarded_text()
                    try:
                        asyncio.run(
                            send_text(
                                bot_token=bot_token,
                                chat_id=inviter.telegram_id,
                                text=bonus_text,
                                reply_markup=get_referral_progress_inline_keyboard(),
                            )
                        )
                    except Exception:
                        logger.exception(
                            "referral_upload_bonus_notification_failed inviter_user_id=%s",
                            inviter.id,
                        )

            asyncio.run(
                send_text(
                    bot_token=bot_token,
                    chat_id=user.telegram_id,
                    text=services.build_check_result_text(problems),
                    reply_markup=get_check_result_inline_keyboard(),
                )
            )

            logger.info("check_stub_done request_id=%s user_id=%s", request.id, user.id)

            if input_path and input_path.exists():
                services.cleanup_temp_files(input_path)

            return True

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

        _, formatting_warnings = run_format_with_timeout(
            guide_code=guide_code,
            input_path=str(input_path),
            output_path=str(output_path),
            timeout_seconds=FORMAT_TIMEOUT_SECONDS,
        )

        request.status = "done"
        request.result_file_path = str(output_path)
        request.error_message = None
        request.completed_at = utcnow_naive()
        db.commit()

        services.track_event(
            db,
            event_name="processing_completed",
            user_id=user.id,
            source="worker",
            payload_json=f'{{"request_id": {request.id}, "document_id": {document.id}}}',
        )

        logger.info(
            "formatting_done request_id=%s user_id=%s output_path=%s",
            request.id,
            user.id,
            output_path,
        )

        caption = services.build_format_success_caption(
            services.get_bot_username_fallback(),
            user,
        )
        # Warnings are sent as a separate text message AFTER the document
        # (Telegram document captions are limited to 1024 chars; text messages
        # allow 4096 — prevents BadRequest: Message caption is too long).

        logger.info("send_document_start request_id=%s warnings=%d", request.id, len(formatting_warnings))
        asyncio.run(
            send_document(
                bot_token,
                user.telegram_id,
                output_path,
                output_path.name,
                caption,
                warnings=formatting_warnings or None,
            )
        )
        asyncio.run(
            send_text(
                bot_token=bot_token,
                chat_id=user.telegram_id,
                text="Что дальше сделать?",
                reply_markup=get_action_inline_keyboard(),
            )
        )

        logger.info(
            "document_sent request_id=%s user_id=%s filename=%s",
            request.id,
            user.id,
            output_path.name,
        )

        if input_path and input_path.exists():
            services.cleanup_temp_files(input_path)
        if output_path and output_path.exists():
            services.cleanup_temp_files(output_path)

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

            silent_fail = False

            if request:
                user = db.query(User).filter(User.id == request.user_id).first()
                telegram_id = user.telegram_id if user else None
                silent_fail = bool(request.silent_fail)

            if user_id is not None:
                fail_request(
                    request_id=request_id,
                    user_id=user_id,
                    error_text=str(e),
                    bot_token=bot_token,
                    telegram_id=telegram_id,
                    silent_fail=silent_fail,
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

            silent_fail = False

            if request:
                user = db.query(User).filter(User.id == request.user_id).first()
                telegram_id = user.telegram_id if user else None
                silent_fail = bool(request.silent_fail)

            if user_id is not None:
                fail_request(
                    request_id=request_id,
                    user_id=user_id,
                    error_text=str(e),
                    bot_token=bot_token,
                    telegram_id=telegram_id,
                    silent_fail=silent_fail,
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
    bot_token = get_bot_token()

    logger.info(
        "worker_started timeout=%s stale_timeout=%s poll_interval=%s",
        FORMAT_TIMEOUT_SECONDS,
        STALE_PROCESSING_SECONDS,
        POLL_INTERVAL_SECONDS,
    )

    empty_polls = 0

    while True:
        try:
            reclaimed = reclaim_stale_processing_requests(bot_token)
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
            process_one_request(request_id, bot_token)

        except Exception:
            logger.exception("worker_loop_crashed")
            time.sleep(POLL_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
