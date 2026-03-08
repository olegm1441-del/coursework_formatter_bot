import os
import uuid
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)
from sqlalchemy.exc import IntegrityError

from db import Base, SessionLocal, engine
from formatter_service import format_docx
from models import CreditLedger, Document, FormattingRequest, Referral, User
import models  # noqa: F401


load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
TEMP_DIR = Path("bot_storage")
TEMP_DIR.mkdir(exist_ok=True)

MENU_KEYBOARD = ReplyKeyboardMarkup(
    [
        ["Контакт", "Какие методички поддерживаются?"],
        ["Моя ссылка"],
    ],
    resize_keyboard=True,
)

METHOD_GUIDE_BASENAME = "Методические рекомендации по подготовке и написанию курсовой работы 2025 КФУ"
ASSETS_DIR = Path("assets")


def find_method_file() -> Path | None:
    for ext in [".docx", ".pdf"]:
        candidate = ASSETS_DIR / f"{METHOD_GUIDE_BASENAME}{ext}"
        if candidate.exists():
            return candidate
    return None


def generate_referral_code() -> str:
    return uuid.uuid4().hex[:10]


def get_user_by_telegram_id(db, telegram_id: int) -> User | None:
    return db.query(User).filter(User.telegram_id == telegram_id).first()


def get_user_by_referral_code(db, referral_code: str) -> User | None:
    return db.query(User).filter(User.referral_code == referral_code).first()


def get_or_create_user(
    db,
    telegram_id: int,
    username: str | None,
    first_name: str | None,
    last_name: str | None,
    referral_code_from_start: str | None = None,
) -> tuple[User, bool]:
    user = get_user_by_telegram_id(db, telegram_id)
    if user:
        return user, False

    referred_by_user_id = None
    inviter_referral_code = None

    if referral_code_from_start:
        inviter = get_user_by_referral_code(db, referral_code_from_start)
        if inviter and inviter.telegram_id != telegram_id:
            referred_by_user_id = inviter.id
            inviter_referral_code = inviter.referral_code

    while True:
        try:
            user = User(
                telegram_id=telegram_id,
                username=username,
                first_name=first_name,
                last_name=last_name,
                referral_code=generate_referral_code(),
                referred_by_user_id=referred_by_user_id,
            )
            db.add(user)
            db.flush()

            welcome_credit = CreditLedger(
                user_id=user.id,
                operation_type="welcome_bonus",
                amount=1,
                source_type="system",
                source_id=str(user.id),
                idempotency_key=f"welcome_bonus_{user.id}",
            )
            db.add(welcome_credit)

            if referred_by_user_id and inviter_referral_code:
                referral = Referral(
                    inviter_user_id=referred_by_user_id,
                    invited_user_id=user.id,
                    referral_code=inviter_referral_code,
                )
                db.add(referral)

            db.commit()
            db.refresh(user)
            return user, True

        except IntegrityError:
            db.rollback()


def get_referral_link(bot_username: str, referral_code: str) -> str:
    return f"https://t.me/{bot_username}?start=ref_{referral_code}"


def get_user_credit_balance(db, user_id: int) -> int:
    rows = db.query(CreditLedger).filter(CreditLedger.user_id == user_id).all()
    return sum(row.amount for row in rows)


def debit_one_credit(db, user_id: int, source_id: str) -> bool:
    balance = get_user_credit_balance(db, user_id)
    if balance <= 0:
        return False

    existing = (
        db.query(CreditLedger)
        .filter(CreditLedger.idempotency_key == f"format_debit_{source_id}")
        .first()
    )
    if existing:
        return True

    entry = CreditLedger(
        user_id=user_id,
        operation_type="format_debit",
        amount=-1,
        source_type="formatting_request",
        source_id=source_id,
        idempotency_key=f"format_debit_{source_id}",
    )
    db.add(entry)
    db.commit()
    return True


def refund_one_credit_if_needed(db, user_id: int, source_id: str) -> None:
    existing_refund = (
        db.query(CreditLedger)
        .filter(CreditLedger.idempotency_key == f"format_refund_{source_id}")
        .first()
    )
    if existing_refund:
        return

    debit_exists = (
        db.query(CreditLedger)
        .filter(CreditLedger.idempotency_key == f"format_debit_{source_id}")
        .first()
    )
    if not debit_exists:
        return

    refund = CreditLedger(
        user_id=user_id,
        operation_type="format_refund",
        amount=1,
        source_type="formatting_request",
        source_id=source_id,
        idempotency_key=f"format_refund_{source_id}",
    )
    db.add(refund)
    db.commit()


def grant_referral_upload_bonus_if_needed(db, invited_user_id: int) -> None:
    referral = (
        db.query(Referral)
        .filter(Referral.invited_user_id == invited_user_id)
        .filter(Referral.qualified_upload_at.is_(None))
        .first()
    )

    if not referral:
        return

    existing_bonus = (
        db.query(CreditLedger)
        .filter(CreditLedger.idempotency_key == f"referral_upload_bonus_{referral.id}")
        .first()
    )
    if existing_bonus:
        return

    bonus = CreditLedger(
        user_id=referral.inviter_user_id,
        operation_type="referral_upload_bonus",
        amount=1,
        source_type="referral",
        source_id=str(referral.id),
        idempotency_key=f"referral_upload_bonus_{referral.id}",
    )
    db.add(bonus)
    referral.qualified_upload_at = datetime.utcnow()
    db.commit()


async def send_referral_message(update: Update, context: ContextTypes.DEFAULT_TYPE, user: User) -> None:
    if not update.message:
        return

    bot_username = (await context.bot.get_me()).username
    referral_link = get_referral_link(bot_username, user.referral_code)

    await update.message.reply_text(
        "Твоя реферальная ссылка:\n"
        f"{referral_link}\n\n"
        "Если новый пользователь впервые успешно оформит документ по этой ссылке, ты получишь +1 оформление.\n"
        "Если он впервые оплатит, ты тоже получишь +1 оформление.",
        reply_markup=MENU_KEYBOARD,
    )


async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_user:
        return

    telegram_user = update.effective_user

    referral_code = None
    if context.args:
        raw_arg = context.args[0]
        if raw_arg.startswith("ref_"):
            referral_code = raw_arg.replace("ref_", "", 1)

    db = SessionLocal()
    try:
        user, is_new = get_or_create_user(
            db=db,
            telegram_id=telegram_user.id,
            username=telegram_user.username,
            first_name=telegram_user.first_name,
            last_name=telegram_user.last_name,
            referral_code_from_start=referral_code,
        )

        balance = get_user_credit_balance(db, user.id)

        if is_new:
            text = (
                "Здравствуйте! Это бот для форматирования курсовых работ КФУ.\n\n"
                "У вас уже есть 1 бесплатное оформление.\n"
                f"Сейчас доступно оформлений: {balance}\n\n"
                "Можно сразу отправить .docx-файл на обработку."
            )
        else:
            text = (
                "С возвращением!\n"
                f"Сейчас доступно оформлений: {balance}\n\n"
                "Можно сразу отправить .docx-файл на обработку."
            )

        await update.message.reply_text(text, reply_markup=MENU_KEYBOARD)

    finally:
        db.close()


async def referral_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.effective_user:
        return

    db = SessionLocal()
    try:
        user, _ = get_or_create_user(
            db=db,
            telegram_id=update.effective_user.id,
            username=update.effective_user.username,
            first_name=update.effective_user.first_name,
            last_name=update.effective_user.last_name,
            referral_code_from_start=None,
        )
        await send_referral_message(update, context, user)
    finally:
        db.close()


async def contact_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    text = (
        "Контакт разработчика:\n"
        "@aelart\n\n"
        "Бот разработан @aelart."
    )
    await update.message.reply_text(text, reply_markup=MENU_KEYBOARD)


async def method_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    text = (
        "Сейчас поддерживается методичка:\n"
        "«Методические рекомендации по подготовке и написанию курсовой работы 2025 КФУ»"
    )
    await update.message.reply_text(text, reply_markup=MENU_KEYBOARD)

    method_file = find_method_file()
    if method_file is not None:
        await update.message.reply_document(document=method_file.open("rb"))
    else:
        await update.message.reply_text(
            "Файл методички пока не найден в папке assets.",
            reply_markup=MENU_KEYBOARD,
        )


async def text_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.message.text:
        return

    text = update.message.text.strip()

    if text == "Контакт":
        await contact_handler(update, context)
        return

    if text == "Какие методички поддерживаются?":
        await method_handler(update, context)
        return

    if text == "Моя ссылка":
        await referral_handler(update, context)
        return

    await update.message.reply_text(
        "Можно отправить .docx-файл на обработку или воспользоваться кнопками ниже.",
        reply_markup=MENU_KEYBOARD,
    )


async def handle_doc(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.message.document or not update.effective_user:
        return

    file = update.message.document

    if not file.file_name or not file.file_name.lower().endswith(".docx"):
        await update.message.reply_text(
            "Пожалуйста, отправьте файл в формате .docx",
            reply_markup=MENU_KEYBOARD,
        )
        return

    db = SessionLocal()
    input_path = None
    output_path = None
    formatting_request = None
    user = None

    try:
        user, _ = get_or_create_user(
            db=db,
            telegram_id=update.effective_user.id,
            username=update.effective_user.username,
            first_name=update.effective_user.first_name,
            last_name=update.effective_user.last_name,
            referral_code_from_start=None,
        )

        balance = get_user_credit_balance(db, user.id)
        if balance <= 0:
            await update.message.reply_text(
                "У вас нет доступных оформлений.\n"
                "Используйте свою реферальную ссылку или купите пакет позже.",
                reply_markup=MENU_KEYBOARD,
            )
            return

        await update.message.reply_text("Файл получен. Форматирую...", reply_markup=MENU_KEYBOARD)

        job_id = str(uuid.uuid4())
        original_name = Path(file.file_name)
        safe_name = f"{original_name.stem}_safe.docx"

        input_path = TEMP_DIR / f"{job_id}_input.docx"
        output_path = TEMP_DIR / safe_name

        telegram_file = await file.get_file()
        await telegram_file.download_to_drive(str(input_path))

        document = Document(
            user_id=user.id,
            original_filename=file.file_name,
            storage_path=str(input_path),
        )
        db.add(document)
        db.flush()

        formatting_request = FormattingRequest(
            user_id=user.id,
            document_id=document.id,
            service_type="format",
            university_code="kfu",
            document_type="coursework",
            guideline_version="2025",
            status="processing",
        )
        db.add(formatting_request)
        db.commit()
        db.refresh(formatting_request)

        if not debit_one_credit(db, user.id, str(formatting_request.id)):
            formatting_request.status = "failed"
            formatting_request.error_message = "Недостаточно кредитов"
            db.commit()
            await update.message.reply_text(
                "Недостаточно доступных оформлений.",
                reply_markup=MENU_KEYBOARD,
            )
            return

        format_docx(str(input_path), str(output_path))

        formatting_request.status = "done"
        formatting_request.result_file_path = str(output_path)
        formatting_request.completed_at = datetime.utcnow()
        db.commit()

        with output_path.open("rb") as f:
            await update.message.reply_document(
                document=f,
                filename=safe_name,
                reply_markup=MENU_KEYBOARD,
            )

        grant_referral_upload_bonus_if_needed(db, user.id)

        user = get_user_by_telegram_id(db, update.effective_user.id)
        await send_referral_message(update, context, user)

    except Exception as e:
        db.rollback()

        if formatting_request is not None:
            formatting_request = db.query(FormattingRequest).filter(
                FormattingRequest.id == formatting_request.id
            ).first()

            if formatting_request is not None:
                formatting_request.status = "failed"
                formatting_request.error_message = str(e)
                db.commit()

                refund_one_credit_if_needed(db, user.id, str(formatting_request.id))

        await update.message.reply_text(
            f"Ошибка обработки: {e}",
            reply_markup=MENU_KEYBOARD,
        )

    finally:
        db.close()
        if input_path and input_path.exists():
            input_path.unlink()
        if output_path and output_path.exists():
            output_path.unlink()


def main() -> None:
    if not TOKEN:
        raise RuntimeError("Переменная BOT_TOKEN не задана")

    Base.metadata.create_all(bind=engine)

    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("referral", referral_handler))
    app.add_handler(CommandHandler("contact", contact_handler))
    app.add_handler(CommandHandler("method", method_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_doc))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_menu_handler))

    print("Бот запущен")
    app.run_polling()


if __name__ == "__main__":
    main()
