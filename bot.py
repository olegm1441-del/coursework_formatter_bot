import os
import uuid
from pathlib import Path
from db import Base, engine, SessionLocal
from repositories import get_user_by_telegram_id, get_user_by_referral_code, create_user
from dotenv import load_dotenv
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from db import Base, engine
import models

from formatter_service import format_docx

load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")
TEMP_DIR = Path("bot_storage")
TEMP_DIR.mkdir(exist_ok=True)

MENU_KEYBOARD = ReplyKeyboardMarkup(
    [
        ["Контакт", "Какие методички поддерживаются?"],
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


async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    telegram_user = update.effective_user
    if telegram_user is None:
        return

    referral_code = None
    if context.args:
        raw_arg = context.args[0]
        if raw_arg.startswith("ref_"):
            referral_code = raw_arg.replace("ref_", "", 1)

    db = SessionLocal()
    try:
        user = get_user_by_telegram_id(db, telegram_user.id)

        if not user:
            referred_by_user_id = None

            if referral_code:
                inviter = get_user_by_referral_code(db, referral_code)
                if inviter and inviter.telegram_id != telegram_user.id:
                    referred_by_user_id = inviter.id

            user = create_user(
                db=db,
                telegram_id=telegram_user.id,
                username=telegram_user.username,
                first_name=telegram_user.first_name,
                last_name=telegram_user.last_name,
                referred_by_user_id=referred_by_user_id,
            )

        await update.message.reply_text(
            "Привет! Отправь .docx файл курсовой, и я оформлю его по методичке КФУ."
        )
    finally:
        db.close()


async def contact_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    text = (
        "Контакт разработчика:\n"
        "@aelart\n\n"
        "Бот разработан @aelart."
    )
    await update.message.reply_text(text, reply_markup=MENU_KEYBOARD)


async def method_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
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

    await update.message.reply_text(
        "Можно отправить .docx-файл на обработку или воспользоваться кнопками ниже.",
        reply_markup=MENU_KEYBOARD,
    )


async def handle_doc(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.message.document:
        return

    file = update.message.document

    if not file.file_name or not file.file_name.lower().endswith(".docx"):
        await update.message.reply_text("Пожалуйста, отправьте файл в формате .docx", reply_markup=MENU_KEYBOARD)
        return

    await update.message.reply_text("Файл получен. Форматирую...", reply_markup=MENU_KEYBOARD)

    job_id = str(uuid.uuid4())
    original_name = Path(file.file_name)
    safe_name = f"{original_name.stem}_safe.docx"

    input_path = TEMP_DIR / f"{job_id}_input.docx"
    output_path = TEMP_DIR / safe_name

    telegram_file = await file.get_file()
    await telegram_file.download_to_drive(str(input_path))

    try:
        format_docx(str(input_path), str(output_path))

        with output_path.open("rb") as f:
            await update.message.reply_document(
                document=f,
                filename=safe_name,
                reply_markup=MENU_KEYBOARD,
            )

    except Exception as e:
        await update.message.reply_text(
            f"Ошибка обработки: {e}",
            reply_markup=MENU_KEYBOARD,
        )

    finally:
        if input_path.exists():
            input_path.unlink()
        if output_path.exists():
            output_path.unlink()


def main() -> None:
    if not TOKEN:
        raise RuntimeError("Переменная BOT_TOKEN не задана")

    Base.metadata.create_all(bind=engine)
    
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("contact", contact_handler))
    app.add_handler(CommandHandler("method", method_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_doc))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_menu_handler))

    print("Бот запущен")
    app.run_polling()


if __name__ == "__main__":
    main()
