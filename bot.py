import os
import uuid
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, ContextTypes, filters

from formatter_service import format_docx

load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")

TEMP_DIR = "bot_storage"
os.makedirs(TEMP_DIR, exist_ok=True)

async def handle_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):

    file = update.message.document

    if not file.file_name.endswith(".docx"):
        await update.message.reply_text("Отправьте файл .docx")
        return

    await update.message.reply_text("Файл получен. Форматирую...")

    file_id = str(uuid.uuid4())

    input_path = f"{TEMP_DIR}/{file_id}_input.docx"
    output_path = f"{TEMP_DIR}/{file_id}_output.docx"

    telegram_file = await file.get_file()
    await telegram_file.download_to_drive(input_path)

    try:
        format_docx(input_path, output_path)

        await update.message.reply_document(
            document=open(output_path, "rb"),
            filename="formatted.docx"
        )

    except Exception as e:
        await update.message.reply_text(f"Ошибка обработки: {e}")

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)
        if os.path.exists(output_path):
            os.remove(output_path)


def main():

    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(MessageHandler(filters.Document.ALL, handle_doc))

    print("Бот запущен")

    app.run_polling()


if __name__ == "__main__":
    main()
