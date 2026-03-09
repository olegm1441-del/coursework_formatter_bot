import logging
import os
import threading

import uvicorn
from dotenv import load_dotenv
from telegram.ext import ApplicationBuilder

from db import Base, engine
import models  # noqa: F401
from handlers import register_handlers
from payments_api import app as payments_app


def setup_logging() -> None:
    logging.basicConfig(
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        level=logging.INFO,
    )


def start_api() -> None:
    uvicorn.run(payments_app, host="0.0.0.0", port=8000)


def main() -> None:
    load_dotenv()
    setup_logging()

    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("Переменная BOT_TOKEN не задана")

    Base.metadata.create_all(bind=engine)

    api_thread = threading.Thread(target=start_api, daemon=True)
    api_thread.start()

    telegram_app = ApplicationBuilder().token(token).build()
    register_handlers(telegram_app)

    print("Бот запущен")
    telegram_app.run_polling()


if __name__ == "__main__":
    main()
