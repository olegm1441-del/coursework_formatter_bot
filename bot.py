import logging
import os
import threading
import time

import uvicorn
from dotenv import load_dotenv
from telegram.error import Conflict
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


async def on_error(update, context):
    logging.getLogger(__name__).exception(
        "Unhandled telegram error",
        exc_info=context.error,
    )


def run_polling_forever(app):
    while True:
        try:
            app.run_polling(drop_pending_updates=True)
            return
        except Conflict:
            logging.getLogger(__name__).warning(
                "Telegram polling conflict: another instance is still shutting down. Retry in 10s..."
            )
            time.sleep(10)


def start_api() -> None:
    port = int(os.getenv("PORT", "8000"))
    uvicorn.run(payments_app, host="0.0.0.0", port=port)


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
    telegram_app.add_error_handler(on_error)

    print("Бот запущен")
    run_polling_forever(telegram_app)


if __name__ == "__main__":
    main()
