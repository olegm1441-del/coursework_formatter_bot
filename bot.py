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
from telegram.error import Conflict
import logging

logger = logging.getLogger(__name__)

async def on_error(update, context):
    err = context.error
    if isinstance(err, Conflict):
        logger.error("Polling conflict (another instance is running). Waiting...")
        # Просто ждём — Updater сам продолжит/переподключится
        return
    logger.exception("Unhandled error in telegram app", exc_info=err)

telegram_app.add_error_handler(on_error)


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
import time
from telegram.error import Conflict

def run_polling_forever(app):
    while True:
        try:
            app.run_polling(drop_pending_updates=True)
        except Conflict:
            # другой инстанс еще жив (обычно при деплое). Подождем и попробуем снова.
            time.sleep(10)

run_polling_forever(telegram_app)


if __name__ == "__main__":
    main()
