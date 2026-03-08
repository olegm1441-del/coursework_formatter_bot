import logging
import os

from dotenv import load_dotenv
from telegram.ext import ApplicationBuilder

from db import Base, engine
import models  # noqa: F401
from handlers import register_handlers


def setup_logging() -> None:
    logging.basicConfig(
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        level=logging.INFO,
    )


def main() -> None:
    load_dotenv()
    setup_logging()

    token = os.getenv("BOT_TOKEN")
    if not token:
        raise RuntimeError("Переменная BOT_TOKEN не задана")

    Base.metadata.create_all(bind=engine)

    app = ApplicationBuilder().token(token).build()

    register_handlers(app)

    print("Бот запущен")
    app.run_polling()


if __name__ == "__main__":
    main()
