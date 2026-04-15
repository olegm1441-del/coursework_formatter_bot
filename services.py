import os
import uuid
import unicodedata
import logging
from datetime import datetime, timezone
from pathlib import Path

from dotenv import load_dotenv
from sqlalchemy import func
from sqlalchemy.exc import IntegrityError

from db import SessionLocal
from models import (
    User,
    Referral,
    Payment,
    CreditLedger,
    Document,
    FormattingRequest,
    AnalyticsEvent,
)
from guides.coursework_kfu_2025.formatter_service import format_docx


load_dotenv()

logger = logging.getLogger(__name__)

# =========================
# Базовые настройки проекта
# =========================

BOT_USERNAME_FALLBACK = os.getenv("BOT_USERNAME", "").strip()
TEMP_DIR = Path("bot_storage")
TEMP_DIR.mkdir(exist_ok=True)

ASSETS_DIR = Path("assets")

DEFAULT_GUIDE_CODE = "kfu_coursework_2025"
REFERRAL_UPLOAD_BONUS_SIZE = 3
REFERRAL_UPLOAD_BONUS_ENABLED_AT_DEFAULT = "2026-04-14T20:52:03Z"
REFERRAL_UPLOAD_BONUS_ENABLED_AT_RAW = os.getenv(
    "REFERRAL_UPLOAD_BONUS_ENABLED_AT",
    os.getenv("REFERRAL_PAIRS_ENABLED_AT", REFERRAL_UPLOAD_BONUS_ENABLED_AT_DEFAULT),
).strip()

TARIFFS = {
    "one": {
        "title": "1 оформление",
        "price_rub": 200,
        "credits": 1,
    },
    "three": {
        "title": "3 оформления",
        "price_rub": 500,
        "credits": 3,
    },
}

GUIDES = {
    "kfu_coursework_2025": {
        "title": "КФУ — курсовая 2025",
        "university_code": "kfu",
        "document_type": "coursework",
        "guideline_version": "2025",
        "method_basename": "КР КФУ 12.2025",
        "formatter": format_docx,
        "is_active": True,
    }
}


# =========================
# Пользователи и рефералка
# =========================

def generate_referral_code() -> str:
    return uuid.uuid4().hex[:10]


def get_user_by_telegram_id(db, telegram_id: int) -> User | None:
    return db.query(User).filter(User.telegram_id == telegram_id).first()


def get_user_by_id(db, user_id: int) -> User | None:
    return db.query(User).filter(User.id == user_id).first()


def get_user_by_referral_code(db, referral_code: str) -> User | None:
    return db.query(User).filter(User.referral_code == referral_code).first()


def get_referral_by_invited_user_id(db, invited_user_id: int) -> Referral | None:
    return (
        db.query(Referral)
        .filter(Referral.invited_user_id == invited_user_id)
        .first()
    )


def _user_model_supports_selected_guide() -> bool:
    return hasattr(User, "selected_guide_code")


def get_user_selected_guide_code(user: User) -> str:
    """
    Работает безопасно даже если колонка selected_guide_code
    ещё не добавлена в models.py / БД.
    """
    if _user_model_supports_selected_guide():
        value = getattr(user, "selected_guide_code", None)
        if value in GUIDES:
            return value
    return DEFAULT_GUIDE_CODE


def set_user_selected_guide_code(db, user: User, guide_code: str) -> str:
    """
    Если selected_guide_code уже добавлен в модель — сохраняем в БД.
    Если ещё нет — просто возвращаем guide_code без падения.
    """
    if guide_code not in GUIDES:
        raise ValueError("Неизвестная методичка")

    if _user_model_supports_selected_guide():
        setattr(user, "selected_guide_code", guide_code)
        db.commit()
        db.refresh(user)

    return guide_code


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
        if not referral_code_from_start:
            return user, False

        logger.info(
            "referral_existing_user_link_attempt telegram_id=%s db_user_id=%s code=%s",
            telegram_id,
            user.id,
            referral_code_from_start,
        )

        if user.referred_by_user_id:
            logger.info(
                "referral_existing_user_skip reason=already_referred telegram_id=%s db_user_id=%s inviter_id=%s",
                telegram_id,
                user.id,
                user.referred_by_user_id,
            )
            return user, False

        existing_referral = get_referral_by_invited_user_id(db, user.id)
        if existing_referral:
            logger.info(
                "referral_existing_user_skip reason=already_referred telegram_id=%s db_user_id=%s inviter_id=%s",
                telegram_id,
                user.id,
                existing_referral.inviter_user_id,
            )
            return user, False

        inviter = get_user_by_referral_code(db, referral_code_from_start)
        if not inviter:
            logger.info(
                "referral_existing_user_skip reason=inviter_not_found telegram_id=%s db_user_id=%s code=%s",
                telegram_id,
                user.id,
                referral_code_from_start,
            )
            return user, False

        if inviter.telegram_id == telegram_id:
            logger.info(
                "referral_existing_user_skip reason=self_referral telegram_id=%s db_user_id=%s inviter_id=%s",
                telegram_id,
                user.id,
                inviter.id,
            )
            return user, False

        user.referred_by_user_id = inviter.id
        referral = Referral(
            inviter_user_id=inviter.id,
            invited_user_id=user.id,
            referral_code=inviter.referral_code,
        )
        db.add(referral)
        db.commit()
        db.refresh(user)
        logger.info(
            "referral_existing_user_linked inviter_id=%s invited_id=%s invited_telegram_id=%s",
            inviter.id,
            user.id,
            telegram_id,
        )
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
            user_kwargs = {
                "telegram_id": telegram_id,
                "username": username,
                "first_name": first_name,
                "last_name": last_name,
                "referral_code": generate_referral_code(),
                "referred_by_user_id": referred_by_user_id,
            }

            if _user_model_supports_selected_guide():
                user_kwargs["selected_guide_code"] = DEFAULT_GUIDE_CODE

            user = User(**user_kwargs)
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
                logger.info(
                    "referral_linked inviter_id=%s invited_id=%s invited_telegram_id=%s",
                    referred_by_user_id,
                    user.id,
                    telegram_id,
                )

            db.commit()
            db.refresh(user)
            return user, True

        except IntegrityError:
            db.rollback()


def get_referral_link(bot_username: str, referral_code: str) -> str:
    bot_username = bot_username.strip().lstrip("@")
    return f"https://t.me/{bot_username}?start=ref_{referral_code}"


def build_referral_text(
    bot_username: str,
    user: User,
    balance: int | None = None,
    progress: int | None = None,
    target: int | None = None,
) -> str:
    referral_link = get_referral_link(bot_username, user.referral_code)
    balance_text = ""
    if balance is not None:
        balance_text = f"Доступно оформлений: {balance}\n\n"
    progress_text = ""
    if progress is not None and target is not None:
        progress_text = f"\n{build_referral_progress_text(progress, target)}\n"

    return (
        f"{balance_text}"
        "Твоя реферальная ссылка:\n"
        f"{referral_link}\n\n"
        "За каждых 3 новых друзей, которые впервые загрузят .docx на автопроверку по твоей ссылке, начисляется +1 оформление.\n"
        "Ещё +1 оформление, когда приглашённый друг впервые оплатит.\n"
        f"{progress_text}\n"
        "Можно отправить ссылку одногруппнику, которому тоже скоро сдавать курсовую."
    )
    
def build_referral_bonus_notification_text(balance: int, trigger: str) -> str:
    if trigger == "upload":
        reason = "трое приглашённых друзей впервые загрузили .docx на автопроверку"
    elif trigger == "payment":
        reason = "приглашённый пользователь впервые оплатил"
    else:
        reason = "сработал реферальный бонус"

    return (
        "🎉 Начислен реферальный бонус!\n\n"
        f"Причина: {reason}.\n"
        "Вы получили +1 оформление.\n"
        f"Ваш баланс: {balance} оформлений."
    )

# =========================
# Баланс и кредиты
# =========================

def get_user_credit_balance(db, user_id: int) -> int:
    total = (
        db.query(func.sum(CreditLedger.amount))
        .filter(CreditLedger.user_id == user_id)
        .scalar()
    )
    return total or 0


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
        operation_type="formatting_spent",
        amount=-1,
        source_type="formatting_request",
        source_id=source_id,
        idempotency_key=f"format_debit_{source_id}",
    )
    db.add(entry)
    db.commit()
    return True


def refund_one_credit_in_new_session(user_id: int, source_id: str) -> None:
    db = SessionLocal()
    try:
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
            operation_type="refund",
            amount=1,
            source_type="formatting_request",
            source_id=source_id,
            idempotency_key=f"format_refund_{source_id}",
        )
        db.add(refund)
        db.commit()
    finally:
        db.close()


def grant_admin_bonus(db, target_user_id: int, amount: int, admin_source_id: str) -> int:
    entry = CreditLedger(
        user_id=target_user_id,
        operation_type="admin_grant",
        amount=amount,
        source_type="admin",
        source_id=admin_source_id,
        idempotency_key=f"admin_grant_{target_user_id}_{amount}_{uuid.uuid4().hex[:8]}",
    )
    db.add(entry)
    db.commit()
    return get_user_credit_balance(db, target_user_id)


def get_referral_upload_bonus_enabled_at() -> datetime:
    raw = REFERRAL_UPLOAD_BONUS_ENABLED_AT_RAW
    try:
        value = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        if value.tzinfo is not None:
            value = value.astimezone(timezone.utc).replace(tzinfo=None)
        return value
    except ValueError:
        return datetime.fromisoformat(REFERRAL_UPLOAD_BONUS_ENABLED_AT_DEFAULT)


def get_referral_upload_bonus_progress(db, inviter_user_id: int) -> tuple[int, int, int]:
    enabled_at = get_referral_upload_bonus_enabled_at()
    qualified_count = (
        db.query(func.count(Referral.id))
        .filter(Referral.inviter_user_id == inviter_user_id)
        .filter(Referral.qualified_upload_at.isnot(None))
        .filter(Referral.qualified_upload_at >= enabled_at)
        .scalar()
    ) or 0

    completed_bonuses = qualified_count // REFERRAL_UPLOAD_BONUS_SIZE
    progress = qualified_count % REFERRAL_UPLOAD_BONUS_SIZE
    return progress, REFERRAL_UPLOAD_BONUS_SIZE, completed_bonuses


def grant_referral_upload_bonus_if_needed(db, invited_user_id: int):
    """
    Mark the invited user's referral as qualified (first upload) if not yet done.
    Grant +1 credit to the inviter for every 3rd friend who qualifies.
    Returns inviter_user_id if a bonus was granted, else None.
    """
    referral = (
        db.query(Referral)
        .filter(Referral.invited_user_id == invited_user_id)
        .filter(Referral.qualified_upload_at.is_(None))
        .first()
    )

    if not referral:
        return None   # already marked, or no referral record

    # Mark this referral as qualified
    referral.qualified_upload_at = datetime.utcnow()
    db.commit()

    # Count total qualified referrals for this inviter
    qualified_count = (
        db.query(Referral)
        .filter(Referral.inviter_user_id == referral.inviter_user_id)
        .filter(Referral.qualified_upload_at.isnot(None))
        .count()
    )

    # Grant bonus for every 3rd qualified friend (3, 6, 9, ...)
    if qualified_count % 3 != 0:
        return None

    # Unique idempotency key per milestone
    idem_key = f"referral_upload_bonus_{referral.inviter_user_id}_count{qualified_count}"
    existing = (
        db.query(CreditLedger)
        .filter(CreditLedger.idempotency_key == idem_key)
        .first()
    )
    if existing:
        return None

    bonus = CreditLedger(
        user_id=referral.inviter_user_id,
        operation_type="referral_upload_bonus",
        amount=1,
        source_type="referral",
        source_id=str(referral.id),
        idempotency_key=idem_key,
    )
    db.add(bonus)
    db.commit()

    return referral.inviter_user_id


def build_referral_progress_notification_text(progress: int, target: int) -> str:
    return (
        "🎉 Новый друг загрузил файл по твоей ссылке.\n"
        f"Реферальный прогресс: {progress}/{target}."
    )


def build_referral_upload_bonus_awarded_text() -> str:
    return "🎉 Ты получил +1 оформление за приглашённых друзей."

def grant_referral_payment_bonus_if_needed(db, invited_user_id: int) -> None:
    referral = (
        db.query(Referral)
        .filter(Referral.invited_user_id == invited_user_id)
        .filter(Referral.first_payment_at.is_(None))
        .first()
    )

    if not referral:
        return

    existing_bonus = (
        db.query(CreditLedger)
        .filter(CreditLedger.idempotency_key == f"referral_payment_bonus_{referral.id}")
        .first()
    )
    if existing_bonus:
        return

    bonus = CreditLedger(
        user_id=referral.inviter_user_id,
        operation_type="referral_payment_bonus",
        amount=1,
        source_type="referral",
        source_id=str(referral.id),
        idempotency_key=f"referral_payment_bonus_{referral.id}",
    )
    db.add(bonus)
    referral.first_payment_at = datetime.utcnow()
    db.commit()


def apply_successful_payment(db, paid_user_id: int, credits: int, provider: str = "manual") -> None:
    payment_entry = CreditLedger(
        user_id=paid_user_id,
        operation_type="payment_bonus",
        amount=credits,
        source_type=provider,
        source_id=str(paid_user_id),
        idempotency_key=f"payment_bonus_{paid_user_id}_{credits}_{uuid.uuid4().hex[:8]}",
    )
    db.add(payment_entry)
    db.commit()

    grant_referral_payment_bonus_if_needed(db, paid_user_id)

def track_event(
    db,
    event_name: str,
    user_id: int | None = None,
    source: str | None = None,
    payload_json: str | None = None,
) -> None:
    event = AnalyticsEvent(
        user_id=user_id,
        event_name=event_name,
        source=source,
        payload_json=payload_json,
    )
    db.add(event)
    db.commit()
    
# =========================
# Методички
# =========================

def get_available_guides() -> list[dict]:
    result = []
    for guide_code, guide in GUIDES.items():
        if not guide.get("is_active", False):
            continue

        result.append({
            "guide_code": guide_code,
            "title": guide["title"],
            "university_code": guide["university_code"],
            "document_type": guide["document_type"],
            "guideline_version": guide["guideline_version"],
        })
    return result


def get_guide(guide_code: str) -> dict:
    guide = GUIDES.get(guide_code)
    if not guide or not guide.get("is_active", False):
        raise ValueError("Методичка недоступна")
    return guide

def _normalize_filename_text(value: str) -> str:
    value = unicodedata.normalize("NFC", value)
    return value.strip().lower()


def find_method_file(guide_code: str) -> Path | None:
    guide = get_guide(guide_code)
    basename = _normalize_filename_text(guide["method_basename"])

    if not ASSETS_DIR.exists():
        return None

    candidates = []
    for path in ASSETS_DIR.iterdir():
        if not path.is_file():
            continue
        suffix = path.suffix.lower()
        name_normalized = _normalize_filename_text(path.name)
        if suffix not in {".docx", ".pdf"} and name_normalized != basename:
            continue

        filename_part = path.stem if suffix in {".docx", ".pdf"} else path.name
        stem_normalized = _normalize_filename_text(filename_part)
        candidates.append((path, stem_normalized))

    # 1. Точное совпадение
    for path, stem_normalized in candidates:
        if stem_normalized == basename:
            return path

    # 2. Частичное совпадение в обе стороны
    for path, stem_normalized in candidates:
        if basename in stem_normalized or stem_normalized in basename:
            return path

    return None


def build_guide_selection_text(user: User) -> str:
    current_code = get_user_selected_guide_code(user)
    current_guide = get_guide(current_code)

    return (
        "Выберите методичку для оформления.\n\n"
        f"Сейчас активна:\n{current_guide['title']}\n\n"
        "Эта методичка будет использоваться для следующей проверки или обработки .docx."
    )


def build_guide_selected_text(guide_code: str) -> str:
    guide = get_guide(guide_code)
    return (
        "Методичка выбрана.\n\n"
        f"Активная методичка: {guide['title']}\n\n"
        "Теперь можно:\n"
        "• проверить оформление .docx\n"
        "• или оформить его по этой методичке"
    )


def build_method_file_missing_text(guide_code: str) -> str:
    guide = get_guide(guide_code)
    return (
        "Файл методички пока не найден в папке assets.\n\n"
        f"Ожидалась методичка:\n{guide['title']}"
    )


# =========================
# Тексты интерфейса
# =========================

def build_tariffs_text() -> str:
    return (
        "Тарифы:\n"
        f"• {TARIFFS['one']['title']} — {TARIFFS['one']['price_rub']} ₽\n"
        f"• {TARIFFS['three']['title']} — {TARIFFS['three']['price_rub']} ₽"
    )


def build_referral_progress_text(progress: int, target: int) -> str:
    left = max(target - progress, 0)
    if left == target:
        return f"Реферальный прогресс: {progress}/{target}. Пригласи 3 друзей на автопроверку, чтобы получить +1 оформление."
    if left == 1:
        return f"Реферальный прогресс: {progress}/{target}. Остался 1 новый друг до +1 оформления."
    return f"Реферальный прогресс: {progress}/{target}. Осталось {left} новых друга до +1 оформления."


def build_start_text(
    balance: int,
    is_new: bool,
    active_guide_title: str,
    referral_progress: int,
    referral_target: int,
) -> str:
    balance_line = (
        "У тебя есть 1 бесплатное оформление."
        if is_new
        else f"Доступно оформлений: {balance}."
    )
    return (
        "Что можно сделать:\n\n"
        "🔍 <b>Проверить оформление</b>\n"
        "— загрузи файл и получи список ошибок\n\n"
        "✍️ <b>Оформить работу</b>\n"
        "— автоматически исправлю документ по методичке КФУ\n\n"
        f"{balance_line}\n\n"
        "Продолжая использование бота, вы соглашаетесь с "
        "<a href=\"https://docs.google.com/document/d/14Sk5N1ow-x30Dh2dLqYQtUU5-LbdUlemkTleL-THJDc/edit?usp=drivesdk\">Политикой обработки персональных данных</a> "
        "и <a href=\"https://docs.google.com/document/d/1x4OYZzURefM4RWWSRipuX0QQfRpmhBooUjZa11IY8Ck/edit?usp=sharing\">Условиями использования сервиса</a>.\n\n"
        "Сервис выполняет автоматическое форматирование и не гарантирует 100% соответствие требованиям преподавателя. "
        "Загруженные документы обрабатываются автоматически и не просматриваются вручную. "
        "Сервис предназначен только для форматирования документов."
    )


def build_balance_text(user: User, balance: int, bot_username: str) -> str:
    guide_code = get_user_selected_guide_code(user)
    guide = get_guide(guide_code)
    referral_text = build_referral_text(bot_username, user, balance=balance)

    return (
        f"Ваш баланс: {balance} оформлений\n"
        f"Активная методичка: {guide['title']}\n\n"
        f"{build_tariffs_text()}\n\n"
        f"{referral_text}"
    )


def build_no_credits_text(user: User, bot_username: str) -> str:
    referral_text = build_referral_text(bot_username, user)
    return (
        "У вас нет доступных оформлений.\n\n"
        "Автооформление сейчас не запущено. Можно купить оформление или бесплатно проверить файл по методичке.\n\n"
        f"{build_tariffs_text()}\n\n"
        f"{referral_text}"
    )


def build_text_fallback_text() -> str:
    return (
        "Отправь .docx-файл прямо сюда.\n\n"
        "Я могу:\n"
        "• бесплатно проверить оформление\n"
        "• или оформить работу по активной методичке\n\n"
        "Если просто отправить файл, я запущу бесплатную проверку. Для оформления сначала нажми «Оформить работу»."
    )


def build_top_up_balance_text() -> str:
    return (
        "Тарифы:\n"
        "• 1 оформление — 200 ₽\n"
        "• 3 оформления — 500 ₽\n\n"
        "После оплаты оформления начислятся на баланс.\n\n"
        "Если не хочешь покупать сразу, можно бесплатно проверить оформление файла."
    )


def build_check_selected_text() -> str:
    return (
        "Отправь .docx-файл.\n"
        "По умолчанию я бесплатно проверю оформление по методичке КФУ."
    )


def build_format_selected_text(balance: int) -> str:
    if balance <= 0:
        return (
            "На балансе нет доступных оформлений.\n\n"
            f"{build_tariffs_text()}\n\n"
            "Можно купить оформление или получить +1 оформление за каждых 3 новых друзей, которые впервые загрузят .docx на автопроверку по твоей ссылке."
        )

    return (
        "Оформление выбрано.\n\n"
        f"Доступно оформлений: {balance}\n\n"
        "Отправь .docx-файл ещё раз для автооформления.\n"
        "Оформление спишет 1 оформление с баланса.\n"
        "Если перед этим была проверка, она не списывала оформление."
    )


def build_file_received_text(mode: str) -> str:
    if mode == "check":
        return (
            "Файл получен.\n\n"
            "Начинаю проверку оформления по методичке КФУ."
        )

    return (
        "Файл получен.\n\n"
        "Начинаю оформление по методичке КФУ.\n"
        "Готовый .docx-файл отправлю сюда автоматически."
    )


def build_check_result_text(problems: list[str]) -> str:
    if not problems:
        return (
            "Проверка завершена.\n\n"
            "Я не нашёл явных проблем по основным правилам оформления.\n\n"
            "Если хочешь, всё равно можешь запустить автооформление, чтобы привести документ к единому виду по методичке."
        )

    lines = "\n".join(f"• {problem}" for problem in problems)
    return (
        "Проверка завершена.\n\n"
        "Найдены замечания по оформлению:\n"
        f"{lines}\n\n"
        "Чтобы исправить это автоматически, нажми «Оформить работу»."
    )


def build_check_another_text() -> str:
    return (
        "Отправь .docx-файл.\n"
        "Я бесплатно проверю оформление по методичке КФУ."
    )


def build_format_success_caption(bot_username: str, user: User) -> str:
    referral_link = get_referral_link(bot_username, user.referral_code)
    return (
        "Готово, курсовая оформлена.\n\n"
        "Если у тебя есть одногруппники, которые тоже сдают курсовую — скинь им бота по рефералке:\n"
        f"{referral_link}\n\n"
        "Ты получишь:\n"
        "• +1 оформление за каждых 3 новых друзей, которые впервые загрузят .docx на автопроверку\n"
        "• ещё +1 оформление, когда друг впервые оплатит"
    )


# =========================
# Форматирование документа
# =========================

def create_document_record(db, user_id: int, original_filename: str, storage_path: str) -> Document:
    document = Document(
        user_id=user_id,
        original_filename=original_filename,
        storage_path=storage_path,
    )
    db.add(document)
    db.flush()
    return document


def create_formatting_request(
    db,
    user_id: int,
    document_id: int,
    guide_code: str,
    service_type: str = "format",
) -> FormattingRequest:
    guide = get_guide(guide_code)

    formatting_request = FormattingRequest(
        user_id=user_id,
        document_id=document_id,
        service_type=service_type,
        university_code=guide["university_code"],
        document_type=guide["document_type"],
        guideline_version=guide["guideline_version"],
        status="queued",
    )
    db.add(formatting_request)
    db.commit()
    db.refresh(formatting_request)
    return formatting_request


def mark_formatting_processing(db, request_id: int) -> bool:
    formatting_request = (
        db.query(FormattingRequest)
        .filter(FormattingRequest.id == request_id)
        .first()
    )
    if not formatting_request:
        return False

    if formatting_request.status != "queued":
        return False

    formatting_request.status = "processing"
    db.commit()
    return True

def get_formatting_request_by_id(db, request_id: int) -> FormattingRequest | None:
    return (
        db.query(FormattingRequest)
        .filter(FormattingRequest.id == request_id)
        .first()
    )


def mark_formatting_done(db, request_id: int, result_file_path: str) -> None:
    formatting_request = (
        db.query(FormattingRequest)
        .filter(FormattingRequest.id == request_id)
        .first()
    )
    if not formatting_request:
        return

    formatting_request.status = "done"
    formatting_request.result_file_path = result_file_path
    formatting_request.completed_at = datetime.utcnow()
    db.commit()


def mark_formatting_failed_in_new_session(request_id: int, error_text: str) -> None:
    db = SessionLocal()
    try:
        formatting_request = (
            db.query(FormattingRequest)
            .filter(FormattingRequest.id == request_id)
            .first()
        )
        if formatting_request:
            formatting_request.status = "failed"
            formatting_request.error_message = error_text[:1000]
            db.commit()
    finally:
        db.close()


def format_document_by_guide(
    guide_code: str, input_path: str, output_path: str
) -> tuple[str, list[str]]:
    """
    Run the formatter for *guide_code* and return (output_path, warnings).
    *warnings* is a list of short Russian strings for the trouble-report.
    """
    guide = get_guide(guide_code)
    formatter = guide["formatter"]
    return formatter(input_path, output_path)


def build_processing_paths(original_filename: str) -> tuple[str, Path, Path]:
    job_id = str(uuid.uuid4())
    original_name = Path(original_filename)
    safe_name = f"{original_name.stem}.docx"

    input_path = TEMP_DIR / f"{job_id}_input.docx"
    output_path = TEMP_DIR / safe_name

    return safe_name, input_path, output_path


def cleanup_temp_files(*paths: Path | None) -> None:
    for path in paths:
        try:
            if path and path.exists():
                path.unlink()
        except Exception:
            pass


# =========================
# Простые хелперы для handlers
# =========================

def ensure_user(
    db,
    telegram_id: int,
    username: str | None,
    first_name: str | None,
    last_name: str | None,
    referral_code_from_start: str | None = None,
) -> tuple[User, bool]:
    return get_or_create_user(
        db=db,
        telegram_id=telegram_id,
        username=username,
        first_name=first_name,
        last_name=last_name,
        referral_code_from_start=referral_code_from_start,
    )


def get_bot_username_fallback() -> str:
    return BOT_USERNAME_FALLBACK or "your_bot_username"


def get_userinfo_text(db, user: User) -> str:
    balance = get_user_credit_balance(db, user.id)
    selected_guide_code = get_user_selected_guide_code(user)
    referral_progress, referral_target, completed_bonuses = get_referral_upload_bonus_progress(
        db,
        user.id,
    )

    return (
        f"user_id: {user.id}\n"
        f"telegram_id: {user.telegram_id}\n"
        f"referral_code: {user.referral_code}\n"
        f"balance: {balance}\n"
        f"selected_guide_code: {selected_guide_code}\n"
        f"referral_upload_bonus_progress: {referral_progress}/{referral_target}\n"
        f"referral_upload_bonuses_completed: {completed_bonuses}"
    )
