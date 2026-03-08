import uuid

from sqlalchemy.orm import Session

from models import User, CreditLedger, Referral


def generate_referral_code() -> str:
    return uuid.uuid4().hex[:10]


def get_user_by_telegram_id(db: Session, telegram_id: int):
    return db.query(User).filter(User.telegram_id == telegram_id).first()


def get_user_by_referral_code(db: Session, referral_code: str):
    return db.query(User).filter(User.referral_code == referral_code).first()


def create_user(
    db: Session,
    telegram_id: int,
    username: str | None,
    first_name: str | None,
    last_name: str | None,
    referred_by_user_id: int | None = None,
):
    referral_code = generate_referral_code()

    user = User(
        telegram_id=telegram_id,
        username=username,
        first_name=first_name,
        last_name=last_name,
        referral_code=referral_code,
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

    if referred_by_user_id:
        referral = Referral(
            inviter_user_id=referred_by_user_id,
            invited_user_id=user.id,
            referral_code="linked",
        )
        db.add(referral)

    db.commit()
    db.refresh(user)
    return user
