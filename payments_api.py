import os
import hmac
import hashlib
import logging
from datetime import datetime

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from sqlalchemy.orm import Session

from db import SessionLocal
from models import Payment, CreditLedger, User, Referral

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TRIBUTE_API_KEY = os.getenv("TRIBUTE_API_KEY")
SUCCESS_URL = os.getenv("PAYMENT_SUCCESS_URL")
FAIL_URL = os.getenv("PAYMENT_FAIL_URL")

# Готовые ссылки на товары Tribute
BUY1_LINK = "https://t.me/tribute/app?startapp=psvI"
BUY3_LINK = "https://t.me/tribute/app?startapp=psvJ"

app = FastAPI()


def _normalize_currency(currency: str | None) -> str:
    return (currency or "").strip().upper()


def _resolve_tariff(
    amount: int | None,
    currency: str | None,
    product_name: str | None = None,
    product_id: int | None = None,
) -> tuple[str | None, int]:
    if product_id == 109598:
        return "one_format", 1
    if product_id == 109599:
        return "three_formats", 3

    name = (product_name or "").lower()

    if "3" in name and "формат" in name:
        return "three_formats", 3
    if "формат" in name:
        return "one_format", 1

    return None, 0

def _create_payment_link(tariff_code: str) -> tuple[str | None, int]:
    if tariff_code == "three_formats":
        return BUY3_LINK, 349
    if tariff_code == "one_format":
        return BUY1_LINK, 149
    return None, 0


def _parse_paid_at(value: str | None) -> datetime:
    if not value:
        return datetime.utcnow()
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except Exception:
        return datetime.utcnow()


def _apply_first_payment_referral_bonus(db: Session, invited_user_id: int, paid_at: datetime) -> None:
    referral = (
        db.query(Referral)
        .filter(
            Referral.invited_user_id == invited_user_id,
            Referral.first_payment_at.is_(None),
        )
        .first()
    )
    if not referral:
        return

    referral.first_payment_at = paid_at

    inviter_bonus = CreditLedger(
        user_id=referral.inviter_user_id,
        operation_type="referral_bonus",
        amount=1,
        source_type="referral_first_payment",
        source_id=str(invited_user_id),
        idempotency_key=f"referral:first_payment:{invited_user_id}",
    )
    db.add(inviter_bonus)


@app.get("/payment-success")
async def payment_success():
    return JSONResponse(
        content={"status": "ok", "message": "Оплата прошла. Вернитесь в Telegram."},
        media_type="application/json; charset=utf-8",
    )


@app.get("/payment-fail")
async def payment_fail():
    return JSONResponse(
        content={"status": "fail", "message": "Оплата не завершена. Попробуйте снова."},
        media_type="application/json; charset=utf-8",
    )


@app.post("/tribute/webhook")
async def tribute_webhook(request: Request):
    body = await request.body()
    signature = request.headers.get("trbt-signature")

    logger.info("tribute_webhook_received")
    logger.info("signature_present=%s", bool(signature))

    if not TRIBUTE_API_KEY:
        logger.error("TRIBUTE_API_KEY is empty")
        raise HTTPException(status_code=500, detail="TRIBUTE_API_KEY is not set")

    check = hmac.new(
        TRIBUTE_API_KEY.encode(),
        body,
        hashlib.sha256
    ).hexdigest()

    if signature != check:
        logger.info("signature_invalid")
        raise HTTPException(status_code=401, detail="Invalid signature")

    data = await request.json()
    event_name = data.get("name")
    logger.info("tribute_webhook_name=%s", event_name)

    if event_name != "new_digital_product":
        logger.info("webhook_ignored_not_new_digital_product")
        return {"status": "ignored"}

    payload = data.get("payload") or {}
    product_id = payload.get("product_id")

    purchase_id = payload.get("purchase_id")
    telegram_user_id = payload.get("telegram_user_id")
    amount = payload.get("amount")
    currency = payload.get("currency")
    product_name = payload.get("product_name")
    paid_at = _parse_paid_at(payload.get("purchase_created_at"))

    if not purchase_id or not telegram_user_id:
        logger.info("webhook_missing_required_fields purchase_id=%s telegram_user_id=%s", purchase_id, telegram_user_id)
        return {"status": "ignored"}

    tariff_code, credits = _resolve_tariff(amount, currency, product_name, product_id)
    if not tariff_code:
        logger.info(
            "webhook_unknown_product purchase_id=%s amount=%s currency=%s product_name=%s",
            purchase_id,
            amount,
            currency,
            product_name,
        )
        return {"status": "unknown_product"}

    amount_rub = 349 if tariff_code == "three_formats" else 149

    db: Session = SessionLocal()
    try:
        user = db.query(User).filter(User.telegram_id == int(telegram_user_id)).first()
        if not user:
            logger.info("payment_user_not_found telegram_user_id=%s purchase_id=%s", telegram_user_id, purchase_id)
            return {"status": "user_not_found"}

        existing_payment = db.query(Payment).filter(
            Payment.external_payment_id == str(purchase_id)
        ).first()

        if existing_payment and existing_payment.status == "paid":
            logger.info("payment_already_processed purchase_id=%s", purchase_id)
            return {"status": "already_processed"}

        if not existing_payment:
            existing_payment = Payment(
                user_id=user.id,
                provider="tribute",
                tariff_code=tariff_code,
                amount_rub=amount_rub,
                status="paid",
                external_payment_id=str(purchase_id),
                paid_at=paid_at,
            )
            db.add(existing_payment)
        else:
            existing_payment.status = "paid"
            existing_payment.tariff_code = tariff_code
            existing_payment.amount_rub = amount_rub
            existing_payment.paid_at = paid_at

        credit = CreditLedger(
            user_id=user.id,
            operation_type="purchase",
            amount=credits,
            source_type="tribute_payment",
            source_id=str(purchase_id),
            idempotency_key=f"tribute:{purchase_id}",
        )
        db.add(credit)

        _apply_first_payment_referral_bonus(db, user.id, paid_at)

        db.commit()

        logger.info(
            "payment_processed purchase_id=%s user_id=%s tariff=%s credits=%s",
            purchase_id,
            user.id,
            tariff_code,
            credits,
        )
        return {"status": "ok"}

    finally:
        db.close()


@app.post("/create-payment")
async def create_payment(user_id: int, tariff_code: str = "one_format"):
    # Оставляем endpoint для совместимости с текущим handlers.py:
    # он по-прежнему получает {"ok": true, "payment_url": "..."}
    payment_url, amount_rub = _create_payment_link(tariff_code)

    if not payment_url:
        logger.info("unknown_tariff_code tariff_code=%s user_id=%s", tariff_code, user_id)
        return {"ok": False, "error": "unknown_tariff_code"}

    logger.info(
        "payment_link_created user_id=%s tariff=%s amount_rub=%s",
        user_id,
        tariff_code,
        amount_rub,
    )

    return {
        "ok": True,
        "payment_url": payment_url,
        "tariff_code": tariff_code,
        "amount_rub": amount_rub,
    }
