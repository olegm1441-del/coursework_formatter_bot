import os
import hmac
import hashlib
import logging
import uuid
from decimal import Decimal, InvalidOperation
from telegram import Bot
from sqlalchemy import func
from sqlalchemy.exc import IntegrityError
from datetime import datetime

import httpx
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse, PlainTextResponse
from sqlalchemy.orm import Session

from db import SessionLocal
from models import Payment, CreditLedger, User, Referral

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TRIBUTE_API_KEY = os.getenv("TRIBUTE_API_KEY")
SUCCESS_URL = os.getenv("PAYMENT_SUCCESS_URL")
FAIL_URL = os.getenv("PAYMENT_FAIL_URL")
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "").rstrip("/")
YOOKASSA_SHOP_ID = os.getenv("YOOKASSA_SHOP_ID")
YOOKASSA_SECRET_KEY = os.getenv("YOOKASSA_SECRET_KEY")
YOOKASSA_API_BASE_URL = "https://api.yookassa.ru/v3"

VK_SECRET = os.getenv("VK_SECRET")
VK_CONFIRMATION_CODE = os.getenv("VK_CONFIRMATION_CODE")
VK_GROUP_ID = os.getenv("VK_GROUP_ID")
VK_GROUP_TOKEN = os.getenv("VK_GROUP_TOKEN")
VK_API_VERSION = os.getenv("VK_API_VERSION", "5.199")

# Готовые ссылки на товары Tribute
BUY1_LINK = "https://t.me/tribute/app?startapp=psvI"
BUY3_LINK = "https://t.me/tribute/app?startapp=psvJ"

YOOKASSA_TARIFFS = {
    "one_format": {
        "amount_rub": 200,
        "credits": 1,
        "description": "1 оформление курсовой КФУ",
    },
    "three_formats": {
        "amount_rub": 500,
        "credits": 3,
        "description": "3 оформления курсовой КФУ",
    },
}

app = FastAPI()
logger.info("public_base_url=%s", PUBLIC_BASE_URL or "<unset>")


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
        return BUY3_LINK, 380
    if tariff_code == "one_format":
        return BUY1_LINK, 150
    return None, 0


def _get_yookassa_tariff(tariff_code: str) -> dict | None:
    return YOOKASSA_TARIFFS.get(tariff_code)


def _format_rub(amount_rub: int) -> str:
    return f"{amount_rub}.00"


def _parse_amount_rub(value: str | None) -> int | None:
    try:
        return int(Decimal(str(value)).quantize(Decimal("1")))
    except (InvalidOperation, TypeError, ValueError):
        return None


def _get_yookassa_return_url() -> str | None:
    if not PUBLIC_BASE_URL:
        return None
    return f"{PUBLIC_BASE_URL}/payment-success"


def _get_yookassa_auth() -> tuple[str, str] | None:
    if not YOOKASSA_SHOP_ID or not YOOKASSA_SECRET_KEY:
        return None
    return YOOKASSA_SHOP_ID, YOOKASSA_SECRET_KEY


def _safe_yookassa_payment_log_fields(data: dict) -> dict:
    payment_object = data.get("object") or {}
    amount = payment_object.get("amount") or {}
    return {
        "event": data.get("event"),
        "object_id": payment_object.get("id"),
        "object_status": payment_object.get("status"),
        "object_paid": payment_object.get("paid"),
        "amount_value": amount.get("value"),
        "metadata": payment_object.get("metadata") or {},
    }


async def _create_yookassa_payment(
    user_id: int,
    tariff_code: str,
) -> tuple[str | None, int, str | None]:
    tariff = _get_yookassa_tariff(tariff_code)
    if not tariff:
        return None, 0, "unknown_tariff_code"

    auth = _get_yookassa_auth()
    if not auth:
        return None, 0, "yookassa_config_missing"

    return_url = _get_yookassa_return_url()
    if not return_url:
        return None, 0, "public_base_url_missing"

    amount_rub = int(tariff["amount_rub"])
    credits = int(tariff["credits"])
    request_body = {
        "amount": {
            "value": _format_rub(amount_rub),
            "currency": "RUB",
        },
        "capture": True,
        "confirmation": {
            "type": "redirect",
            "return_url": return_url,
        },
        "description": tariff["description"],
        "metadata": {
            "user_id": str(user_id),
            "tariff_code": tariff_code,
            "credits": str(credits),
        },
    }

    headers = {"Idempotence-Key": f"create_payment:{uuid.uuid4().hex}"}
    async with httpx.AsyncClient(timeout=30.0) as client:
        try:
            response = await client.post(
                f"{YOOKASSA_API_BASE_URL}/payments",
                auth=auth,
                headers=headers,
                json=request_body,
            )
            response.raise_for_status()
        except httpx.HTTPStatusError as exc:
            logger.warning(
                "yookassa_create_payment_http_error status=%s response=%s",
                exc.response.status_code,
                exc.response.text[:500],
            )
            return None, amount_rub, "yookassa_http_error"
        except httpx.HTTPError as exc:
            logger.warning("yookassa_create_payment_request_error error=%s", exc)
            return None, amount_rub, "yookassa_request_error"

    data = response.json()
    confirmation = data.get("confirmation") or {}
    payment_url = confirmation.get("confirmation_url")
    if not payment_url:
        logger.warning(
            "yookassa_create_payment_missing_confirmation payment_id=%s status=%s",
            data.get("id"),
            data.get("status"),
        )
        return None, amount_rub, "missing_confirmation_url"

    logger.info(
        "yookassa_payment_created payment_id=%s user_id=%s tariff=%s amount_rub=%s",
        data.get("id"),
        user_id,
        tariff_code,
        amount_rub,
    )
    return payment_url, amount_rub, None


async def _fetch_yookassa_payment(payment_id: str) -> dict | None:
    auth = _get_yookassa_auth()
    if not auth:
        logger.error("yookassa_config_missing_for_verify payment_id=%s", payment_id)
        return None

    async with httpx.AsyncClient(timeout=30.0) as client:
        try:
            response = await client.get(
                f"{YOOKASSA_API_BASE_URL}/payments/{payment_id}",
                auth=auth,
            )
            response.raise_for_status()
        except httpx.HTTPStatusError as exc:
            logger.warning(
                "yookassa_verify_payment_http_error payment_id=%s status=%s response=%s",
                payment_id,
                exc.response.status_code,
                exc.response.text[:500],
            )
            return None
        except httpx.HTTPError as exc:
            logger.warning(
                "yookassa_verify_payment_request_error payment_id=%s error=%s",
                payment_id,
                exc,
            )
            return None

    return response.json()


def _parse_paid_at(value: str | None) -> datetime:
    if not value:
        return datetime.utcnow()
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except Exception:
        return datetime.utcnow()


def _apply_first_payment_referral_bonus(
    db: Session,
    invited_user_id: int,
    paid_at: datetime,
) -> int | None:
    referral = (
        db.query(Referral)
        .filter(
            Referral.invited_user_id == invited_user_id,
            Referral.first_payment_at.is_(None),
        )
        .first()
    )
    if not referral:
        return None

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

    return referral.inviter_user_id


async def _apply_yookassa_payment(payment: dict) -> dict:
    payment_id = str(payment.get("id") or "")
    metadata = payment.get("metadata") or {}
    tariff_code = metadata.get("tariff_code")
    tariff = _get_yookassa_tariff(tariff_code)
    if not payment_id:
        return {"status": "ignored", "reason": "missing_payment_id"}
    if not tariff:
        logger.info(
            "yookassa_webhook_unknown_tariff payment_id=%s tariff_code=%s",
            payment_id,
            tariff_code,
        )
        return {"status": "unknown_tariff"}

    try:
        user_id = int(metadata.get("user_id"))
    except (TypeError, ValueError):
        logger.info(
            "yookassa_webhook_invalid_user_id payment_id=%s metadata=%s",
            payment_id,
            metadata,
        )
        return {"status": "ignored", "reason": "invalid_user_id"}

    amount = payment.get("amount") or {}
    amount_rub = _parse_amount_rub(amount.get("value")) or int(tariff["amount_rub"])
    expected_amount_rub = int(tariff["amount_rub"])
    if amount_rub != expected_amount_rub:
        logger.warning(
            "yookassa_amount_mismatch payment_id=%s amount_rub=%s expected_amount_rub=%s",
            payment_id,
            amount_rub,
            expected_amount_rub,
        )
        amount_rub = expected_amount_rub

    credits = int(tariff["credits"])
    try:
        metadata_credits = int(metadata.get("credits"))
    except (TypeError, ValueError):
        metadata_credits = credits
    if metadata_credits != credits:
        logger.warning(
            "yookassa_credits_mismatch payment_id=%s metadata_credits=%s expected_credits=%s",
            payment_id,
            metadata_credits,
            credits,
        )

    paid_at = _parse_paid_at(payment.get("captured_at") or payment.get("created_at"))

    db: Session = SessionLocal()
    try:
        user = db.query(User).filter(User.id == user_id).first()
        if not user:
            logger.info(
                "yookassa_payment_user_not_found user_id=%s payment_id=%s",
                user_id,
                payment_id,
            )
            return {"status": "user_not_found"}

        existing_payment = db.query(Payment).filter(
            Payment.external_payment_id == payment_id
        ).first()

        if existing_payment and existing_payment.status == "paid":
            logger.info("yookassa_payment_already_processed payment_id=%s", payment_id)
            return {"status": "already_processed"}

        if not existing_payment:
            existing_payment = Payment(
                user_id=user.id,
                provider="yookassa",
                tariff_code=tariff_code,
                amount_rub=amount_rub,
                status="paid",
                external_payment_id=payment_id,
                paid_at=paid_at,
            )
            db.add(existing_payment)
        else:
            existing_payment.provider = "yookassa"
            existing_payment.status = "paid"
            existing_payment.tariff_code = tariff_code
            existing_payment.amount_rub = amount_rub
            existing_payment.paid_at = paid_at

        credit = CreditLedger(
            user_id=user.id,
            operation_type="purchase",
            amount=credits,
            source_type="yookassa_payment",
            source_id=payment_id,
            idempotency_key=f"yookassa:{payment_id}",
        )
        db.add(credit)

        inviter_user_id = _apply_first_payment_referral_bonus(db, user.id, paid_at)

        try:
            db.commit()
        except IntegrityError as e:
            db.rollback()
            logger.warning(
                "yookassa_payment_commit_integrity_error payment_id=%s error=%s",
                payment_id,
                e,
            )

            duplicate_payment = db.query(Payment).filter(
                Payment.external_payment_id == payment_id
            ).first()
            duplicate_credit = db.query(CreditLedger).filter(
                CreditLedger.idempotency_key == f"yookassa:{payment_id}"
            ).first()

            if duplicate_payment or duplicate_credit:
                logger.info(
                    "yookassa_payment_already_processed_by_constraint payment_id=%s",
                    payment_id,
                )
                return {"status": "already_processed"}

            raise

        balance = (
            db.query(CreditLedger)
            .filter(CreditLedger.user_id == user.id)
            .with_entities(func.sum(CreditLedger.amount))
            .scalar()
        ) or 0

        try:
            bot = Bot(token=os.getenv("BOT_TOKEN"))
            await bot.send_message(
                chat_id=user.telegram_id,
                text=(
                    f"✅ Оплата получена!\n\n"
                    f"Начислено: {credits} оформлений.\n"
                    f"Ваш баланс: {balance} оформлений.\n\n"
                    "Можно сразу отправить следующий .docx-файл в этот чат.\n"
                    "Или пригласить друга по реферальной ссылке и получить ещё бонус."
                ),
            )

            if inviter_user_id:
                inviter = db.query(User).filter(User.id == inviter_user_id).first()
                if inviter:
                    inviter_balance = (
                        db.query(CreditLedger)
                        .filter(CreditLedger.user_id == inviter.id)
                        .with_entities(func.sum(CreditLedger.amount))
                        .scalar()
                    ) or 0

                    await bot.send_message(
                        chat_id=inviter.telegram_id,
                        text=(
                            "🎉 Начислен реферальный бонус!\n\n"
                            "Причина: приглашённый пользователь впервые оплатил.\n"
                            "Вы получили +1 оформление.\n"
                            f"Ваш баланс: {inviter_balance} оформлений."
                        ),
                    )

        except Exception as e:
            logger.error("yookassa_payment_notification_failed %s", e)

        logger.info(
            "yookassa_payment_processed payment_id=%s user_id=%s tariff=%s credits=%s",
            payment_id,
            user.id,
            tariff_code,
            credits,
        )
        return {"status": "ok"}
    finally:
        db.close()


@app.get("/health")
async def health():
    return {"status": "ok"}


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
        hashlib.sha256,
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
        logger.info(
            "webhook_missing_required_fields purchase_id=%s telegram_user_id=%s",
            purchase_id,
            telegram_user_id,
        )
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

    amount_rub = 380 if tariff_code == "three_formats" else 150

    db: Session = SessionLocal()
    try:
        user = db.query(User).filter(User.telegram_id == int(telegram_user_id)).first()
        if not user:
            logger.info(
                "payment_user_not_found telegram_user_id=%s purchase_id=%s",
                telegram_user_id,
                purchase_id,
            )
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

        inviter_user_id = _apply_first_payment_referral_bonus(db, user.id, paid_at)

        try:
            db.commit()
        except IntegrityError as e:
            db.rollback()
            logger.warning(
                "payment_commit_integrity_error purchase_id=%s error=%s",
                purchase_id,
                e,
            )

            duplicate_payment = db.query(Payment).filter(
                Payment.external_payment_id == str(purchase_id)
            ).first()
            duplicate_credit = db.query(CreditLedger).filter(
                CreditLedger.idempotency_key == f"tribute:{purchase_id}"
            ).first()

            if duplicate_payment or duplicate_credit:
                logger.info(
                    "payment_already_processed_by_constraint purchase_id=%s",
                    purchase_id,
                )
                return {"status": "already_processed"}

            raise

        balance = (
            db.query(CreditLedger)
            .filter(CreditLedger.user_id == user.id)
            .with_entities(func.sum(CreditLedger.amount))
            .scalar()
        ) or 0

        try:
            bot = Bot(token=os.getenv("BOT_TOKEN"))
            await bot.send_message(
                chat_id=user.telegram_id,
                text=(
                    f"✅ Оплата получена!\n\n"
                    f"Начислено: {credits} оформлений.\n"
                    f"Ваш баланс: {balance} оформлений.\n\n"
                    "Можно сразу отправить следующий .docx-файл в этот чат.\n"
                    "Или пригласить друга по реферальной ссылке и получить ещё бонус."
                ),
            )

            if inviter_user_id:
                inviter = db.query(User).filter(User.id == inviter_user_id).first()
                if inviter:
                    inviter_balance = (
                        db.query(CreditLedger)
                        .filter(CreditLedger.user_id == inviter.id)
                        .with_entities(func.sum(CreditLedger.amount))
                        .scalar()
                    ) or 0

                    await bot.send_message(
                        chat_id=inviter.telegram_id,
                        text=(
                            "🎉 Начислен реферальный бонус!\n\n"
                            "Причина: приглашённый пользователь впервые оплатил.\n"
                            "Вы получили +1 оформление.\n"
                            f"Ваш баланс: {inviter_balance} оформлений."
                        ),
                    )

        except Exception as e:
            logger.error("payment_notification_failed %s", e)

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


@app.post("/yookassa/webhook")
async def yookassa_webhook(request: Request):
    data = await request.json()
    logger.info("yookassa_webhook_received %s", _safe_yookassa_payment_log_fields(data))

    event_name = data.get("event")
    payment = data.get("object") or {}
    payment_id = payment.get("id")
    payment_status = payment.get("status")
    payment_paid = payment.get("paid")

    if event_name != "payment.succeeded":
        logger.info("yookassa_webhook_ignored_event event=%s", event_name)
        return {"status": "ignored"}

    if payment_status != "succeeded" or payment_paid is not True:
        logger.info(
            "yookassa_webhook_ignored_not_succeeded payment_id=%s status=%s paid=%s",
            payment_id,
            payment_status,
            payment_paid,
        )
        return {"status": "ignored"}

    if not payment_id:
        logger.info("yookassa_webhook_missing_payment_id")
        return {"status": "ignored"}

    verified_payment = await _fetch_yookassa_payment(str(payment_id))
    if not verified_payment:
        raise HTTPException(status_code=503, detail="Payment verification failed")

    if (
        verified_payment.get("status") != "succeeded"
        or verified_payment.get("paid") is not True
    ):
        logger.info(
            "yookassa_verified_payment_not_succeeded payment_id=%s status=%s paid=%s",
            payment_id,
            verified_payment.get("status"),
            verified_payment.get("paid"),
        )
        return {"status": "ignored"}

    return await _apply_yookassa_payment(verified_payment)


@app.post("/create-payment")
async def create_payment(user_id: int, tariff_code: str = "one_format"):
    # Оставляем endpoint для совместимости с текущим handlers.py:
    # он по-прежнему получает {"ok": true, "payment_url": "..."}
    payment_url, amount_rub, error = await _create_yookassa_payment(user_id, tariff_code)

    if not payment_url:
        logger.info(
            "create_payment_failed user_id=%s tariff=%s error=%s",
            user_id,
            tariff_code,
            error,
        )
        return {"ok": False, "error": error or "payment_url_missing"}

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

@app.post("/vk/webhook")
async def vk_webhook(request: Request):
    data = await request.json()

    logger.info("vk_webhook_received type=%s", data.get("type"))

    if data.get("type") == "confirmation":
        return PlainTextResponse(VK_CONFIRMATION_CODE or "")

    if data.get("secret") != VK_SECRET:
        logger.warning("vk_webhook_invalid_secret")
        return PlainTextResponse("forbidden", status_code=403)

    if data.get("type") == "message_new":
        message_obj = ((data.get("object") or {}).get("message") or {})
        logger.info(
            "vk_message_new from_id=%s text=%s",
            message_obj.get("from_id"),
            message_obj.get("text"),
        )
        return PlainTextResponse("ok")

    return PlainTextResponse("ok")
