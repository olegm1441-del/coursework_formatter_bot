import os
import hmac
import hashlib
import logging
from datetime import datetime

import httpx
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from sqlalchemy.orm import Session

from db import SessionLocal
from models import Payment, CreditLedger

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

TRIBUTE_API_KEY = os.getenv("TRIBUTE_API_KEY")
SUCCESS_URL = os.getenv("PAYMENT_SUCCESS_URL")
FAIL_URL = os.getenv("PAYMENT_FAIL_URL")

# ID продуктов Tribute
TRIBUTE_PRODUCT_ID_ONE = "1aafbe06-f3bc-45bf-8fd0-1e7cd138a68d"
TRIBUTE_PRODUCT_ID_THREE = "ВСТАВЬ_СЮДА_ID_ПАКЕТА_3_ОФОРМЛЕНИЙ"

app = FastAPI()


def _safe_json(response: httpx.Response):
    try:
        return response.json()
    except Exception:
        return {"raw_text": response.text}


def _credits_for_tariff(tariff_code: str) -> int:
    if tariff_code == "three_formats":
        return 3
    return 1


def _amount_for_tariff(tariff_code: str) -> int:
    if tariff_code == "three_formats":
        return 349
    return 149


def _product_id_for_tariff(tariff_code: str) -> str:
    if tariff_code == "three_formats":
        return TRIBUTE_PRODUCT_ID_THREE
    return TRIBUTE_PRODUCT_ID_ONE


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
    logger.info("webhook_type=%s", data.get("type"))

    if data.get("type") != "shop_order":
        logger.info("webhook_ignored_not_shop_order")
        return {"status": "ignored"}

    payload = data.get("payload") or {}
    order_status = payload.get("status")
    tribute_id = payload.get("uuid")

    if order_status != "paid":
        logger.info("webhook_ignored_not_paid tribute_id=%s status=%s", tribute_id, order_status)
        return {"status": "ignored"}

    if not tribute_id:
        logger.info("webhook_missing_uuid")
        return {"status": "ignored"}

    db: Session = SessionLocal()
    try:
        payment = db.query(Payment).filter(
            Payment.external_payment_id == tribute_id
        ).first()

        if not payment:
            logger.info("payment_not_found tribute_id=%s", tribute_id)
            return {"status": "payment_not_found"}

        if payment.status == "paid":
            logger.info("payment_already_processed tribute_id=%s", tribute_id)
            return {"status": "already_processed"}

        payment.status = "paid"
        payment.paid_at = datetime.utcnow()

        credits = _credits_for_tariff(payment.tariff_code)

        credit = CreditLedger(
            user_id=payment.user_id,
            operation_type="purchase",
            amount=credits,
            source_type="tribute_payment",
            source_id=tribute_id,
            idempotency_key=f"tribute:{tribute_id}",
        )

        db.add(credit)
        db.commit()

        logger.info(
            "payment_processed tribute_id=%s user_id=%s tariff=%s credits=%s",
            tribute_id,
            payment.user_id,
            payment.tariff_code,
            credits,
        )
        return {"status": "ok"}

    finally:
        db.close()


@app.post("/create-payment")
async def create_payment(user_id: int, tariff_code: str = "one_format"):
    if not TRIBUTE_API_KEY:
        logger.error("TRIBUTE_API_KEY is empty")
        raise HTTPException(status_code=500, detail="TRIBUTE_API_KEY is not set")

    product_id = _product_id_for_tariff(tariff_code)
    amount_rub = _amount_for_tariff(tariff_code)

    if not product_id or product_id.startswith("ВСТАВЬ_СЮДА"):
        logger.error("Tribute product id is not configured for tariff_code=%s", tariff_code)
        return {
            "ok": False,
            "error": "product_id_not_configured",
            "message": f"Не настроен productId для тарифа {tariff_code}",
        }

    async with httpx.AsyncClient(timeout=30.0) as client:
        response = await client.post(
            "https://tribute.tg/api/v1/shop/orders",
            headers={"Api-Key": TRIBUTE_API_KEY},
            json={
                "productId": product_id,
                "successUrl": SUCCESS_URL,
                "failUrl": FAIL_URL,
            }
        )

    data = _safe_json(response)

    logger.info(
        "tribute_create_payment status_code=%s tariff_code=%s response=%s",
        response.status_code,
        tariff_code,
        data,
    )

    payment_url = data.get("paymentUrl")
    tribute_id = data.get("uuid")

    if response.status_code >= 400 or not payment_url or not tribute_id:
        return {
            "ok": False,
            "error": "tribute_create_order_failed",
            "status_code": response.status_code,
            "response": data,
        }

    db: Session = SessionLocal()
    try:
        payment = Payment(
            user_id=user_id,
            provider="tribute",
            tariff_code=tariff_code,
            amount_rub=amount_rub,
            status="pending",
            external_payment_id=tribute_id,
        )
        db.add(payment)
        db.commit()
    finally:
        db.close()

    logger.info(
        "payment_created tribute_id=%s user_id=%s tariff=%s",
        tribute_id,
        user_id,
        tariff_code,
    )

    return {
        "ok": True,
        "payment_url": payment_url,
        "tribute_id": tribute_id,
        "tariff_code": tariff_code,
    }
