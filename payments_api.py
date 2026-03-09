import os
import hmac
import hashlib
import logging

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

app = FastAPI()


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

    order = data["payload"]

    if order["status"] != "paid":
        logger.info("webhook_ignored_not_paid")
        return {"status": "ignored"}

    tribute_id = order["uuid"]

    db: Session = SessionLocal()

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

    credit = CreditLedger(
        user_id=payment.user_id,
        operation_type="purchase",
        amount=1,
        source_type="tribute_payment",
        source_id=tribute_id,
        idempotency_key=f"tribute:{tribute_id}",
    )

    db.add(credit)
    db.commit()

    logger.info("payment_processed tribute_id=%s", tribute_id)

    return {"status": "ok"}


@app.post("/create-payment")
async def create_payment(user_id: int):

    async with httpx.AsyncClient() as client:

        response = await client.post(
            "https://tribute.tg/api/v1/shop/orders",
            headers={
                "Api-Key": TRIBUTE_API_KEY
            },
            json={
                "amount": 14900,
                "currency": "rub",
                "title": "1 оформление курсовой",
                "description": "Автоматическое оформление курсовой по методичке КФУ",
                "successUrl": SUCCESS_URL,
                "failUrl": FAIL_URL
            }
        )

    data = response.json()

    payment_url = data["paymentUrl"]
    tribute_id = data["uuid"]

    db: Session = SessionLocal()

    payment = Payment(
        user_id=user_id,
        provider="tribute",
        tariff_code="one_format",
        amount_rub=149,
        status="pending",
        external_payment_id=tribute_id,
    )

    db.add(payment)
    db.commit()

    logger.info("payment_created tribute_id=%s", tribute_id)

    return {"payment_url": payment_url}
