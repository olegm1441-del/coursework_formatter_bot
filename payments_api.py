import os
import hmac
import hashlib
import httpx

from fastapi import FastAPI, Request, HTTPException
from sqlalchemy.orm import Session

from db import SessionLocal
from models import Payment, CreditLedger

TRIBUTE_API_KEY = os.getenv("TRIBUTE_API_KEY")
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL")
SUCCESS_URL = os.getenv("PAYMENT_SUCCESS_URL")
FAIL_URL = os.getenv("PAYMENT_FAIL_URL")

app = FastAPI()


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


@app.get("/payment-success")
async def payment_success():
    return {"status": "ok", "message": "Оплата прошла. Вернитесь в Telegram."}


@app.get("/payment-fail")
async def payment_fail():
    return {"status": "fail", "message": "Оплата не завершена. Попробуйте снова."}


@app.post("/tribute/webhook")
async def tribute_webhook(request: Request):

    body = await request.body()
    signature = request.headers.get("trbt-signature")

    check = hmac.new(
        TRIBUTE_API_KEY.encode(),
        body,
        hashlib.sha256
    ).hexdigest()

    if signature != check:
        raise HTTPException(status_code=401, detail="Invalid signature")

    data = await request.json()

    if data.get("type") != "shop_order":
        return {"status": "ignored"}

    order = data["payload"]

    if order["status"] != "paid":
        return {"status": "ignored"}

    tribute_id = order["uuid"]

    db: Session = SessionLocal()

    payment = db.query(Payment).filter(
        Payment.external_payment_id == tribute_id
    ).first()

    if not payment:
        return {"status": "payment_not_found"}

    if payment.status == "paid":
        return {"status": "already_processed"}

    payment.status = "paid"

    credit = CreditLedger(
        user_id=payment.user_id,
        operation_type="purchase",
        amount=1,
        source_type="tribute_payment",
        source_id=tribute_id,
        idempotency_key=f"tribute:{tribute_id}"
    )

    db.add(credit)
    db.commit()

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
                "description": "Автоформатирование курсовой по методичке КФУ",
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
        external_payment_id=tribute_id
    )

    db.add(payment)
    db.commit()

    return {"payment_url": payment_url}
