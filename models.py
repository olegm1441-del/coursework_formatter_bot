from datetime import datetime

from sqlalchemy import (
    BigInteger,
    DateTime,
    ForeignKey,
    Integer,
    String,
    Text,
)
from sqlalchemy.orm import relationship

from db import Base


class User(Base):
    __tablename__ = "users"

    id = Integer(primary_key=True)
    telegram_id = BigInteger(unique=True, nullable=False, index=True)
    username = String(255, nullable=True)
    first_name = String(255, nullable=True)
    last_name = String(255, nullable=True)
    referral_code = String(64, unique=True, nullable=False, index=True)
    referred_by_user_id = ForeignKey("users.id", ondelete="SET NULL")
    created_at = DateTime(default=datetime.utcnow, nullable=False)


class Referral(Base):
    __tablename__ = "referrals"

    id = Integer(primary_key=True)
    inviter_user_id = ForeignKey("users.id", ondelete="CASCADE")
    invited_user_id = ForeignKey("users.id", ondelete="CASCADE")
    referral_code = String(64, nullable=False)
    linked_at = DateTime(default=datetime.utcnow, nullable=False)
    qualified_upload_at = DateTime(nullable=True)
    first_payment_at = DateTime(nullable=True)


class Payment(Base):
    __tablename__ = "payments"

    id = Integer(primary_key=True)
    user_id = ForeignKey("users.id", ondelete="CASCADE")
    provider = String(100, nullable=False)
    tariff_code = String(100, nullable=False)
    amount_rub = Integer(nullable=False)
    status = String(50, nullable=False)
    external_payment_id = String(255, nullable=True)
    created_at = DateTime(default=datetime.utcnow, nullable=False)
    paid_at = DateTime(nullable=True)


class CreditLedger(Base):
    __tablename__ = "credit_ledger"

    id = Integer(primary_key=True)
    user_id = ForeignKey("users.id", ondelete="CASCADE")
    operation_type = String(100, nullable=False)
    amount = Integer(nullable=False)
    source_type = String(100, nullable=True)
    source_id = String(100, nullable=True)
    idempotency_key = String(255, unique=True, nullable=False)
    created_at = DateTime(default=datetime.utcnow, nullable=False)


class Document(Base):
    __tablename__ = "documents"

    id = Integer(primary_key=True)
    user_id = ForeignKey("users.id", ondelete="CASCADE")
    original_filename = String(255, nullable=False)
    storage_path = Text(nullable=False)
    created_at = DateTime(default=datetime.utcnow, nullable=False)


class FormattingRequest(Base):
    __tablename__ = "formatting_requests"

    id = Integer(primary_key=True)
    user_id = ForeignKey("users.id", ondelete="CASCADE")
    document_id = ForeignKey("documents.id", ondelete="CASCADE")
    service_type = String(50, nullable=False, default="format")
    university_code = String(50, nullable=False, default="kfu")
    document_type = String(50, nullable=False, default="coursework")
    guideline_version = String(50, nullable=False, default="2025")
    status = String(50, nullable=False, default="created")
    result_file_path = Text(nullable=True)
    error_message = Text(nullable=True)
    created_at = DateTime(default=datetime.utcnow, nullable=False)
    completed_at = DateTime(nullable=True)
