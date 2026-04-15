import json
import logging
import os
from pathlib import Path

import httpx
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from db import SessionLocal
from keyboards import (
    BTN_CHECK,
    BTN_FORMAT,
    BTN_REFERRAL,
    BTN_SELECT_GUIDE,
    BTN_TOP_UP_BALANCE,
    CB_ACTION_BUY,
    CB_ACTION_CHECK,
    CB_ACTION_FORMAT,
    CB_ACTION_REFERRAL,
    CB_BACK_TO_MENU,
    CB_CHECK_ANOTHER,
    CB_SELECT_GUIDE_KFU_COURSEWORK_2025,
    CB_SHOW_GUIDE_KFU_COURSEWORK_2025_FILE,
    get_action_inline_keyboard,
    get_back_to_menu_inline_keyboard,
    get_guides_inline_keyboard,
    get_main_menu_keyboard,
    get_no_credits_inline_keyboard,
)
import services


logger = logging.getLogger(__name__)

PAYMENTS_API_BASE_URL = "https://courseworkformatterbot-production.up.railway.app"
PENDING_DOCX_ACTION_KEY = "pending_docx_action"
DOCX_ACTION_CHECK = "check"
DOCX_ACTION_FORMAT = "format"

_ADMIN_IDS: set[int] = {
    int(x) for x in os.getenv("ADMIN_TELEGRAM_IDS", "").split(",") if x.strip()
}


def _is_admin(update: Update, db) -> bool:
    """Проверяет права администратора по Telegram ID ИЛИ по внутреннему user_id."""
    tg_id = update.effective_user.id
    logger.info("admin_check tg_id=%s admin_ids=%s", tg_id, _ADMIN_IDS)
    if tg_id in _ADMIN_IDS:
        return True
    user = services.get_user_by_telegram_id(db, tg_id)
    logger.info("admin_check user=%s user_id=%s", user, getattr(user, "id", None))
    if user and (user.id == 1 or user.id in _ADMIN_IDS):
        return True
    if not _ADMIN_IDS:
        logger.warning("admin_check_failed: ADMIN_TELEGRAM_IDS is empty and user_id is not 1")
    return False


def _extract_referral_code_from_start(text: str | None) -> str | None:
    if not text:
        return None

    parts = text.strip().split(maxsplit=1)
    if len(parts) < 2:
        return None

    payload = parts[1].strip()
    if payload.startswith("ref_") and len(payload) > 4:
        return payload.replace("ref_", "", 1)

    return None


def _get_bot_username(context: ContextTypes.DEFAULT_TYPE) -> str:
    username = ""
    if context.bot and context.bot.username:
        username = context.bot.username

    if not username:
        username = services.get_bot_username_fallback()

    return username


def _ensure_current_user(
    db,
    update: Update,
    referral_code_from_start: str | None = None,
):
    tg_user = update.effective_user
    return services.ensure_user(
        db=db,
        telegram_id=tg_user.id,
        username=tg_user.username,
        first_name=tg_user.first_name,
        last_name=tg_user.last_name,
        referral_code_from_start=referral_code_from_start,
    )


async def start_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    db = SessionLocal()
    try:
        referral_code = _extract_referral_code_from_start(update.message.text)
        if referral_code:
            logger.info(
                "referral_start_detected code=%s telegram_id=%s",
                referral_code,
                update.effective_user.id,
            )
        user, is_new = _ensure_current_user(
            db,
            update,
            referral_code_from_start=referral_code,
        )
        services.track_event(
            db,
            event_name="start_clicked",
            user_id=user.id,
            source="telegram_start",
        )

        balance = services.get_user_credit_balance(db, user.id)
        guide_code = services.get_user_selected_guide_code(user)
        guide = services.get_guide(guide_code)
        referral_progress, referral_target, _ = services.get_referral_upload_bonus_progress(
            db,
            user.id,
        )

        text = services.build_start_text(
            balance=balance,
            is_new=is_new,
            active_guide_title=guide["title"],
            referral_progress=referral_progress,
            referral_target=referral_target,
        )

        await update.message.reply_text(
            text,
            parse_mode="HTML",
            disable_web_page_preview=True,
            reply_markup=get_main_menu_keyboard(),
        )
    finally:
        db.close()


async def balance_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    db = SessionLocal()
    try:
        user, _ = _ensure_current_user(db, update)
        balance = services.get_user_credit_balance(db, user.id)
        bot_username = _get_bot_username(context)

        text = services.build_balance_text(
            user=user,
            balance=balance,
            bot_username=bot_username,
        )

        await update.message.reply_text(
            text,
            reply_markup=get_main_menu_keyboard(),
        )
    finally:
        db.close()


async def referral_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    db = SessionLocal()
    try:
        user, _ = _ensure_current_user(db, update)
        bot_username = _get_bot_username(context)
        balance = services.get_user_credit_balance(db, user.id)
        progress, target, _ = services.get_referral_upload_bonus_progress(db, user.id)
        text = services.build_referral_text(
            bot_username,
            user,
            balance=balance,
            progress=progress,
            target=target,
        )

        await update.message.reply_text(
            text,
            reply_markup=get_main_menu_keyboard(),
        )
    finally:
        db.close()


async def choose_guide_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    db = SessionLocal()
    try:
        user, _ = _ensure_current_user(db, update)
        text = services.build_guide_selection_text(user)

        await update.message.reply_text(
            text,
            reply_markup=get_guides_inline_keyboard(),
        )
    finally:
        db.close()


async def top_up_balance_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    await update.message.reply_text(
        services.build_top_up_balance_text(),
        reply_markup=get_no_credits_inline_keyboard(),
    )


async def _create_payment_and_send_link(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    tariff_code: str,
    button_text: str,
    reply_text: str,
) -> None:
    message = update.message
    if not message:
        return

    db = SessionLocal()
    try:
        user, _ = _ensure_current_user(db, update)
    finally:
        db.close()

    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.post(
            f"{PAYMENTS_API_BASE_URL}/create-payment",
            params={
                "user_id": user.id,
                "tariff_code": tariff_code,
            },
        )

    try:
        data = resp.json()
    except Exception:
        await message.reply_text(
            f"Не удалось создать ссылку на оплату. Код ответа: {resp.status_code}",
            reply_markup=get_main_menu_keyboard(),
        )
        return

    if not data.get("ok"):
        logger.info("create_payment_failed response=%s", data)
        await message.reply_text(
            "Не удалось создать ссылку на оплату.\n"
            f"Техническая причина: {data.get('error', 'unknown_error')}",
            reply_markup=get_main_menu_keyboard(),
        )
        return

    payment_url = data.get("payment_url")
    if not payment_url:
        await message.reply_text(
            "Ссылка на оплату не была получена. Попробуйте ещё раз через минуту.",
            reply_markup=get_main_menu_keyboard(),
        )
        return

    await message.reply_text(
        reply_text,
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton(button_text, url=payment_url)]]
        ),
    )


async def buy1_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _create_payment_and_send_link(
        update=update,
        context=context,
        tariff_code="one_format",
        button_text="💳 Оплатить 200 ₽",
        reply_text="Ссылка на оплату 1 оформления:",
    )


async def buy3_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _create_payment_and_send_link(
        update=update,
        context=context,
        tariff_code="three_formats",
        button_text="📦 Оплатить 500 ₽",
        reply_text="Ссылка на оплату пакета из 3 оформлений:",
    )


async def guide_callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query:
        return

    await query.answer()
    data = query.data or ""

    if data == "buy:one":
        fake_update = Update(update.update_id, message=query.message)
        await buy1_handler(fake_update, context)
        return

    if data == "buy:three":
        fake_update = Update(update.update_id, message=query.message)
        await buy3_handler(fake_update, context)
        return

    if data == CB_BACK_TO_MENU:
        await query.edit_message_text(
            "Главное меню открыто. Можно выбрать действие с обычной клавиатуры ниже.",
            reply_markup=get_back_to_menu_inline_keyboard(),
        )
        return

    if data in {CB_ACTION_CHECK, CB_ACTION_FORMAT, CB_CHECK_ANOTHER, CB_ACTION_BUY, CB_ACTION_REFERRAL}:
        if data == CB_ACTION_BUY:
            await query.message.reply_text(
                services.build_top_up_balance_text(),
                reply_markup=get_no_credits_inline_keyboard(),
            )
            return

        db = SessionLocal()
        try:
            user = services.get_user_by_telegram_id(db, query.from_user.id)
            if not user:
                await query.edit_message_text(
                    "Пользователь не найден. Нажмите /start.",
                    reply_markup=get_back_to_menu_inline_keyboard(),
                )
                return

            if data == CB_ACTION_REFERRAL:
                bot_username = _get_bot_username(context)
                balance = services.get_user_credit_balance(db, user.id)
                progress, target, _ = services.get_referral_upload_bonus_progress(db, user.id)
                text = services.build_referral_text(
                    bot_username,
                    user,
                    balance=balance,
                    progress=progress,
                    target=target,
                )
                reply_markup = get_action_inline_keyboard()
            elif data == CB_CHECK_ANOTHER:
                context.user_data[PENDING_DOCX_ACTION_KEY] = DOCX_ACTION_CHECK
                text = services.build_check_another_text()
                reply_markup = get_action_inline_keyboard()
            elif data == CB_ACTION_CHECK:
                context.user_data[PENDING_DOCX_ACTION_KEY] = DOCX_ACTION_CHECK
                text = services.build_check_selected_text()
                reply_markup = get_action_inline_keyboard()
            else:
                balance = services.get_user_credit_balance(db, user.id)
                if balance > 0:
                    context.user_data[PENDING_DOCX_ACTION_KEY] = DOCX_ACTION_FORMAT
                else:
                    context.user_data.pop(PENDING_DOCX_ACTION_KEY, None)
                text = services.build_format_selected_text(balance)
                reply_markup = get_action_inline_keyboard() if balance > 0 else get_no_credits_inline_keyboard()
                current_message_text = getattr(query.message, "text", "") or ""
                if "Проверка завершена" in current_message_text:
                    services.track_event(
                        db,
                        event_name="check_to_format_clicked",
                        user_id=user.id,
                        source="telegram_callback",
                        payload_json=json.dumps({"balance": balance}),
                    )

            if data == CB_ACTION_CHECK:
                services.track_event(
                    db,
                    event_name="check_selected",
                    user_id=user.id,
                    source="telegram_callback",
                )

            await query.message.reply_text(text, reply_markup=reply_markup)
        finally:
            db.close()
        return

    if data == CB_SELECT_GUIDE_KFU_COURSEWORK_2025:
        db = SessionLocal()
        try:
            user = services.get_user_by_telegram_id(db, query.from_user.id)
            if not user:
                await query.edit_message_text(
                    "Пользователь не найден. Нажмите /start.",
                    reply_markup=get_back_to_menu_inline_keyboard(),
                )
                return

            services.set_user_selected_guide_code(db, user, "kfu_coursework_2025")
            text = services.build_guide_selected_text("kfu_coursework_2025")

            await query.edit_message_text(
                text,
                reply_markup=get_back_to_menu_inline_keyboard(),
            )
        finally:
            db.close()
        return

    if data == CB_SHOW_GUIDE_KFU_COURSEWORK_2025_FILE:
        method_file = services.find_method_file("kfu_coursework_2025")
        if not method_file:
            await query.edit_message_text(
                services.build_method_file_missing_text("kfu_coursework_2025"),
                reply_markup=get_back_to_menu_inline_keyboard(),
            )
            return

        with open(method_file, "rb") as f:
            filename = Path(method_file).name
            if Path(method_file).suffix.lower() not in {".docx", ".pdf"}:
                filename = f"{filename}.pdf"
            await query.message.reply_document(
                document=f,
                filename=filename,
                caption="Файл методички",
            )

        await query.edit_message_text(
            "Файл методички отправлен.",
            reply_markup=get_back_to_menu_inline_keyboard(),
        )
        return


async def docx_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.message.document:
        return

    telegram_document = update.message.document
    filename = telegram_document.file_name or ""

    if not filename.lower().endswith(".docx"):
        await update.message.reply_text(
            "Принимаются только .docx файлы.",
            reply_markup=get_main_menu_keyboard(),
        )
        return

    db = SessionLocal()
    input_path = None
    request = None
    user = None

    try:
        user, _ = _ensure_current_user(db, update)
        balance = services.get_user_credit_balance(db, user.id)
        service_type = context.user_data.pop(PENDING_DOCX_ACTION_KEY, DOCX_ACTION_CHECK)
        if service_type not in {DOCX_ACTION_CHECK, DOCX_ACTION_FORMAT}:
            service_type = DOCX_ACTION_CHECK

        logger.info(
            "docx_received user_id=%s filename=%s balance=%s service_type=%s",
            user.id,
            filename,
            balance,
            service_type,
        )

        if service_type == DOCX_ACTION_FORMAT and balance <= 0:
            text = services.build_no_credits_text(user, _get_bot_username(context))
            await update.message.reply_text(
                text,
                reply_markup=get_no_credits_inline_keyboard(),
            )
            return

        guide_code = services.get_user_selected_guide_code(user)
        _, input_path, _ = services.build_processing_paths(filename)

        logger.info(
            "download_start user_id=%s filename=%s input_path=%s",
            user.id,
            filename,
            input_path,
        )

        tg_file = await telegram_document.get_file()
        await tg_file.download_to_drive(custom_path=str(input_path))

        logger.info(
            "download_done user_id=%s filename=%s",
            user.id,
            filename,
        )

        doc_record = services.create_document_record(
            db,
            user_id=user.id,
            original_filename=filename,
            storage_path=tg_file.file_path or str(input_path),
        )
        services.track_event(
            db,
            event_name="file_uploaded",
            user_id=user.id,
            source="telegram_docx",
            payload_json=json.dumps(
                {
                    "document_id": doc_record.id,
                    "filename": filename,
                    "service_type": service_type,
                }
            ),
        )

        request = services.create_formatting_request(
            db,
            user_id=user.id,
            document_id=doc_record.id,
            guide_code=guide_code,
            service_type=service_type,
        )

        logger.info(
            "request_queued request_id=%s user_id=%s guide=%s service_type=%s",
            request.id,
            user.id,
            guide_code,
            service_type,
        )

        if service_type == DOCX_ACTION_CHECK:
            services.track_event(
                db,
                event_name="check_started",
                user_id=user.id,
                source="telegram_docx",
                payload_json=json.dumps(
                    {"request_id": request.id, "document_id": doc_record.id}
                ),
            )
            await update.message.reply_text(
                services.build_file_received_text("check"),
                reply_markup=get_main_menu_keyboard(),
            )
            return

        debited = services.debit_one_credit(
            db,
            user.id,
            source_id=str(request.id),
        )

        if not debited:
            services.mark_formatting_failed_in_new_session(
                request.id,
                "enqueue failed: not enough credits",
            )

            text = services.build_no_credits_text(user, _get_bot_username(context))
            await update.message.reply_text(
                text,
                reply_markup=get_no_credits_inline_keyboard(),
            )
            return

        logger.info(
            "credit_debited request_id=%s user_id=%s",
            request.id,
            user.id,
        )

        await update.message.reply_text(
            services.build_file_received_text("format"),
            reply_markup=get_main_menu_keyboard(),
        )

    except Exception as e:
        logger.exception(
            "queue_enqueue_failed request_id=%s filename=%s",
            getattr(request, "id", None),
            filename,
        )

        if request is not None:
            services.mark_formatting_failed_in_new_session(
                request.id,
                f"enqueue failed: {str(e)}",
            )

        if request is not None and user is not None:
            services.refund_one_credit_in_new_session(
                user.id,
                str(request.id),
            )

        if update.message:
            await update.message.reply_text(
                f"Не удалось поставить документ в очередь: {str(e)[:300]}",
                reply_markup=get_main_menu_keyboard(),
            )

    finally:
        db.close()


async def userinfo_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    db = SessionLocal()
    try:
        user, _ = _ensure_current_user(db, update)
        text = services.get_userinfo_text(db, user)
        await update.message.reply_text(
            text,
            reply_markup=get_main_menu_keyboard(),
        )
    finally:
        db.close()


async def give_credits_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    db = SessionLocal()
    try:
        if not _is_admin(update, db):
            return

        if len(context.args) != 2:
            await update.message.reply_text("Формат: /givecredits user_id amount")
            return

        try:
            target_user_id = int(context.args[0])
            amount = int(context.args[1])
        except ValueError:
            await update.message.reply_text("Формат: /givecredits user_id amount")
            return

        balance = services.grant_admin_bonus(
            db,
            target_user_id=target_user_id,
            amount=amount,
            admin_source_id=str(update.effective_user.id),
        )
        await update.message.reply_text(
            f"Начислено {amount} оформлений пользователю {target_user_id}\n"
            f"Текущий баланс: {balance}"
        )
    finally:
        db.close()


async def markpaid_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    db = SessionLocal()
    try:
        if not _is_admin(update, db):
            return

        if len(context.args) != 2:
            await update.message.reply_text("Формат: /markpaid user_id credits")
            return

        try:
            paid_user_id = int(context.args[0])
            credits = int(context.args[1])
        except ValueError:
            await update.message.reply_text("Формат: /markpaid user_id credits")
            return

        services.apply_successful_payment(
            db,
            paid_user_id=paid_user_id,
            credits=credits,
            provider="manual",
        )
        balance = services.get_user_credit_balance(db, paid_user_id)
        await update.message.reply_text(
            f"Оплата отмечена. Баланс пользователя {paid_user_id}: {balance}"
        )
    finally:
        db.close()


async def text_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message or not update.message.text:
        return

    text = update.message.text.strip()

    if text == BTN_CHECK:
        context.user_data[PENDING_DOCX_ACTION_KEY] = DOCX_ACTION_CHECK
        await update.message.reply_text(
            services.build_check_selected_text(),
            reply_markup=get_main_menu_keyboard(),
        )
        return

    if text == BTN_FORMAT:
        db = SessionLocal()
        try:
            user, _ = _ensure_current_user(db, update)
            balance = services.get_user_credit_balance(db, user.id)
            if balance > 0:
                context.user_data[PENDING_DOCX_ACTION_KEY] = DOCX_ACTION_FORMAT
                reply_markup = get_main_menu_keyboard()
            else:
                context.user_data.pop(PENDING_DOCX_ACTION_KEY, None)
                reply_markup = get_no_credits_inline_keyboard()
            await update.message.reply_text(
                services.build_format_selected_text(balance),
                reply_markup=reply_markup,
            )
        finally:
            db.close()
        return

    if text == BTN_TOP_UP_BALANCE:
        await top_up_balance_handler(update, context)
        return

    if text == BTN_REFERRAL:
        await referral_handler(update, context)
        return

    if text == BTN_SELECT_GUIDE:
        await choose_guide_handler(update, context)
        return

    await update.message.reply_text(
        services.build_text_fallback_text(),
        reply_markup=get_action_inline_keyboard(),
    )


def register_handlers(app: Application) -> None:
    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("balance", balance_handler))
    app.add_handler(CommandHandler("referral", referral_handler))
    app.add_handler(CommandHandler("buy1", buy1_handler))
    app.add_handler(CommandHandler("buy3", buy3_handler))
    app.add_handler(CommandHandler("userinfo", userinfo_handler))
    app.add_handler(CommandHandler("user_info", userinfo_handler))
    app.add_handler(CommandHandler("method", choose_guide_handler))
    app.add_handler(CommandHandler("givecredits", give_credits_handler))
    app.add_handler(CommandHandler("give_credits", give_credits_handler))
    app.add_handler(CommandHandler("markpaid", markpaid_handler))

    app.add_handler(
        CallbackQueryHandler(
            guide_callback_handler,
            pattern=r"^(guide:|guide_file:|menu:back|buy:|action:|check:)",
        )
    )
    app.add_handler(MessageHandler(filters.Document.FileExtension("docx"), docx_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_menu_handler))
