import logging
from pathlib import Path

from telegram import Update
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
    BTN_BALANCE,
    BTN_CONTACT,
    BTN_SELECT_GUIDE,
    CB_BACK_TO_MENU,
    CB_SELECT_GUIDE_KFU_COURSEWORK_2025,
    CB_SHOW_GUIDE_KFU_COURSEWORK_2025_FILE,
    get_back_to_menu_inline_keyboard,
    get_guides_inline_keyboard,
    get_main_menu_keyboard,
)
import services


logger = logging.getLogger(__name__)


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

        text = services.build_start_text(
            balance=balance,
            is_new=is_new,
            active_guide_title=guide["title"],
        )

        await update.message.reply_text(
            text,
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
        text = services.build_referral_text(bot_username, user)

        await update.message.reply_text(
            text,
            reply_markup=get_main_menu_keyboard(),
        )
    finally:
        db.close()


async def contact_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
        return

    await update.message.reply_text(
        services.get_contact_text(),
        reply_markup=get_main_menu_keyboard(),
    )


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


async def guide_callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query:
        return

    await query.answer()
    data = query.data or ""

    if data == CB_BACK_TO_MENU:
        await query.edit_message_text(
            "Главное меню открыто. Можно выбрать действие с обычной клавиатуры ниже.",
            reply_markup=get_back_to_menu_inline_keyboard(),
        )
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
            await query.message.reply_document(
                document=f,
                filename=Path(method_file).name,
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

        logger.info(
            "docx_received user_id=%s filename=%s balance=%s",
            user.id,
            filename,
            balance,
        )

        if balance <= 0:
            text = services.build_no_credits_text(user, _get_bot_username(context))
            await update.message.reply_text(
                text,
                reply_markup=get_main_menu_keyboard(),
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
            storage_path=str(input_path),
        )
        services.track_event(
            db,
            event_name="file_uploaded",
            user_id=user.id,
            source="telegram_docx",
            payload_json=f'{{"document_id": {doc_record.id}, "filename": "{filename}"}}',
        )

        request = services.create_formatting_request(
            db,
            user_id=user.id,
            document_id=doc_record.id,
            guide_code=guide_code,
        )

        logger.info(
            "request_queued request_id=%s user_id=%s guide=%s",
            request.id,
            user.id,
            guide_code,
        )

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
                reply_markup=get_main_menu_keyboard(),
            )
            return

        logger.info(
            "credit_debited request_id=%s user_id=%s",
            request.id,
            user.id,
        )

        await update.message.reply_text(
            (
                "Документ принят в очередь на оформление.\n"
                f"Номер заявки: {request.id}\n\n"
                "Когда обработка завершится, результат будет отправлен автоматически."
            ),
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

    if len(context.args) != 2:
        await update.message.reply_text("Формат: /givecredits user_id amount")
        return

    try:
        target_user_id = int(context.args[0])
        amount = int(context.args[1])
    except ValueError:
        await update.message.reply_text("Формат: /givecredits user_id amount")
        return

    db = SessionLocal()
    try:
        balance = services.grant_admin_bonus(
            db,
            target_user_id=target_user_id,
            amount=amount,
            admin_source_id=str(update.effective_user.id),
        )
        await update.message.reply_text(
            f"Готово. Баланс пользователя {target_user_id}: {balance}"
        )
    finally:
        db.close()


async def markpaid_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not update.message:
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

    db = SessionLocal()
    try:
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

    if text == BTN_BALANCE:
        await balance_handler(update, context)
        return

    if text == BTN_SELECT_GUIDE:
        await choose_guide_handler(update, context)
        return

    if text == BTN_CONTACT:
        await contact_handler(update, context)
        return

    await update.message.reply_text(
        "Можно нажать кнопку в меню или отправить .docx-файл на обработку.",
        reply_markup=get_main_menu_keyboard(),
    )


def register_handlers(app: Application) -> None:
    app.add_handler(CommandHandler("start", start_handler))
    app.add_handler(CommandHandler("balance", balance_handler))
    app.add_handler(CommandHandler("referral", referral_handler))
    app.add_handler(CommandHandler("userinfo", userinfo_handler))
    app.add_handler(CommandHandler("user_info", userinfo_handler))
    app.add_handler(CommandHandler("contact", contact_handler))
    app.add_handler(CommandHandler("method", choose_guide_handler))
    app.add_handler(CommandHandler("givecredits", give_credits_handler))
    app.add_handler(CommandHandler("give_credits", give_credits_handler))
    app.add_handler(CommandHandler("markpaid", markpaid_handler))

    app.add_handler(
        CallbackQueryHandler(
            guide_callback_handler,
            pattern=r"^(guide:|guide_file:|menu:back)",
        )
    )
    app.add_handler(MessageHandler(filters.Document.FileExtension("docx"), docx_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_menu_handler))
