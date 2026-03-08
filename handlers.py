from aiogram import types, Router, F
from aiogram.filters import CommandStart, Command

from db import SessionLocal
from keyboards import main_menu_keyboard, guides_keyboard
import services

router = Router()


# =========================
# /start
# =========================

@router.message(CommandStart())
async def cmd_start(message: types.Message):

    start_payload = message.text.replace("/start", "").strip()

    referral_code = None
    if start_payload.startswith("ref_"):
        referral_code = start_payload.replace("ref_", "")

    db = SessionLocal()

    user, is_new = services.ensure_user(
        db=db,
        telegram_id=message.from_user.id,
        username=message.from_user.username,
        first_name=message.from_user.first_name,
        last_name=message.from_user.last_name,
        referral_code_from_start=referral_code,
    )

    balance = services.get_user_credit_balance(db, user.id)

    guide_code = services.get_user_selected_guide_code(user)
    guide = services.get_guide(guide_code)

    text = services.build_start_text(
        balance=balance,
        is_new=is_new,
        active_guide_title=guide["title"],
    )

    await message.answer(
        text,
        reply_markup=main_menu_keyboard()
    )

    db.close()


# =========================
# Баланс
# =========================

@router.message(F.text == "Баланс")
async def balance_handler(message: types.Message):

    db = SessionLocal()

    user = services.get_user_by_telegram_id(db, message.from_user.id)

    if not user:
        await message.answer("Ошибка пользователя")
        return

    balance = services.get_user_credit_balance(db, user.id)

    bot_username = services.get_bot_username_fallback()

    text = services.build_balance_text(
        user=user,
        balance=balance,
        bot_username=bot_username
    )

    await message.answer(text)

    db.close()


# =========================
# Контакт
# =========================

@router.message(F.text == "Контакт")
async def contact_handler(message: types.Message):

    await message.answer(
        services.get_contact_text()
    )


# =========================
# Выбор методички
# =========================

@router.message(F.text == "Выбрать методичку")
async def choose_guide_handler(message: types.Message):

    db = SessionLocal()

    user = services.get_user_by_telegram_id(db, message.from_user.id)

    text = services.build_guide_selection_text(user)

    await message.answer(
        text,
        reply_markup=guides_keyboard()
    )

    db.close()


@router.callback_query(F.data.startswith("guide:"))
async def guide_selected(callback: types.CallbackQuery):

    guide_code = callback.data.split(":")[1]

    db = SessionLocal()

    user = services.get_user_by_telegram_id(db, callback.from_user.id)

    services.set_user_selected_guide_code(
        db,
        user,
        guide_code
    )

    text = services.build_guide_selected_text(guide_code)

    await callback.message.edit_text(text)

    await callback.answer()

    db.close()


# =========================
# DOCX обработка
# =========================

@router.message(F.document)
async def docx_handler(message: types.Message):

    document = message.document

    if not document.file_name.endswith(".docx"):
        await message.answer("Принимаются только .docx файлы")
        return

    db = SessionLocal()

    user = services.get_user_by_telegram_id(db, message.from_user.id)

    balance = services.get_user_credit_balance(db, user.id)

    if balance <= 0:

        text = services.build_no_credits_text(
            user,
            services.get_bot_username_fallback()
        )

        await message.answer(text)

        db.close()
        return

    guide_code = services.get_user_selected_guide_code(user)

    safe_name, input_path, output_path = services.build_processing_paths(
        document.file_name
    )

    file = await message.bot.get_file(document.file_id)

    await message.bot.download_file(
        file.file_path,
        input_path
    )

    doc = services.create_document_record(
        db,
        user_id=user.id,
        original_filename=document.file_name,
        storage_path=str(input_path)
    )

    request = services.create_formatting_request(
        db,
        user_id=user.id,
        document_id=doc.id,
        guide_code=guide_code
    )

    services.debit_one_credit(
        db,
        user.id,
        source_id=str(request.id)
    )

    try:

        services.format_document_by_guide(
            guide_code,
            str(input_path),
            str(output_path)
        )

        services.mark_formatting_done(
            db,
            request.id,
            str(output_path)
        )

        services.grant_referral_upload_bonus_if_needed(
            db,
            user.id
        )

        await message.answer_document(
            types.FSInputFile(output_path),
            caption="Документ оформлен"
        )

    except Exception as e:

        services.mark_formatting_failed_in_new_session(
            request.id,
            str(e)
        )

        services.refund_one_credit_in_new_session(
            user.id,
            str(request.id)
        )

        await message.answer(
            "Ошибка форматирования документа"
        )

    finally:

        services.cleanup_temp_files(
            input_path,
            output_path
        )

        db.close()


# =========================
# ADMIN команды
# =========================

@router.message(Command("user_info"))
async def user_info(message: types.Message):

    db = SessionLocal()

    user = services.get_user_by_telegram_id(
        db,
        message.from_user.id
    )

    text = services.get_userinfo_text(db, user)

    await message.answer(text)

    db.close()


@router.message(Command("give_credits"))
async def give_credits(message: types.Message):

    try:
        _, user_id, amount = message.text.split()

        user_id = int(user_id)
        amount = int(amount)

    except:
        await message.answer("Формат: /give_credits user_id amount")
        return

    db = SessionLocal()

    balance = services.grant_admin_bonus(
        db,
        user_id,
        amount,
        admin_source_id=str(message.from_user.id)
    )

    await message.answer(
        f"Баланс пользователя {user_id}: {balance}"
    )

    db.close()
