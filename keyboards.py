from telegram import KeyboardButton, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup

BTN_REFERRAL = "Реферальная ссылка"
BTN_TOP_UP_BALANCE = "Купить оформления"
BTN_SELECT_GUIDE = "Выбрать методичку"
BTN_CONTACT = "Контакт"

BTN_GUIDE_KFU_COURSEWORK_2025 = "КФУ — курсовая 2025"
BTN_SHOW_GUIDE_FILE = "Открыть файл методички"
BTN_BACK_TO_MENU = "Назад"

BTN_BUY_1 = "💳 1 оформление — 200 ₽"
BTN_BUY_3 = "📦 3 оформления — 500 ₽"

CB_SELECT_GUIDE_KFU_COURSEWORK_2025 = "guide:kfu_coursework_2025"
CB_SHOW_GUIDE_KFU_COURSEWORK_2025_FILE = "guide_file:kfu_coursework_2025"
CB_BACK_TO_MENU = "menu:back"


def get_main_menu_keyboard() -> ReplyKeyboardMarkup:
    keyboard = [
        [KeyboardButton(BTN_REFERRAL)],
        [KeyboardButton(BTN_TOP_UP_BALANCE)],
        [KeyboardButton(BTN_SELECT_GUIDE)],
        [KeyboardButton(BTN_CONTACT)],
    ]
    return ReplyKeyboardMarkup(
        keyboard=keyboard,
        resize_keyboard=True,
        one_time_keyboard=False,
        selective=False,
    )


def get_guides_inline_keyboard() -> InlineKeyboardMarkup:
    keyboard = [
        [InlineKeyboardButton(BTN_GUIDE_KFU_COURSEWORK_2025, callback_data=CB_SELECT_GUIDE_KFU_COURSEWORK_2025)],
        [InlineKeyboardButton(BTN_SHOW_GUIDE_FILE, callback_data=CB_SHOW_GUIDE_KFU_COURSEWORK_2025_FILE)],
        [InlineKeyboardButton(BTN_BACK_TO_MENU, callback_data=CB_BACK_TO_MENU)],
    ]
    return InlineKeyboardMarkup(keyboard)


def get_top_up_balance_inline_keyboard() -> InlineKeyboardMarkup:
    keyboard = [
        [InlineKeyboardButton(BTN_BUY_1, callback_data="buy:one")],
        [InlineKeyboardButton(BTN_BUY_3, callback_data="buy:three")],
        [InlineKeyboardButton(BTN_BACK_TO_MENU, callback_data=CB_BACK_TO_MENU)],
    ]
    return InlineKeyboardMarkup(keyboard)


def get_back_to_menu_inline_keyboard() -> InlineKeyboardMarkup:
    keyboard = [
        [InlineKeyboardButton(BTN_BACK_TO_MENU, callback_data=CB_BACK_TO_MENU)],
    ]
    return InlineKeyboardMarkup(keyboard)
