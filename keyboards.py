from telegram import KeyboardButton, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup

BTN_REFERRAL = "Реферальная ссылка"
BTN_TOP_UP_BALANCE = "Купить оформления"
BTN_SELECT_GUIDE = "Выбрать методичку"

BTN_GUIDE_KFU_COURSEWORK_2025 = "КФУ — курсовая 2025"
BTN_SHOW_GUIDE_FILE = "Открыть файл методички"
BTN_BACK_TO_MENU = "Назад"
BTN_CHECK = "🔍 Проверить оформление"
BTN_FORMAT = "✍️ Оформить работу"
BTN_FORMAT_THIS = "✍️ Оформить эту работу"
BTN_CHECK_ANOTHER = "🔍 Проверить другой файл"
BTN_BUY_MENU = "💳 Купить оформления"
BTN_REFERRAL_LINK = "🔗 Реферальная ссылка"
BTN_REFERRAL_UNLOCK = "Получить оформление за друзей"

BTN_BUY_1 = "💳 1 оформление — 200 ₽"
BTN_BUY_3 = "📦 3 оформления — 500 ₽"

CB_SELECT_GUIDE_KFU_COURSEWORK_2025 = "guide:kfu_coursework_2025"
CB_SHOW_GUIDE_KFU_COURSEWORK_2025_FILE = "guide_file:kfu_coursework_2025"
CB_BACK_TO_MENU = "menu:back"
CB_ACTION_CHECK = "action:check"
CB_ACTION_FORMAT = "action:format"
CB_ACTION_BUY = "action:buy"
CB_ACTION_REFERRAL = "action:referral"
CB_CHECK_ANOTHER = "check:another"


def get_main_menu_keyboard() -> ReplyKeyboardMarkup:
    keyboard = [
        [KeyboardButton(BTN_CHECK), KeyboardButton(BTN_FORMAT)],
        [KeyboardButton(BTN_TOP_UP_BALANCE)],
        [KeyboardButton(BTN_REFERRAL)],
        [KeyboardButton(BTN_SELECT_GUIDE)],
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


def get_action_inline_keyboard() -> InlineKeyboardMarkup:
    keyboard = [
        [InlineKeyboardButton(BTN_CHECK, callback_data=CB_ACTION_CHECK)],
        [InlineKeyboardButton(BTN_FORMAT, callback_data=CB_ACTION_FORMAT)],
        [InlineKeyboardButton(BTN_BUY_MENU, callback_data=CB_ACTION_BUY)],
        [InlineKeyboardButton(BTN_REFERRAL_LINK, callback_data=CB_ACTION_REFERRAL)],
    ]
    return InlineKeyboardMarkup(keyboard)


def get_no_credits_inline_keyboard() -> InlineKeyboardMarkup:
    keyboard = [
        [InlineKeyboardButton(BTN_CHECK, callback_data=CB_ACTION_CHECK)],
        [InlineKeyboardButton(BTN_BUY_1, callback_data="buy:one")],
        [InlineKeyboardButton(BTN_BUY_3, callback_data="buy:three")],
    ]
    return InlineKeyboardMarkup(keyboard)


def get_check_result_inline_keyboard() -> InlineKeyboardMarkup:
    keyboard = [
        [InlineKeyboardButton(BTN_FORMAT_THIS, callback_data=CB_ACTION_FORMAT)],
        [InlineKeyboardButton(BTN_CHECK_ANOTHER, callback_data=CB_CHECK_ANOTHER)],
        [InlineKeyboardButton(BTN_BUY_MENU, callback_data=CB_ACTION_BUY)],
        [InlineKeyboardButton(BTN_REFERRAL_LINK, callback_data=CB_ACTION_REFERRAL)],
    ]
    return InlineKeyboardMarkup(keyboard)


def get_back_to_menu_inline_keyboard() -> InlineKeyboardMarkup:
    keyboard = [
        [InlineKeyboardButton(BTN_BACK_TO_MENU, callback_data=CB_BACK_TO_MENU)],
    ]
    return InlineKeyboardMarkup(keyboard)
