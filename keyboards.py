from telegram import KeyboardButton, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup

# Тексты кнопок главного меню
BTN_BALANCE = "Баланс"
BTN_SELECT_GUIDE = "Выбрать методичку"
BTN_CONTACT = "Контакт"

# Тексты inline-кнопок
BTN_GUIDE_KFU_COURSEWORK_2025 = "КФУ — курсовая 2025"
BTN_SHOW_GUIDE_FILE = "Открыть файл методички"
BTN_BACK_TO_MENU = "Назад"

# Callback data
CB_SELECT_GUIDE_KFU_COURSEWORK_2025 = "guide:kfu_coursework_2025"
CB_SHOW_GUIDE_KFU_COURSEWORK_2025_FILE = "guide_file:kfu_coursework_2025"
CB_BACK_TO_MENU = "menu:back"


def get_main_menu_keyboard() -> ReplyKeyboardMarkup:
    """
    Главное меню бота.
    .docx пользователь присылает просто файлом, без отдельной кнопки.
    """
    keyboard = [
        [KeyboardButton(BTN_BALANCE)],
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
    """
    Клавиатура выбора методички.
    Пока доступен один вариант, но структура уже готова к масштабированию.
    """
    keyboard = [
        [InlineKeyboardButton(BTN_GUIDE_KFU_COURSEWORK_2025, callback_data=CB_SELECT_GUIDE_KFU_COURSEWORK_2025)],
        [InlineKeyboardButton(BTN_SHOW_GUIDE_FILE, callback_data=CB_SHOW_GUIDE_KFU_COURSEWORK_2025_FILE)],
        [InlineKeyboardButton(BTN_BACK_TO_MENU, callback_data=CB_BACK_TO_MENU)],
    ]
    return InlineKeyboardMarkup(keyboard)


def get_back_to_menu_inline_keyboard() -> InlineKeyboardMarkup:
    """
    Универсальная inline-кнопка 'Назад'.
    """
    keyboard = [
        [InlineKeyboardButton(BTN_BACK_TO_MENU, callback_data=CB_BACK_TO_MENU)],
    ]
    return InlineKeyboardMarkup(keyboard)
