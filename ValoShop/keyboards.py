from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardButton, InlineKeyboardMarkup

kbMain = [[KeyboardButton(text='Профиль'), KeyboardButton(text='Инвентарь')],
          [KeyboardButton(text='Магазин'), KeyboardButton(text='Играть')]]

main = ReplyKeyboardMarkup(keyboard=kbMain, resize_keyboard=True, input_field_placeholder='Выберите пункт ниже')















