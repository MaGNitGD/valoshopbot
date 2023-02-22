import time
import logging
import datetime

import aiogram
import openpyxl

from aiogram import Bot, Dispatcher, executor, types


API_TOKEN = '5819234032:AAFxPINrKJkPokcnwpI2hPDk_bacvLyUhnc' #токен бота
logging.basicConfig(level=logging.INFO) #сбор логов (куда?)

bot = Bot(token=API_TOKEN)
disp = Dispatcher(bot=bot)

def findCell(value, column, raw):
    bd = openpyxl.load_workbook('users.xlsx')  # открываю бд
    ws = bd['usersInfo']  # выбираю лист usersInfo
    i = 1  # образ raw в цикле
    currentCell = column + str(raw) # определяю проверяему ячейку
    currentCellValue = ws[currentCell].value # определяю значение проверяемой ячейки
    none = False # пустая ли ячейка?
    x = 0  # отключение while
    while x != 1:
        raw = str(i)  # преобразую i в номер строки
        if ws[currentCell].value == value: # если ячейка найдена
            x = 1 # отключение while
            return [currentCell, currentCellValue] # координаты ячейки и её значение
        elif ws[currentCell].value is None: # если ячейки нет
            none = True
            return [currentCell, none] # координаты ячеёки и ин-я, что она пуста
        else:
            i += 1




@disp.message_handler(commands=['start', 'старт']) #обработчик команды /start
async def cmdStart(message: types.Message):
        bd = openpyxl.load_workbook('users.xlsx')  # открываю бд
        ws = bd['usersInfo']  # выбираю лист usersInfo
        regtime = f'{datetime.datetime.now().day}.{datetime.datetime.now().month}.{datetime.datetime.now().year}' # дата рег-ии [17 02 2023]

        tgidCell = findCell(message.from_user.username, 'A', 2)
        botidCell = 'B' + tgidCell[0][1]
        regtimeCell = 'C' + tgidCell[0][1]
        if tgidCell[1] == message.from_user.username:
            await message.reply(f'Вы уже зарегистрированы как {tgidCell[1]}')
        elif tgidCell[1] == True:
            ws[tgidCell[0]] = message.from_user.username # установка tgid
            ws[botidCell] = str(int(tgidCell[0][1])-1) + ws[tgidCell[0]].value[0] # установка botid формата [номер][первая буква]
            ws[regtimeCell] = regtime
            bd.save('users.xlsx') # сохранение бд

            await message.reply(f'Добро пожаловать в ValoShop!\nВы успешно зарегистрировались как {ws[tgidCell[0]].value}\n\nПомощь - /help')

@disp.message_handler(commands=['profile', 'профиль']) # обработка команды /profile
async def cmdProfile(message: types.Message):
    bd = openpyxl.load_workbook('users.xlsx')  # открываю бд
    ws = bd['usersInfo']  # выбираю лист usersInfo
    tgid = findCell(message.from_user.username, 'A', 2) # поиск юзера в базе
    regtime = 'C' + tgid[0][1] # ячейка для даты
    if tgid[1] == True: # если юзер найден
        await message.reply('Вы ещё не зарегистрированы. Для регистрации введите /start')
    else: # если юзер не зареган
        await message.reply(f'Профиль пользователя @{tgid[1]}:\n\nДата регистрации: {ws[regtime].value}')

@disp.message_handler(commands=['help', 'помощь', 'команды', 'cmd']) # обработка команды /help
async def cmdHelp(message: types.Message):
    await message.reply(f'Команды бота:\n\n/start - Регистрация\n/profile - Профиль\n/help - Помощь\n/inv - инвентарь\n{"_"*23}\nСоздатель: @magnitgd')

@disp.message_handler(commands=['inv', 'inventory', 'инвентарь']) # обработка команды /inventory
async def cmdInv(message: types.Message):
    a = 1



if __name__ == '__main__':
    executor.start_polling(disp)
