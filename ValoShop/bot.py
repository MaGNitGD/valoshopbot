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
    uinf = bd['usersInfo']  # выбираю лист usersInfo
    i = raw  # образ raw в цикле
    currentCell = column + str(raw) # определяю проверяему ячейку
    currentCellValue = uinf[currentCell].value # определяю значение проверяемой ячейки
    none = False # пустая ли ячейка?
    x = 0  # отключение while
    while x != 1:
        raw = str(i)  # преобразую i в номер строки
        currentCell = column + str(raw) # определяю проверяемую ячейку
        currentCellValue = uinf[currentCell].value  # определяю значение проверяемой ячейки
        if currentCellValue == value: # если ячейка найдена
            x = 1 # отключение while
            bd.save('users.xlsx')
            return [currentCell, currentCellValue] # координаты ячейки и её значение
        elif uinf[currentCell].value is None: # если ячейки нет
            none = True
            bd.save('users.xlsx')
            return [currentCell, none] # координаты ячеёки и ин-я, что она пуста
        else:
            i += 1

def setDefaultSkins(botId):
    botidCell = findCell(botId, 'A', 1)
    bd = openpyxl.load_workbook('users.xlsx')
    ueqiup = bd['usersEquipped']
    ueqiup['B' + str(int(botidCell[0][1]) - 1)] = 'ke0'
    ueqiup['C' + str(int(botidCell[0][1]) - 1)] = 'сс0'
    ueqiup['D' + str(int(botidCell[0][1]) - 1)] = 'sy0'
    ueqiup['E' + str(int(botidCell[0][1]) - 1)] = 'fy0'
    ueqiup['F' + str(int(botidCell[0][1]) - 1)] = 'gt0'
    ueqiup['G' + str(int(botidCell[0][1]) - 1)] = 'sf0'
    ueqiup['H' + str(int(botidCell[0][1]) - 1)] = 'by0'
    ueqiup['I' + str(int(botidCell[0][1]) - 1)] = 'ml0'
    ueqiup['J' + str(int(botidCell[0][1]) - 1)] = 'sr0'
    ueqiup['K' + str(int(botidCell[0][1]) - 1)] = 'as0'
    ueqiup['L' + str(int(botidCell[0][1]) - 1)] = 'se0'
    ueqiup['M' + str(int(botidCell[0][1]) - 1)] = 'je0'
    ueqiup['N' + str(int(botidCell[0][1]) - 1)] = 'bg0'
    ueqiup['O' + str(int(botidCell[0][1]) - 1)] = 'gn0'
    ueqiup['P' + str(int(botidCell[0][1]) - 1)] = 'pm0'
    ueqiup['Q' + str(int(botidCell[0][1]) - 1)] = 'vl0'
    ueqiup['R' + str(int(botidCell[0][1]) - 1)] = 'on0'
    ueqiup['S' + str(int(botidCell[0][1]) - 1)] = 'or0'
    bd.save('users.xlsx')

@disp.message_handler(commands=['start', 'старт']) #обработчик команды /start
async def cmdStart(message: types.Message):
        bd = openpyxl.load_workbook('users.xlsx')  # открываю бд
        uinf = bd['usersInfo']  # выбираю лист usersInfo
        regtime = f'{datetime.datetime.now().day}.{datetime.datetime.now().month}.{datetime.datetime.now().year}' # дата рег-ции [17 02 2023]


        tgidCell = findCell(message.from_user.id, 'B', 2) # формат переменной [координаты ячейки, содержимое]
        tgidCell[0] = 'A' + tgidCell[0][1]
        botidCell = 'B' + (tgidCell[0][1])  # координаты ботид
        username = message.from_user.username
        if message.from_user.username is None:
            username = message.from_user.first_name + str(int(tgidCell[0][1]) - 1)
        if uinf[botidCell].value == message.from_user.id: # если уже зарегестрирован
            await message.reply(f'Вы уже зарегистрированы как {uinf["A" + tgidCell[0][1]].value}')
        elif tgidCell[1] == True: # регистрация

            botidCell = 'B' + tgidCell[0][1] # координаты ботид
            regtimeCell = 'C' + tgidCell[0][1] # координаты времени
            uinf[tgidCell[0]] = username # установка tgid
            uinf[botidCell] = message.from_user.id # установка botid формата [номер][первая буква]
            uinf[regtimeCell] = regtime

            uequip = bd['usersEquipped']
            botidEquip = findCell(uinf[tgidCell[0]].value, 'A', 2)
            uequip[botidEquip[0]] = uinf[botidCell].value
            bd.save('users.xlsx') # сохранение бд
            setDefaultSkins(uinf[botidCell].value)

            await message.reply(f'Добро пожаловать в ValoShop!\nВы успешно зарегистрировались как {uinf[tgidCell[0]].value}\n\nПомощь - /help')

@disp.message_handler(commands=['profile', 'профиль']) # обработка команды /profile
async def cmdProfile(message: types.Message):
    bd = openpyxl.load_workbook('users.xlsx')  # открываю бд
    uinf = bd['usersInfo']  # выбираю лист usersInfo
    botid = findCell(message.from_user.id, 'B', 2) # поиск юзера в базе
    tgid = uinf['A' + botid[0][1]].value
    regtime = 'C' + botid[0][1] # ячейка для даты
    if botid[1] == True: # если юзер не зареган
        await message.reply('Вы ещё не зарегистрированы. Для регистрации введите /start')
    else: # если юзер найден
        await message.reply(f'Профиль пользователя @{tgid}:\n\nДата регистрации: {uinf[regtime].value}')

@disp.message_handler(commands=['help', 'помощь', 'команды', 'cmd']) # обработка команды /help
async def cmdHelp(message: types.Message):
    await message.reply(f'Команды бота:\n\n/start - Регистрация\n/profile - Профиль\n/help - Помощь\n/inv - инвентарь\n{"_"*23}\nСоздатель: @magnitgd')

@disp.message_handler(commands=['inv', 'inventory', 'инвентарь']) # обработка команды /inventory
async def cmdInv(message: types.Message):
    a = 1



if __name__ == '__main__':
    executor.start_polling(disp)
