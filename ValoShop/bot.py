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

def findCell(value, column, raw, bdin, uinfin):
    bd = openpyxl.load_workbook(bdin)  # открываю бд
    uinf = bd[uinfin]  # выбираю лист usersInfo
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
            bd.save(bdin)
            return [currentCell, currentCellValue] # координаты ячейки и её значение
        elif uinf[currentCell].value is None: # если ячейки нет
            none = True
            bd.save(bdin)
            return [currentCell, none] # координаты ячеёки и ин-я, что она пуста
        else:
            i += 1

def setDefaultSkins(botId):
    botidCell = findCell(botId, 'A', 1, 'users.xlsx', 'usersInfo')
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

def getUserStats(botId):
    botIdCell = findCell(botId, 'A', 2, 'users.xlsx', 'usersEquipped')
    ubd = openpyxl.load_workbook('users.xlsx')
    sbd = openpyxl.load_workbook('skins.xlsx')
    ueqip = ubd['usersEquipped']
    skins = sbd['skins']
    ke = ueqip['B' + botIdCell[0][1]].value
    ke = skins['D' + findCell(ke, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    bg = ueqip['N' + botIdCell[0][1]].value
    bg = skins['D' + findCell(bg, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    gn = ueqip['O' + botIdCell[0][1]].value
    gn = skins['D' + findCell(gn, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    hp = ke + bg + (ke + bg)/100*gn
    fy = ueqip['E' + botIdCell[0][1]].value
    fy = skins['D' + findCell(fy, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    ass = ueqip['K' + botIdCell[0][1]].value
    ass = skins['D' + findCell(ass, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    on = ueqip['R' + botIdCell[0][1]].value
    on = skins['D' + findCell(on, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    armor = fy + ass + (fy + ass)/100*on
    cc = ueqip['C' + botIdCell[0][1]].value
    cc = skins['D' + findCell(cc, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    sr = ueqip['J' + botIdCell[0][1]].value
    sr = skins['D' + findCell(sr, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    vl = ueqip['Q' + botIdCell[0][1]].value
    vl = skins['D' + findCell(vl, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    damage = cc + sr + (cc + sr)/100*vl
    sf = ueqip['G' + botIdCell[0][1]].value
    sf = skins['D' + findCell(sf, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    ml = ueqip['I' + botIdCell[0][1]].value
    ml = skins['D' + findCell(ml, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    orr = ueqip['S' + botIdCell[0][1]].value
    orr = skins['D' + findCell(orr, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    accuracy = sf + orr + (sf+orr)/100*ml
    gt = ueqip['F' + botIdCell[0][1]].value
    gt = skins['D' + findCell(gt, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    se = ueqip['L' + botIdCell[0][1]].value
    se = skins['D' + findCell(se, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    pm = ueqip['P' + botIdCell[0][1]].value
    pm = skins['D' + findCell(pm, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    hs = gt + pm + (gt + pm)/100*se
    sy = ueqip['D' + botIdCell[0][1]].value
    sy = skins['D' + findCell(sy, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    by = ueqip['H' + botIdCell[0][1]].value
    by = skins['D' + findCell(by, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    je = ueqip['M' + botIdCell[0][1]].value
    je = skins['D' + findCell(je, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value
    dodge = sy + je + (sy + je)/100*by

    return hp, armor, damage, accuracy, hs, dodge

@disp.message_handler(commands=['start', 'старт']) #обработчик команды /start
async def cmdStart(message: types.Message):
        bd = openpyxl.load_workbook('users.xlsx')  # открываю бд
        uinf = bd['usersInfo']  # выбираю лист usersInfo
        regtime = f'{datetime.datetime.now().day}.{datetime.datetime.now().month}.{datetime.datetime.now().year}' # дата рег-ции [17 02 2023]


        tgidCell = findCell(message.from_user.id, 'B', 2, 'users.xlsx', 'usersInfo') # формат переменной [координаты ячейки, содержимое]
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
            balanceCell = 'D' + tgidCell[0][1] # координаты баланса
            uinf[tgidCell[0]] = username # установка tgid
            uinf[botidCell] = message.from_user.id # установка botid формата [номер][первая буква]
            uinf[balanceCell] = 17000
            uinf[regtimeCell] = regtime


            uequip = bd['usersEquipped']
            botidEquip = findCell(uinf[tgidCell[0]].value, 'A', 2, 'users.xlsx', 'usersInfo')
            uequip[botidEquip[0]] = uinf[botidCell].value
            bd.save('users.xlsx') # сохранение бд
            setDefaultSkins(uinf[botidCell].value)

            await message.reply(f'Добро пожаловать в ValoShop!\nВы успешно зарегистрировались как {uinf[tgidCell[0]].value}\n\nПомощь - /help')
            await message.reply(f'Вам начислено 17000 VP')

@disp.message_handler(commands=['profile', 'профиль']) # обработка команды /profile
async def cmdProfile(message: types.Message):
    bd = openpyxl.load_workbook('users.xlsx')  # открываю бд
    uinf = bd['usersInfo']  # выбираю лист usersInfo
    botid = findCell(message.from_user.id, 'B', 2, 'users.xlsx', 'usersInfo') # поиск юзера в базе
    tgid = uinf['A' + botid[0][1]].value
    regtime = 'C' + botid[0][1] # ячейка для даты
    if botid[1] == True: # если юзер не зареган
        await message.reply('Вы ещё не зарегистрированы. Для регистрации введите /start')
    else: # если юзер найден
        stats = getUserStats(message.from_user.id)
        await message.reply(f'Профиль пользователя @{tgid}:\n\n'
                            f'Показатели:\n'
                            f'Здоровье: {stats[0]}\n'
                            f'Броня: {stats[1]}\n'
                            f'Урон: {stats[2]}\n'
                            f'Меткость: {stats[3]}%\n'
                            f'Шанс попадания в голову: {stats[4]}%\n'
                            f'Уклонение: {stats[5]}%\n\n'
                            f'Дата регистрации: {uinf[regtime].value}\n')

@disp.message_handler(commands=['help', 'помощь', 'команды', 'cmd']) # обработка команды /help
async def cmdHelp(message: types.Message):
    await message.reply(f'Команды бота:\n\n/start - Регистрация\n/profile - Профиль\n/help - Помощь\n/inv - инвентарь\n{"_"*23}\nСоздатель: @magnitgd')

@disp.message_handler(commands=['inv', 'inventory', 'инвентарь']) # обработка команды /inventory
async def cmdInv(message: types.Message):
    tgidCell = findCell(message.from_user.id, 'B', 2, 'users.xlsx', 'usersInfo')  # формат переменной [координаты ячейки, содержимое]
    bd = openpyxl.load_workbook('users.xlsx')
    uinfo = bd['usersInfo']
    username = uinfo['A' + tgidCell[0][1]].value
    botIdCell = findCell(message.from_user.id, 'A', 2, 'users.xlsx', 'usersEquipped')
    ubd = openpyxl.load_workbook('users.xlsx')
    sbd = openpyxl.load_workbook('skins.xlsx')
    ueqip = ubd['usersEquipped']
    skins = sbd['skins']
    ke = ueqip['B' + botIdCell[0][1]].value
    raw = str(int(skins['C' + findCell(ke, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value[2:]) + 2)
    kestring = [skins['D' + raw].value, skins['A' + raw].value, skins['B' + raw].value]
    cc = ueqip['C' + botIdCell[0][1]].value
    raw = str(int(skins['C' + findCell(cc, 'C', 1, 'skins.xlsx', 'skins')[0][1:]].value[2:]) + 2)
    ccstring = [skins['D' + raw].value, skins['A' + raw].value, skins['B' + raw].value]
    await message.reply(f'Инвентарь {username}:\n'
                        f'\n'
                        f'"{kestring[2][:-1]}" {kestring[1]} | +{kestring[0]} хп\n'
                        f'"{ccstring[2][:-1]}" {kestring[1]} | +{kestring[0]} урона\n'
                        f'')



if __name__ == '__main__':
    executor.start_polling(disp)
