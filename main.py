import logging
from aiogram import Bot, Dispatcher, types, executor
from aiogram.dispatcher.filters import Text
import openpyxl
import config
from datetime import date

bot = Bot(token=config.API_TOKEN)
organis = config.org
gabo = config.gabo
kost = config.kost
dmin = config.dmin
fileexcel = openpyxl.load_workbook('Организации 2021.xlsx')
sheet = fileexcel['Октябрь21']



disp = Dispatcher(bot)

logging.basicConfig(level=logging.INFO)
logging.basicConfig(level=logging.DEBUG)


@disp.message_handler(commands='start')
async def Botik(message: types.Message):
    await message.answer('Напиши /job для того, чтобы отметиться...')


@disp.message_handler(commands='job')
async def asver(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [organis[0], organis[1], organis[2], organis[3], organis[4], organis[5], organis[6], organis[7],
               organis[8], organis[9], organis[10], organis[11], organis[12], organis[13], organis[14], organis[15],
               organis[16], organis[17], organis[18], organis[19], organis[20], organis[21], organis[22], organis[23],
               organis[24], organis[25]]
    keyboard.add(*buttons)
    await message.answer("В какой организации ты был?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[0]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [gabo[0], gabo[1]]
    keyboard.add(*buttons)
    await message.answer("Где именно?", reply_markup=keyboard)

@disp.message_handler(Text(equals=gabo[0]))
async def with_puree(message: types.Message):
    sheet['E3'] = message.from_user.first_name + " был " + str(date.today())
    fileexcel.save('Организации 2021.xlsx')
    await message.reply("Записал")

@disp.message_handler(Text(equals=gabo[1]))
async def with_puree(message: types.Message):
    sheet['E4'] = message.from_user.first_name + " был " + str(date.today())
    fileexcel.save('Организации 2021.xlsx')
    await message.reply("Записал")


@disp.message_handler(Text(equals=organis[1]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.dmin[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[2]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [kost[0], kost[1], kost[2], kost[3], kost[4]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[3]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.rog[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[4]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.novs[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[5]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.dia[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[6]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.dksin[0], config.dksin[1], config.dksin[2], config.dksin[3], config.dksin[4], config.dksin[5],
               config.dksin[6], config.dksin[7], config.dksin[8]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[7]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.stro[0], config.stro[1], config.stro[2], config.stro[3], config.stro[4]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[8]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.dkya[0], config.dkya[1], config.dkya[2]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[9]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.dd[0], config.dd[1]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[10]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.ksp[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[11]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.on[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[12]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.skb[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[13]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.strem[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[14]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.tyg[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[15]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.tyn[0], config.tyn[1]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[16]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.tyd[0], config.tyd[1]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[17]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.tyr[0], config.tyr[1]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[18]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.tys[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[19]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.tyk[0], config.tyk[1], config.tyk[2]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[20]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.tyya[0], config.tyya[1], config.tyya[2], config.tyya[3]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[21]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.csps[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[22]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.torp[0], config.torp[1], config.torp[2], config.torp[3], config.torp[4]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[23]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.csa[0]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[24]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.vod[0], config.vod[1], config.vod[2]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


@disp.message_handler(Text(equals=organis[25]))
async def with_puree(message: types.Message):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    buttons = [config.pso[0], config.pso[1]]
    keyboard.add(*buttons)
    await message.reply("Где именно?", reply_markup=keyboard)


if __name__ == '__main__':
    executor.start_polling(disp, skip_updates=True)
