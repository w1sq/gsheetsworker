import asyncio
from aiogram import Bot, Dispatcher,types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup
from gsheets import Google_Sheets
from db.__all_models import Users, Notifications
from db.db_session import global_init, create_session
from datetime import datetime
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton



menu_keyboard = ReplyKeyboardMarkup(resize_keyboard=True).row(KeyboardButton('Товары'),KeyboardButton('Маркетплейсы'))\
    .row(KeyboardButton('Кроссплатформенная'),KeyboardButton('Маркетинг')).row(KeyboardButton('Road Map'),KeyboardButton('Платежи'))\
        .row(KeyboardButton('Уведомления'),KeyboardButton('Записать данные'))
google_sheets = Google_Sheets()
bot = Bot(token='5410398881:AAEDNmqdHiGYo5GLd8p3VdjNLmOVMK0jNEM')
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)
global_init()
@dp.message_handler(commands=['start'])
async def start(message):
    password = message.text.split()[-1]
    if password == 'semi':
        db_sess = create_session()
        user = db_sess.query(Users).get(message.chat.id)
        if not user:
            user = Users(
                id = message.chat.id
            )
            db_sess.add(user)
            db_sess.commit()
            db_sess.close()
        await message.answer('Меню',reply_markup = menu_keyboard)


@dp.message_handler(text='Товары')
async def send_statistics(message):
    keyb= InlineKeyboardMarkup()
    for product in google_sheets.products:
        keyb.add(InlineKeyboardButton(text=product,callback_data=f'show_product |{product}'))
    keyb.add(InlineKeyboardButton(text='Вывести все товары',callback_data=f'show_product |all'))
    await message.answer('Выберите товар:',reply_markup=keyb)


async def show_product(call):
    name = call.data.split('|')[1]
    await bot.send_chat_action(call.message.chat.id,'typing')
    if name =='all':
        for product in google_sheets.products:
            await call.message.answer(google_sheets.get_item(product))
    else:
        await call.message.answer(google_sheets.get_item(name))

@dp.message_handler(text='Маркетплейсы')
async def send_statistics(message):
    keyb= InlineKeyboardMarkup()
    for marketplace in google_sheets.marketplaces:
        keyb.add(InlineKeyboardButton(text=marketplace,callback_data=f'show_marketplace |{marketplace}'))
    keyb.add(InlineKeyboardButton(text='Вывести все',callback_data=f'show_marketplace |all'))
    await message.answer('Маркетплейсы:',reply_markup=keyb)

async def show_marketplace(call):
    name = call.data.split('|')[1]
    await bot.send_chat_action(call.message.chat.id,'typing')
    if name =='all': 
        for marketplace in google_sheets.marketplaces:
            await call.message.answer(google_sheets.get_marketplace(marketplace))
    else:
        await call.message.answer(google_sheets.get_marketplace(name))


@dp.message_handler(text='Кроссплатформенная')
async def send_crossplatform(message):
    await bot.send_chat_action(message.chat.id,'typing')
    await message.answer(google_sheets.get_crossplatform())

@dp.message_handler(text='Маркетинг')
async def send_marketing(message):
    await bot.send_chat_action(message.chat.id,'typing')
    await message.answer(google_sheets.get_marketing())

@dp.message_handler(text='Road Map')
async def send_roadmap(message):
    await bot.send_chat_action(message.chat.id,'typing')
    await message.answer(google_sheets.get_roadmap())

@dp.message_handler(text='Платежи')
async def send_bills(message):
    await bot.send_chat_action(message.chat.id,'typing')
    await message.answer(google_sheets.get_bills())

@dp.message_handler(text='Записать данные')
async def send_bills(message):
    await bot.send_chat_action(message.chat.id,'typing')
    google_sheets.write_data()
    await message.answer('Данные успешно записаны')

@dp.message_handler(text='Уведомления')
async def send_bills(message):
    db_sess = create_session()
    user = db_sess.query(Users).get(message.chat.id)
    notifications = db_sess.query(Notifications).all()
    for notification in notifications:
        reply_markup = InlineKeyboardMarkup()
        if 'заказать' in notification.text:
            if str(notification.id) in user.muted_notifications:
                reply_markup.add(InlineKeyboardButton(text='🔕',callback_data=f'unmutenotification {notification.id} {user.id}'))
            else:
                reply_markup.add(InlineKeyboardButton(text='🔔',callback_data=f'mutenotification {notification.id} {user.id}'))
        reply_markup.add(InlineKeyboardButton(text='❌',callback_data=f'deletenotification {notification.id}'))
        await message.answer(f'Уведомление от {notification.date_added.strftime("%d.%m.%Y")}\n' + notification.text, reply_markup=reply_markup)
    db_sess.close()

async def unmutenotification(call):
    notification_id = call.data.split()[1]
    user_id = call.data.split()[2]
    db_sess = create_session()
    user = db_sess.query(Users).get(user_id)
    reply_markup = InlineKeyboardMarkup().add(InlineKeyboardButton(text='🔔',callback_data=f'mutenotification {notification_id} {user_id}'))
    await bot.edit_message_reply_markup(chat_id=call.message.chat.id,message_id=call.message.message_id,reply_markup=reply_markup)
    user.muted_notifications.replace(notification_id,'')
    db_sess.commit()
    db_sess.close()

async def deletenotification(call):
    notification_id = call.data.split()[1]
    db_sess = create_session()
    users = db_sess.query(Users).all()
    for user in users:
        user.muted_notifications.replace(notification_id,'')
    notification = db_sess.query(Notifications).get(notification_id)
    db_sess.delete(notification)
    await call.answer('Уведомление успешно удалено')
    db_sess.commit()
    db_sess.close()

async def mutenotification(call):
    notification_id = call.data.split()[1]
    user_id = call.data.split()[2]
    db_sess = create_session()
    reply_markup = InlineKeyboardMarkup().add(InlineKeyboardButton(text='🔕',callback_data=f'unmutenotification {notification_id} {user_id}'))
    await bot.edit_message_reply_markup(chat_id=call.message.chat.id,message_id=call.message.message_id,reply_markup=reply_markup)
    user = db_sess.query(Users).get(user_id)
    user.muted_notifications += f' {notification_id}'
    db_sess.commit()
    db_sess.close()

commands = {
    'show_product' : show_product,
    'show_marketplace' : show_marketplace,
    'unmutenotification' : unmutenotification,
    'mutenotification' : mutenotification,
    'deletenotification' : deletenotification
}

async def check_notifications():
    db_sess = create_session()
    while True:
        if datetime.now().hour == 11:
            notifications = google_sheets.get_updates()
            users = db_sess.query(Users).all()
            for user in users:
                for notification in notifications:
                    notification = db_sess.query(Notifications).filter(Notifications.text == notification).first()
                    if str(notification.id) not in user.muted_notifications:
                        if 'заказать' in notification.text:
                            reply_markup = InlineKeyboardMarkup().add(InlineKeyboardButton(text='🔔',callback_data=f'mutenotification {notification.id} {user.id}'))
                            await bot.send_message(user.id,notification.text,reply_markup=reply_markup)
                        else:
                            await bot.send_message(user.id,notification.text)
            await asyncio.sleep(3600 * 23 + 60*55)
        await asyncio.sleep(60)


@dp.callback_query_handler(lambda call: True)
async def ans(call):
    await commands[call.data.split()[0]](call)

async def main():
    await asyncio.gather(dp.start_polling(),check_notifications())

if __name__ == '__main__':
    print('Bot has started')

    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())