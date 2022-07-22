import asyncio
from glob import glob
from typing import List
from aiogram import Bot, Dispatcher,types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup
from gsheets import Google_Sheets
from db.__all_models import Users, Notifications
from db.db_session import global_init, create_session
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
import aiogram
from config import tg_bot_token
import aioschedule as schedule
from aiogram.dispatcher.filters.state import State, StatesGroup
menu_keyboard = ReplyKeyboardMarkup(resize_keyboard=True).row(KeyboardButton('Товары'),KeyboardButton('Маркетплейсы'))\
    .row(KeyboardButton('Кроссплатформенная'), KeyboardButton('Поставки'), KeyboardButton('Маркетинг')).row(KeyboardButton('Road Map'),KeyboardButton('Отзывы'),KeyboardButton('Платежи'))\
        .row(KeyboardButton('Уведомления'),KeyboardButton('Записать данные'),(KeyboardButton('Конверсия')))

class Answer(StatesGroup):
    review_answer = State()

google_sheets = Google_Sheets()
bot = Bot(token=tg_bot_token)
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

@dp.message_handler(commands=['send_message'])
async def send_message(message):
    text = ' '.join(message.text.split()[1:])
    db_sess = create_session()
    users = db_sess.query(Users).all()
    db_sess.close()
    for user in users:
        try:
            await bot.send_message(user.id, text)
        except aiogram.utils.exceptions.ChatNotFound:
            pass

@dp.message_handler(text='Отзывы')
async def send_reviews(message):
    global local_reviews
    await bot.send_chat_action(message.chat.id,'typing')
    reviews = google_sheets.get_reviews()
    if reviews:
        for review in reviews:
            answer_keyboard = InlineKeyboardMarkup(one_time_keyboard=True).row(InlineKeyboardButton(text='Ввести текст', callback_data=f'answer_review {review[1]}'))
            await message.answer(review[0], reply_markup=answer_keyboard)
    else:
        await message.answer('Негативных отзывов пока не обнаружено')

@dp.message_handler(text='Конверсия')
async def send_conversion(message):
    await bot.send_chat_action(message.chat.id,'typing')
    await message.answer(google_sheets.get_conversions())

@dp.message_handler(text='Товары')
async def send_items(message):
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
async def send_markets(message):
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

@dp.message_handler(text='Поставки')
async def send_supplies(message):
    supplies_keyb = InlineKeyboardMarkup().row(InlineKeyboardButton(text='Региональные Wildberries',callback_data='regional wb'))\
        .row(InlineKeyboardButton(text='Региональные OZON', callback_data='regional ozon'))\
            .row(InlineKeyboardButton('Лимиты Wb', callback_data='limits wb'))
    await message.answer(text='Выберите нужные данные:', reply_markup=supplies_keyb)

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
async def write_data(message):
    await bot.send_chat_action(message.chat.id,'typing')
    google_sheets.write_data()
    await message.answer('Данные успешно записаны')

@dp.message_handler(text='Уведомления')
async def send_notifications(message):
    db_sess = create_session()
    user = db_sess.query(Users).get(message.chat.id)
    notifications = db_sess.query(Notifications).all()
    if notifications:
        for notification in notifications:
            reply_markup = InlineKeyboardMarkup()
            if 'заказать' in notification.text:
                if str(notification.id) in user.muted_notifications:
                    reply_markup.add(InlineKeyboardButton(text='🔕',callback_data=f'unmutenotification {notification.id} {user.id}'))
                else:
                    reply_markup.add(InlineKeyboardButton(text='🔔',callback_data=f'mutenotification {notification.id} {user.id}'))
            reply_markup.add(InlineKeyboardButton(text='❌',callback_data=f'deletenotification {notification.id}'))
            await message.answer(f'Уведомление от {notification.date_added.strftime("%d.%m.%Y")}\n' + notification.text, reply_markup=reply_markup)
    else:
        await message.answer('Уведомлений пока нет')
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
    await call.message.delete()
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

async def answer_review(call):
    await call.message.answer('Благодарю. Пожалуйста, отправьте мне ответ в поддержку следующим сообщением')
    await Answer.review_answer.set()
    current_state = dp.get_current().current_state()
    await current_state.update_data(answer_id=call.data.split()[1])

@dp.message_handler(state=Answer.review_answer)
async def process_answer(message: types.Message, state):
    state_data = await state.get_data()
    await state.finish()
    review_text = google_sheets.get_review_by_appeal_num(state_data['answer_id'])
    confirm_keyboard = InlineKeyboardMarkup(one_time_keyboard=True).row(InlineKeyboardButton(text='Подтвердить', callback_data=f"confirm_answer {state_data['answer_id']}"))
    await message.answer(f"Супер! Я прямо сейчас создам обращение в Wildberries на этот отзыв:\n\n{review_text}\n\n…со следующим текстом:\n\n{message.text}", reply_markup=confirm_keyboard)
    
async def confirm_answer(call):
    answer_id = call.data.split()[1]
    answer = call.message.text.split('\n\n')[3]
    google_sheets.send_answer(answer_id, answer)
    await call.message.answer('Обращение отослано')

async def review_delete_success(call):
    google_sheets.change_review_status(call.data.split()[1], 'удалён')
    await call.message.answer('Успешно подтверждено удаление отзыва.')

async def review_help_needed(call):
    google_sheets.change_review_status(call.data.split()[1], 'нужна помощь')
    await call.message.answer('Статус сменен на "нужна помощь".')

async def regional(call):
    platform = call.data.split()[1]
    await call.message.answer(google_sheets.get_regional(platform))

async def limits(call):
    platform = call.data.split()[1]
    await call.message.answer(google_sheets.get_limits(platform))

commands = {
    'show_product' : show_product,
    'show_marketplace' : show_marketplace,
    'unmutenotification' : unmutenotification,
    'mutenotification' : mutenotification,
    'deletenotification' : deletenotification,
    'answer_review' : answer_review,
    'confirm_answer' : confirm_answer,
    'review_delete_success' : review_delete_success,
    'review_help_needed' : review_help_needed,
    'regional' : regional,
    'limits' : limits
}

@dp.message_handler(commands=['conversion'])
async def send_conversion_notifications(message=''):
    db_sess = create_session()
    conversions = google_sheets.get_conversions_notifications()
    users = db_sess.query(Users).all()
    for user in users:
        try:
            await bot.send_message(user.id,conversions)
        except aiogram.utils.exceptions.ChatNotFound:
            pass
    db_sess.close()


@dp.message_handler(commands=['supply'])
async def send_supply_notifications(message=''):
    db_sess = create_session()
    supply_notifications = google_sheets.get_supply_notifications()
    users = db_sess.query(Users).all()
    for user in users:
        for supply_notification_place in supply_notifications.keys():
            if len(supply_notifications[supply_notification_place]) > 0:
                message = f'⚡️ Внимание: рекомендую сделать поставку на Ozon {supply_notification_place} в следующем количестве:\n\n'
                for supply in supply_notifications[supply_notification_place]:
                    message += supply
                try:
                    await bot.send_message(user.id,message)
                except aiogram.utils.exceptions.ChatNotFound:
                    pass
    db_sess.close()

@dp.message_handler(commands=['main'])
async def send_main_notifications(message=''):
    db_sess = create_session()
    notifications = google_sheets.get_updates()
    users = db_sess.query(Users).all()
    for notification_chunk in notifications:
        if type(notification_chunk) == str:
            for user in users:
                try:
                    await bot.send_message(user.id,notification_chunk)
                except aiogram.utils.exceptions.ChatNotFound and aiogram.utils.exceptions.MessageTextIsEmpty and aiogram.utils.exceptions.BotBlocked:
                    pass
        else:
            for notification in notification_chunk:
                notification = db_sess.query(Notifications).filter(Notifications.text == notification).first()
                for user in users:
                    if str(notification.id) not in user.muted_notifications:
                        try:
                            reply_markup = InlineKeyboardMarkup().add(InlineKeyboardButton(text='🔔',callback_data=f'mutenotification {notification.id} {user.id}'))
                            await bot.send_message(user.id,notification.text,reply_markup=reply_markup)
                        except aiogram.utils.exceptions.ChatNotFound and aiogram.utils.exceptions.MessageTextIsEmpty and aiogram.utils.exceptions.BotBlocked:
                            pass
        await asyncio.sleep(60*15)
    db_sess.close()

@dp.message_handler(commands=['test'])
async def send_test_main_notifications(message=''):
    db_sess = create_session()
    notifications = google_sheets.get_updates()
    users = [db_sess.query(Users).get(5546230210)]
    for notification_chunk in notifications:
        if type(notification_chunk) == str:
            for user in users:
                try:
                    await bot.send_message(user.id,notification_chunk)
                except aiogram.utils.exceptions.ChatNotFound and aiogram.utils.exceptions.MessageTextIsEmpty and aiogram.utils.exceptions.BotBlocked:
                    pass
        else:
            for notification in notification_chunk:
                notification = db_sess.query(Notifications).filter(Notifications.text == notification).first()
                for user in users:
                    if str(notification.id) not in user.muted_notifications:
                        try:
                            reply_markup = InlineKeyboardMarkup().add(InlineKeyboardButton(text='🔔',callback_data=f'mutenotification {notification.id} {user.id}'))
                            await bot.send_message(user.id,notification.text,reply_markup=reply_markup)
                        except aiogram.utils.exceptions.ChatNotFound and aiogram.utils.exceptions.MessageTextIsEmpty and aiogram.utils.exceptions.BotBlocked:
                            pass
    db_sess.close()

@dp.callback_query_handler(lambda call: True)
async def ans(call):
    await commands[call.data.split()[0]](call)


async def main():
    await dp.start_polling()


async def check_schedule():
    while True:
        await schedule.run_pending()
        await asyncio.sleep(1)


if __name__ == '__main__':
    print('Bot has started')
    loop = asyncio.get_event_loop()
    schedule.every().day.at("11:00").do(send_main_notifications)
    schedule.every().day.at("13:00").do(send_supply_notifications)
    schedule.every().monday.at("10:00").do(send_conversion_notifications)
    schedule.every().thursday.at("10:00").do(send_conversion_notifications)
    loop.create_task(check_schedule())
    loop.run_until_complete(main())