import asyncio
from aiogram import Bot, Dispatcher,types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardButton, InlineKeyboardMarkup
from gsheets import Google_Sheets
from db.__all_models import Users, Notifications, Limits
from db.db_session import global_init, create_session
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
import aiogram
from config import tg_bot_token
from datetime import datetime,timedelta
import aioschedule as schedule
from aiogram.dispatcher.filters.state import State, StatesGroup
menu_keyboard = ReplyKeyboardMarkup(resize_keyboard=True).row(KeyboardButton('Товары'),KeyboardButton('Маркетплейсы'))\
    .row(KeyboardButton('Кроссплатформенная'), KeyboardButton('Поставки'), KeyboardButton('Маркетинг')).row(KeyboardButton('Road Map'),KeyboardButton('Отзывы'),KeyboardButton('Платежи'))\
        .row(KeyboardButton('Уведомления'),KeyboardButton('Записать данные'),(KeyboardButton('Конверсия')))

class Answer(StatesGroup):
    review_answer = State()

class GetLimitAmount(StatesGroup):
    limit_amount = State()

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

async def main_menu(call):
    await call.message.answer('Вы в меню',reply_markup = menu_keyboard)

@dp.message_handler(commands=['send_message'])
async def send_message(message):
    text = ' '.join(message.text.split()[1:])
    db_sess = create_session()
    users = db_sess.query(Users).all()
    db_sess.close()
    for user in users:
        try:
            await bot.send_message(user.id, text)
        except (aiogram.utils.exceptions.ChatNotFound, aiogram.utils.exceptions.MessageTextIsEmpty, aiogram.utils.exceptions.BotBlocked):
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
    for product_id in range(len(google_sheets.products)):
        keyb.add(InlineKeyboardButton(text=google_sheets.products[product_id],callback_data=f'show_product |{product_id}'))
    keyb.add(InlineKeyboardButton(text='Вывести все товары',callback_data=f'show_product |all'))
    await message.answer('Выберите товар:',reply_markup=keyb)

async def show_product(call):
    product_id = call.data.split('|')[1]
    await bot.send_chat_action(call.message.chat.id,'typing')
    if product_id =='all':
        for product in google_sheets.keys_coords.keys():
            await call.message.answer(google_sheets.get_item(product))
    else:
        await call.message.answer(google_sheets.get_item(google_sheets.products[int(product_id)]))

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
            .row(InlineKeyboardButton('Лимиты Wb', callback_data='limits_wb'))
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

async def limits_wb(call):
    db_sess = create_session()
    limits = db_sess.query(Limits).all()
    limits_add_keyboard = InlineKeyboardMarkup().row(InlineKeyboardButton(text='Добавить новые склады',callback_data='add_limits')).row(InlineKeyboardButton(text="Главное меню", callback_data='main_menu'))
    if not limits:
        await call.message.answer("Привет! Сейчас я не отслеживаю лимитов Wildberries. Но могу начать это делать 24/7. Нажимай кнопку внизу.", reply_markup = limits_add_keyboard)
    else:
        message = "Привет! Сейчас я 24/7 ищу лимиты на эти склады:\n\n"
        for limit in limits:
            message += f'{list(google_sheets.warehouses.keys())[list(google_sheets.warehouses.values()).index(limit.warehouse)]}\n'
        await call.message.answer(message, reply_markup = limits_add_keyboard)

async def add_limits(call):
    message = 'На какие склады мне отслеживать поставки?'
    warehouses_keyboard = InlineKeyboardMarkup()
    for i in google_sheets.warehouses.keys():
        warehouses_keyboard.row(InlineKeyboardButton(text=i, callback_data=f'process_limits_warehouse {google_sheets.warehouses[i]}'))
    warehouses_keyboard.row(InlineKeyboardButton(text='Назад', callback_data='limits_wb')).row(InlineKeyboardButton(text='В главное меню', callback_data='main_menu'))
    await call.message.answer(message, reply_markup=warehouses_keyboard)

async def process_limits_warehouse(call):
    warehouse = call.data.split()[1]
    message = 'Какой вид поставки?'
    warehouses_keyboard = InlineKeyboardMarkup()
    containers = ["Короба", "Монопалеты", "Суперсейф"]
    for i in containers:
        warehouses_keyboard.row(InlineKeyboardButton(text=i, callback_data=f'pl_amount {warehouse} {i}'))
    warehouses_keyboard.row(InlineKeyboardButton(text='Выбрать другой склад', callback_data='add_limits')).row(InlineKeyboardButton(text='В главное меню', callback_data='main_menu'))
    await call.message.answer(message, reply_markup=warehouses_keyboard)

async def process_limits_amount(call):
    warehouse,container = call.data.split()[1:]
    message = 'Далее необходимое число для отслеживания:\nУкажите необходимый лимит в условном значении. Указывайте точное количества товара в вашей поставке. Это количество указано в скобках рядом с номером заказа в разделе - Управление поставками\n\nОтправьте мне число ниже 👇'
    await GetLimitAmount.limit_amount.set()
    state = dp.get_current().current_state()
    await state.update_data(warehouse=warehouse)
    await state.update_data(container=container)
    await call.message.answer(message)

async def choose_date(call):
    warehouse, limit_type, amount = call.data.split()[1:]
    text = f"Хорошо, буду искать лимит на поставку {amount} штук типа {limit_type} в {list(google_sheets.warehouses.keys())[list(google_sheets.warehouses.values()).index(int(warehouse))]}. Теперь выберите даты для поисков лимитов:"
    dates_keyboard = InlineKeyboardMarkup()
    dates = {"Сегодня":0, "Завтра":1, "Неделя":7, "Месяц":30, "Искать пока не найдется":-1}
    for i in dates.keys():
        dates_keyboard.row(InlineKeyboardButton(text=i, callback_data=f"pl_dates {warehouse} {limit_type} {amount} {dates[i]}"))
    dates_keyboard.row(InlineKeyboardButton(text='Изменить лимит', callback_data=f"pl_amount {warehouse} {limit_type}")).row(InlineKeyboardButton(text='В главное меню', callback_data='main_menu'))
    await call.message.answer(text=text, reply_markup=dates_keyboard)

@dp.message_handler(state=GetLimitAmount.limit_amount)
async def process_limits_amount_confirm(message: types.Message, state):
    if message.text == "ОТМЕНА":
        await state.finish()
        await message.answer('Вы в меню',reply_markup = menu_keyboard)
    elif message.text.isdigit():
        state_data = await state.get_data()
        amount = int(message.text)
        text = f"Хорошо, буду искать лимит на поставку {amount} штук типа {state_data['container']} в {list(google_sheets.warehouses.keys())[list(google_sheets.warehouses.values()).index(int(state_data['warehouse']))]}. Теперь выберите даты для поисков лимитов:"
        dates_keyboard = InlineKeyboardMarkup()
        dates = {"Сегодня":0, "Завтра":1, "Неделя":7, "Месяц":30, "Искать пока не найдется":-1}
        for i in dates.keys():
            dates_keyboard.row(InlineKeyboardButton(text=i, callback_data=f"pl_dates {state_data['warehouse']} {state_data['container']} {amount} {dates[i]}"))
        dates_keyboard.row(InlineKeyboardButton(text='Изменить лимит', callback_data=f"pl_amount {state_data['warehouse']} {state_data['container']}")).row(InlineKeyboardButton(text='В главное меню', callback_data='main_menu'))
        await message.answer(text=text, reply_markup=dates_keyboard)
        await state.finish()
    else:
        await message.answer('Я принимаю только числа, либо отправьте число, либо напишите "ОТМЕНА" для отмены')

async def process_limits_dates(call):
    warehouse, container, amount, time_range = call.data.split()[1:]
    dates = {"Сегодня":0, "Завтра":1, "Неделя":7, "Месяц":30, "Искать пока не найдется":-1}
    confirm_keyboard = InlineKeyboardMarkup().row(InlineKeyboardButton(text='Да, добавить еще отслеживание', callback_data='add_limits')).row(InlineKeyboardButton(text='Спасибо, не нужно', callback_data='main_menu')).row(InlineKeyboardButton(text='В главное меню', callback_data='main_menu'))
    message = f"Хорошо. Будут усиленно искать поставку {amount} штук {list(google_sheets.warehouses.keys())[list(google_sheets.warehouses.values()).index(int(warehouse))]} типа {container} по этому заданию «{list(dates.keys())[list(dates.values()).index(int(time_range))]}»"
    db_sess = create_session()
    if int(time_range) != -1:
        new_limit = Limits(warehouse=warehouse, type=container, amount=int(amount),time_range= datetime.now() + timedelta(days=int(time_range)))
    else:
        new_limit = Limits(warehouse=warehouse, type=container, amount=int(amount),forever=True)
    db_sess.add(new_limit)
    db_sess.commit()
    db_sess.close()
    await call.message.answer(text=message, reply_markup = confirm_keyboard)

async def review_restored_success(call):
    google_sheets.review_recover(call.data.split()[1], 'восстановили отзыв')
    await call.message.answer('Статус отзыва стал "восстановили отзыв".')

async def review_unrestored_needed(call):
    google_sheets.review_recover_and_date(call.data.split()[1], 'написать в поддержку')
    await call.message.answer('Статус отзыва стал "написать в поддержку".')

async def book_a_limit(call):
    await call.message.answer("Пока меня не научили это делать. Поэтому попрошу Вас занять поставку самостоятельно")

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
    'limits' : limits,
    'limits_wb' : limits_wb,
    'main_menu' : main_menu,
    'add_limits' : add_limits,
    'process_limits_warehouse' : process_limits_warehouse,
    'pl_amount' : process_limits_amount,
    'pl_dates' : process_limits_dates,
    'review_restored_success' : review_restored_success,
    "review_unrestored_needed" : review_unrestored_needed,
    'book_a_limit' : book_a_limit,
    'cd': choose_date
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
                except (aiogram.utils.exceptions.ChatNotFound, aiogram.utils.exceptions.MessageTextIsEmpty, aiogram.utils.exceptions.BotBlocked):
                    pass
    db_sess.close()

@dp.message_handler(commands=['main'])
async def send_main_notifications(message=''):
    db_sess = create_session()
    notifications = google_sheets.get_updates()
    await bot.send_message(5546230210, str(notifications))
    users = db_sess.query(Users).all()
    for notification_chunk in notifications:
        if type(notification_chunk) == str:
            for user in users:
                try:
                    await bot.send_message(user.id,notification_chunk)
                except (aiogram.utils.exceptions.ChatNotFound, aiogram.utils.exceptions.MessageTextIsEmpty, aiogram.utils.exceptions.BotBlocked):
                    pass
        else:
            for notification in notification_chunk:
                notification = db_sess.query(Notifications).filter(Notifications.text == notification).first()
                for user in users:
                    if str(notification.id) not in user.muted_notifications:
                        try:
                            reply_markup = InlineKeyboardMarkup().add(InlineKeyboardButton(text='🔔',callback_data=f'mutenotification {notification.id} {user.id}'))
                            await bot.send_message(user.id,notification.text,reply_markup=reply_markup)
                        except (aiogram.utils.exceptions.ChatNotFound, aiogram.utils.exceptions.MessageTextIsEmpty, aiogram.utils.exceptions.BotBlocked):
                            pass
        await asyncio.sleep(60*15)
    db_sess.close()

@dp.message_handler(commands=['test'])
async def send_test_main_notifications(message=''):
    db_sess = create_session()
    notifications = google_sheets.get_updates()
    await bot.send_message(5546230210, str(notifications))
    users = [db_sess.query(Users).get(5546230210)]
    for notification_chunk in notifications:
        if type(notification_chunk) == str:
            for user in users:
                try:
                    await bot.send_message(user.id,notification_chunk)
                except (aiogram.utils.exceptions.ChatNotFound, aiogram.utils.exceptions.MessageTextIsEmpty, aiogram.utils.exceptions.BotBlocked):
                    pass
        else:
            for notification in notification_chunk:
                notification = db_sess.query(Notifications).filter(Notifications.text == notification).first()
                for user in users:
                    if str(notification.id) not in user.muted_notifications:
                        try:
                            reply_markup = InlineKeyboardMarkup().add(InlineKeyboardButton(text='🔔',callback_data=f'mutenotification {notification.id} {user.id}'))
                            await bot.send_message(user.id,notification.text,reply_markup=reply_markup)
                        except (aiogram.utils.exceptions.ChatNotFound, aiogram.utils.exceptions.MessageTextIsEmpty, aiogram.utils.exceptions.BotBlocked):
                            pass
    db_sess.close()

async def send_limits_notifications():
    db_sess = create_session()
    limits = db_sess.query(Limits).all()
    messages = {}
    users = db_sess.query(Users).all()
    for limit in limits:
        if not limit.forever and limit.time_range < datetime.now() + timedelta(days=-1):
            db_sess.delete(limit)
        else:
            date, amount = await google_sheets.get_warehouse_limits(limit)
            if date:
                messages[limit] = f"🚚  Место для поставки на склад {list(google_sheets.warehouses.keys())[list(google_sheets.warehouses.values()).index(limit.warehouse)]} на {amount} мест с типом {limit.type} найдено на дату {date}"
            db_sess.delete(limit)
    if messages:
        for limit in messages.keys():
            for user in users:
                try:
                    keyb = InlineKeyboardMarkup().row(InlineKeyboardButton(text='Да, займи поставку на эти даты', callback_data='book_a_limit')).row(InlineKeyboardButton(text='Ищи другие даты', callback_data=f'cd {limit.warehouse} {limit.type} {limit.amount}'))
                    await bot.send_message(user.id,messages[limit],reply_markup=keyb)
                except (aiogram.utils.exceptions.ChatNotFound, aiogram.utils.exceptions.MessageTextIsEmpty, aiogram.utils.exceptions.BotBlocked):
                    pass
    db_sess.commit()
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
    schedule.every(5).minutes.do(send_limits_notifications)
    schedule.every().monday.at("10:00").do(send_conversion_notifications)
    schedule.every().thursday.at("10:00").do(send_conversion_notifications)
    loop.create_task(check_schedule())
    loop.run_until_complete(main())