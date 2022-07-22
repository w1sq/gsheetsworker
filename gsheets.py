import json
from typing import List
import gspread
from datetime import datetime,timedelta
from db.__all_models import Users, Notifications
from db.db_session import global_init, create_session
from sqlalchemy.orm import Session
import re
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from config import gsheet_token
from time import sleep  
import xlrd
from selenium.webdriver.common.by import By
import asyncio, aiohttp
def float1(string:str):
    try:
        string = float(string.replace(',','.'))
        return string
    except Exception:
        return 0

def float2(string:str):
    try:
        string = float(string.replace(',','.'))
        return string
    except Exception:
        return 

def int1(string:str):
    try:
        string = int(string)
        return string
    except Exception:
        return 0

def int2(string:str):
    try:
        string = int(string)
        return string
    except Exception:
        return

def get_plus(number:float):
    if number > 0:
        return '+'+str(round(number,3))
    return str(round(number,3)) 

def get_plus2(number):
    if number < 0:
        return '+'+str(number * (-1))
    return str(number*-1)

class Google_Sheets():
    def __init__(self) -> None:
        self.gc = gspread.service_account(filename='service_key.json')
        self.sheets = self.gc.open_by_key(gsheet_token)
        self.worksheet = self.sheets.get_worksheet(0)
        self.worksheet_roadmap = self.sheets.get_worksheet(2)
        self.worksheet_bills = self.sheets.get_worksheet(3)
        self.need_names = ['Хабаровск', 'Самара', 'Тверь', 'Хоругвино', 'Казань', 'Санкт-Петербург', 'Новосибирск', 'Екатеринбург', 'Ростов-на-Дону', 'Калининград', 'Красноярск', 'Нижний Новгород', 'Новая Рига']
        self.warehouses = {
            "Алексин": 206348,
            "Екатеринбург" : 1733,
            "Казань": 117986,
            "Коледино" : 507,
            "Крёкшино КБТ" : 124731, 
            "Новосибирск" : 686,
            "Подольск" : 117501,
            "СЦ Астрахань" : 169872,
            "СЦ Белая Дача" : 205228,
            "СЦ Владимир" : 144649,
            "СЦ Волгоград" : 6144,
            "СЦ Иваново" : 203632,
            "СЦ Калуга" : 117442,
            "СЦ Комсомольская" : 154371,
            "СЦ Красногорск" : 6159,
            "СЦ Курск" : 140302,
            "СЦ Курьяновская" : 156814,
            "СЦ Липецк" : 160030,
            "СЦ Лобня" : 117289,
            "СЦ Минск" : 117393,
            "СЦ Мытищи" : 115650,
            "СЦ Набережные Челны": 204952,
            "СЦ Нижний Новгород" : 118535,
            "СЦ Новокосино" : 141637,
            "СЦ Подрезково" : 124716,
            "СЦ Рязань" : 6156,
            "СЦ Серов" : 169537,
            "СЦ Симферополь" : 144154,
            "СЦ Тамбов" : 117866,
            "СЦ Тверь" : 117456,
            "СЦ Уфа" : 149445,
            "СЦ Чебоксары" : 203799,
            "СЦ Южные Ворота" : 158328,
            "СЦ Ярославль" : 6154,
            "Санкт-Петербург Уткина Заводь 4к4" : 2737,
            "Санкт-Петербург Шушары" : 159402,
            "Склад Казахстан" : 204939,
            "Склад Краснодар" : 130744,
            "Электросталь" : 120762,
            "Электросталь КБТ" : 121709
        }
        self.products = ['Массажное масло', 'Спрей для волос', 'Масло для волос', 'Крем для тела', 'Крем для ног', 'Маска для волос', 'Кератолитик']
        self.marketplaces = ['Wildberries', 'OZON', 'Yandex', 'Остальное']
        self.weekdays = ['Понедельник','Вторник','Среда','Четверг','Пятница','Суббота','Воскресенье']
        self.keys_coords = {'Массажное масло':'R1', 'Спрей для волос':'AG1', 'Масло для волос':'AV1', 'Крем для тела':'BK1', 'Крем для ног':'BZ1', 'Маска для волос':'CO1', 'Кератолитик':'DD1'}
        self.chrome_options = Options()
        self.chrome_options.add_argument(
                    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36")
        self.chrome_options.add_argument('--disable-extensions')
        self.chrome_options.add_experimental_option("prefs", {"download.default_directory": r"C:\Users\79152\Google Drive\all_python\telegram\gsheetsworker"})
        self.chrome_options.add_argument('--no-sandbox')
        self.chrome_options.add_argument('--disable-dev-shm-usage')
        prefs = {"profile.managed_default_content_settings.images": 2}
        self.chrome_options.add_experimental_option("prefs", prefs)
        self.chrome_options.add_argument("--headless")
        self.chrome_options.add_argument('--log-level=3')
        self.chrome_options.add_argument("start-maximized")
        self.chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.chrome_options.add_experimental_option('useAutomationExtension', False)
        self.chrome_options.add_argument('--disable-blink-features=AutomationControlled')


    async def get_warehouse_limits(self, warehouse_id : int):
        warehouses_url = 'https://seller.wildberries.ru/ns/sm/supply-manager/api/v1/plan/listLimits'
        headers = {
            "Host": "seller.wildberries.ru",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:102.0) Gecko/20100101 Firefox/102.0",
            "Accept": "*/*",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate, br",
            "Content-Type": "application/json",
            "Content-Length" : f"{108+len(str(warehouse_id))}",
            "Origin": "https://seller.wildberries.ru",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode" : "cors",
            "Sec-Fetch-Site" : "same-origin",
            "Referer": "https://seller.wildberries.ru/supplies-management/warehouses-limits",
            "Connection" : "keep-alive",
            "Cookie": "___wbu=732543b3-2248-4059-95fb-5c4efea375e8.1629749510; _wbauid=10159857131629749509; _ga=GA1.2.558243196.1629749510; locale=ru; WBToken=AseSkyr40sCtDPiOqq4MQsxU7Ql9wXKTaYgvvFIEVw5LGAUyOmL5a2O1BHQmAm69_jaxSYm6SPWFs04NNXiITxF3rt_dUZpnbdLf8GUWDHq3tA; x-supplier-id=fa9c5339-9cc8-4029-b2ee-bfd61bbf9221; __wbl=cityId%3D0%26regionId%3D0%26city%3D%D0%9C%D0%BE%D1%81%D0%BA%D0%B2%D0%B0%26phone%3D84957755505%26latitude%3D55%2C755787%26longitude%3D37%2C617634%26src%3D1; __store=117673_122258_122259_125238_125239_125240_507_3158_117501_120602_120762_6158_121709_124731_130744_159402_2737_117986_1733_686_132043_161812_1193; __region=68_64_83_4_38_80_33_70_82_86_75_30_69_22_66_31_40_1_48_71; __pricemargin=1.0--; __cpns=12_3_18_15_21; __sppfix=; __dst=-1029256_-102269_-2162196_-1257786; __tm=1658338552"
        }
        now = datetime.now()
        monday = now - timedelta(days = now.weekday())
        monday_str = monday.strftime("%Y-%m-%d")
        weekend_str = (monday + timedelta(days=6)).strftime("%Y-%m-%d")
        payload = '{{"params":{{"dateFrom":"{}","dateTo":"{}","warehouseId":{}}},"jsonrpc":"2.0","id":"json-rpc_20"}}'.format(monday_str, weekend_str, warehouse_id)
        async with aiohttp.ClientSession(headers=headers) as session:
            response = await session.post(url=warehouses_url, data=payload, headers=headers)
            responce_json = await response.json()
            return responce_json
    
    def get_last_date_conversions(self, worksheet):
        i = 0
        while True:
            delta = timedelta(days=i)
            date = (datetime.today()- delta).strftime('%d.%m.%Y')
            cell = worksheet.find(date)
            if worksheet.acell(f'E{cell.row}').value:
                return date, int(cell.row) -1
            i += 1

    def get_last_date_main(self):
        i = 0
        while True:
            delta = timedelta(days=i)
            date = (datetime.today()- delta).strftime('%d.%m.%Y')
            cell = self.worksheet.find(date)
            if self.worksheet.acell(f'S{cell.row}').value:
                return date, int(cell.row) -1
            i += 1
    
    def write_data(self):
        sheet = self.gc.open_by_key('1GdMuPn71NNhKm03Pg0sno_gok3Hg3AwfvetFCfpjsG4')
        worksheet = sheet.worksheet('trigger')
        worksheet.update('A1', 'on')

    def get_ozon(self, browser):
        browser.get('https://seller.ozon.ru/app/analytics/fulfillment-reports/forecast')
        sleep(1)
        email = browser.find_element("name", "email")
        email.send_keys('semily.cosmetic@gmail.com')
        sleep(1)
        password = browser.find_element("name", "password")
        password.send_keys('cycbAz-suxta4-sesrer')
        sleep(1)
        login = browser.find_element(By.TAG_NAME, "button")
        login.click()
        sleep(2)
        login = browser.find_element(By.TAG_NAME, "button")
        login.click()
        sleep(10)
        login = browser.find_element(By.TAG_NAME, "button")
        login.click()
        sleep(30)

    def get_reviews(self):
        sheet = self.gc.open_by_key('1LMt-hlMhaDq0iyMenlC6QdWWfQ7sgVZ4tl9_ka14HVY')
        worksheet = sheet.worksheet('Отзывы')
        all_data = worksheet.get_all_values()
        reviews = []
        for row in all_data:
            if row[7] in ['нужна помощь', 'не определил']:
                reviews.append((f"🧐 На товар {row[1]} пришел плохой отзыв {row[4]} звезды. Хочу его удалить через поддержку. Прошу помочь мне с текстом обращения:\n\nТекст отзыва:\n{row[3]}\n\nИспользуйте шаблон для этого клиента:\n\nДобрый день. Прошу удалить отзыв пользователя {row[2]} от {row[0]}. Он не относится к оценке качества товара. Кроме того, ответственность за хранение и транспортировку несет Wildberries. (………….) . Клиент не полностью ознакомился с карточкой товара на Wildberries. Требуем удалить отзыв, так как он вводит в заблуждение других покупателей и несправедливо снижает рейтинг карточки товара!",row[6]))
        return reviews
    
    def get_review_by_appeal_num(self, appeal_number:int):
        sheet = self.gc.open_by_key('1LMt-hlMhaDq0iyMenlC6QdWWfQ7sgVZ4tl9_ka14HVY')
        worksheet = sheet.worksheet('Отзывы')
        row = worksheet.find(appeal_number).row
        return worksheet.acell(f'D{row}').value

    def get_regional(self, platform:str):
        sheet = self.gc.open_by_key('1lCp3Myysw5kekRL3CTXnhuGwvQLT348V9q5rM735Uvg')
        worksheet = sheet.worksheet('Уведомления')
        if platform == 'wb':
            return worksheet.acell('A2').value
        elif platform == 'ozon':
            return worksheet.acell('B2').value
    
    def get_limits(self, platform:str):
        return "Лимиты пока не готовы"

    def change_review_status(self, appeal_number:int, status:str):
        sheet = self.gc.open_by_key('1LMt-hlMhaDq0iyMenlC6QdWWfQ7sgVZ4tl9_ka14HVY')
        worksheet = sheet.worksheet('Отзывы')
        row = worksheet.find(appeal_number).row
        worksheet.update(f'H{row}', status)

    def send_answer(self, answer_id, answer):
        sheet = self.gc.open_by_key('1LMt-hlMhaDq0iyMenlC6QdWWfQ7sgVZ4tl9_ka14HVY')
        worksheet = sheet.worksheet('Отзывы')
        row = worksheet.find(answer_id).row
        worksheet.update(f'F{row}', answer)
        worksheet.update(f'H{row}', 'написать в поддержку')


    def get_supply_notifications(self):
        browser = webdriver.Chrome(executable_path='./chromedriver',options=self.chrome_options)
        self.get_ozon(browser)
        browser.quit()
        sheet = self.gc.open_by_key('11c6uAwJF1crfad7fpGsLbuC9U1pCMupkNxmv2BfSbxM')
        gworksheet = sheet.get_worksheet(0)
        date = datetime.today().strftime('%d.%m.%Y')
        workbook = xlrd.open_workbook(f"Demands_forecast_{date}.xlsx")
        needed = {}
        for name in self.need_names:
            try:
                worksheet = workbook.sheet_by_name(name)
                needed[name] = []
                row = 10
                while True:
                    try:
                        need = worksheet.cell_value(row , 6)
                        if need != 'Нет прогноза' and int(need) > 50:
                            art = worksheet.cell_value(row , 1)
                            others = worksheet.cell_value(row , 5)
                            product_name = gworksheet.find(art)
                            product_name = gworksheet.cell(product_name.row -1, product_name.col).value
                            needed[name].append(f'{product_name} {need} шт. (остаток на других складах {others} шт.)')
                        row += 1
                    except IndexError:
                        break
            except Exception:
                pass
        return needed

    def get_conversions_notifications(self):
        sheet = self.gc.open_by_key('11c6uAwJF1crfad7fpGsLbuC9U1pCMupkNxmv2BfSbxM')
        worksheet = sheet.get_worksheet(0)
        date, row = self.get_last_date_conversions(worksheet)
        wb_message = '⚡️ Рекомендую что-то изменить в карточках на маркетплейсах. У нас низкие конверсии и слабая динамика изменений у товаров:\nWildberries :\nэталон 10% 👉  40% 👉 25%\n\n'
        ozon_message = '\n\nOZON :\nэталон 50% 👉 5% 👉 40%\n\n'
        all_data = worksheet.get_all_values()
        for i in range(2,34):
            item_name = all_data[1][i]
            if item_name:
                all_conversion = all_data[row+6][i+1]
                ozon_all_conversion = all_data[row+6][i+1+32]
                if all_conversion and ozon_all_conversion and float(all_conversion[:-1].replace(',','.'))<7 and float(ozon_all_conversion[:-1].replace(',','.'))<2.5:
                    views = all_data[row][i+1]
                    clicks = all_data[row+1][i+1]
                    cart = all_data[row+2][i+1]
                    ozon_views = all_data[row][i+1+32]
                    ozon_clicks = all_data[row+1][i+1+32]
                    ozon_cart = all_data[row+2][i+1+32]
                    wb_message += f'- {item_name} :\n{views} 👉 {clicks} 👉 {cart}\n'
                    ozon_message += f'- {item_name} :\n{ozon_views} 👉 {ozon_clicks} 👉 {ozon_cart}\n'
        return wb_message + ozon_message

    def get_conversions(self):
        sheet = self.gc.open_by_key('11c6uAwJF1crfad7fpGsLbuC9U1pCMupkNxmv2BfSbxM')
        worksheet = sheet.get_worksheet(0)
        marketplaces = {"Wildberries": "Wildberries 10%\n10% 👉  40% 👉 25%","OZON":"OZON 5%\n50% 👉 5% 👉 40%"}
        date, row = self.get_last_date_conversions(worksheet)
        main_message = f'Конверсии за {date}\n\n📍 Эталон\nWildberries 10%\n10% 👉  40% 👉 25%\nOZON 5%\n50% 👉 5% 👉 40%\n'
        all_data = worksheet.get_all_values()
        for i in range(2,34):
            item_name = all_data[1][i]
            if item_name:
                product_message = f"\n\n{item_name}"
                views = all_data[row][i+1]
                clicks = all_data[row+1][i+1]
                cart = all_data[row+2][i+1]
                all_conversion = all_data[row+6][i+1]
                if views and clicks and cart and all_conversion:
                    product_message += f"\nWildberries {all_conversion}\n{views} 👉 {clicks} 👉 {cart}"
                else:
                    product_message += f"\nWildberries данных нет"
                ozon_views = all_data[row][i+1+32]
                ozon_clicks = all_data[row+1][i+1+32]
                ozon_cart = all_data[row+2][i+1+32]
                ozon_all_conversion = all_data[row+6][i+1+32]
                if ozon_views and ozon_clicks and ozon_cart and ozon_all_conversion:
                    product_message += f"\nOZON {ozon_all_conversion}\n{ozon_views} 👉 {ozon_clicks} 👉 {ozon_cart}\n"
                else:
                    product_message += f"\nOZON данных нет"
                main_message += product_message
        return main_message

    def add_to_db(self, db_sess:Session, notifications):
        if type(notifications) == str:
            db_notification = Notifications(text=notifications)
            db_sess.add(db_notification)
            db_sess.commit()
        elif type(notifications) == List:
            for notification in notifications:
                db_notification = Notifications(text=notification)
                db_sess.add(db_notification)
                db_sess.commit()

    def add_to_db_smart(self, db_sess:Session, notifications):
        db_notifications = db_sess.query(Notifications).all()
        for key in notifications.keys():
            old  = False
            for db_notification in db_notifications:
                if key in db_notification.text:
                    db_notification.text = key + notifications[key]
                    old = True
            if not old:
                db_notification = Notifications(text=key + notifications[key])
                db_sess.add(db_notification)
        db_sess.commit()

    def get_updates(self):
        db_sess = create_session()
        notifications = db_sess.query(Notifications).all()
        date, row = self.get_last_date_main()
        all_data = self.worksheet.get_all_values()
        for notification in notifications:
            if (datetime.now() - notification.date_added).days > 7:
                notifications.remove(notification)
                db_sess.delete(notification)
        notifications = []
        notifications.append(self.get_supply_notification(db_sess, all_data, row))
        notifications.append(self.get_search_pos_notification(db_sess, all_data, row))
        notifications.append(self.get_sell_pos_notification(db_sess, all_data, row))
        notifications.append(self.get_other_notification(db_sess, all_data, row))
        db_sess.commit()
        db_sess.close()
        return notifications

    def get_other_notification(self, db_sess, all_data, row):
        other_notifications = []
        for rating_notification in self.get_rating_notification(db_sess, all_data, row):
            other_notifications.append(rating_notification)
        for vk_and_inst_notification in self.get_vk_and_inst_notification(all_data, row):
            other_notifications.append(vk_and_inst_notification)
        return "\n\n".join(other_notifications)

    def get_rating_notification(self, db_sess, all_data, row):
        rating_notifications = []
        for product in self.products:
            alph_delta = self.products.index(product) * 15
            for marketplace in self.marketplaces:
                delta = self.marketplaces.index(marketplace)
                rating = float2(all_data[row+delta][29+alph_delta])
                rating_old = float2(all_data[row-5+delta][29+alph_delta])
                if rating and rating_old and rating_old > rating:
                    rating_notifications.append(f"⚡️ Внимание: у товара «{product}» упал рейтинг на {marketplace} с {rating_old} до {rating}")
        self.add_to_db(db_sess, rating_notifications)
        return rating_notifications

    def get_search_pos_notification(self, db_sess, all_data, row):
        search_pos_notifications = []
        for product in self.products:
            alph_delta = self.products.index(product) * 15
            for marketplace in self.marketplaces:
                delta = self.marketplaces.index(marketplace)
                search_pos = int2(all_data[row+delta][30+alph_delta])
                search_pos_old = int2(all_data[row-5+delta][30+alph_delta])
                if search_pos and search_pos_old and search_pos_old < search_pos:
                    key = re.search(re.compile(r'\".+\"'), self.worksheet.acell(self.keys_coords[product]).value).group(0).replace('"','')
                    search_pos_notifications.append(f"⚡️ Внимание: товар «{product}» упал в поиске на {marketplace} по запросу «{key}» с {search_pos_old} места на {search_pos} место")
        msg = "\n\n".join(search_pos_notifications)
        self.add_to_db(db_sess, msg)
        return msg
    
    def get_sell_pos_notification(self, db_sess:Session, all_data, row):
        sell_pos_notifications = []
        for product in self.products:
            alph_delta = self.products.index(product) * 15
            for marketplace in self.marketplaces:
                delta = self.marketplaces.index(marketplace)
                sell_pos = int2(all_data[row+delta][31+alph_delta])
                sell_pos_old = int2(all_data[row-5+delta][31+alph_delta])
                if sell_pos and sell_pos_old and sell_pos_old < sell_pos:
                    sell_pos_notifications.append(f"⚡️ Внимание: товар «{product}» стал продаваться хуже конкурентов на {marketplace}. Его рыночное место изменилось с {sell_pos_old} на {sell_pos}")
        msg = "\n\n".join(sell_pos_notifications)
        self.add_to_db(db_sess, msg)
        return msg

    def get_market_supply_notification(self, db_sess:Session, all_data, row):
        local_notifications = {}
        for product in self.products:
            alph_delta = self.products.index(product) * 15
            for marketplace in self.marketplaces:
                delta = self.marketplaces.index(marketplace)
                left_for = int2(all_data[row+delta][25+alph_delta])
                if left_for and left_for < 15:
                    local_notifications[f"⚡️ Внимание: нужно срочно поставить товар «{product}» на маркетплейс «{marketplace}». При текущей скорости продаж его хватит всего на "] = str(left_for) + ' суток'
        self.add_to_db_smart(db_sess, local_notifications)
        notification_list = []
        for key in local_notifications.keys():
            notification_list.append(key + local_notifications[key])
        return notification_list

    def get_supply_notification(self, db_sess, all_data, row):
        supply_notifications = []
        for market_supply_notification in self.get_market_supply_notification(db_sess, all_data, row):
            supply_notifications.append(market_supply_notification)
        for fabric_supply_notification in self.get_fabric_supply_notification(db_sess, all_data, row):
            supply_notifications.append(fabric_supply_notification)
        return supply_notifications

    def get_fabric_supply_notification(self, db_sess:Session, all_data, row):
        local_notifications = {}
        for product in self.products:
            alph_delta = self.products.index(product) * 15
            left_for = int2(all_data[row+4][25+alph_delta])
            if left_for and left_for < 50:
                local_notifications[f"⚡️ Внимание: нужно срочно заказать товар «{product}» на производстве. Он будет в дефиците на складах через "] = str(left_for)+' суток'
        self.add_to_db_smart(db_sess, local_notifications)
        notification_list = []
        for key in local_notifications.keys():
            notification_list.append(key + local_notifications[key])
        return notification_list
    
    def get_vk_and_inst_notification(self, all_data, row):
        vk_and_inst_notifications = []
        for product in self.products:
            alph_delta = self.products.index(product) * 15
            inst = int2(all_data[row+4][25+alph_delta])
            inst_old = int2(all_data[row+4][25+alph_delta])
            vk = int2(all_data[row+4][25+alph_delta])
            vk_old = int2(all_data[row+4][25+alph_delta])
            if inst and inst_old and inst_old > inst:
                vk_and_inst_notifications.append(f"⚡️ Внимание: упало число подписчиков в Instagram на {inst_old - inst} человек с {inst_old} до {inst}")
            if vk and vk_old and vk_old > vk:
                vk_and_inst_notifications.append(f"⚡️ Внимание: упало число подписчиков в ВКонтакте на {inst_old - inst} человек с {inst_old} до {inst}")
        return vk_and_inst_notifications

    def get_bills(self):
        all_data = self.worksheet_bills.get_all_values()[1:]
        data_dict = {}
        message = ''
        for row in all_data:
            if row[0] and row[1] and row[3]:
                date = datetime.strptime(row[3],'%d.%m.%Y')
                if date > datetime.now():
                    if row[1] in data_dict.keys():
                        data_dict[date].append(f'{row[1]} {row[0]}р')
                    else:
                        data_dict[date] = [f'{row[1]} {row[0]}р']
        for item in sorted(data_dict.items(), key=lambda p: p[0]):
            message += f"{item[0].strftime('%d.%m.%Y')} {self.weekdays[item[0].weekday()]} {item[1][0]} (осталось {(item[0] - datetime.now()).days} суток)\n"
        return message


    def get_roadmap(self):
        message = ''
        dates = ['Сегодня', 'Завтра','Послезавтра']
        for day in dates:
            delta = timedelta(days=dates.index(day))
            date = (datetime.today()+ delta).strftime('%d.%m.%Y')
            cells = self.worksheet_roadmap.findall(date)
            if not cells:
                message += f'{day} - задач не запланировано\n'
            else:
                local_message = ''
                for cell in cells:
                    if self.worksheet_roadmap.cell(cell.row,cell.col+2).value == 'FALSE':
                        name = ''
                        i = 0
                        while not name:
                            name = self.worksheet_roadmap.cell(cell.row+i,1).value
                            i -= 1
                        local_message += self.worksheet_roadmap.cell(cell.row,cell.col-1).value + f'({name})' + '\n'
                if local_message:
                    message += f'{day}:\n'+local_message
                else:
                    message += f'{day} - задач не запланировано\n'
        return message
    
    def get_marketing(self):
        all_data = self.worksheet.get_all_values()
        date, row = self.get_last_date_main()
        message = f'Данные продвижения за {date}\n\n'
        message += f"1️⃣ ВК подписчиков {int1(all_data[row][8])} ({get_plus(int1(all_data[row][8]) - int1(all_data[row-5][8]))})\n"
        message += f"2️⃣ Вк рекламный бюджет {int1(all_data[row][9])} ({get_plus(int1(all_data[row][9]) - int1(all_data[row-5][9]))})\n"
        message += f"3️⃣ Inst подписчиков {int1(all_data[row][10])} ({get_plus(int1(all_data[row][10]) - int1(all_data[row-5][10]))})\n\n"
        message += f"4️⃣ Баланс на всех картах {int1(all_data[row][11])} ({get_plus(int1(all_data[row][11]) - int1(all_data[row-5][11]))})\n"
        message += f"5️⃣ Удалили отзывов {int1(all_data[row][12])} ({get_plus(int1(all_data[row][12]) - int1(all_data[row-5][12]))})\n"
        message += f"6️⃣ Раздачи за отзывы {int1(all_data[row][14])} ({get_plus(int1(all_data[row][14]) - int1(all_data[row-5][14]))}) шт. на {int1(all_data[row][13])} р ({get_plus(int1(all_data[row][13]) - int1(all_data[row-5][13]))})\n"
        message += f"7️⃣ Прочие затраты {int1(all_data[row][15])} ({get_plus(int1(all_data[row][15]) - int1(all_data[row-5][15]))})"
        return message

    def get_crossplatform(self):
        all_data = self.worksheet.get_all_values()
        date, row = self.get_last_date_main()
        message = f'Кроссплатформенная аналитика за {date}\n\n'
        no_bloger_order_message = f'\n1️⃣ Чистые заказы без выкупов и блогеров\nВсего {int1(all_data[row+4][305])} ({get_plus(int1(all_data[row+4][305]) - int1(all_data[row-1][305]))}) шт. на {int1(all_data[row+4][304])} р ({get_plus(int1(all_data[row+4][304]) - int1(all_data[row-1][304]))})\n'
        price_message = f'\n2️⃣ Себестоимость\nВсего {int1(all_data[row+4][306])} ({get_plus(int1(all_data[row+4][306]) - int1(all_data[row-1][306]))}) р\n'
        income_message = f'\n3️⃣ Прибыль\nВсего {int1(all_data[row+4][307])} ({get_plus(int1(all_data[row+4][307]) - int1(all_data[row-1][307]))}) р\n'
        margin_message = f'\n4️⃣ Маржа\nВсего {float1(all_data[row+4][308][:-1])}% ({get_plus(float1(all_data[row+4][308][:-1]) - float1(all_data[row-1][308][:-1]))}%)\n'
        order_message = f'\n5️⃣ Заказов:\nВсего {int1(all_data[row+4][310])} ({get_plus(int1(all_data[row+4][310]) - int1(all_data[row-1][310]))}) шт. на {int1(all_data[row+4][309])} р ({get_plus(int1(all_data[row+4][309]) - int1(all_data[row-1][309]))})\n'
        buy_message = f'\n6️⃣ Выкупы всего {int1(all_data[row+4][312])} ({get_plus(int1(all_data[row+4][312]) - int1(all_data[row-1][312]))}) шт. на {int1(all_data[row+4][311])} р ({get_plus(int1(all_data[row+4][311]) - int1(all_data[row-1][311]))})\n\n'
        giveaway_message = f'\n7️⃣ Бесплатные раздачи всего {int1(all_data[row+4][314])} ({get_plus(int1(all_data[row+4][314]) - int1(all_data[row-1][314]))}) шт. на {int1(all_data[row+4][313])} р ({get_plus(int1(all_data[row+4][313]) - int1(all_data[row-1][313]))})\n\n'
        sell_pos_message = f'\n8️⃣ Конкурентная позиция в продажах:\nВсего {int1(all_data[row+4][322])} ({get_plus2(int1(all_data[row+4][322]) - int1(all_data[row-1][322]))})\n'
        search_pos_message = f'\n9️⃣ Конкурентная позиция в поиске:\nВсего {int1(all_data[row+4][321])} ({get_plus2(int1(all_data[row+4][321]) - int1(all_data[row-1][321]))})\n'
        reviews_message = f'\n🔟 Отзывов:\nВсего {int1(all_data[row+4][319])} ({get_plus(int1(all_data[row+4][319]) - int1(all_data[row-1][319]))})\n'
        rating_message =  f'\n1️⃣1️⃣ Рейтинг:\nВсего {int1(all_data[row+4][320])} ({get_plus(int1(all_data[row+4][320]) - int1(all_data[row-1][320]))})\n'
        left_message = f'\n1️⃣2️⃣ Остаток:\nВсего {int1(all_data[row+4][315])} ({get_plus(int1(all_data[row+4][315]) - int1(all_data[row-1][315]))})\n'
        enough_message = f'\n1️⃣3️⃣ Хватит \nВсего {int1(all_data[row+4][316])} ({get_plus(int1(all_data[row+4][316]) - int1(all_data[row-1][316]))})\n'
        warehouse_message = f'\n1️⃣4️⃣ ВСЕГО на Мой склад {int1(all_data[row+4][317])} ({get_plus(int1(all_data[row+4][317]) - int1(all_data[row-1][317]))})\n\n'
        for marketplace in self.marketplaces:
            delta = self.marketplaces.index(marketplace)
            no_bloger_order = int2(all_data[row+delta][305])
            if no_bloger_order:
                no_bloger_order_old = int1(all_data[row-5+delta][305])
                no_bloger_order_rub = int2(all_data[row+delta][304])
                no_bloger_order_rub_old = int1(all_data[row-5+delta][304])
                no_bloger_order_message += f'На {marketplace} {no_bloger_order} ({get_plus(no_bloger_order - no_bloger_order_old)}) шт. на {no_bloger_order_rub} р ({get_plus(no_bloger_order_rub - no_bloger_order_rub_old)})\n'
            price = int2(all_data[row+delta][306])
            if price:
                price_old = int1(all_data[row-5+delta][306])
                price_message += f'На {marketplace} {price} ({get_plus(price - price_old)}) р\n'
            income= int2(all_data[row+delta][307])
            if income:
                income_old = int1(all_data[row-5+delta][307])
                income_message += f'На {marketplace} {income} ({get_plus(income - income_old)}) р\n'
            margin= float2(all_data[row+delta][308][:-1])
            if margin:
                margin_old = float1(all_data[row-5+delta][308][:-1])
                margin_message += f'На {marketplace} {margin}% ({get_plus(margin - margin_old)}%) р\n'
            order = int2(all_data[row+delta][310])
            if order:
                order_old = int1(all_data[row-5+delta][310])
                order_rub = int2(all_data[row+delta][309])
                order_rub_old = int1(all_data[row-5+delta][310])
                order_message += f'На {marketplace} {order} ({get_plus(order - order_old)}) шт. на {order_rub} р ({get_plus(order_rub - order_rub_old)})\n'
            sell_pos = int2(all_data[row+delta][322])
            if sell_pos:
                sell_pos_old = int1(all_data[row-5+delta][322])
                sell_pos_message += f'{marketplace} {sell_pos} ({get_plus2(sell_pos - sell_pos_old)})\n'
            search_pos = int2(all_data[row+delta][321])
            if search_pos:
                search_pos_old = int1(all_data[row-5+delta][321])
                search_pos_message += f'{marketplace} {search_pos} ({get_plus2(search_pos - search_pos_old)})\n'
            reviews = float2(all_data[row+delta][319])
            if reviews:
                reviews_old = float1(all_data[row-5+delta][319])
                reviews_message += f'{marketplace} {reviews} ({get_plus(reviews - reviews_old)})\n'
            rating = float2(all_data[row+delta][320])
            if rating:
                rating_old = float1(all_data[row-5+delta][320])
                rating_message += f'{marketplace} {rating} ({get_plus(rating - rating_old)})\n'
            left = int2(all_data[row+delta][315])
            if left:
                left_old = int1(all_data[row-5+delta][315])
                left_message += f'{marketplace} {left} ({get_plus(left - left_old)})\n'
            enough = int2(all_data[row+delta][316])
            if enough:
                enough_old = int1(all_data[row-5+delta][316])
                enough_message += f'{marketplace} {enough} ({get_plus(enough - enough_old)})\n'
        message += no_bloger_order_message + price_message + income_message + margin_message + order_message + buy_message + giveaway_message + sell_pos_message + search_pos_message + reviews_message + rating_message + left_message + enough_message + warehouse_message
        return message

    def get_marketplace(self,marketplace):
        date, row = self.get_last_date_main()
        all_data = self.worksheet.get_all_values()
        row += self.marketplaces.index(marketplace)
        message = f'Всего на {marketplace} за {date}\n\n'
        order_message = f'1️⃣ Заказов {int1(all_data[row][310])} ({get_plus(int1(all_data[row][310]) - int1(all_data[row-5][310]))}) шт. на {int1(all_data[row][309])} р ({get_plus(int1(all_data[row][309]) - int1(all_data[row-5][309]))})\n\n'
        buy_message = f'\n2️⃣ Выкупы {int1(all_data[row][312])} ({get_plus(int1(all_data[row][312]) - int1(all_data[row-5][312]))})\n\n'
        giveaway_message = '\n3️⃣ Бесплатные раздачи\n\n'
        sell_pos_message = '\n4️⃣ Позиция в продажах\n\n'
        search_pos_message = '\n5️⃣ Позиция в поиске\n\n'
        reviews_message = '\n6️⃣ Отзывов\n\n'
        rating_message = f'\n7️⃣ Рейтинг {float1(all_data[row][320])} ({get_plus(float1(all_data[row][320]) - float1(all_data[row-5][320]))})\n\n'
        left_message = f'\n8️⃣ Остаток на {marketplace}\n\n'
        enough_message = '\n9️⃣ Хватит на дней\n\n'
        warehouse_message = '\n🔟 Склад\n\n'
        for product in self.products:
            alph_delta = self.products.index(product) * 15
            order = int2(all_data[row][19+alph_delta])
            if order:
                order_old = int1(all_data[row-5][19+alph_delta])
                order_rub = int2(all_data[row][18+alph_delta])
                order_rub_old = int1(all_data[row-5][18+alph_delta])
                order_message += f'{product} {order} ({get_plus(order - order_old)}) шт. на {order_rub} р ({get_plus(order_rub - order_rub_old)})\n'
            buy = int2(all_data[row][21+alph_delta])
            if buy:
                buy_old = int1(all_data[row-5][21+alph_delta])
                buy_rub = int2(all_data[row][20+alph_delta])
                buy_rub_old = int1(all_data[row-5][20+alph_delta])
                buy_message += f'{product} {buy} ({get_plus(buy - buy_old)}) шт. на {buy_rub} р ({get_plus(buy_rub - buy_rub_old)})\n'
            giveaway = int2(all_data[row][23+alph_delta])
            if giveaway:
                giveaway_old = int1(all_data[row-5][23+alph_delta])
                giveaway_rub = int2(all_data[row][22+alph_delta])
                giveaway_rub_old = int1(all_data[row-5][22+alph_delta])
                giveaway_message += f'{product} {giveaway} ({get_plus(giveaway - giveaway_old)}) шт. на {giveaway_rub} р ({get_plus(giveaway_rub - giveaway_rub_old)})\n'
            sell_pos = int2(all_data[row][31+alph_delta])
            if sell_pos:
                sell_pos_old = int1(all_data[row-5][31+alph_delta])
                sell_pos_message += f'{product} {sell_pos} ({get_plus2(sell_pos - sell_pos_old)})\n'
            search_pos = int2(all_data[row][30+alph_delta])
            if search_pos:
                search_pos_old = int1(all_data[row-5][30+alph_delta])
                search_pos_message += f'{product} {search_pos} ({get_plus2(search_pos - search_pos_old)})\n'
            reviews = int2(all_data[row][28+alph_delta])
            if reviews:
                reviews_old = int1(all_data[row-5][28+alph_delta])
                reviews_message += f'{product} {reviews} ({get_plus(reviews - reviews_old)})\n'
            rating = float2(all_data[row][29+alph_delta])
            if rating:
                rating_old = float1(all_data[row-5][29+alph_delta])
                rating_message += f'{product} {rating} ({get_plus(rating - rating_old)})\n'
            left = int2(all_data[row][24+alph_delta])
            if left:
                left_old = int1(all_data[row-5][24+alph_delta])
                left_message += f'{product} {left} ({get_plus(left - left_old)})\n'
            enough = int2(all_data[row][25+alph_delta])
            if enough:
                enough_old = int1(all_data[row-5][25+alph_delta])
                enough_message += f'{product} {enough} ({get_plus(enough - enough_old)})\n'
            warehouse = int2(all_data[row][26+alph_delta])
            if warehouse:
                warehouse_old = int1(all_data[row-5][26+alph_delta])
                warehouse_message += f'{product} {warehouse} ({get_plus(warehouse - warehouse_old)})\n'
        message += order_message + buy_message + giveaway_message + sell_pos_message + search_pos_message + reviews_message + rating_message + left_message + enough_message + warehouse_message
        return message

    def get_item(self,product):
        date, row = self.get_last_date_main()
        alph_delta = self.products.index(product) * 15
        all_data = self.worksheet.get_all_values()

        order_wb = int1(all_data[row][19 + alph_delta])
        order_ozon = int1(all_data[row+1][19 + alph_delta])
        order_yndx = int1(all_data[row+2][19 + alph_delta])
        order_ost = int1(all_data[row+3][19 + alph_delta])
        order_vsego = int1(all_data[row+4][19 + alph_delta])
        order_wb_old = int1(all_data[row-5][19 + alph_delta])
        order_ozon_old = int1(all_data[row-4][19 + alph_delta])
        order_yndx_old = int1(all_data[row-3][19 + alph_delta])
        order_ost_old = int1(all_data[row-2][19 + alph_delta])
        order_vsego_old = int1(all_data[row-1][19 + alph_delta])
        order_wb_rub = int1(all_data[row][18 + alph_delta])
        order_ozon_rub = int1(all_data[row+1][18 + alph_delta])
        order_yndx_rub = int1(all_data[row+2][18 + alph_delta])
        order_ost_rub = int1(all_data[row+3][18 + alph_delta])
        order_vsego_rub = int1(all_data[row+4][18 + alph_delta])
        order_wb_rub_old = int1(all_data[row-5][18 + alph_delta])
        order_ozon_rub_old = int1(all_data[row-4][18 + alph_delta])
        order_yndx_rub_old = int1(all_data[row-3][18 + alph_delta])
        order_ost_rub_old = int1(all_data[row-2][18 + alph_delta])
        order_vsego_rub_old = int1(all_data[row-1][18 + alph_delta])

        buy = int1(all_data[row][21 + alph_delta])
        buy_old = int1(all_data[row-5][21 + alph_delta])
        buy_rub = int1(all_data[row][20 + alph_delta])
        buy_rub_old = int1(all_data[row-5][20 + alph_delta])

        giveaway = int1(all_data[row][23 + alph_delta])
        giveaway_old = int1(all_data[row-5][23 + alph_delta])
        giveaway_rub = int1(all_data[row][22 + alph_delta])
        giveaway_rub_old = int1(all_data[row-5][22 + alph_delta])

        sell_pos_wb = int1(all_data[row][31 + alph_delta])
        sell_pos_ozon = int1(all_data[row+1][31 + alph_delta])
        sell_pos_yndx = int1(all_data[row+2][31 + alph_delta])
        sell_pos_ost = int1(all_data[row+3][31 + alph_delta])
        sell_pos_vsego = int1(all_data[row+4][31 + alph_delta])
        sell_pos_wb_old = int1(all_data[row-5][31 + alph_delta])
        sell_pos_ozon_old = int1(all_data[row-4][31 + alph_delta])
        sell_pos_yndx_old = int1(all_data[row-3][31 + alph_delta])
        sell_pos_ost_old = int1(all_data[row-2][31 + alph_delta])
        sell_pos_vsego_old = int1(all_data[row-1][31 + alph_delta])

        search_pos_wb = int1(all_data[row][30 + alph_delta])
        search_pos_ozon = int1(all_data[row+1][30 + alph_delta])
        search_pos_yndx = int1(all_data[row+2][30 + alph_delta])
        search_pos_ost = int1(all_data[row+3][30 + alph_delta])
        search_pos_vsego = int1(all_data[row+4][30 + alph_delta])
        search_pos_wb_old = int1(all_data[row-5][30 + alph_delta])
        search_pos_ozon_old = int1(all_data[row-4][30 + alph_delta])
        search_pos_yndx_old = int1(all_data[row-3][30 + alph_delta])
        search_pos_ost_old = int1(all_data[row-2][30 + alph_delta])
        search_pos_vsego_old = int1(all_data[row-1][30 + alph_delta])

        reviews_wb = int1(all_data[row][28 + alph_delta])
        reviews_ozon = int1(all_data[row+1][28 + alph_delta])
        reviews_yndx = int1(all_data[row+2][28 + alph_delta])
        reviews_ost = int1(all_data[row+2][28 + alph_delta])
        reviews_vsego = int1(all_data[row+4][28 + alph_delta])
        reviews_wb_old = int1(all_data[row-5][28 + alph_delta])
        reviews_ozon_old = int1(all_data[row-4][28 + alph_delta])
        reviews_yndx_old = int1(all_data[row-3][28 + alph_delta])
        reviews_ost_old = int1(all_data[row-2][28 + alph_delta])
        reviews_vsego_old = int1(all_data[row-1][28 + alph_delta])

        rating_wb = float1(all_data[row][29 + alph_delta])
        rating_ozon = float1(all_data[row+1][29 + alph_delta])
        rating_yndx = float1(all_data[row+2][29 + alph_delta])
        rating_ost = float1(all_data[row+3][29 + alph_delta])
        rating_vsego = float1(all_data[row+4][29 + alph_delta])
        rating_wb_old = float1(all_data[row-5][29 + alph_delta])
        rating_ozon_old = float1(all_data[row-4][29 + alph_delta])
        rating_yndx_old = float1(all_data[row-3][29 + alph_delta])
        rating_ost_old = float1(all_data[row-2][29 + alph_delta])
        rating_vsego_old = float1(all_data[row-1][29 + alph_delta])

        left_wb = int1(all_data[row][24 + alph_delta])
        left_ozon = int1(all_data[row+1][24 + alph_delta])
        left_yndx = int1(all_data[row+2][24 + alph_delta])
        left_ost = int1(all_data[row+3][24 + alph_delta])
        left_vsego = int1(all_data[row+4][24 + alph_delta])
        left_wb_old = int1(all_data[row-5][24 + alph_delta])
        left_ozon_old = int1(all_data[row-4][24 + alph_delta])
        left_yndx_old = int1(all_data[row-3][24 + alph_delta])
        left_ost_old = int1(all_data[row-2][24 + alph_delta])
        left_vsego_old = int1(all_data[row-1][24 + alph_delta])

        enough_wb = int1(all_data[row][25 + alph_delta])
        enough_ozon = int1(all_data[row+1][25 + alph_delta])
        enough_yndx = int1(all_data[row+2][25 + alph_delta])
        enough_ost = int1(all_data[row+3][25 + alph_delta])
        enough_vsego = int1(all_data[row+4][25 + alph_delta])
        enough_wb_old = int1(all_data[row-5][25 + alph_delta])
        enough_ozon_old = int1(all_data[row-4][25 + alph_delta])
        enough_yndx_old = int1(all_data[row-3][25 + alph_delta])
        enough_ost_old = int1(all_data[row-2][25 + alph_delta])
        enough_vsego_old = int1(all_data[row-1][25 + alph_delta])

        warehouse = int1(all_data[row][26 + alph_delta])
        warehouse_old = int1(all_data[row-5][26 + alph_delta])

        message = f'''
Аналитика по товару {product} за {date}
1️⃣ Заказов:
Всего  {order_vsego} ({get_plus(order_vsego - order_vsego_old)}) шт. на {order_vsego_rub} р ({get_plus(order_vsego_rub-order_vsego_rub_old)})
На Wildberries  {order_wb} ({get_plus(order_wb - order_wb_old)}) шт. на {order_wb_rub} р ({get_plus(order_wb_rub - order_wb_rub_old)})
На Ozon  {order_ozon} ({get_plus(order_ozon - order_ozon_old)}) шт. на {order_ozon_rub} р ({get_plus(order_ozon_rub - order_ozon_rub_old)})
На Yandex  {order_yndx} ({get_plus(order_yndx - order_yndx_old)}) шт. на {order_yndx_rub} р ({get_plus(order_yndx_rub - order_yndx_rub_old)})
На Остальное  {order_ost} ({get_plus(order_ost - order_ost_old)}) шт. на {order_ost_rub} р ({get_plus(order_ost_rub - order_ost_rub_old)})

2️⃣ Выкупы всего  {buy} ({get_plus(buy - buy_old)}) шт. на {buy_rub} р ({get_plus(buy_rub - buy_rub_old)})

3️⃣ Бесплатные раздачи всего  {giveaway} ({get_plus(giveaway - giveaway_old)}) шт. на {giveaway_rub} р ({get_plus(giveaway_rub - giveaway_rub_old)})

4️⃣ Конкурентная позиция в продажах: 
ВСЕГО {sell_pos_vsego} ({get_plus2(sell_pos_vsego - sell_pos_vsego_old)})
На Wildberries {sell_pos_wb} ({get_plus2(sell_pos_wb - sell_pos_wb_old)})
На Ozon {sell_pos_ozon} ({get_plus2(sell_pos_ozon - sell_pos_ozon_old)})
На Yandex {sell_pos_yndx} ({get_plus2(sell_pos_yndx - sell_pos_yndx_old)})
На Остальное {sell_pos_ost} ({get_plus2(sell_pos_ost- sell_pos_ost_old)})

5️⃣Конкурентная позиция в поиске:
ВСЕГО {search_pos_vsego} ({get_plus2(search_pos_vsego - search_pos_vsego_old)})
На Wildberries {search_pos_wb} ({get_plus2(search_pos_wb - search_pos_wb_old)})
На Ozon {search_pos_ozon} ({get_plus2(search_pos_ozon - search_pos_ozon_old)})
На Yandex {search_pos_yndx} ({get_plus2(search_pos_yndx - search_pos_yndx_old)})
На Остальное {search_pos_ost} ({get_plus2(search_pos_ost - search_pos_ost_old)})

6️⃣ Отзывов:
ВСЕГО {reviews_vsego} ({get_plus(reviews_vsego - reviews_vsego_old)})
На Wildberries {reviews_wb} ({get_plus(reviews_wb - reviews_wb_old)})
На Ozon {reviews_ozon} ({get_plus(reviews_ozon - reviews_ozon_old)})
На Yandex {reviews_yndx} ({get_plus(reviews_yndx - reviews_yndx_old)})
На Остальное {reviews_ost} ({get_plus(reviews_ost - reviews_ost_old)})

7️⃣ Рейтинг:
ВСЕГО {rating_vsego} ({get_plus(rating_vsego - rating_vsego_old)})
На Wildberries {rating_wb} ({get_plus(rating_wb - rating_wb_old)})
На Ozon {rating_ozon} ({get_plus(rating_ozon - rating_ozon_old)})
На Yandex {rating_yndx} ({get_plus(rating_yndx - rating_yndx_old)})
На Остальное {rating_ost} ({get_plus(rating_ost - rating_ost_old)})

8️⃣ Остаток:
ВСЕГО {left_vsego} ({get_plus(left_vsego - left_vsego_old)})
На Wildberries {left_wb} ({get_plus(left_wb - left_wb_old)})
На Ozon {left_ozon} ({get_plus(left_ozon - left_ozon_old)})
На Yandex {left_yndx} ({get_plus(left_yndx - left_yndx_old)})
На Остальное {left_ost} ({get_plus(left_ost - left_ost_old)})

9️⃣ Хватит 
ВСЕГО {enough_vsego} ({get_plus(enough_vsego - enough_vsego_old)})
На Wildberries {enough_wb} ({get_plus(enough_wb - enough_wb_old)})
На Ozon {enough_ozon} ({get_plus(enough_ozon - enough_ozon_old)})
На Yandex {enough_yndx} ({get_plus(enough_yndx - enough_yndx_old)})
На Остальное {enough_ost} ({get_plus(enough_ost - enough_ost_old)})

🔟 ВСЕГО на Мой склад  {warehouse} ({get_plus(warehouse - warehouse_old)})
'''
        return message

if __name__ == '__main__':
    g_sheets = Google_Sheets()
    loop = asyncio.get_event_loop()
    for i in g_sheets.warehouses.keys():
        print(i,"\n\n\n", loop.run_until_complete(g_sheets.get_warehouse_limits(g_sheets.warehouses[i])),"\n\n\n")
