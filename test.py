from datetime import datetime, timedelta
import aiohttp, asyncio
warehouses = {
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

async def get_warehouse_limits(warehouse_id : int):
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
    weekend_str = (monday + timedelta(days=90)).strftime("%Y-%m-%d")
    payload = '{{"params":{{"dateFrom":"{}","dateTo":"{}","warehouseId":{}}},"jsonrpc":"2.0","id":"json-rpc_20"}}'.format(monday_str, weekend_str, warehouse_id)
    async with aiohttp.ClientSession(headers=headers) as session:
        response = await session.post(url=warehouses_url, data=payload, headers=headers)
        responce_json = await response.json()
        return responce_json

if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    for i in warehouses.keys():
        print(i,"\n\n\n", loop.run_until_complete(get_warehouse_limits(warehouses[i])),"\n\n\n")