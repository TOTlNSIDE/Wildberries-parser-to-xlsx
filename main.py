import requests
import pandas as pd
import time
import random
# import openpyxl

QUERY = "пальто из натуральной шерсти"
MAX_PAGES = 5
RETRIES = 15
DELAY_RANGE = (1, 5)
COLUMNS = [
    "Ссылка на товар",
    "Артикул",
    "Название",
    "Цена",
    "Описание",
    "Изображения",
    "Характеристики",
    "Селлер",
    "Ссылка на селлера",
    "Размеры",
    "Остатки",
    "Рейтинг",
    "Количество отзывов"
]
pd.DataFrame(columns=COLUMNS).to_csv("ParsedData.csv", index=False, encoding="utf-8-sig")

def get_page(page):
    url = f"https://search.wb.ru/exactmatch/ru/common/v18/search?appType=1&curr=rub&lang=ru&page={page}&query={QUERY}&resultset=catalog"

    for attempt in range(RETRIES):
        if attempt != 0:
            time.sleep(random.uniform(*DELAY_RANGE))

        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            return data

    return None

def get_basket(vol):
    if 0 <= vol <= 143:
        return "01"
    elif vol <= 287:
        return "02"
    elif vol <= 431:
        return "03"
    elif vol <= 719:
        return "04"
    elif vol <= 1007:
        return "05"
    elif vol <= 1061:
        return "06"
    elif vol <= 1115:
        return "07"
    elif vol <= 1169:
        return "08"
    elif vol <= 1313:
        return "09"
    elif vol <= 1601:
        return "10"
    elif vol <= 1655:
        return "11"
    elif vol <= 1919:
        return "12"
    elif vol <= 2045:
        return "13"
    elif vol <= 2189:
        return "14"
    elif vol <= 2405:
        return "15"
    elif vol <= 2621:
        return "16"
    elif vol <= 2837:
        return "17"
    elif vol <= 3053:
        return "18"
    elif vol <= 3269:
        return "19"
    elif vol <= 3485:
        return "20"
    elif vol <= 3701:
        return "21"
    elif vol <= 3917:
        return "22"
    elif vol <= 4133:
        return "23"
    elif vol <= 4349:
        return "24"
    elif vol <= 4565:
        return "25"
    elif vol <= 4877:
        return "26"
    elif vol <= 5189:
        return "27"
    elif vol <= 5501:
        return "28"
    elif vol <= 5813:
        return "29"
    elif vol <= 6125:
        return "30"
    elif vol <= 6437:
        return "31"
    elif vol <= 6749:
        return "32"
    elif vol <= 7061:
        return "33"
    elif vol <= 7373:
        return "34"
    elif vol <= 7685:
        return "35"
    elif vol <= 7997:
        return "36"
    elif vol <= 8309:
        return "37"
    elif vol <= 8741:
        return "38"
    elif vol <= 9173:
        return "39"
    elif vol <= 9605:
        return "40"
    else:
        return "41"

def get_card(id):
    vol = id // 100000
    part = id // 1000
    number = get_basket(vol)
    url = f"https://basket-{number}.wbbasket.ru/vol{vol}/part{part}/{id}/info/ru/card.json"

    for attempt in range(RETRIES):
        if attempt != 0:
            time.sleep(random.uniform(*DELAY_RANGE))

        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            return data

    return None

def format_card(card):
    groups = card.get("grouped_options")
    lines = []

    for group in groups:
        group_name = group.get("group_name", "Без группы")

        lines.append(f"[{group_name}]")

        for option in group.get("options", []):
            name = option.get("name", "")
            value = option.get("value", "")

            lines.append(f"{name}: {value}")

        lines.append("")

    return "\n".join(lines).strip()

def get_images(id, count):
    vol = id // 100000
    part = id // 1000
    number = get_basket(vol)
    images = []

    for c in range(1, count - 1):
        url = f"https://basket-{number}.wbbasket.ru/vol{vol}/part{part}/{id}/images/big/{c}.webp"

        images.append(url)

    return ", ".join(images)

def get_detail(id):
    cookies = {
        'x_wbaas_token': '1.1000.8bf1be1176c8458d8e45d59f2320e142.MHw3OC4xMTEuMTU1LjIzM3xNb3ppbGxhLzUuMCAoV2luZG93cyBOVCAxMC4wOyBXaW42NDsgeDY0KSBBcHBsZVdlYktpdC81MzcuMzYgKEtIVE1MLCBsaWtlIEdlY2tvKSBDaHJvbWUvMTQ2LjAuMC4wIFNhZmFyaS81MzcuMzZ8MTc3NTQxMjYwMXxyZXVzYWJsZXwyfGV5Sm9ZWE5vSWpvaUluMD18MHwzfDE3NzQ4MDc4MDF8MQ==.MEQCIB8Xs1U/X3sR1wNa9Eg/qZzp9CjTnvh1Ouo9/gzmEELJAiB+HKZEN7plFwzHTrXnxGZPZQiVeySbfpbUrJazQE59Xw==',
        '_wbauid': '7306236281774203003',
    }
    params = {
        'appType': '1',
        'curr': 'rub',
        'dest': '-1257786',
        'spp': '30',
        'ab_testing': 'false',
        'lang': 'ru',
        'nm': f'{id}',
    }
    headers = {
        'accept': '*/*',
        'accept-language': 'ru-RU,ru;q=0.9',
        'deviceid': 'site_4c7fe2cdff46471a98bdb3e9449e92d8',
        'priority': 'u=1, i',
        'referer': f'https://www.wildberries.ru/catalog/{id}/detail.aspx',
        'sec-ch-ua': '"Chromium";v="146", "Not-A.Brand";v="24", "Google Chrome";v="146"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36',
        'x-requested-with': 'XMLHttpRequest',
        'x-spa-version': '14.2.3',
    }

    for attempt in range(RETRIES):
        if attempt != 0:
            time.sleep(random.uniform(*DELAY_RANGE))

        response = requests.get(
            "https://www.wildberries.ru/__internal/u-card/cards/v4/detail",
            params=params,
            cookies=cookies,
            headers=headers
        )
        if response.status_code == 200:
            data = response.json()
            return data

    return None

for page in range(1, MAX_PAGES + 1):
    print(f"\nСтраница {page}")

    if page != 1:
        time.sleep(random.uniform(*DELAY_RANGE))

    data = get_page(page)
    if not data:
        print(f"Пропуск страницы {page}")
        continue

    products = data.get("products")
    if not products:
        print(f"Товаров нет")
        continue

    print(f"Найдено товаров: {len(products)}")

    items = []

    for product in products:
        id = product.get("id")
        product_url = f"https://www.wildberries.ru/catalog/{id}/detail.aspx"
        product_detail = get_detail(id).get("products", [{}])[0]
        sizes = [
            size.get("name")
            for size in product.get("sizes", [])
            if size.get("name")
        ]
        card = get_card(id)
        description = card.get("description", "")
        characteristics = format_card(card)
        item = {
            "Ссылка на товар": product_url,
            "Артикул": id,
            "Название": product.get("name"),
            "Цена": product.get("sizes", [{}])[0].get("price", {}).get("product", 0) / 100,
            "Описание": description,
            "Изображения": get_images(id, card.get("media").get("photo_count")),
            "Характеристики": characteristics,
            "Селлер": product.get("supplier"),
            "Ссылка на селлера": f"https://www.wildberries.ru/seller/{product.get('supplierId')}",
            "Размеры": ", ".join(sizes),
            "Остатки": product_detail.get("totalQuantity", 0),
            "Рейтинг": product_detail.get("reviewRating", 0),
            "Количество отзывов": product_detail.get("feedbacks", 0)
        }

        items.append(item)

    df = pd.DataFrame(items)
    df.to_csv("ParsedData.csv", mode="a", header=False, index=False, encoding="utf-8-sig")
    print(f"Сохранено {len(items)} товаров")

df = pd.read_csv("ParsedData.csv", encoding="utf-8-sig")
df.to_excel("ParsedData.xlsx", index=False)
print("\nПарсинг завершён")