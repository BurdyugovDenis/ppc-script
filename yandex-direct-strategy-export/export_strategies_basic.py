import os
import requests
import pandas as pd

TOKEN = os.getenv("YANDEX_DIRECT_TOKEN", "")
CLIENT_LOGIN = os.getenv("YANDEX_DIRECT_CLIENT_LOGIN", "")
API_URL = "https://api.direct.yandex.com/json/v5/campaigns"

if not TOKEN or not CLIENT_LOGIN:
    raise RuntimeError("Set YANDEX_DIRECT_TOKEN and YANDEX_DIRECT_CLIENT_LOGIN")

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Client-Login": CLIENT_LOGIN,
    "Accept-Language": "ru",
    "Content-Type": "application/json"
}

body = {
    "method": "get",
    "params": {
        "SelectionCriteria": {
            "Types": ["TEXT_CAMPAIGN"]  # Исправлена опечатка: TEXT_CAMPAIN → TEXT_CAMPAIGN
        },
        "FieldNames": [
            "Id",
            "Name",
            "Status",
            "State"
        ],
        "TextCampaignFieldNames": [
            "BiddingStrategy"
        ]
    }
}

response = requests.post(API_URL, headers=headers, json=body)

if response.status_code != 200:
    print(f"Ошибка HTTP: {response.status_code}")
    print(response.text)
    exit()

data = response.json()

if 'error' in data:
    print("Ошибка API:")
    print(data['error'])
    exit()

if 'result' not in data:
    print("В ответе отсутствует ключ 'result'")
    print(data)
    exit()


# Функция для преобразования статуса с учетом State
def get_status_text(status, state):
    # Если статус ARCHIVED - всегда в архиве
    if status == "ARCHIVED":
        return "в архиве"

    # Если статус ENDED и состояние CONVERTED - в архиве
    if status == "ENDED" and state == "CONVERTED":
        return "в архиве"

    # Если статус ACCEPTED и состояние ON - активна
    if status == "ACCEPTED" and state == "ON":
        return "активна"

    # Во всех остальных случаях - приостановлена
    return "приостановлена"


rows = []
for camp in data.get("result", {}).get("Campaigns", []):
    text_campaign = camp.get("TextCampaign")
    if not text_campaign:
        continue

    bidding_strategy = text_campaign.get("BiddingStrategy", {})

    # Получаем текстовое представление статуса с учетом State
    status = camp.get("Status")
    state = camp.get("State")
    status_text = get_status_text(status, state)

    item = {
        "CampaignId": camp.get("Id"),
        "Название": camp.get("Name"),
        "Статус активности": status_text,
        "Поиск": bidding_strategy.get("Search"),
        "Сеть": bidding_strategy.get("Network")
    }
    rows.append(item)

if rows:
    df = pd.DataFrame(rows)
    df.to_csv("direct_strategies.csv", index=False, encoding='utf-8-sig')
    print(f"Успешно выгружено {len(rows)} кампаний")
    print("Данные сохранены в direct_strategies.csv")
else:
    print("Нет данных для сохранения")