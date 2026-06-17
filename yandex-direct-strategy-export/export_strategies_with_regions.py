import os
import requests
import pandas as pd

TOKEN = os.getenv("YANDEX_DIRECT_TOKEN", "")
CLIENT_LOGIN = os.getenv("YANDEX_DIRECT_CLIENT_LOGIN", "")
API_URL = "https://api.direct.yandex.com/json/v5/campaigns"
API_URL_ADGROUPS = "https://api.direct.yandex.com/json/v5/adgroups"
API_URL_DICTIONARIES = "https://api.direct.yandex.com/json/v5/dictionaries"

if not TOKEN or not CLIENT_LOGIN:
    raise RuntimeError("Set YANDEX_DIRECT_TOKEN and YANDEX_DIRECT_CLIENT_LOGIN")

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Client-Login": CLIENT_LOGIN,
    "Accept-Language": "ru",
    "Content-Type": "application/json"
}


# Функция для выполнения запросов к API
def make_api_request(url, body):
    response = requests.post(url, headers=headers, json=body)
    if response.status_code != 200:
        raise Exception(f"Ошибка HTTP {response.status_code}: {response.text}")
    data = response.json()
    if 'error' in data:
        raise Exception(f"Ошибка API: {data['error']}")
    return data


# Получаем список кампаний
campaigns_body = {
    "method": "get",
    "params": {
        "SelectionCriteria": {
            "Types": ["TEXT_CAMPAIGN"]
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

campaigns_data = make_api_request(API_URL, campaigns_body)
campaigns = campaigns_data.get('result', {}).get('Campaigns', [])

# Получаем справочник регионов
dictionaries_body = {
    "method": "get",
    "params": {
        "DictionaryNames": ["GeoRegions"]
    }
}
regions_data = make_api_request(API_URL_DICTIONARIES, dictionaries_body)
geo_regions = regions_data.get('result', {}).get('GeoRegions', [])
region_id_to_name = {region['RegionId']: region['RegionName'] for region in geo_regions}

# Получаем группы объявлений для всех кампаний
campaign_ids = [campaign['Id'] for campaign in campaigns]
adgroups_body = {
    "method": "get",
    "params": {
        "SelectionCriteria": {
            "CampaignIds": campaign_ids
        },
        "FieldNames": ["CampaignId", "RegionIds"],
        "Page": {"Limit": 10000}
    }
}

adgroups_data = make_api_request(API_URL_ADGROUPS, adgroups_body)
adgroups = adgroups_data.get('result', {}).get('AdGroups', [])

# Собираем регионы для каждой кампании
campaign_regions = {}
for adgroup in adgroups:
    campaign_id = adgroup['CampaignId']
    region_ids = adgroup.get('RegionIds', [])
    if campaign_id not in campaign_regions:
        campaign_regions[campaign_id] = set()
    campaign_regions[campaign_id].update(region_ids)


# Функция для преобразования статуса
def get_status_text(status, state):
    if status == "ARCHIVED":
        return "в архиве"
    if status == "ENDED" and state == "CONVERTED":
        return "в архиве"
    if status == "ACCEPTED" and state == "ON":
        return "активна"
    return "приостановлена"


# Формируем данные для выгрузки
rows = []
for campaign in campaigns:
    text_campaign = campaign.get("TextCampaign")
    if not text_campaign:
        continue

    bidding_strategy = text_campaign.get("BiddingStrategy", {})
    status = campaign.get("Status")
    state = campaign.get("State")
    status_text = get_status_text(status, state)

    # Получаем регионы для кампании
    region_ids = campaign_regions.get(campaign['Id'], set())
    region_names = [region_id_to_name.get(rid, f"Неизвестный регион ({rid})") for rid in region_ids]
    regions_str = ", ".join(region_names) if region_names else "Регионы не указаны"

    item = {
        "CampaignId": campaign.get("Id"),
        "Название": campaign.get("Name"),
        "Статус активности": status_text,
        "Поиск": bidding_strategy.get("Search"),
        "Сеть": bidding_strategy.get("Network"),
        "Регионы": regions_str  # Добавляем регионы
    }
    rows.append(item)

# Сохраняем в CSV
if rows:
    df = pd.DataFrame(rows)
    df.to_csv("direct_strategies.csv", index=False, encoding='utf-8-sig')
    print(f"Успешно выгружено {len(rows)} кампаний")
    print("Данные сохранены в direct_strategies.csv")
else:
    print("Нет данных для сохранения")