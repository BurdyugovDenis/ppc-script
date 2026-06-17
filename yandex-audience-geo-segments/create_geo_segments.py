import os
import requests
import json
import sys
import io
import time
import math
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

TOKEN = os.getenv("YANDEX_AUDIENCE_TOKEN", "")

headers = {
    "Authorization": f"OAuth {TOKEN}",
    "Content-Type": "application/json"
}


def read_all_locations_from_file(file_path):
    """Читает все координаты из файла"""
    locations = []
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if line:
                    locations.append(line)
        print(f"Прочитано {len(locations)} координат из файла")
        return locations
    except Exception as e:
        print(f"Ошибка при чтении файла: {e}")
        return []


def create_geo_segment(segment_name, locations_batch):
    """Создает гео-сегмент для пакета координат"""
    url = "https://api-audience.yandex.ru/v1/management/segments/create_geo"

    points = []
    for loc in locations_batch:
        parts = loc.split(',')
        if len(parts) < 2:
            continue

        try:
            lat = float(parts[0].strip())
            lon = float(parts[1].strip())

            if not (-90 <= lat <= 90) or not (-180 <= lon <= 180):
                continue

            points.append({
                "latitude": lat,
                "longitude": lon
            })
        except ValueError:
            continue

    if not points:
        print(f"Пакет '{segment_name}' не содержит валидных точек")
        return None

    payload = {
        "segment": {
            "name": segment_name,
            "geo_segment_type": "condition",
            "radius": 500,
            "points": points,
            "times_quantity": 3,
            "period_length": 30
        }
    }

    try:
        response = requests.post(
            url,
            headers=headers,
            json=payload
        )

        if response.status_code in (200, 201):
            response_data = response.json()
            segment_id = response_data.get("id")
            print(f"Сегмент '{segment_name}' создан. ID: {segment_id} | Точек: {len(points)}")
            return True
        else:
            print(f"Ошибка при создании сегмента '{segment_name}': {response.status_code}")
            print("Тело ответа:", response.text)
            return False
    except Exception as e:
        print(f"Ошибка при отправке запроса '{segment_name}': {e}")
        return False


def process_locations_in_batches(locations, batch_size=1000):
    """Обрабатывает координаты пакетами с последовательной нумерацией"""
    if not locations:
        print("Нет координат для обработки")
        return

    total_points = len(locations)
    batches = math.ceil(total_points / batch_size)
    print(f"Всего точек: {total_points} | Пакетов: {batches} | Размер пакета: {batch_size}")

    # Добавляем префикс с датой для уникальности
    date_prefix = datetime.now().strftime("%Y%m%d")
    base_name = "GeoSegment_sanatorii_rf"

    successful_segments = 0
    for i in range(batches):
        start_idx = i * batch_size
        end_idx = min((i + 1) * batch_size, total_points)
        batch = locations[start_idx:end_idx]

        # Простое последовательное имя: GeoSegment_1, GeoSegment_2 и т.д.
        segment_name = f"{base_name}_{i + 1}"

        print(f"\nОбработка пакета {i + 1}/{batches} [{start_idx + 1}-{end_idx}] -> '{segment_name}'...")

        if create_geo_segment(segment_name, batch):
            successful_segments += 1

        # Пауза между запросами
        if i < batches - 1:
            time.sleep(1)

    print(f"\nОбработка завершена. Успешно создано сегментов: {successful_segments}/{batches}")


if __name__ == "__main__":
    file_path = "locations.txt"
    all_locations = read_all_locations_from_file(file_path)

    if all_locations:
        process_locations_in_batches(all_locations, batch_size=1000)
    else:
        print("Не удалось прочитать координаты из файла.")