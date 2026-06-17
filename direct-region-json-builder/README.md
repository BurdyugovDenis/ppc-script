# Формирование JSON регионов Директа

Преобразует CSV со строками `id,регион` в JSON-массив формата, который можно вставлять в геотаргетинг Яндекс Директа.

## Запуск

```bash
python build_direct_region_json.py
```

## Входные файлы
- `input.csv — `id,name` без заголовка`

## Выходные файлы
- `output.txt — JSON-массив регионов`

## Источник
- https://vc.ru/marketing/2143954-skript-na-python-dlya-formirovaniya-formata-yandeks-direkta-po-regionam

В репозитории оставлены только демонстрационные данные. Перед рабочим запуском замените примерные файлы своими выгрузками.
