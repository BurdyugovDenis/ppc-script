# Проверка релевантности выдачи Яндекса

Selenium-скрипт открывает выдачу Яндекса по запросам и считает упоминания основного слова/словоформ в топе результатов.

## Запуск

```bash
ANTICAPTCHA_KEY=your_key python check_yandex_serp_mentions.py
```

## Входные файлы
- `input.csv — колонки main;Query`

## Выходные файлы
- `output.csv — main, Query, Mentions`

## Зависимости
- `selenium`
- `webdriver-manager`
- `anticaptchaofficial`
- `requests`
- `pymorphy3`

## Источник
- https://vc.ru/marketing/2174767-proverka-relevantnosti-vydachi-yandeksa-na-python

## Важно
- Ключ AntiCaptcha берется из переменной окружения `ANTICAPTCHA_KEY`.

В репозитории оставлены только демонстрационные данные. Перед рабочим запуском замените примерные файлы своими выгрузками.
