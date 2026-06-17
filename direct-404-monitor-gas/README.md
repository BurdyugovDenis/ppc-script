# Проверка ссылок Директа на 404 в Google Sheets

Набор Google Apps Script-файлов выгружает активные ссылки объявлений и быстрых ссылок из Яндекс Директа, проверяет HTTP-коды чанками и отправляет найденные 404 в Telegram.

## Запуск

```bash
Создайте Google Sheet с листом config, вставьте .gs-файлы в Apps Script и настройте dailyActiveUrlsJob(), dailySitelinksJob(), dailyUrlCheckJob(), send404FromColumn().
```

## Входные файлы
- `Лист config — пример в config_example.csv`
- `Листы urls_unique, sitelinks_unique, content_ads_unique_hrefs`

## Выходные файлы
- `Коды ответов в Google Sheets`
- `Сообщение в Telegram при наличии 404`

## Источник
- https://vc.ru/marketing/2655479-avtomaticheskaya-proverka-silok-yandeks-direkta-na-404-s-otpravkoy-v-tg

## Важно
- BOT_TOKEN и CHAT_ID замените своими значениями или перенесите в Script Properties.

В репозитории оставлены только демонстрационные данные. Перед рабочим запуском замените примерные файлы своими выгрузками.
