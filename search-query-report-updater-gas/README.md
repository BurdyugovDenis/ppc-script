# Обновление отчетов по поисковым запросам

Google Apps Script выгружает SEARCH_QUERY_PERFORMANCE_REPORT из Яндекс Директа, агрегирует данные, хранит исторический лист и ежедневные новые запросы.

## Запуск

```bash
Откройте Google Sheets -> Apps Script, вставьте search_query_report_updater.gs. Сначала запустите initAllWithoutYesterday(), затем dailyUpdate() по триггеру.
```

## Входные файлы
- `Лист SETTINGS — пример в settings_example.csv`

## Выходные файлы
- `ALL_RAW`
- `YESTERDAY_RAW`
- `NEW_YESTERDAY`
- `LOG`

## Источник
- Локальная папка `Чистка поисковых запросов`

В репозитории оставлены только демонстрационные данные. Перед рабочим запуском замените примерные файлы своими выгрузками.
