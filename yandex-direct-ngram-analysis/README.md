# N-gram анализ поисковых запросов

Строит Excel-отчет с исходными данными, словарем лемм, униграммами, биграммами и триграммами по каждому тегу.

## Запуск

```bash
python analyze_ngrams.py
```

## Входные файлы
- `input.csv — фраза;показы;клики;расход;конверсии;тег`

## Выходные файлы
- `output.xlsx — Excel-отчет`
- `output_example.csv — текстовый пример структуры одного листа`

## Зависимости
- `pandas`
- `openpyxl`
- `pymorphy3`

## Источник
- https://vc.ru/marketing/1854076-skripty-python-dlya-yandeks-direkta-n-gram-analiz-klasterizator-chastotnyi-slovar

В репозитории оставлены только демонстрационные данные. Перед рабочим запуском замените примерные файлы своими выгрузками.
