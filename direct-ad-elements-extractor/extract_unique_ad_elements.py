import csv
import os
import sys


def print_debug_info():
    print("\n=== Диагностика ===")
    print(f"Текущая директория: {os.getcwd()}")
    print("Содержимое папки:")
    for file in os.listdir('.'):
        print(f"- {file}")
    print("===================\n")


def extract_unique_elements():
    print_debug_info()

    input_file = 'input.csv'

    if not os.path.isfile(input_file):
        raise FileNotFoundError(f"Файл {input_file} не найден")

    elements = {
        'Ссылка': set(),
        'Адреса быстрых ссылок': set(),
        'Заголовок 1': set(),
        'Заголовок 2': set(),
        'Текст': set(),
        'Отображаемая ссылка': set(),
        'Заголовки быстрых ссылок': set(),
        'Описания быстрых ссылок': set(),
        'Уточнения': set()
    }

    with open(input_file, 'r', encoding='utf-16') as csvfile:
        reader = csv.reader(csvfile, delimiter='\t')

        # Пропускаем первые две строки
        for _ in range(2):
            next(reader, None)

        headers = next(reader, None)
        if not headers:
            raise ValueError("Не найдена строка с заголовками (3-я строка файла)")

        normalized_headers = [h.strip().lower() for h in headers]

        try:
            columns = {
                'Ссылка': normalized_headers.index('ссылка'),
                'Адреса быстрых ссылок': normalized_headers.index('адреса быстрых ссылок'),
                'Заголовок 1': normalized_headers.index('заголовок 1'),
                'Заголовок 2': normalized_headers.index('заголовок 2'),
                'Текст': normalized_headers.index('текст'),
                'Отображаемая ссылка': normalized_headers.index('отображаемая ссылка'),
                'Заголовки быстрых ссылок': normalized_headers.index('заголовки быстрых ссылок'),
                'Описания быстрых ссылок': normalized_headers.index('описания быстрых ссылок'),
                'Уточнения': normalized_headers.index('уточнения')
            }
        except ValueError as e:
            raise Exception(f"Один из нужных столбцов не найден: {str(e)}")

        # Чтение строк с данными
        for row in reader:
            for col_name, index in columns.items():
                if len(row) <= index:
                    continue
                value = row[index].strip()
                if not value or value.lower() in {'', 'nan', 'none'}:
                    continue

                if col_name in ['Адреса быстрых ссылок', 'Заголовки быстрых ссылок',
                                'Описания быстрых ссылок', 'Уточнения']:
                    for item in value.split('||'):
                        cleaned = item.strip()
                        if cleaned:
                            elements[col_name].add(f"{cleaned};{col_name}")
                else:
                    elements[col_name].add(f"{value};{col_name}")

    # Объединяем, сортируем и сохраняем
    all_items = sorted([item for group in elements.values() for item in group])

    with open('output.txt', 'w', encoding='utf-8') as f:
        for item in all_items:
            f.write(f"{item}\n")


if __name__ == "__main__":
    try:
        extract_unique_elements()
        print("\nГотово! Результат в output.txt")
        print("Каждая строка: <значение>;<название столбца>")
    except Exception as e:
        print(f"\nОшибка: {str(e)}")
        print("\nПроверь:")
        print("- Названия столбцов в 3-й строке файла")
        print("- Что файл сохранён в кодировке UTF-16 (так указано в скрипте)")
        input("\nНажми Enter для выхода...")
        sys.exit(1)
