import csv
import os
import sys


def print_debug_info():
    print("\n=== Диагностика ===")
    print(f"Текущая директория: {os.getcwd()}")
    print("Содержимое папки:")

    for file_name in os.listdir("."):
        print(f"- {file_name}")

    print("===================\n")


def find_column(headers, column_name, occurrence=1):
    """
    Возвращает индекс указанного вхождения столбца.

    occurrence=1 — первое вхождение.
    occurrence=2 — второе вхождение.
    """
    matching_indices = [
        index
        for index, header in enumerate(headers)
        if header == column_name.lower()
    ]

    if len(matching_indices) < occurrence:
        raise ValueError(
            f'Столбец "{column_name}" встречается '
            f"{len(matching_indices)} раз, требуется вхождение №{occurrence}"
        )

    return matching_indices[occurrence - 1]


def extract_unique_elements():
    print_debug_info()

    input_file = "input.csv"
    output_file = "output.txt"

    if not os.path.isfile(input_file):
        raise FileNotFoundError(f"Файл {input_file} не найден")

    column_settings = [
        # Название в output.txt, название в CSV, номер вхождения
        ("Ссылка", "Ссылка", 1),
        ("Адреса быстрых ссылок", "Адреса быстрых ссылок", 1),

        # Берём второй столбец с названием «Заголовок 1»
        ("Заголовок 1", "Заголовок 1", 2),

        ("Заголовок 2", "Заголовок 2", 1),

        # Старый столбец «Текст» больше не используется
        ("Текст 1", "Текст 1", 1),
        ("Текст 2", "Текст 2", 1),
        ("Текст 3", "Текст 3", 1),

        ("Отображаемая ссылка", "Отображаемая ссылка", 1),
        ("Заголовки быстрых ссылок", "Заголовки быстрых ссылок", 1),
        ("Описания быстрых ссылок", "Описания быстрых ссылок", 1),
        ("Уточнения", "Уточнения", 1),
    ]

    multiple_value_columns = {
        "Адреса быстрых ссылок",
        "Заголовки быстрых ссылок",
        "Описания быстрых ссылок",
        "Уточнения",
    }

    elements = {
        output_name: set()
        for output_name, _, _ in column_settings
    }

    with open(input_file, "r", encoding="utf-16", newline="") as csvfile:
        reader = csv.reader(csvfile, delimiter="\t")

        # Пропускаем первые две строки
        for _ in range(2):
            next(reader, None)

        headers = next(reader, None)

        if not headers:
            raise ValueError(
                "Не найдена строка с заголовками — ожидается 3-я строка файла"
            )

        normalized_headers = [
            header.strip().lower()
            for header in headers
        ]

        columns = {}

        for output_name, csv_column_name, occurrence in column_settings:
            columns[output_name] = find_column(
                normalized_headers,
                csv_column_name,
                occurrence,
            )

        for row in reader:
            for column_name, column_index in columns.items():
                if column_index >= len(row):
                    continue

                value = row[column_index].strip()

                if not value or value.lower() in {"nan", "none"}:
                    continue

                if column_name in multiple_value_columns:
                    values = value.split("||")
                else:
                    values = [value]

                for item in values:
                    cleaned = item.strip()

                    if cleaned:
                        elements[column_name].add(
                            f"{cleaned};{column_name}"
                        )

    all_items = sorted(
        item
        for group in elements.values()
        for item in group
    )

    with open(output_file, "w", encoding="utf-8", newline="") as file:
        for item in all_items:
            file.write(f"{item}\n")

    print(f"\nГотово! Результат сохранён в {output_file}")
    print("Формат каждой строки: <значение>;<название столбца>")


if __name__ == "__main__":
    try:
        extract_unique_elements()

    except Exception as error:
        print(f"\nОшибка: {error}")
        print("\nПроверь:")
        print("- названия столбцов в 3-й строке файла;")
        print("- наличие двух столбцов «Заголовок 1»;")
        print("- наличие столбцов «Текст 1», «Текст 2» и «Текст 3»;")
        print("- что файл сохранён в кодировке UTF-16.")

        input("\nНажми Enter для выхода...")
        sys.exit(1)
