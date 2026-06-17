import pandas as pd
import os


def filter_platforms(input_path, output_path, remove_zero=True):
    # Определяем реальный разделитель в CSV файле
    with open(input_path, 'r', encoding='utf-8-sig') as f:
        first_line = f.readline().strip()
        delimiter = ';' if ';' in first_line else ','
        print(f"Определен разделитель входного файла: {repr(delimiter)}")

    # Читаем файл с правильным разделителем
    try:
        df = pd.read_csv(input_path, delimiter=delimiter, encoding='utf-8-sig')  # Исправлено здесь
        print(f"Файл успешно прочитан, колонок: {len(df.columns)}")
    except Exception as e:
        print(f"Ошибка при чтении файла: {e}")
        return

    # Переименовываем колонки, если они содержат лишние пробелы
    df.columns = df.columns.str.strip()

    # Проверка наличия необходимых столбцов
    required_columns = ['Площадка', 'Расход (руб.)']
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        print(f"Ошибка: В файле отсутствуют необходимые колонки: {', '.join(missing)}")
        print(f"Найденные колонки: {', '.join(df.columns)}")
        print("Попытка нормализовать названия колонок...")

        # Попробуем найти колонки по похожим названиям
        platform_col = next((col for col in df.columns if 'площад' in col.lower()), None)
        cost_col = next((col for col in df.columns if 'расход' in col.lower() and 'руб' in col.lower()), None)

        if platform_col and cost_col:
            print(f"Найдены альтернативные колонки: '{platform_col}', '{cost_col}'")
            df = df.rename(columns={
                platform_col: 'Площадка',
                cost_col: 'Расход (руб.)'
            })
            print("Колонки переименованы")
        else:
            print("Не удалось найти подходящие колонки")
            return

    # Преобразуем расход в числа
    df['Расход (руб.)'] = df['Расход (руб.)'].astype(str).str.replace(',', '.').str.replace(' ', '').astype(float)

    # Удаляем строки с нулевым расходом (если выбрано)
    original_count = len(df)
    if remove_zero:
        df = df[df['Расход (руб.)'] > 0]
        print(f"Удалено строк с нулевым расходом: {original_count - len(df)}")
    else:
        print("Сохранение площадок с нулевым расходом")

    print(f"Осталось строк для обработки: {len(df)}")

    # Маски для поиска (в нижнем регистре)
    masks = [
        'dsp', 'game', 'vpn', 'ege', 'oge', 'gdz',
        'puzzle', 'soliter', 'com.'
    ]

    # Исключающие подстроки
    exclude = ['vk', 'dzen', 'zen', 'youla', 'edadeal']

    # Функция для проверки площадки
    def should_include(platform):
        if not isinstance(platform, str):
            return False

        platform_lower = platform.lower()

        # Проверка на исключения
        if any(exc in platform_lower for exc in exclude):
            return False

        # Проверка на совпадение с масками
        return any(mask in platform_lower for mask in masks)

    # Фильтрация данных
    filtered = df[df['Площадка'].apply(should_include)]
    print(f"Найдено площадок по маскам: {len(filtered)}")

    # Выбор нужных столбцов и сортировка
    result = filtered[['Площадка', 'Расход (руб.)']]
    result = result.sort_values('Расход (руб.)', ascending=False)

    # Преобразуем расход в строки с запятой вместо точки
    result['Расход (руб.)'] = result['Расход (руб.)'].apply(
        lambda x: f"{x:.2f}".replace('.', ',') if pd.notnull(x) else "0,00"
    )

    # Сохранение результатов с разделителем ; и запятой в числах
    result.to_csv(output_path, sep=';', index=False, header=False)
    print(f"Результат сохранен в: {os.path.abspath(output_path)}")
    print(f"Всего отфильтровано площадок: {len(result)}")
    print("Первые 5 строк результата:")
    print(result.head(5).to_string(index=False))


def ask_yes_no(question, default=True):
    """Задает вопрос с ответом Д/н и возвращает boolean"""
    choices = " [Д/н] " if default else " [д/Н] "
    prompt = question + choices

    while True:
        response = input(prompt).strip().lower()
        if response == '':
            return default
        elif response in ['д', 'y', 'yes']:
            return True
        elif response in ['н', 'n', 'no']:
            return False
        else:
            print("Пожалуйста, ответьте 'д' или 'н'")


if __name__ == "__main__":
    # Пути относительно папки проекта
    project_dir = os.path.dirname(os.path.abspath(__file__))
    input_file = os.path.join(project_dir, 'input.csv')
    output_file = os.path.join(project_dir, 'output.csv')

    # Проверка существования входного файла
    if not os.path.exists(input_file):
        print(f"Ошибка: Входной файл не найден по пути: {input_file}")
        print("Пожалуйста, создайте файл input.csv в папке проекта")
        input("Нажмите Enter для выхода...")
        exit()

    # Запрос о нулевых расходах
    print("\n" + "=" * 50)
    print(" Фильтрация площадок по маскам")
    print("=" * 50)
    remove_zero = ask_yes_no("Удалять площадки с нулевым расходом?", default=True)

    # Запускаем обработку
    print(f"\nОбработка файла: {os.path.basename(input_file)}")
    filter_platforms(input_file, output_file, remove_zero)

    # Завершение работы
    print("\nОбработка завершена успешно!")
    input("Нажмите Enter для выхода...")