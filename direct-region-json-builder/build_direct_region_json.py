import csv

# Чтение данных из CSV
with open('input.csv', 'r', encoding='utf-8') as csv_file:
    csv_reader = csv.reader(csv_file)
    data = []
    for row in csv_reader:
        if len(row) < 2:  # Пропускаем неполные строки
            continue

        # Формируем строку в нужном формате
        formatted_str = f'{{"id":"{row[0].strip()}","name":"{row[1].strip()}"}}'
        data.append(formatted_str)

# Сохранение в TXT
with open('output.txt', 'w', encoding='utf-8') as txt_file:
    # Записываем все элементы через запятую
    txt_file.write('[' + ','.join(data) + ']')

print("Конвертация завершена! Результат в output.txt")