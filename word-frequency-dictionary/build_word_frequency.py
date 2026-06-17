import csv

# Читаем фразы из CSV-файла
input_filename = "phrases.csv"
output_filename = "word_counts.csv"

phrases = []

with open(input_filename, "r", encoding="utf-8") as file:
    reader = csv.reader(file)

    # Пропускаем заголовок (если есть)
    header = next(reader, None)

    # Собираем фразы из первого столбца
    for row in reader:
        if row:  # Игнорируем пустые строки
            phrases.append(row[0].strip())

# Подсчитываем слова
word_counts = {}
for phrase in phrases:
    for word in phrase.split():
        cleaned_word = word.lower().strip(".,!?()\"'")  # Очистка от знаков препинания
        if cleaned_word:
            word_counts[cleaned_word] = word_counts.get(cleaned_word, 0) + 1

# Сортируем слова по алфавиту
sorted_words = sorted(word_counts.items())

# Сохраняем результат в CSV
with open(output_filename, "w", encoding="utf-8", newline="") as file:
    writer = csv.writer(file)
    writer.writerow(["word", "count"])  # Заголовок

    for word, count in sorted_words:
        writer.writerow([word, count])

print(f"Результат сохранен в файл: {output_filename}")