import csv
import pymorphy3
import os
from collections import defaultdict

morph = pymorphy3.MorphAnalyzer()


def process_phrase(phrase):
    words = phrase.split()
    processed_words = []
    for word in words:
        clean_word = word.strip('.,!?;:()[]{}"\'')
        if not clean_word:
            continue

        parsed = morph.parse(clean_word)[0]
        if 'PREP' in parsed.tag:
            continue

        processed_words.append(parsed.normal_form)

    return ' '.join(processed_words)


def load_excluded_words():
    excluded = set()
    if os.path.exists('excluded_words.csv'):
        with open('excluded_words.csv', 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=';')
            next(reader)  # Пропускаем заголовок
            for row in reader:
                if row:
                    excluded.add(row[0].strip())
    return excluded


def main():
    excluded_words = load_excluded_words()

    # Чтение и обработка данных
    rows = []
    with open('input.csv', 'r', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile, delimiter=';')
        for row in reader:
            if len(row) < 2:
                continue
            original_phrase = row[0].strip()
            count_str = row[1].strip()
            try:
                count = int(count_str)
            except ValueError:
                continue
            processed_phrase = process_phrase(original_phrase)
            if processed_phrase:
                rows.append({
                    'original': original_phrase,
                    'processed': processed_phrase,
                    'count': count
                })

    # Расчет весов слов
    word_weights = defaultdict(float)
    for row in rows:
        words = row['processed'].split()
        word_count = len(words)
        if word_count == 0:
            continue
        freq_share = row['count'] / word_count
        for word in words:
            word_weights[word] += freq_share

    # Определение групп с учетом исключений
    group_freq = defaultdict(int)
    for row in rows:
        words = [w for w in row['processed'].split() if w not in excluded_words]
        sorted_words = sorted(words, key=lambda w: (-word_weights.get(w, 0), w))
        group = ' | '.join(sorted_words) if sorted_words else 'No Group'
        row['group'] = group
        group_freq[group] += row['count']

    # Подготовка данных для записи
    output_data = []
    for row in rows:
        words = row['processed'].split()
        valid_words = [w for w in words if w not in excluded_words]
        freq_share = row['count'] / len(valid_words) if len(valid_words) > 0 else 0
        output_data.append({
            'group': row['group'],
            'original': row['original'],
            'processed': row['processed'],
            'count': row['count'],
            'freq_share': freq_share,
            'group_total': group_freq[row['group']]
        })

    # Сортировка по группе и запись
    output_data.sort(key=lambda x: x['group'])

    with open('output.csv', 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile, delimiter=';')
        writer.writerow(['Group', 'Original Phrase', 'Processed Phrase', 'Frequency', 'Frequency Share', 'Group Total'])
        for item in output_data:
            writer.writerow([
                item['group'],
                item['original'],
                item['processed'],
                item['count'],
                int(round(item['freq_share'], 0)),
                item['group_total']
            ])

    # Запись файла весов ключевых слов
    sorted_words = sorted(word_weights.items(), key=lambda x: (-x[1], x[0]))
    with open('value_keyword.csv', 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile, delimiter=';')
        writer.writerow(['Keyword', 'Weight'])
        for word, weight in sorted_words:
            writer.writerow([word, int(round(weight, 0))])


if __name__ == "__main__":
    main()