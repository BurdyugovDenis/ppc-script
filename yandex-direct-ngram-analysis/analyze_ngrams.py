import string
import re
import pandas as pd
from collections import defaultdict
import pymorphy3
from itertools import combinations  # Импортируем модуль для генерации комбинаций

# Инициализация анализатора pymorphy3
morph = pymorphy3.MorphAnalyzer()

# Список русских предлогов
PREPOSITIONS = {
    'в', 'во', 'на', 'за', 'под', 'над', 'с', 'со', 'из', 'у', 'от', 'до',
    'по', 'о', 'об', 'обо', 'про', 'к', 'ко', 'перед', 'при', 'через',
    'сквозь', 'между', 'для', 'без', 'ради', 'из-за', 'из-под', 'близ',
    'вблизи', 'вглубь', 'вдоль', 'возле', 'вокруг', 'впереди', 'после',
    'внутри', 'около', 'посредь', 'среди', 'благодаря', 'ввиду', 'вместо',
    'вне', 'насчёт', 'вроде', 'выше', 'далее', 'для', 'до', 'ежели', 'затем',
    'кроме', 'меж', 'наверху', 'ниже', 'относительно', 'помимо', 'после',
    'прежде', 'против', 'путём', 'сверх', 'снизу', 'согласно', 'соответственно',
    'спустя', 'среди', 'супротив', 'у', 'черех', 'и'
}


def clean_word(word):
    """Удаляет пунктуацию и приводит к нижнему регистру."""
    chars = string.punctuation + '«»„“'
    return word.strip(chars).lower()


def remove_prepositions(phrase):
    """Удаляет предлоги из фразы."""
    if not isinstance(phrase, str):
        return ''
    words = phrase.split()
    cleaned_words = [word for word in words if clean_word(word) not in PREPOSITIONS]
    return ' '.join(cleaned_words)


def normalize_word(word):
    """Нормализует слово с использованием pymorphy3 и пользовательских исключений."""
    CUSTOM_LEMMAS = {
        'дети': 'дети',
        'детей': 'дети',
        'детишек': 'дети',
        'детей11': 'дети'
    }

    match = re.match(r"^(\W*)(.*?)(\W*)$", word)
    if not match:
        return word
    prefix, stem, suffix = match.groups()

    if not stem:
        return word

    was_title = stem[0].isupper()
    lower_word = stem.lower()

    if lower_word in CUSTOM_LEMMAS:
        normalized_stem = CUSTOM_LEMMAS[lower_word]
    else:
        parsed = morph.parse(lower_word)
        normalized_stem = parsed[0].normal_form if parsed else lower_word

    if was_title:
        normalized_stem = normalized_stem.capitalize()
    else:
        normalized_stem = normalized_stem.lower()

    return f"{prefix}{normalized_stem}{suffix}"


def normalize_phrase(phrase):
    """Приводит морфологию фразы к нормальной форме."""
    if not isinstance(phrase, str):
        return ''
    tokens = re.findall(r"\S+|\s+", phrase)
    processed = []
    for token in tokens:
        if re.match(r"^\s*$", token):
            processed.append(token)
        else:
            processed.append(normalize_word(token))
    return "".join(processed)


def create_result_df(stats_dict):
    """Создает DataFrame из словаря статистики."""
    df = pd.DataFrame([
        {'Слово': gram, **stats} for gram, stats in stats_dict.items()
    ]).rename(columns={
        'показы': 'Показы', 'клики': 'Клики', 'расход': 'Расход', 'конверсии': 'Конверсии'
    })

    numeric_cols = ['Показы', 'Клики', 'Расход', 'Конверсии']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).round(0).astype(int)

    df['CPL'] = df.apply(
        lambda x: '-' if x['Конверсии'] == 0 else round(x['Расход'] / x['Конверсии']),
        axis=1
    )

    df = df.sort_values('Расход', ascending=False)
    return df


def process_csv(input_file, output_file):
    """Обрабатывает CSV и сохраняет результаты в Excel с разными листами."""
    df = pd.read_csv(input_file, sep=';', decimal=',', engine='python')
    df.columns = df.columns.str.lower()
    required_columns = ['фраза', 'показы', 'клики', 'расход', 'конверсии', 'тег']
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Отсутствуют колонки: {', '.join(missing)}")

    # Приводим 'фраза' к строкам, чтобы избежать float NaN проблем
    df['фраза'] = df['фраза'].astype(str).fillna('')

    # Приводим числовые поля к числам
    for col in ['показы', 'клики', 'расход', 'конверсии']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Добавляем столбец с лемматизированной фразой
    df['лемматизированная_фраза'] = df['фраза'].apply(normalize_phrase)

    # Переупорядочиваем столбцы: фраза -> лемматизированная_фраза -> остальные
    original_columns = df.columns.tolist()
    new_order = ['фраза', 'лемматизированная_фраза'] + [col for col in original_columns if col not in {'фраза', 'лемматизированная_фраза'}]
    df = df[new_order]

    # Собираем словарь исходных слов и их лемм
    all_cleaned_words = set()
    for phrase in df['фраза']:
        words = phrase.split()
        for word in words:
            cleaned = clean_word(word)
            if cleaned:
                all_cleaned_words.add(cleaned)

    word_lemma_pairs = []
    for word in all_cleaned_words:
        if word in PREPOSITIONS:
            lemma = word
        else:
            normalized = normalize_word(word)
            lemma = clean_word(normalized)
        word_lemma_pairs.append((word, lemma))

    lemma_df = pd.DataFrame(word_lemma_pairs, columns=['Исходное слово', 'Лемма'])
    lemma_df = lemma_df.drop_duplicates().sort_values('Исходное слово')

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Исходные данные', index=False)
        lemma_df.to_excel(writer, sheet_name='Словарь лемм', index=False)

        tags = df['тег'].unique()

        for raw_tag in tags:
            if pd.isna(raw_tag):
                tag = 'no_tag'
            else:
                tag = str(raw_tag)
            safe_tag = re.sub(r'[\\/*?:"<>|]', '_', tag).replace(' ', '_')[:31]

            if pd.isna(raw_tag):
                tag_df = df[df['тег'].isna()].copy()
            else:
                tag_df = df[df['тег'] == raw_tag].copy()

            if tag_df.empty:
                continue

            tag_df['обработанная_фраза'] = tag_df['фраза'].apply(remove_prepositions).apply(normalize_phrase)

            word_stats = defaultdict(lambda: {'показы': 0, 'клики': 0, 'расход': 0.0, 'конверсии': 0})
            bigram_stats = defaultdict(lambda: {'показы': 0, 'клики': 0, 'расход': 0.0, 'конверсии': 0})
            trigram_stats = defaultdict(lambda: {'показы': 0, 'клики': 0, 'расход': 0.0, 'конверсии': 0})

            for _, row in tag_df.iterrows():
                processed_phrase = row['обработанная_фраза']
                words = processed_phrase.split()
                cleaned_words = [clean_word(word) for word in words if clean_word(word)]

                # 1-граммы (оставляем как есть)
                for word in cleaned_words:
                    word_stats[word]['показы'] += row['показы']
                    word_stats[word]['клики'] += row['клики']
                    word_stats[word]['расход'] += row['расход']
                    word_stats[word]['конверсии'] += row['конверсии']

                # 2-граммы: все комбинации по 2 слова (порядок не важен)
                if len(cleaned_words) >= 2:
                    for combo in combinations(cleaned_words, 2):
                        # Сортируем комбинацию для единообразия
                        sorted_combo = sorted(combo)
                        bigram = ' '.join(sorted_combo)
                        bigram_stats[bigram]['показы'] += row['показы']
                        bigram_stats[bigram]['клики'] += row['клики']
                        bigram_stats[bigram]['расход'] += row['расход']
                        bigram_stats[bigram]['конверсии'] += row['конверсии']

                # 3-граммы: все комбинации по 3 слова (порядок не важен)
                if len(cleaned_words) >= 3:
                    for combo in combinations(cleaned_words, 3):
                        sorted_combo = sorted(combo)
                        trigram = ' '.join(sorted_combo)
                        trigram_stats[trigram]['показы'] += row['показы']
                        trigram_stats[trigram]['клики'] += row['клики']
                        trigram_stats[trigram]['расход'] += row['расход']
                        trigram_stats[trigram]['конверсии'] += row['конверсии']

            if word_stats:
                df_1gram = create_result_df(word_stats)
                sheet_name = f"{safe_tag}_1gram"[:31]
                df_1gram.to_excel(writer, sheet_name=sheet_name, index=False)

            if bigram_stats:
                df_2gram = create_result_df(bigram_stats)
                sheet_name = f"{safe_tag}_2gram"[:31]
                df_2gram.to_excel(writer, sheet_name=sheet_name, index=False)

            if trigram_stats:
                df_3gram = create_result_df(trigram_stats)
                sheet_name = f"{safe_tag}_3gram"[:31]
                df_3gram.to_excel(writer, sheet_name=sheet_name, index=False)


if __name__ == '__main__':
    process_csv('input.csv', 'output.xlsx')