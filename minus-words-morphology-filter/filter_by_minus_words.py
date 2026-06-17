import pymorphy3
import re


def load_minus_words(filename):
    """Загрузка и парсинг минус-слов с учетом кавычек и квадратных скобок"""
    simple_words = []  # обычные слова
    quoted_phrases = []  # фразы в кавычках
    bracketed_phrases = []  # фразы в квадратных скобках

    with open(filename, 'r', encoding='utf-8') as file:
        for line in file:
            line = line.strip()
            if not line:
                continue

            # Обработка фраз в кавычках
            if line.startswith('"') and line.endswith('"'):
                phrase = line[1:-1].strip()
                if phrase:
                    quoted_phrases.append(phrase)
            # Обработка фраз в квадратных скобках
            elif line.startswith('[') and line.endswith(']'):
                phrase = line[1:-1].strip()
                if phrase:
                    bracketed_phrases.append(phrase)
            # Обработка обычных слов
            else:
                simple_words.append(line)

    return simple_words, quoted_phrases, bracketed_phrases


def normalize_word(word, morph):
    """Нормализация слова до его леммы"""
    cleaned = re.sub(r'[^\w-]', '', word.lower())
    if cleaned:
        parsed = morph.parse(cleaned)[0]
        return parsed.normal_form
    return word


def check_simple_words(phrase, simple_lemmas):
    """Проверка на наличие простых минус-слов в фразе"""
    words = re.findall(r'[\w-]+', phrase.lower())
    for word in words:
        if word in simple_lemmas:
            return True
    return False


def check_quoted_phrases(phrase, quoted_lemmas_list, morph):
    """Проверка фраз в кавычках (точное соответствие, но с любыми окончаниями)"""
    phrase_words = [normalize_word(word, morph) for word in re.findall(r'[\w-]+', phrase.lower())]

    for quoted_lemmas in quoted_lemmas_list:
        # Проверяем, содержит ли фраза все слова из минус-фразы в любом порядке
        if all(lemma in phrase_words for lemma in quoted_lemmas):
            return True
    return False


def check_bracketed_phrases(phrase, bracketed_lemmas_list, morph):
    """Проверка фраз в квадратных скобках (строгий порядок слов с любыми окончаниями)"""
    phrase_words = [normalize_word(word, morph) for word in re.findall(r'[\w-]+', phrase.lower())]

    for bracketed_lemmas in bracketed_lemmas_list:
        # Ищем точное соответствие последовательности лемм
        for i in range(len(phrase_words) - len(bracketed_lemmas) + 1):
            if phrase_words[i:i + len(bracketed_lemmas)] == bracketed_lemmas:
                return True
    return False


def main():
    morph = pymorphy3.MorphAnalyzer()

    # Загрузка и обработка минус-слов
    simple_words, quoted_phrases, bracketed_phrases = load_minus_words('minus.txt')

    # Нормализация простых слов
    simple_lemmas = set()
    for word in simple_words:
        lemma = normalize_word(word, morph)
        if lemma:
            simple_lemmas.add(lemma)

    # Нормализация фраз в кавычках
    quoted_lemmas_list = []
    for phrase in quoted_phrases:
        lemmas = [normalize_word(word, morph) for word in re.findall(r'[\w-]+', phrase.lower())]
        if lemmas:
            quoted_lemmas_list.append(lemmas)

    # Нормализация фраз в квадратных скобках
    bracketed_lemmas_list = []
    for phrase in bracketed_phrases:
        lemmas = [normalize_word(word, morph) for word in re.findall(r'[\w-]+', phrase.lower())]
        if lemmas:
            bracketed_lemmas_list.append(lemmas)

    # Обработка входных фраз
    with open('input.txt', 'r', encoding='utf-8') as input_file:
        phrases = [line.strip() for line in input_file if line.strip()]

    result_phrases = []
    white_phrases = []

    for phrase in phrases:
        # Проверяем все типы минус-слов
        has_simple = check_simple_words(phrase, simple_lemmas)
        has_quoted = check_quoted_phrases(phrase, quoted_lemmas_list, morph)
        has_bracketed = check_bracketed_phrases(phrase, bracketed_lemmas_list, morph)

        if has_simple or has_quoted or has_bracketed:
            result_phrases.append(phrase)
        else:
            white_phrases.append(phrase)

    # Сохранение результатов
    with open('result.txt', 'w', encoding='utf-8') as res_file:
        res_file.write('\n'.join(result_phrases))

    with open('white.txt', 'w', encoding='utf-8') as white_file:
        white_file.write('\n'.join(white_phrases))

    print(f"Обработано фраз: {len(phrases)}")
    print(f"Отфильтровано фраз: {len(result_phrases)}")
    print(f"Чистых фраз: {len(white_phrases)}")


if __name__ == '__main__':
    main()