import os
import re
import ssl
from dataclasses import dataclass
from typing import List, Tuple, Dict, Set, Iterable

import pandas as pd
import pymorphy3

import nltk
from nltk.stem import WordNetLemmatizer


# ================== НАСТРОЙКИ ==================
LIST1_PATH = "list1.csv"
LIST2_PATH = "list2.csv"
OUT_PATH = "result_matches.csv"

# Названия колонок
COL_QUERY = "query"   # список 1
COL_TERM = "term"     # список 2

# CSV параметры
CSV_SEP = ","         # если у вас ';' — поменяйте на ';'
CSV_ENCODING = "utf-8"

MIN_TOKEN_LEN = 2

# Фразы:
# False: совпадение последовательности лемм (строго)
# True: совпадение "все леммы фразы есть в запросе" (шире)
ALLOW_BAG_OF_WORDS_PHRASE_MATCH = False

SAVE_QUERY_LEMMAS = True

# Куда скачивать NLTK-данные
NLTK_DATA_DIR = os.path.expanduser("~/nltk_data")
# ===============================================


_RE_TOKEN = re.compile(r"[A-Za-z]+|[А-Яа-яЁё]+|\d+", flags=re.UNICODE)

morph_ru = pymorphy3.MorphAnalyzer(lang="ru")
wnl = WordNetLemmatizer()


def disable_ssl_verification_for_downloads():
    """
    Workaround: отключает SSL-проверку для HTTPS-контекста (чтобы nltk.download не падал).
    Ровно тот паттерн с ssl._create_unverified_context [web:54].
    """
    try:
        _create_unverified_https_context = ssl._create_unverified_context
    except AttributeError:
        return
    else:
        ssl._create_default_https_context = _create_unverified_https_context  # [web:54]


def ensure_nltk_data_dir():
    os.makedirs(NLTK_DATA_DIR, exist_ok=True)
    if NLTK_DATA_DIR not in nltk.data.path:
        nltk.data.path.append(NLTK_DATA_DIR)  # NLTK ищет ресурсы по nltk.data.path [web:45]


def ensure_wordnet() -> bool:
    """
    Пытается гарантировать наличие wordnet + omw-1.4.
    Возвращает True если wordnet доступен, иначе False (тогда EN-лемматизация будет "как есть").
    """
    ensure_nltk_data_dir()

    try:
        nltk.data.find("corpora/wordnet")
        return True
    except LookupError:
        pass

    # Отключаем SSL-проверку и пробуем скачать [web:54]
    disable_ssl_verification_for_downloads()

    try:
        nltk.download("wordnet", download_dir=NLTK_DATA_DIR, quiet=True)
        nltk.download("omw-1.4", download_dir=NLTK_DATA_DIR, quiet=True)
        nltk.data.find("corpora/wordnet")
        return True
    except Exception:
        return False


WORDNET_OK = ensure_wordnet()


def tokenize(text: str) -> List[str]:
    if not isinstance(text, str):
        return []
    return _RE_TOKEN.findall(text.lower())


def is_cyrillic(token: str) -> bool:
    return bool(re.fullmatch(r"[а-яё]+", token))


def is_latin(token: str) -> bool:
    return bool(re.fullmatch(r"[a-z]+", token))


def lemma_ru(token: str) -> str:
    return morph_ru.parse(token)[0].normal_form


def lemma_en(token: str) -> str:
    if not WORDNET_OK:
        return token

    # API: lemmatize(word, pos=...), допустимые pos перечислены в документации [web:9]
    cands = [
        wnl.lemmatize(token, pos="n"),
        wnl.lemmatize(token, pos="v"),
        wnl.lemmatize(token, pos="a"),
        wnl.lemmatize(token, pos="r"),
    ]
    return min(cands, key=len)


def lemmatize_tokens(tokens: Iterable[str]) -> List[str]:
    out: List[str] = []
    for t in tokens:
        if len(t) < MIN_TOKEN_LEN:
            continue
        if is_cyrillic(t):
            out.append(lemma_ru(t))
        elif is_latin(t):
            out.append(lemma_en(t))
        else:
            out.append(t)
    return out


def phrase_match_window(query_lemmas: List[str], term_lemmas: List[str]) -> bool:
    m = len(term_lemmas)
    if m == 0 or m > len(query_lemmas):
        return False
    for i in range(len(query_lemmas) - m + 1):
        if query_lemmas[i:i + m] == term_lemmas:
            return True
    return False


def phrase_match_bow(query_set: Set[str], term_lemmas: List[str]) -> bool:
    if not term_lemmas:
        return False
    return all(l in query_set for l in term_lemmas)


@dataclass(frozen=True)
class TermItem:
    term_raw: str
    term_lemmas: Tuple[str, ...]


def build_terms_index(terms: Iterable[str]) -> List[TermItem]:
    index: List[TermItem] = []
    for term in terms:
        term = str(term).strip()
        if not term:
            continue
        lemmas = tuple(lemmatize_tokens(tokenize(term)))
        if not lemmas:
            continue
        index.append(TermItem(term_raw=term, term_lemmas=lemmas))
    return index


def match_one_query(query: str, terms_index: List[TermItem]) -> Tuple[List[Dict], List[str]]:
    q_lemmas = lemmatize_tokens(tokenize(query))
    q_set = set(q_lemmas)

    matches: List[Dict] = []
    for item in terms_index:
        t_lem = list(item.term_lemmas)

        if len(t_lem) == 1:
            if t_lem[0] in q_set:
                matches.append({
                    "term": item.term_raw,
                    "term_lemmas": " ".join(t_lem),
                    "match_type": "single_lemma_in_query",
                })
        else:
            if ALLOW_BAG_OF_WORDS_PHRASE_MATCH:
                ok = phrase_match_bow(q_set, t_lem)
                mt = "phrase_bow"
            else:
                ok = phrase_match_window(q_lemmas, t_lem)
                mt = "phrase_window"

            if ok:
                matches.append({
                    "term": item.term_raw,
                    "term_lemmas": " ".join(t_lem),
                    "match_type": mt,
                })

    return matches, q_lemmas


def read_csv_smart(path: str) -> pd.DataFrame:
    try:
        return pd.read_csv(path, sep=CSV_SEP, encoding=CSV_ENCODING)
    except Exception:
        return pd.read_csv(path, sep=";", encoding=CSV_ENCODING)


def main():
    df_q = read_csv_smart(LIST1_PATH)
    df_t = read_csv_smart(LIST2_PATH)

    if COL_QUERY not in df_q.columns:
        raise ValueError(f"В {LIST1_PATH} нет колонки '{COL_QUERY}'. Есть: {list(df_q.columns)}")
    if COL_TERM not in df_t.columns:
        raise ValueError(f"В {LIST2_PATH} нет колонки '{COL_TERM}'. Есть: {list(df_t.columns)}")

    terms = df_t[COL_TERM].dropna().astype(str).tolist()
    terms_index = build_terms_index(terms)

    out_rows: List[Dict] = []
    queries = df_q[COL_QUERY].fillna("").astype(str).tolist()

    for row_id, query in enumerate(queries, start=1):
        matches, q_lemmas = match_one_query(query, terms_index)

        row = {
            "row_id": row_id,
            "query": query,
            "matched": 1 if matches else 0,
            "terms": " | ".join(m["term"] for m in matches) if matches else "",
            "terms_lemmas": " | ".join(m["term_lemmas"] for m in matches) if matches else "",
            "match_types": " | ".join(m["match_type"] for m in matches) if matches else "",
        }
        if SAVE_QUERY_LEMMAS:
            row["query_lemmas"] = " ".join(q_lemmas)

        out_rows.append(row)

    pd.DataFrame(out_rows).to_csv(OUT_PATH, index=False, encoding="utf-8")
    print(f"OK: saved -> {OUT_PATH}")
    if not WORDNET_OK:
        print("NOTE: WordNet не доступен, английские слова не лемматизировались (матчинг по исходным токенам).")


if __name__ == "__main__":
    main()
