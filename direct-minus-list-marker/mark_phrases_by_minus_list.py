#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv
import io
import os
import re
from typing import List, Tuple, Optional, Set

# ===== ФАЙЛЫ =====
CSV_INPUT = "direct.csv"
CSV_OUTPUT = "output.csv"
MINUS_TXT = "minus_words.txt"

# ===== КОЛОНКИ =====
PHRASE_COL_PREFERRED = "Фраза (с минус-словами)"
PHRASE_COL_FALLBACK = "Фраза"
PARAM1_COL_NAME_STRICT = "Параметр 1"

# Что пишем в "Параметр 1"
SET_VALUE = "удалить"

# возможные разделители (у тебя был \t)
DELIMS = ["\t", ";", ",", "|"]

# оставляем только буквы/цифры/пробел
_re_keep = re.compile(r"[^0-9a-zа-яё\s]+", re.IGNORECASE)

# РЕЖЕМ “минус-хвост” только по шаблону: пробел + -слово (после '-' не '*')
_re_minus_token_space = re.compile(r"\s-[^\s*][^\s]*")

def normalize_minus_line(s: str) -> str:
    """Нормализация строки из minus_words.txt (без отрезания по ' -слово')."""
    if s is None:
        return ""
    s = str(s).lower()
    if "-*" in s:
        s = s.split("-*", 1)[0]
    s = _re_keep.sub(" ", s)
    s = " ".join(s.split()).strip()
    return s

def clean_phrase(s: str) -> str:
    """
    Нормализация фразы из CSV:
    - lower
    - отрезать всё после '-*'
    - отрезать всё начиная с первого ' -слово'
    - убрать спецсимволы, схлопнуть пробелы
    """
    if s is None:
        return ""
    s = str(s).lower()

    if "-*" in s:
        s = s.split("-*", 1)[0]

    m = _re_minus_token_space.search(s)
    if m:
        s = s[:m.start()].strip()

    s = _re_keep.sub(" ", s)
    s = " ".join(s.split()).strip()
    return s

def load_minus_set(path: str) -> Set[str]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Не найден файл минус-слов: {path}")

    out = set()
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            w = normalize_minus_line(line)
            if w:
                out.add(w)
    return out

def detect_encoding(path: str) -> str:
    with open(path, "rb") as f:
        start = f.read(4)
    if start.startswith(b"\xff\xfe") or start.startswith(b"\xfe\xff"):
        return "utf-16"
    if start.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    return "utf-8"

def detect_delimiter_by_count(line: str) -> str:
    best = "\t"
    best_n = -1
    for d in DELIMS:
        n = line.count(d)
        if n > best_n:
            best_n = n
            best = d
    return best

def pick_table_header(lines: List[str]) -> Tuple[int, str]:
    best_i = 0
    best_cols = 0
    best_delim = "\t"

    scan_limit = min(len(lines), 5000)
    for i in range(scan_limit):
        line = lines[i]
        if not line.strip():
            continue
        delim = detect_delimiter_by_count(line)
        cols = line.count(delim) + 1
        if cols > best_cols:
            best_cols = cols
            best_i = i
            best_delim = delim

    if best_cols < 5:
        raise ValueError("Не смог найти табличную часть: нигде нет строки хотя бы с 5 колонками.")
    return best_i, best_delim

def k(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

def split_row(raw: str, delim: str) -> List[str]:
    return next(csv.reader([raw], delimiter=delim, quotechar='"', escapechar='\\'))

def join_row(row: List[str], delim: str) -> str:
    s = io.StringIO()
    w = csv.writer(s, delimiter=delim, quotechar='"', lineterminator="")
    w.writerow(row)
    return s.getvalue()

def find_phrase_idx(header: List[str]) -> Optional[int]:
    hk = [k(x) for x in header]

    pref = k(PHRASE_COL_PREFERRED)
    if pref in hk:
        return hk.index(pref)

    for i, name in enumerate(hk):
        if "фраз" in name and "минус" in name:
            return i

    fb = k(PHRASE_COL_FALLBACK)
    if fb in hk:
        return hk.index(fb)

    for i, name in enumerate(hk):
        if "фраз" in name:
            return i

    return None

def find_param1_idx_strict(header: List[str]) -> int:
    hk = [k(x) for x in header]
    target = k(PARAM1_COL_NAME_STRICT)
    if target not in hk:
        params = [(i, header[i]) for i, name in enumerate(hk) if "параметр" in name]
        raise ValueError(f"Не нашел строго '{PARAM1_COL_NAME_STRICT}'. Найденные 'Параметр*': {params}")
    return hk.index(target)

def main():
    minus_set = load_minus_set(MINUS_TXT)
    print(f"Минус-строк (уникальных) загружено: {len(minus_set)}")

    enc = detect_encoding(CSV_INPUT)
    print(f"CSV encoding: {enc}")

    with open(CSV_INPUT, "r", encoding=enc, newline="") as f:
        lines = f.readlines()

    header_i, delim = pick_table_header(lines)
    print(f"Найдена табличная шапка на строке {header_i+1}, delimiter={repr(delim)}")

    header = split_row(lines[header_i].rstrip("\n"), delim)

    phrase_idx = find_phrase_idx(header)
    param1_idx = find_param1_idx_strict(header)

    if phrase_idx is None:
        raise ValueError(f"Не нашел колонку 'Фраза...' в заголовке: {header}")

    print(f"Колонка фразы: '{header[phrase_idx]}' (idx={phrase_idx})")
    print(f"Колонка параметр 1: '{header[param1_idx]}' (idx={param1_idx})")

    changed = 0
    total = 0

    with open(CSV_OUTPUT, "w", encoding=enc, newline="") as out:
        for i in range(header_i):
            out.write(lines[i])
        out.write(lines[header_i])

        for raw in lines[header_i + 1:]:
            if not raw.strip():
                out.write(raw)
                continue

            row = split_row(raw.rstrip("\n"), delim)
            if len(row) < len(header):
                row = row + [""] * (len(header) - len(row))

            total += 1
            phrase_clean = clean_phrase(row[phrase_idx])

            # ТОЧНОЕ совпадение
            if phrase_clean and phrase_clean in minus_set:
                row[param1_idx] = SET_VALUE
                changed += 1

            out.write(join_row(row, delim) + "\n")

    print(f"Готово: {CSV_OUTPUT}")
    print(f"Строк обработано: {total}, строк помечено: {changed}")

if __name__ == "__main__":
    main()
