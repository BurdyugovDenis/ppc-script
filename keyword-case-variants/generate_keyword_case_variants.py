# -*- coding: utf-8 -*-
"""
pymorphy3-версия генератора вариаций ключевых фраз:
- читает input.txt (1 ключ/фраза на строку)
- склоняет каждое слово по падежам/числам
- делает декартово произведение форм по всем словам (максимум вариантов)
- для каждой базовой вариации добавляет 4 режима:
  1) plain: купить шубу
  2) !: !купить !шубу
  3) []: [купить шубу]
  4) []+!: [!купить !шубу]
- пишет output.txt (1 строка = 1 ключ)

Зависимости:
pip install pymorphy3 pymorphy3-dicts-ru
"""

from __future__ import annotations

import re
from itertools import product
from pathlib import Path
from typing import Iterable, List, Set, Tuple

import pymorphy3

# --- файлы
INPUT_TXT = "input.txt"
OUTPUT_TXT = "output.txt"

# --- граммемы
CASES = ("nomn", "gent", "datv", "accs", "ablt", "loct")
NUMBERS = ("sing", "plur")

# --- поведение
WRITE_VARIANTS = True          # писать ли 4 варианта (plain/!/[]/[]+!)
GLOBAL_DEDUP = True            # убирать ли дубли глобально по output
MAX_BASE_PER_KEY = None        # ограничение на базовые вариации для одной фразы (например 5000) или None

# токенизация: слова/цифры + дефисы (для ключей это обычно практично)
TOKEN_RE = re.compile(r"[A-Za-zА-Яа-яЁё0-9]+(?:-[A-Za-zА-Яа-яЁё0-9]+)*", re.UNICODE)

# Важно: lang='ru' (а словари ставятся pymorphy3-dicts-ru) [web:69]
morph = pymorphy3.MorphAnalyzer(lang="ru")


def tokenize(text: str) -> List[str]:
    return TOKEN_RE.findall(text.strip().lower())


def word_forms(word: str) -> Set[str]:
    """
    Все формы по падежам/числам.
    В pymorphy3 остаётся привычный API parse(...)->Parse, затем Parse.inflect(set_of_grammemes). [web:60][web:61]
    """
    parses = morph.parse(word)
    if not parses:
        return {word}

    p = parses[0]
    forms: Set[str] = set()

    for case in CASES:
        for number in NUMBERS:
            inf = p.inflect({case, number})
            if inf:
                forms.add(inf.word)

    return forms or {word}


def generate_base_variations(phrase: str) -> Set[str]:
    tokens = tokenize(phrase)
    if not tokens:
        return set()

    forms_by_token = [sorted(word_forms(t)) for t in tokens]

    out: Set[str] = set()
    for combo in product(*forms_by_token):
        out.add(" ".join(combo))
        if MAX_BASE_PER_KEY is not None and len(out) >= MAX_BASE_PER_KEY:
            break
    return out


def add_exclamation_each_word(phrase: str) -> str:
    return " ".join(f"!{w}" for w in phrase.split())


def wrap_brackets(phrase: str) -> str:
    return f"[{phrase}]"


def variants(phrase: str) -> Tuple[str, str, str, str]:
    plain = phrase
    exc = add_exclamation_each_word(phrase)
    br = wrap_brackets(phrase)
    both = wrap_brackets(exc)
    return plain, exc, br, both


def read_keys(path: Path) -> List[str]:
    lines = path.read_text(encoding="utf-8").splitlines()
    keys: List[str] = []
    for line in lines:
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        keys.append(s)
    return keys


def write_lines(path: Path, lines: Iterable[str]) -> None:
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    in_path = Path(INPUT_TXT)
    out_path = Path(OUTPUT_TXT)

    if not in_path.exists():
        raise FileNotFoundError(
            f"Не найден {INPUT_TXT} рядом со скриптом.\n"
            f"Создайте файл и вставьте ключи: по 1 фразе на строку."
        )

    keys = read_keys(in_path)
    if not keys:
        raise ValueError("input.txt пустой (или всё закомментировано).")

    out: List[str] = []
    seen: Set[str] = set()

    for key in keys:
        base_vars = generate_base_variations(key)

        for base in sorted(base_vars):
            if WRITE_VARIANTS:
                for v in variants(base):
                    if GLOBAL_DEDUP:
                        if v not in seen:
                            seen.add(v)
                            out.append(v)
                    else:
                        out.append(v)
            else:
                if GLOBAL_DEDUP:
                    if base not in seen:
                        seen.add(base)
                        out.append(base)
                else:
                    out.append(base)

    write_lines(out_path, out)
    print(f"OK: {in_path.resolve()} -> {out_path.resolve()}")
    print(f"Lines written: {len(out)}")


if __name__ == "__main__":
    main()
