import re
import sys
from pathlib import Path

import pandas as pd

INPUT_CSV = "direct_export_input.csv"
CONTAINS_FILE = "stop_contains.txt"
OUTPUT_CSV = "direct_export_output.csv"

PARAM1_COL = "Параметр 1"
STOP_VALUE = "остановить"


# --- настройки поиска колонки с фразой ---
# сначала пробуем точные совпадения, затем эвристику по подстрокам
PHRASE_COL_EXACT = [
    "Фраза (с минус-словами)",
    "Фраза",
    "Ключевая фраза",
    "Ключевые фразы",
    "Ключевая фраза (с минус-словами)",
]

PHRASE_COL_CONTAINS_ANY = [
    "фраз",          # фраза/фразы
    "ключев",        # ключевая/ключевые
]


def load_contains_list(path: str) -> list[str]:
    items = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if not s or s.startswith("#"):
                continue
            items.append(s)
    return items


def strip_minus_tail(phrase: str) -> str:
    """
    Отрезаем минус-слова: берем всё ДО первого дефиса '-'.
    """
    if phrase is None:
        return ""
    s = str(phrase).strip()
    if "-" in s:
        s = s.split("-", 1)[0].strip()
    return s


def find_phrase_column(columns: list[str]) -> str:
    # 1) точные варианты
    for name in PHRASE_COL_EXACT:
        if name in columns:
            return name

    # 2) эвристика по частичному совпадению
    for col in columns:
        low = col.lower()
        # если в названии есть "фраз"
        if any(token in low for token in PHRASE_COL_CONTAINS_ANY) and ("фраз" in low):
            return col

    raise KeyError(
        "Не найдена колонка с фразами. "
        f"Доступные колонки: {columns}. "
        "Проверь, что это выгрузка по ключевым фразам/поисковым запросам, а не по текстовым блокам."
    )


def read_direct_csv_robust(path: str) -> tuple[pd.DataFrame, dict]:
    """
    Возвращает (df, meta) где meta — с какими параметрами файл прочитался.
    Пробуем типовые кодировки и разделители. pd.read_csv это поддерживает через sep/encoding [web:2].
    """
    seps = ["\t", ";", ","]
    encs = ["utf-16", "cp1251", "utf-8-sig", "utf-8"]

    last_err = None
    for enc in encs:
        for sep in seps:
            try:
                df = pd.read_csv(
                    path,
                    sep=sep,
                    encoding=enc,
                    engine="python",
                )
                # минимальная проверка: должно быть больше 1 колонки
                # (иначе часто это признак неправильного sep)
                if df.shape[1] < 2:
                    continue
                return df, {"encoding": enc, "sep": sep}
            except Exception as e:
                last_err = e
                continue

    raise RuntimeError(f"Не удалось прочитать файл {path}. Последняя ошибка: {last_err}")


def main():
    input_path = Path(INPUT_CSV)
    contains_path = Path(CONTAINS_FILE)

    if not input_path.exists():
        raise FileNotFoundError(f"Не найден входной файл: {input_path.resolve()}")
    if not contains_path.exists():
        raise FileNotFoundError(f"Не найден файл подстрок: {contains_path.resolve()}")

    contains_list = load_contains_list(str(contains_path))
    if not contains_list:
        raise ValueError(f"Файл подстрок пустой: {contains_path.resolve()}")

    df, meta = read_direct_csv_robust(str(input_path))
    print(f"CSV прочитан: encoding={meta['encoding']}, sep={repr(meta['sep'])}. Колонок: {df.shape[1]}")

    phrase_col = find_phrase_column(list(df.columns))
    print(f"Колонка с фразой: {phrase_col}")

    # гарантируем наличие "Параметр 1"
    if PARAM1_COL not in df.columns:
        df[PARAM1_COL] = ""

    cleaned = df[phrase_col].map(strip_minus_tail).fillna("")

    # Подстроки объединяем в regex: (a|b|c), экранируем спецсимволы
    pattern = "|".join(re.escape(x) for x in contains_list)

    # case=False — регистронезависимо; na=False — NaN считаем как False [web:16]
    stop_mask = cleaned.str.contains(pattern, case=False, na=False, regex=True)

    df.loc[stop_mask, PARAM1_COL] = STOP_VALUE

    out_path = Path(OUTPUT_CSV)
    df.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"Готово. Отмечено к остановке: {int(stop_mask.sum())}. Выходной файл: {out_path.resolve()}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        raise
