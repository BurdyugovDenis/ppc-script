from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from anticaptchaofficial.imagecaptcha import imagecaptcha

import csv
import time
import requests
import urllib.parse
import random
import pymorphy3
import re
import os


# =========================
# Конфигурация
# =========================

ANTICAPTCHA_KEY = os.getenv("ANTICAPTCHA_KEY", "")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/115.0.0.0 Safari/537.36"
)

INPUT_FILE = "input.csv"
OUTPUT_FILE = "output.csv"

REGION_CODE = 213  # Москва

MIN_DELAY = 1
MAX_DELAY = 2.5

TOP_RESULTS_LIMIT = 10

# Кэш словоформ, чтобы не считать одно и то же по 100 раз
WORD_FORMS_CACHE = {}


# =========================
# Инициализация pymorphy
# =========================

morph = pymorphy3.MorphAnalyzer()


# =========================
# Капча
# =========================

class CaptchaHandler:
    def __init__(self, driver):
        self.driver = driver
        self.solver = imagecaptcha()
        self.solver.set_key(ANTICAPTCHA_KEY)

    def handle_checkbox(self):
        try:
            checkbox = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable((By.ID, "js-button"))
            )
            self.driver.execute_script("arguments[0].click();", checkbox)
            time.sleep(2)
            return True

        except Exception as e:
            print(f"Ошибка при клике на чекбокс капчи: {short_error(e)}")
            return False

    def solve_image_captcha(self):
        try:
            img = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "img.AdvancedCaptcha-Image")
                )
            )

            img_url = img.get_attribute("src")

            image_content = requests.get(img_url, timeout=20).content
            captcha_text = self.solver.solve_and_return_solution(image_content)

            if captcha_text:
                input_field = self.driver.find_element(
                    By.CSS_SELECTOR,
                    "input.Textinput-Control"
                )
                input_field.clear()
                input_field.send_keys(captcha_text)

                submit_button = self.driver.find_element(
                    By.CSS_SELECTOR,
                    "button[type='submit']"
                )
                submit_button.click()

                time.sleep(3)
                return True

            return False

        except Exception as e:
            print(f"Ошибка решения капчи: {short_error(e)}")
            return False


# =========================
# Вспомогательные функции
# =========================

def short_error(error):
    """
    Укорачивает Selenium-ошибки, чтобы консоль не превращалась в простыню.
    """
    text = str(error)
    first_line = text.splitlines()[0] if text else ""
    return first_line[:300]


def normalize_text(text):
    """
    Нормализация текста:
    - нижний регистр
    - ё -> е
    - лишние пробелы
    """
    if not text:
        return ""

    text = text.lower()
    text = text.replace("ё", "е")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def get_word_forms(word):
    """
    Генерирует словоформы для одного слова.
    Например: сосна, сосны, сосне, сосну.
    """
    word = normalize_text(word)

    if not word:
        return []

    try:
        parsed_word = morph.parse(word)[0]
    except Exception:
        return [word]

    forms = set()

    for form in parsed_word.lexeme:
        forms.add(normalize_text(form.word))

    forms.add(word)

    return list(forms)


def get_phrase_forms(phrase):
    """
    Если main состоит из нескольких слов, например:
    'сосна черная'

    Скрипт добавляет:
    - исходную фразу
    - формы каждого отдельного слова
    """
    phrase = normalize_text(phrase)

    if not phrase:
        return []

    words = phrase.split()
    forms = set()

    forms.add(phrase)

    for word in words:
        for form in get_word_forms(word):
            forms.add(form)

    return list(forms)


def get_cached_forms(main_value):
    """
    Кэширует словоформы, чтобы не гонять pymorphy3 повторно.
    """
    key = normalize_text(main_value)

    if key not in WORD_FORMS_CACHE:
        WORD_FORMS_CACHE[key] = get_phrase_forms(key)

    return WORD_FORMS_CACHE[key]


def count_mentions(text, forms):
    """
    Считает упоминания словоформ в тексте.

    Используем границы слов, чтобы 'сосна' не считалась внутри длинного слова.
    """
    total = 0

    for form in forms:
        if not form:
            continue

        pattern = r"(?<![а-яa-z0-9])" + re.escape(form) + r"(?![а-яa-z0-9])"
        matches = re.findall(pattern, text, flags=re.IGNORECASE)
        total += len(matches)

    return total


def collect_result_text(result):
    """
    Быстрый вариант сбора текста:
    берем весь видимый текст блока результата.

    Это быстрее, чем искать сниппеты по куче CSS-селекторов.
    """
    try:
        return normalize_text(result.text)
    except Exception:
        return ""


def wait_for_results(driver):
    """
    Ждет появления результатов поиска.
    """
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".serp-item"))
    )


def get_search_results(driver):
    """
    Получает поисковые блоки.
    """
    try:
        return driver.find_elements(By.CSS_SELECTOR, ".serp-item")
    except Exception:
        return []


def handle_captcha_if_needed(driver):
    """
    Проверяет наличие капчи и пытается ее пройти.
    """
    if "showcaptcha" not in driver.current_url:
        return True

    print("Обнаружена капча")

    handler = CaptchaHandler(driver)

    if not handler.handle_checkbox():
        return False

    time.sleep(3)

    if "showcaptcha" in driver.current_url:
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "img.AdvancedCaptcha-Image")
                )
            )

            if not handler.solve_image_captcha():
                return False

        except Exception:
            return False

    return "showcaptcha" not in driver.current_url


# =========================
# Основная обработка строки
# =========================

def process_row(driver, main_value, query):
    try:
        search_url = (
            f"https://yandex.ru/search/"
            f"?text={urllib.parse.quote_plus(query)}"
            f"&lr={REGION_CODE}"
        )

        driver.get(search_url)

        captcha_ok = handle_captcha_if_needed(driver)

        if not captcha_ok:
            print(f"Не удалось пройти капчу для запроса: {query}")
            return 0

        wait_for_results(driver)

        results = get_search_results(driver)

        if not results:
            print(f"Не найдены результаты по запросу: {query}")
            return 0

        word_forms = get_cached_forms(main_value)

        total_mentions = 0
        processed_results = 0
        skipped_results = 0

        for result in results[:TOP_RESULTS_LIMIT]:
            result_text = collect_result_text(result)

            if not result_text:
                skipped_results += 1
                continue

            total_mentions += count_mentions(result_text, word_forms)
            processed_results += 1

        print(
            f"{main_value} | {query} | "
            f"упоминаний: {total_mentions} | "
            f"обработано блоков: {processed_results} | "
            f"пропущено: {skipped_results}"
        )

        return total_mentions

    except Exception as e:
        print(f"Ошибка для {main_value} / {query}: {short_error(e)}")
        return 0


# =========================
# Работа с CSV
# =========================

def read_input_rows():
    """
    Читает input.csv.

    Ожидаемый формат:
    main;Query

    Пример:
    сосна;купить сосну в москве
    яблоня;яблоня колоновидная купить
    """
    rows = []

    if not os.path.exists(INPUT_FILE):
        print(f"Файл {INPUT_FILE} не найден.")
        return rows

    with open(INPUT_FILE, "r", encoding="utf-8-sig", newline="") as infile:
        reader = csv.reader(infile, delimiter=";")

        for row in reader:
            if len(row) < 2:
                continue

            main_value = row[0].strip()
            query = row[1].strip()

            if not main_value or not query:
                continue

            # Пропускаем заголовок, если он есть
            if (
                main_value.lower() in ["main", "основной", "название"]
                and query.lower() in ["query", "запрос"]
            ):
                continue

            rows.append((main_value, query))

    return rows


def load_processed_rows():
    """
    Загружает уже обработанные строки из output.csv.

    Благодаря этому при повторном запуске скрипт продолжает работу,
    а не начинает всё заново.
    """
    processed = set()

    if not os.path.exists(OUTPUT_FILE):
        return processed

    if os.path.getsize(OUTPUT_FILE) == 0:
        return processed

    with open(OUTPUT_FILE, "r", encoding="utf-8-sig", newline="") as outfile:
        reader = csv.reader(outfile, delimiter=";")

        # Пропускаем заголовок
        next(reader, None)

        for row in reader:
            if len(row) < 2:
                continue

            main_value = row[0].strip()
            query = row[1].strip()

            if main_value and query:
                processed.add((main_value, query))

    return processed


# =========================
# Selenium driver
# =========================

def create_driver():
    options = webdriver.ChromeOptions()

    options.add_argument(f"user-agent={USER_AGENT}")
    options.add_argument("--disable-blink-features=AutomationControlled")

    # Чтобы браузер не закрывался сразу после завершения скрипта.
    # Когда всё стабильно — можно убрать.
    options.add_experimental_option("detach", True)

    # Снижаем загрузку лишнего: картинки, шрифты, CSS.
    # Для парсинга текста это обычно не нужно.
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.stylesheets": 2,
        "profile.managed_default_content_settings.fonts": 2,
    }

    options.add_experimental_option("prefs", prefs)

    # Немного маскируем автоматизацию
    options.add_experimental_option(
        "excludeSwitches",
        ["enable-automation"]
    )
    options.add_experimental_option(
        "useAutomationExtension",
        False
    )

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    driver.maximize_window()

    return driver


# =========================
# main
# =========================

def main():
    rows = read_input_rows()

    if not rows:
        print("В input.csv не найдено строк для обработки.")
        return

    processed_rows = load_processed_rows()

    total_rows = len(rows)

    remaining_rows = [
        (main_value, query)
        for main_value, query in rows
        if (main_value, query) not in processed_rows
    ]

    print(f"Всего строк во входном файле: {total_rows}")
    print(f"Уже обработано ранее: {len(processed_rows)}")
    print(f"Осталось обработать: {len(remaining_rows)}")

    if not remaining_rows:
        print("Все строки уже обработаны. Повторно запускать нечего.")
        return

    driver = create_driver()

    try:
        file_exists = os.path.exists(OUTPUT_FILE)
        file_is_empty = not file_exists or os.path.getsize(OUTPUT_FILE) == 0

        with open(
            OUTPUT_FILE,
            "a",
            newline="",
            encoding="utf-8-sig"
        ) as outfile:

            writer = csv.writer(outfile, delimiter=";")

            if file_is_empty:
                writer.writerow(["main", "Query", "Mentions"])
                outfile.flush()

            for index, (main_value, query) in enumerate(remaining_rows, start=1):
                print(
                    f"\n[{index}/{len(remaining_rows)}] "
                    f"Обработка: {main_value} | {query}"
                )

                count = process_row(driver, main_value, query)

                # Сохраняем результат сразу после обработки строки
                writer.writerow([main_value, query, count])
                outfile.flush()

                # На всякий случай добавляем в память текущего запуска
                processed_rows.add((main_value, query))

                delay = random.uniform(MIN_DELAY, MAX_DELAY)
                time.sleep(delay)

    except KeyboardInterrupt:
        print("\nСкрипт остановлен вручную. Уже обработанные строки сохранены.")

    except Exception as e:
        print(f"\nКритическая ошибка: {short_error(e)}")
        print("Уже обработанные строки сохранены в output.csv.")

    finally:
        input("\nГотово или остановлено. Нажми Enter, чтобы закрыть браузер...")
        driver.quit()


if __name__ == "__main__":
    main()