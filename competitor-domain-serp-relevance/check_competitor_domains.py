import os
import time
import signal
import sys
import re
from urllib.parse import urlparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


class SearchProcessor:
    def __init__(self):
        self.results_file = "results.txt"
        self.progress_file = "progress.txt"
        self.driver = None
        self.processed_count = 0
        self.shutdown_requested = False
        self.min_domains_threshold = 1

        signal.signal(signal.SIGINT, self.signal_handler)
        signal.signal(signal.SIGTERM, self.signal_handler)

    def signal_handler(self, signum, frame):
        print(f"\nПолучен сигнал остановки. Завершаем работу...")
        self.shutdown_requested = True
        if self.driver:
            self.driver.quit()
        sys.exit(0)

    def setup_driver(self):
        chrome_options = Options()
        # chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_argument(
            "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return self.driver

    def load_phrases(self, filename):
        with open(filename, 'r', encoding='utf-8') as file:
            phrases = [line.strip() for line in file if line.strip()]
        return phrases

    def load_domains(self, filename):
        with open(filename, 'r', encoding='utf-8') as file:
            domains = [line.strip().lower() for line in file if line.strip()]
        return domains

    def load_progress(self):
        processed_phrases = set()
        if os.path.exists(self.progress_file):
            with open(self.progress_file, 'r', encoding='utf-8') as file:
                for line in file:
                    if ';' in line:
                        phrase = line.split(';')[0].strip()
                        processed_phrases.add(phrase)
        return processed_phrases

    def save_result(self, phrase, is_target, found_count, total_domains):
        status = "целевое" if is_target else "не целевое"
        result_line = f"{phrase};{status};{found_count}/{total_domains}"

        with open(self.results_file, 'a', encoding='utf-8') as file:
            file.write(result_line + '\n')

        with open(self.progress_file, 'a', encoding='utf-8') as file:
            file.write(result_line + '\n')

        print(f"Сохранен результат: {result_line}")

    def extract_all_domains_from_page(self, html_content):
        domains_found = set()

        patterns = [
            r'https?://([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
            r'//([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})',
            r'[^a-zA-Z0-9]([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})[^a-zA-Z0-9]',
            r'href="//([^"]+)"',
            r'src="//([^"]+)"',
        ]

        for pattern in patterns:
            matches = re.findall(pattern, html_content, re.IGNORECASE)
            for match in matches:
                domain = match.split('/')[0].split('?')[0].lower()
                if '.' in domain and len(domain) > 3:
                    domains_found.add(domain)

        return domains_found

    def normalize_domain(self, domain):
        domain = domain.lower().strip()
        if domain.startswith('www.'):
            domain = domain[4:]
        return domain

    def get_main_domain(self, domain):
        """Получение основного домена (второго уровня)"""
        parts = domain.split('.')
        if len(parts) >= 2:
            return '.'.join(parts[-2:])
        return domain

    def is_domain_match(self, found_domain, target_domain):
        """Правильное сравнение доменов - ИСПРАВЛЕННАЯ ЛОГИКА"""
        found = self.normalize_domain(found_domain)
        target = self.normalize_domain(target_domain)

        # Точное совпадение
        if found == target:
            return True

        # Совпадение по основному домену (второго уровня)
        found_main = self.get_main_domain(found)
        target_main = self.get_main_domain(target)

        if found_main == target_main:
            return True

        return False

    def search_phrase(self, phrase, domains):
        if self.shutdown_requested:
            return False, 0

        encoded_phrase = phrase.replace(' ', '+')
        url = f"https://yandex.ru/search/?text={encoded_phrase}&lr=213"

        try:
            print(f"  Переход по URL: {url}")
            self.driver.get(url)

            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                                                ".serp-item, .organic, .serp-adv, [class*='organic'], .Link, a"))
            )

            time.sleep(3)

            page_html = self.driver.page_source
            page_text = self.driver.page_source.lower()

            found_domains = self.extract_all_domains_from_page(page_html)

            print(f"  Найдено доменов на странице: {len(found_domains)}")
            if found_domains:
                print(f"  Примеры доменов: {list(found_domains)[:5]}")

            # Проверяем наличие целевых доменов - ИСПРАВЛЕННАЯ ЛОГИКА
            found_target_domains = set()
            for found_domain in found_domains:
                for target_domain in domains:
                    if self.is_domain_match(found_domain, target_domain):
                        found_target_domains.add(target_domain)
                        print(f"  ✓ Найден домен: {target_domain} (как {found_domain})")

            # Дополнительная проверка: ищем домены в тексте страницы
            for target_domain in domains:
                # Более строгая проверка в тексте
                pattern = r'[^a-zA-Z0-9]' + re.escape(target_domain) + r'[^a-zA-Z0-9]'
                if re.search(pattern, page_text):
                    found_target_domains.add(target_domain)
                    print(f"  ✓ Найден домен в тексте: {target_domain}")

            found_count = len(found_target_domains)
            print(f"  Всего найдено целевых доменов: {found_count}")
            for domain in found_target_domains:
                print(f"    - {domain}")

            is_target = found_count >= self.min_domains_threshold
            print(f"  Порог для целевого: {self.min_domains_threshold} домен(а)")
            print(f"  Статус: {'ЦЕЛЕВОЕ' if is_target else 'НЕ ЦЕЛЕВОЕ'}")

            return is_target, found_count

        except Exception as e:
            print(f"Ошибка при поиске фразы '{phrase}': {e}")
            try:
                debug_file = f"debug_{phrase[:20]}.html"
                with open(debug_file, 'w', encoding='utf-8') as f:
                    f.write(self.driver.page_source)
                print(f"  HTML сохранен в {debug_file} для отладки")
            except:
                pass

            try:
                self.driver.quit()
            except:
                pass
            self.setup_driver()
            return False, 0

    def get_remaining_phrases(self, all_phrases, processed_phrases):
        return [phrase for phrase in all_phrases if phrase not in processed_phrases]

    def print_statistics(self, total, processed, remaining):
        print(f"\n=== СТАТИСТИКА ===")
        print(f"Всего фраз: {total}")
        print(f"Обработано: {processed}")
        print(f"Осталось: {remaining}")
        print(f"Прогресс: {(processed / total) * 100:.1f}%")
        print(f"Порог целевого: {self.min_domains_threshold} домен(а)")
        print("==================\n")

    def main(self):
        if not os.path.exists("phrases.txt"):
            print("Ошибка: Файл phrases.txt не найден!")
            return

        if not os.path.exists("domain.txt"):
            print("Ошибка: Файл domain.txt не найден!")
            return

        if not os.path.exists(self.results_file):
            with open(self.results_file, 'w', encoding='utf-8') as file:
                file.write("фраза;статус;найдено_доменов\n")

        try:
            all_phrases = self.load_phrases("phrases.txt")
            domains = self.load_domains("domain.txt")
            processed_phrases = self.load_progress()
        except Exception as e:
            print(f"Ошибка при загрузке файлов: {e}")
            return

        print(f"Текущий порог для 'целевое': {self.min_domains_threshold} домен(а)")
        change_threshold = input("Изменить порог? (y/n): ").lower().strip()
        if change_threshold == 'y':
            try:
                new_threshold = int(input("Введите новый порог: "))
                self.min_domains_threshold = new_threshold
                print(f"Порог изменен на: {self.min_domains_threshold}")
            except ValueError:
                print("Ошибка: введите число. Используется порог по умолчанию.")

        print(f"Загружено фраз: {len(all_phrases)}")
        print(f"Загружено доменов: {len(domains)}")
        print(f"Целевые домены: {domains}")
        print(f"Уже обработано: {len(processed_phrases)}")
        print(f"Порог для целевого статуса: {self.min_domains_threshold} домен(а)")

        remaining_phrases = self.get_remaining_phrases(all_phrases, processed_phrases)

        if not remaining_phrases:
            print("Все фразы уже обработаны!")
            return

        self.setup_driver()

        try:
            total_to_process = len(remaining_phrases)
            self.print_statistics(len(all_phrases), len(processed_phrases), total_to_process)

            print("Начинаем обработку... Для остановки нажмите Ctrl+C")

            for i, phrase in enumerate(remaining_phrases, 1):
                if self.shutdown_requested:
                    break

                print(f"\n[{i}/{total_to_process}] Обработка: '{phrase}'")

                is_target, found_count = self.search_phrase(phrase, domains)
                self.save_result(phrase, is_target, found_count, len(domains))
                self.processed_count += 1

                if i % 5 == 0:
                    self.print_statistics(
                        len(all_phrases),
                        len(processed_phrases) + i,
                        total_to_process - i
                    )

                if not self.shutdown_requested:
                    pause_time = 2 + (i % 4)
                    print(f"  Пауза {pause_time} сек...")
                    time.sleep(pause_time)

            if not self.shutdown_requested:
                print("\nОбработка всех фраз завершена!")
            else:
                print(f"\nОбработка прервана. Обработано {self.processed_count} фраз в этой сессии.")

        finally:
            if self.driver:
                self.driver.quit()

            final_processed = len(processed_phrases) + self.processed_count
            print(f"\n=== ИТОГИ ===")
            print(f"Всего фраз в базе: {len(all_phrases)}")
            print(f"Всего обработано: {final_processed}")
            print(f"Осталось: {len(all_phrases) - final_processed}")
            print(f"Порог целевого: {self.min_domains_threshold} домен(а)")
            print(f"Результаты сохранены в: {self.results_file}")


def main():
    processor = SearchProcessor()
    processor.main()


if __name__ == "__main__":
    main()