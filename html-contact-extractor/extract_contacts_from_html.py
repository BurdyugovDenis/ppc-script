from bs4 import BeautifulSoup
import pandas as pd
import re

# ======== Функции очистки ========

def extract_phones(text):
    """Извлечение всех телефонов из строки"""
    if not text or pd.isna(text):
        return []
    phone_pattern = r'(?:\+7|8)?[\s\-\(\)]?\d{3}[\s\-\(\)]?\d{3}[\s\-\(\)]?\d{2}[\s\-\(\)]?\d{2}'
    return re.findall(phone_pattern, str(text))

def normalize_phone(phone):
    """Приведение телефона к формату 7XXXXXXXXXX"""
    cleaned = re.sub(r'\D', '', phone)
    if len(cleaned) == 10:  # без кода страны
        cleaned = '7' + cleaned
    elif len(cleaned) == 11 and cleaned[0] == '8':
        cleaned = '7' + cleaned[1:]
    if len(cleaned) == 11 and cleaned[0] == '7':
        return cleaned
    return None

def extract_email(text):
    """Извлечение email"""
    if not text or pd.isna(text):
        return None
    email_pattern = r'[\w\.-]+@[\w\.-]+\.\w+'
    match = re.search(email_pattern, str(text))
    return match.group(0).lower() if match else None


# ======== Чтение HTML ========
with open("contacts.html", "r", encoding="utf-8") as f:
    soup = BeautifulSoup(f, "html.parser")

table = soup.find("table")

# Забираем строки
rows = []
for tr in table.find_all("tr"):
    cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
    rows.append(cells)

# Превращаем в DataFrame
df = pd.DataFrame(rows[1:], columns=rows[0])

# ======== Очистка телефонов и email ========
cleaned_phones = set()
cleaned_emails = set()

for col in df.columns:
    # телефоны
    if "тел" in col.lower():
        normalized_list = []
        for val in df[col]:
            phones = []
            for raw in extract_phones(val):
                norm = normalize_phone(raw)
                if norm:
                    phones.append(norm)
                    cleaned_phones.add(norm)
            normalized_list.append(", ".join(phones) if phones else None)
        df[col] = normalized_list

    # e-mail
    if "mail" in col.lower() or "email" in col.lower():
        df[col] = df[col].apply(extract_email)
        for val in df[col].dropna():
            cleaned_emails.add(val)

# ======== Сохраняем ========
df.to_csv("contacts_clean.csv", index=False, encoding="utf-8-sig")

with open("phones.txt", "w") as f:
    for phone in sorted(cleaned_phones):
        f.write(phone + "\n")

with open("emails.txt", "w") as f:
    for email in sorted(cleaned_emails):
        f.write(email + "\n")

print(f"✅ Готово! Файл contacts_clean.csv создан.")
print(f"📱 Телефонов найдено: {len(cleaned_phones)} (phones.txt)")
print(f"📧 Email найдено: {len(cleaned_emails)} (emails.txt)")
