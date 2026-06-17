import pandas as pd

# Чтение данных
df = pd.read_csv('input.csv', header=2, sep=';', encoding='utf-8-sig')

# Предварительная обработка названий групп
df['Название группы'] = df['Название группы'].str.strip().str.lower()

# Поиск целевых строк
autotarget_mask = df['Фраза (с минус-словами)'].str.strip() == '---autotargeting'

# Группировка и фильтрация
def filter_group(group):
    autotarget_count = (group['Фраза (с минус-словами)'].str.strip() == '---autotargeting').sum()
    if autotarget_count >= 2:
        return group[~autotarget_mask]
    return group

filtered_df = df.groupby('Название группы', group_keys=False).apply(filter_group)

# Нормализация номеров групп
first_numbers = filtered_df.groupby('Название группы')['Номер группы'].first().to_dict()
filtered_df['Номер группы'] = filtered_df['Название группы'].map(first_numbers)

# Сохранение
filtered_df.to_csv('output.csv', index=False, sep=';', encoding='utf-8-sig')

print(f"Удалено строк: {len(df) - len(filtered_df)}")
print("Группы с удаленными autotargeting:")
print(df.groupby('Название группы')['Фраза (с минус-словами)']
      .apply(lambda x: x.str.contains('---autotargeting').sum())
      .loc[lambda x: x >= 2])

print(df['Фраза (с минус-словами)'].unique())
