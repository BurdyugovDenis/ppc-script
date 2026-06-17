import pandas as pd

# читаем входной файл с разделителем ;
df = pd.read_csv('input.csv', sep=';')

# убираем дубликаты пар кампания-ключ
df = df.drop_duplicates(subset=['campaign', 'keyword'])

# группируем ключи по кампании
result = (
    df.groupby('campaign')['keyword']
      .apply(lambda x: ', '.join(x))
      .reset_index()
)

# сохраняем результат тоже с ;
result.to_csv('output.csv', index=False, sep=';')

print(result)
