# 爬虫练习

import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE = 'https://books.toscrape.com/catalogue/page-{}.html'
RATING_MAP = {'One':1, 'Two':2, 'Three':3, 'Four':4, 'Five':5}
rows = []

for page in range(1, 6):
    url = BASE.format(page)
    header = {'User-Agent': 'Mozilla/5.0'}
    resp = requests.get(url,timeout=10, headers=header)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, 'html.parser')

    for book in soup.select('article.product_pod'):
        rows.append({
            'title': book.h3.a['title'],
            'price': book.select_one('.price_color').text.replace('Â£',''),
            'rating': book.select_one('.star-rating')['class'][1],
            'link': 'https://books.toscrape.com/catalogue/' + book.h3.a['href'].replace('../', ''),
            'page': page,
        })



df_raw = pd.DataFrame(rows)

df_clean = df_raw.copy()
df_clean['price'] = df_clean['price'].astype(float)
df_clean['rating'] = df_clean['rating'].map(RATING_MAP)
df_clean = df_clean.drop_duplicates(subset='title').sort_values('price')

with pd.ExcelWriter('books.xlsx', engine='openpyxl') as writer:
    df_raw.to_excel(writer, sheet_name='Raw', index=False)
    df_clean.to_excel(writer, sheet_name='Clean', index=False)

print(f'Done! {len(df_raw)} raw rows, {len(df_clean)} clean rows')
