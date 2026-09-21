# 爬虫练习

import requests
from bs4 import BeautifulSoup
import pandas as pd

url = 'https://books.toscrape.com/'
header = {'User-Agent': 'Mozilla/5.0'}
resp = requests.get(url,timeout=10, headers=header)
resp.raise_for_status()
soup = BeautifulSoup(resp.text, 'html.parser')
# print(resp.text)
rows = []
for book in soup.select('article.product_pod'):
    rows.append({
        'title': book.h3.a['title'],
        'price': book.select_one('.price_color').text,
        'rating': book.select_one('.star-rating')['class'][1],
        'link': url + book.h3.a['href'],
    })

df = pd.DataFrame(rows)
df['price'] = df['price'].str.replace('Â£', '').astype(float)
df.to_excel('books.xlsx',index=False)
print(f'Done! {len(df)} rows written to books.xlsx')