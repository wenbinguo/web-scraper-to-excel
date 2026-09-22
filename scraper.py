# 爬虫练习

import time
import argparse
import logging

import requests
from bs4 import BeautifulSoup
import pandas as pd

BASE = 'https://books.toscrape.com/catalogue/page-{}.html'
RATING_MAP = {'One':1, 'Two':2, 'Three':3, 'Four':4, 'Five':5}
header = {'User-Agent': 'Mozilla/5.0'}

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s [%(levelname)s] %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger(__name__)


def fetch_page(url: str, retries: int = 3, delay: int = 2) -> str | None:
    """Fetch a url with retries. Return HTML or None on failure"""
    for attempt in range(1, retries + 1):
        try:
            resp = requests.get(url, timeout=10, headers=header)
            resp.raise_for_status()
            return resp.text
        except requests.RequestException as e:
            logger.warning('Attempt %d/%d failed for %s: %s', attempt, retries, url, e)
            if attempt < retries:
                time.sleep(delay * attempt)
    logger.error('All %d attempts failed for %s', retries, url)
    return None

def parse_page(html: str, page: int) -> list[dict]:
    """Parse one page of books. Skips malformed items."""
    soup = BeautifulSoup(html, 'html.parser')
    rows = []
    for book in soup.select('article.product_pod'):
        try:
            rows.append({
                'title': book.h3.a['title'],
                'price': book.select_one('.price_color').text.replace('Â£', ''),
                'rating': book.select_one('.star-rating')['class'][1],
                'link': 'https://books.toscrape.com/catalogue/' + book.h3.a['href'].replace('../', ''),
                'page': page,
            })
        except (AttributeError, KeyError, IndexError) as e:
            logger.warning('Skipping malformed item on page %d: %s', page ,e)
    return rows

def scrape(pages: int) -> list[dict]:
    """Scrage multiple pages. Skips pages that fail after retries."""
    all_rows = []
    for page in range(1, pages + 1):
        page_url = BASE.format(page)
        logger.info('Fetching page %d/%d: %s', page, pages, page_url)
        html = fetch_page(page_url)
        if html is None:
            logger.error('Skipping page %d due to fetch failure', page)
            continue
        rows = parse_page(html, page)
        all_rows.extend(rows)
        logger.info('Page %d: %d items', page, len(rows))
    return all_rows

def clean(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and deduplicate data."""
    df = df.copy()
    df['price'] = pd.to_numeric(df['price'], errors='coerce')
    df['rating'] = df['rating'].map(RATING_MAP)
    df = df.drop_duplicates(subset='title').sort_values('price')
    return df

def export(rows: list[dict], output: str) -> None:
    """Export raw and clean data to separate Excel sheets."""
    if not rows:
        logger.error('No data to export')
        return
    df_raw = pd.DataFrame(rows)
    df_clean = clean(df_raw)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_raw.to_excel(writer, sheet_name='Raw', index=False)
        df_clean.to_excel(writer, sheet_name='Clean', index=False)
        logger.info('Saved %d raw rows, %d clean rows to %s', len(rows), len(df_clean), output)


def main():
    parser = argparse.ArgumentParser(
        description='Scrape books.toscrape.com and export to excel'
    )
    parser.add_argument(
        '--pages',type=int, default=5,help='Number of pages to scrape (default: 5)'
    )
    parser.add_argument(
        '--output',type=str, default='books.xlsx',help='Output Excel filename (default: books.xlsx)'
    )
    args = parser.parse_args()
    logger.info('Starting scrape: pages=%d, output=%s', args.pages, args.output)
    rows = scrape(args.pages)
    export(rows, args.output)

if __name__ == '__main__':
    main()

