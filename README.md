# Web Scraper to Excel

A Python scraper that extracts book data from books.toscrape.com and exports a clean Excel report.

## Tech Stack
- Python 3
- requests
- BeautifulSoup
- pandas
- openpyxl

## Features
- Scrape book title, price, rating, and link
- Clean price to numeric
- Convert rating text to number
- Export to Excel

## How to Run
pip install -r requirements.txt
python scraper.py

## Output
books.xlsx with 20 rows and 4 columns.

## Notes
Only scrapes public data from a site designed for scraping practice.