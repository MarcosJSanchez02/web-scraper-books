# Books Web Scraper

Python web scraper that extracts 1000 books from [books.toscrape.com](https://books.toscrape.com) and exports the data to a formatted Excel file.

## Features
- Scrapes 50 pages automatically
- Extracts title, price and star rating for each book
- Exports data to Excel with styled headers and auto-adjusted columns

## Tech Stack
- Python 3
- requests
- BeautifulSoup4
- openpyxl

## Installation
```bash
pip install requests beautifulsoup4 openpyxl
```

## Usage
```bash
python scraper.py
```

## Output
| Title | Price | Rating |
|-------|-------|--------|
| A Light in the Attic | £51.77 | ★★★☆☆ |
| Tipping the Velvet | £53.74 | ★☆☆☆☆ |
