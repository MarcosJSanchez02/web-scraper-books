# 📚 Books Web Scraper

Python web scraper that extracts 100 books from [books.toscrape.com](https://books.toscrape.com) and exports the data to a formatted Excel file.

## Features
- Scrapes multiple pages automatically
- Extracts title, price, star rating and category for each book
- Exports data to Excel with styled headers and auto-adjusted columns

## Tech Stack
- Python 3
- requests
- BeautifulSoup4
- openpyxl

## Installation
pip install requests beautifulsoup4 openpyxl

## Usage
python scraper.py

## Output
| Title | Price | Rating | Category |
|-------|-------|--------|----------|
| A Light in the Attic | £51.77 | ★★★☆☆ | Poetry |
| Tipping the Velvet | £53.74 | ★☆☆☆☆ | Historical Fiction |
