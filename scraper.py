import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

RATINGS = {
    "One": "★☆☆☆☆",
    "Two": "★★☆☆☆",
    "Three": "★★★☆☆",
    "Four": "★★★★☆",
    "Five": "★★★★★"
}

books = []

print("⏳ Scraping pages...")

for page_number in range(1, 51):

    url = "https://books.toscrape.com/" if page_number == 1 else f"https://books.toscrape.com/catalogue/page-{page_number}.html"

    response = requests.get(url)
    response.encoding = "utf-8"

    soup = BeautifulSoup(response.text, "html.parser")

    for book in soup.find_all("article", class_="product_pod"):
        title = book.find("h3").find("a")["title"]
        price = book.find("p", class_="price_color").text
        rating = RATINGS[book.find("p", class_="star-rating")["class"][1]]
        books.append([title, price, rating])

    print(f"  ✅ Page {page_number}/50 done")

# Create and style the Excel file
wb = Workbook()
ws = wb.active
ws.title = "Books"

header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", size=12)

for col, header in enumerate(["Title", "Price", "Rating"], start=1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center")

for book in books:
    ws.append(book)

# Auto-adjust column widths
for col in ws.columns:
    max_width = max(len(str(cell.value)) for cell in col if cell.value)
    ws.column_dimensions[col[0].column_letter].width = max_width + 4

wb.save("books.xlsx")

print(f"\n🏆 Done! {len(books)} books saved to books.xlsx")