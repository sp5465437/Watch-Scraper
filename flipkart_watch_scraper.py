import requests
from bs4 import BeautifulSoup
import re
import openpyxl


class WatchScraper:
    def __init__(self):
        self.url = "https://www.flipkart.com/search?q=watches+for+men+under+2000"
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/115.0.0.0 Safari/537.36"
        }
        self.products = []

    def fetch_html(self):
        """Fetch and save HTML content to txt"""
        response = requests.get(self.url, headers=self.headers)
        html_content = response.text
        with open("flipkart_watches_page1.txt", "w", encoding="utf-8") as f:
            f.write(html_content)
        return html_content

    def parse_html(self, html_content):
        """Extract product details from HTML"""
        soup = BeautifulSoup(html_content, "html.parser")

        # Each product link usually has "/p/" in href
        for product_link in soup.find_all("a", href=True):
            if "/p/" not in product_link["href"]:
                continue

            # Get product name
            name = product_link.get_text(strip=True)
            if not name or "watch" not in name.lower():
                continue

            # Navigate upwards to the container div
            container = product_link.find_parent("div")
            if not container:
                continue

            # Find price in nearby elements
            price_tag = container.find_next(string=re.compile(r"₹\s*\d+"))
            if not price_tag:
                continue

            # Clean price and filter
            price = int(re.sub(r"[^\d]", "", price_tag))
            if price > 2000:
                continue

            # Brand is assumed to be first word of product name
            brand = name.split()[0]

            # Availability: Flipkart rarely lists out of stock on search page
            availability = "In Stock"

            # Append to results
            self.products.append([name, brand, price, availability])

    def save_to_excel(self, filename="watches_under_2000.xlsx"):
        """Save extracted products to Excel"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Watches Under 2000"

        headers = ["Watch Name", "Brand", "Price", "Availability"]
        ws.append(headers)

        # Make headers bold
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True)

        for product in self.products:
            ws.append(product)

        wb.save(filename)
        print(f"✅ Data saved to {filename}")


if __name__ == "__main__":
    scraper = WatchScraper()
    print("📡 Fetching Flipkart HTML...")
    html = scraper.fetch_html()

    print("🔍 Parsing products...")
    scraper.parse_html(html)

    print(f"🛒 Total products scraped: {len(scraper.products)}")
    scraper.save_to_excel()
