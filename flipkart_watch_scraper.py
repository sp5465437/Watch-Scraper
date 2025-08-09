import requests
from bs4 import BeautifulSoup
import re
import openpyxl
import time


class WatchScraper:
    def __init__(self):
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/115.0.0.0 Safari/537.36"
        }
        self.products = []

    def scrape_flipkart(self):
        url = "https://www.flipkart.com/search?q=watches+for+men+under+2000"
        response = requests.get(url, headers=self.headers)
        soup = BeautifulSoup(response.text, "html.parser")

        for link in soup.find_all("a", href=True):
            if "/p/" in link['href']:
                name = link.get_text(strip=True)
                if not name or "watch" not in name.lower():
                    continue

                parent = link.find_parent()
                price_tag = parent.find_next("div", string=re.compile(r"₹\d+"))
                if not price_tag:
                    continue

                price = int(re.sub(r"[^\d]", "", price_tag.get_text()))
                if price > 2000:
                    continue

                brand = name.split()[0]
                availability = "In Stock"

                self.products.append([name, brand, price, availability])

    def save_to_excel(self, filename="watches_combined.xlsx"):
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
        print(f"Data saved to {filename}")


if __name__ == "__main__":
    scraper = WatchScraper()
    print("Scraping Flipkart...")
    scraper.scrape_flipkart()
    print(f"Total products scraped: {len(scraper.products)}")
    scraper.save_to_excel()
