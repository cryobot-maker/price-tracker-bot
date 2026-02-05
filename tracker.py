import pandas as pd
import gspread
import requests
import time
import os
import json
import random
from bs4 import BeautifulSoup
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# --- CONFIGURATION ---
EXCEL_FILE = "products.xlsx"
SHEET_NAME = "Price Tracker 2026"
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")

# User Agents to mimic real browsers (prevents blocking)
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
]

def get_price(url):
    """Detects platform and scrapes price."""
    if not isinstance(url, str) or "http" not in url:
        return "Not Available"

    headers = {"User-Agent": random.choice(USER_AGENTS), "Accept-Language": "en-US,en;q=0.9"}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.content, "html.parser")
        price_text = ""

        # --- AMAZON LOGIC ---
        if "amazon" in url:
            element = soup.find("span", {"class": "a-price-whole"})
            if element: price_text = element.get_text()
            else: 
                # Fallback for different Amazon layouts
                element = soup.find("span", {"id": "priceblock_ourprice"})
                if element: price_text = element.get_text()

        # --- FLIPKART LOGIC ---
        elif "flipkart" in url:
            element = soup.find("div", {"class": "_30jeq3 _16Jk6d"})
            if element: price_text = element.get_text()

        # --- TATA 1MG LOGIC ---
        elif "1mg" in url:
            element = soup.find("div", {"class": "Price__price___3NyX9"}) # Class names change often on 1mg
            if not element: element = soup.find("div", {"class": "DrugPriceBox__best-price___32JXw"})
            if element: price_text = element.get_text()

        # --- JIO MART LOGIC ---
        elif "jiomart" in url:
            element = soup.find("div", {"id": "price_section"})
            if element: price_text = element.get_text()

        # --- GENERIC FALLBACK ---
        # Add more blocks for Meesho, Blinkit, etc. here
        
        # CLEANUP PRICE (Remove currency symbols and commas)
        if price_text:
            clean_price = "".join([c for c in price_text if c.isdigit() or c == '.'])
            return f"₹{clean_price}"
        else:
            return "Out of Stock / Error"

    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return "Error"

def main():
    print("Bot: Reading Excel file...")
    # Read Excel - Treating all columns as strings to preserve links
    df = pd.read_excel(EXCEL_FILE, dtype=str)
    
    # Create a list of lists for Google Sheets (Starting with Headers)
    # We will copy the first 3 columns (Static Data) and then fill the rest with Prices
    final_data = []
    headers = df.columns.tolist()
    final_data.append(headers)

    print("Bot: Scraping prices (This may take time)...")
    
    for index, row in df.iterrows():
        row_data = []
        
        # 1. Copy first 3 columns directly (Brand, Product, Pack Size)
        # Assuming they are the first 3 columns in Excel
        row_data.append(str(row[0])) # Brand
        row_data.append(str(row[1])) # Product
        row_data.append(str(row[2])) # Pack Size
        
        # 2. Loop through the rest of the columns (The Links)
        # Starting from column index 3 onwards
        for col_idx in range(3, len(headers)):
            cell_value = row[col_idx]
            
            # If cell has a link, scrape it. If empty, say "Not Available"
            if isinstance(cell_value, str) and "http" in cell_value:
                # Add a small delay to be polite to servers
                time.sleep(1) 
                price = get_price(cell_value)
                row_data.append(price)
                print(f"   -> Scraped {row[1]} on {headers[col_idx]}: {price}")
            else:
                row_data.append("Not Available")
        
        final_data.append(row_data)

    # --- UPLOAD TO GOOGLE SHEETS ---
    print("Bot: Uploading to Google Sheets...")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    if GOOGLE_CREDENTIALS_JSON:
        creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
        
    client = gspread.authorize(creds)
    sheet = client.open(SHEET_NAME).sheet1
    
    # clear and update
    sheet.clear()
    sheet.update(range_name="A1", values=final_data)
    
    # --- FORMATTING ---
    # 1. Header: Dark Blue, White Text, Bold
    sheet.format("A1:Z1", {
        "backgroundColor": {"red": 0.0, "green": 0.2, "blue": 0.6},
        "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}
    })
    # 2. Static Columns (A-C): Light Grey
    sheet.format("A2:C100", {
        "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95}
    })
    
    print("Bot: Success! Prices updated.")

if __name__ == "__main__":
    main()