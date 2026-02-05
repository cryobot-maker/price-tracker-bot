import pandas as pd
import gspread
import time
import os
import json
import re
from bs4 import BeautifulSoup
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

# --- SELENIUM IMPORTS ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Load environment variables
load_dotenv()

# --- CONFIGURATION ---
EXCEL_FILE = "products.xlsx"
SHEET_NAME = "Price Tracker 2026"
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")

def get_driver():
    """Sets up a Stealthy Chrome browser."""
    chrome_options = Options()
    
    # --- STEALTH MODE SETTINGS ---
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    chrome_options.add_argument(f"user-agent={user_agent}")
    
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    # Execute CDP command to hide selenium
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            })
        """
    })
    
    return driver

def clean_price(text):
    """Extracts numbers from price text."""
    if not text: return None
    # Remove everything except digits and dots
    clean = "".join([c for c in str(text) if c.isdigit() or c == '.'])
    if clean:
        try:
            val = float(clean)
            return f"₹{val:.2f}"
        except:
            return text
    return None

def get_smart_price(soup):
    """
    Attempts to find price in hidden JSON-LD or Meta tags.
    This works for Meesho, Snapdeal, Myntra, and many others automatically.
    """
    # 1. Check JSON-LD (The gold standard for bots)
    scripts = soup.find_all('script', type='application/ld+json')
    for script in scripts:
        try:
            data = json.loads(script.string)
            # Schema.org Product/Offer structure
            if isinstance(data, list): data = data[0] # Sometimes it's a list
            
            if 'offers' in data:
                offer = data['offers']
                if isinstance(offer, list): offer = offer[0]
                
                if 'price' in offer: return offer['price']
                if 'lowPrice' in offer: return offer['lowPrice']
        except:
            continue

    # 2. Check Meta Tags (OpenGraph / Twitter Cards)
    meta_price = soup.find("meta", property="og:price:amount")
    if meta_price: return meta_price.get("content")

    meta_price = soup.find("meta", property="product:price:amount")
    if meta_price: return meta_price.get("content")
    
    return None

def get_price(driver, url):
    """Navigates to URL and scrapes price."""
    if not isinstance(url, str) or "http" not in url:
        return "Not Available"

    try:
        driver.get(url)
        time.sleep(5) # Wait for JS to load
        
        soup = BeautifulSoup(driver.page_source, "html.parser")
        price_text = None
        
        # --- STRATEGY 1: SMART HIDDEN DATA (Works for Meesho/Snapdeal) ---
        smart_price = get_smart_price(soup)
        if smart_price:
            return clean_price(smart_price)

        # --- STRATEGY 2: VISUAL FALLBACKS ---
        
        # 1. MEESHO VISUAL
        if "meesho" in url:
            # Look for the large heading with ₹ symbol
            for h in soup.find_all(['h4', 'h5']):
                if '₹' in h.get_text():
                    price_text = h.get_text()
                    break

        # 2. SNAPDEAL VISUAL
        elif "snapdeal" in url:
            element = soup.find("span", {"class": "payBlkBig"})
            if not element: element = soup.find("span", {"class": "pdp-final-price"})
            if not element: element = soup.find("input", {"id": "productPrice"}) # Hidden Input
            if element: 
                price_text = element.get_text() if element.name != "input" else element.get('value')

        # 3. AMAZON VISUAL
        elif "amazon" in url:
            element = soup.find("span", {"class": "a-price-whole"})
            if not element: element = soup.find("span", {"class": "a-offscreen"})
            if element: price_text = element.get_text()

        # 4. FLIPKART VISUAL
        elif "flipkart" in url:
            element = soup.find("div", {"class": "Nx9bqj CxhGGd"})
            if not element: element = soup.find("div", {"class": "_30jeq3 _16Jk6d"})
            if element: price_text = element.get_text()

        # 5. TATA 1MG VISUAL
        elif "1mg" in url:
            element = soup.find("div", {"class": "Price__price___3NyX9"})
            if not element: element = soup.find("div", {"class": "DrugPriceBox__best-price___32JXw"})
            if element: price_text = element.get_text()

        # 6. MOGLIX VISUAL
        elif "moglix" in url:
            element = soup.find("div", {"class": "p-dp-price-amount"})
            if element: price_text = element.get_text()
        
        # 7. BLINKIT
        elif "blinkit" in url:
            element = soup.find("div", {"class": "ProductVariants__Price-sc-1unev4j-2"})
            if element: price_text = element.get_text()

        return clean_price(price_text) if price_text else "Out of Stock / Error"

    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return "Error"

def main():
    print("Bot: Starting Smart Driver...")
    driver = get_driver()
    
    try:
        print("Bot: Reading Excel file...")
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        
        final_data = []
        headers = df.columns.tolist()
        final_data.append(headers)

        print("Bot: Scraping prices...")
        
        for index, row in df.iterrows():
            row_data = []
            
            row_data.append(str(row.iloc[0]))
            row_data.append(str(row.iloc[1]))
            row_data.append(str(row.iloc[2]))
            
            for col_idx in range(3, len(headers)):
                cell_value = row.iloc[col_idx]
                col_name = headers[col_idx]
                
                if isinstance(cell_value, str) and "http" in cell_value:
                    print(f"   -> Scraping {col_name}...")
                    price = get_price(driver, cell_value)
                    print(f"      [Result]: {price}")
                    row_data.append(price)
                else:
                    row_data.append("Not Available")
            
            final_data.append(row_data)

        print("Bot: Uploading to Google Sheets...")
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        
        if GOOGLE_CREDENTIALS_JSON:
            creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        else:
            creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
            
        client = gspread.authorize(creds)
        sheet = client.open(SHEET_NAME).sheet1
        
        sheet.clear()
        sheet.update(range_name="A1", values=final_data)
        
        sheet.format("A1:Z1", {
            "backgroundColor": {"red": 0.0, "green": 0.2, "blue": 0.6},
            "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}
        })
        sheet.format("A2:C100", {
            "backgroundColor": {"red": 0.95, "green": 0.95, "blue": 0.95}
        })
        
        print("Bot: Success! Prices updated.")

    except Exception as e:
        print(f"Fatal Error: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()