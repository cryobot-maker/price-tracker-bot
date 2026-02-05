import pandas as pd
import gspread
import time
import os
import json
import random
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
    chrome_options.add_argument("--window-size=1920,1080") # Fake a real screen size
    chrome_options.add_argument("--disable-blink-features=AutomationControlled") # Hide "Automation" flag
    
    # Fake a real User Agent
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    chrome_options.add_argument(f"user-agent={user_agent}")
    
    # Exclude automation switches
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    # Execute CDP command to further hide selenium
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
    clean = "".join([c for c in text if c.isdigit() or c == '.'])
    if clean:
        try:
            return f"₹{float(clean):.2f}"
        except:
            return text
    return None

def get_price(driver, url):
    """Navigates to URL and scrapes price using Selenium."""
    if not isinstance(url, str) or "http" not in url:
        return "Not Available"

    try:
        driver.get(url)
        time.sleep(5) # Increased wait time for slow loading
        
        # DEBUG: Print title to check if blocked
        page_title = driver.title.strip()
        print(f"      [Page Title]: {page_title[:30]}...") 

        soup = BeautifulSoup(driver.page_source, "html.parser")
        price_text = None

        # --- AMAZON LOGIC ---
        if "amazon" in url:
            element = soup.find("span", {"class": "a-price-whole"})
            if not element: element = soup.find("span", {"class": "a-offscreen"})
            if not element: element = soup.find("span", {"id": "priceblock_ourprice"})
            if not element: element = soup.find("span", {"id": "priceblock_dealprice"})
            if element: price_text = element.get_text()

        # --- FLIPKART LOGIC ---
        elif "flipkart" in url:
            element = soup.find("div", {"class": "_30jeq3 _16Jk6d"})
            if not element: element = soup.find("div", {"class": "_30jeq3"})
            if element: price_text = element.get_text()

        # --- TATA 1MG LOGIC ---
        elif "1mg" in url:
            element = soup.find("div", {"class": "Price__price___3NyX9"})
            if not element: element = soup.find("div", {"class": "DrugPriceBox__best-price___32JXw"})
            if element: price_text = element.get_text()

        # --- JIO MART LOGIC ---
        elif "jiomart" in url:
            element = soup.find("div", {"id": "price_section"})
            if not element: element = soup.find("span", {"class": "final-price"})
            if element: price_text = element.get_text()
            
        # --- BLINKIT LOGIC ---
        elif "blinkit" in url:
            element = soup.find("div", {"class": "ProductVariants__Price-sc-1unev4j-2"})
            if element: price_text = element.get_text()

        # --- MOGLIX LOGIC ---
        elif "moglix" in url:
             element = soup.find("div", {"class": "p-dp-price-amount"})
             if element: price_text = element.get_text()
        
        # --- MEESHO LOGIC ---
        elif "meesho" in url:
             element = soup.find("h4", {"class": "sc-eDvSVe"})
             if element: price_text = element.get_text()

        return clean_price(price_text) if price_text else "Out of Stock / Error"

    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return "Error"

def main():
    print("Bot: Starting Stealth Driver...")
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
            
            # Use iloc for safer access
            row_data.append(str(row.iloc[0])) # Brand
            row_data.append(str(row.iloc[1])) # Product
            row_data.append(str(row.iloc[2])) # Pack Size
            
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