from datetime import datetime
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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load environment variables
load_dotenv()

# --- CONFIGURATION ---
EXCEL_FILE = "products.xlsx"
SHEET_NAME = "Price Tracker 2026"
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON")

def get_driver():
    """Sets up a Stealthy Chrome browser with Referer spoofing."""
    chrome_options = Options()
    
    # --- STEALTH SETTINGS ---
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--enable-javascript")
    
    # Fake User Agent
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    chrome_options.add_argument(f"user-agent={user_agent}")
    
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    # Hide WebDriver property
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
    """Checks JSON-LD (Hidden Data) for Price."""
    scripts = soup.find_all('script', type='application/ld+json')
    for script in scripts:
        try:
            if not script.string: continue
            data = json.loads(script.string)
            if isinstance(data, list): data = data[0]
            
            if 'offers' in data:
                offer = data['offers']
                if isinstance(offer, list): offer = offer[0]
                if 'price' in offer: return str(offer['price'])
                if 'lowPrice' in offer: return str(offer['lowPrice'])
        except:
            continue
    return None

def get_price(driver, url, product_name="Unknown"):
    """Navigates to URL, Scrolls, and Scrapes with Debugging."""
    if not isinstance(url, str) or "http" not in url:
        return "Not Available"

    try:
        driver.get(url)
        
        # SCROLL LOGIC (Triggers Lazy Load)
        driver.execute_script("window.scrollBy(0, 500);")
        time.sleep(2)
        driver.execute_script("window.scrollBy(0, 500);")
        time.sleep(3) 

        soup = BeautifulSoup(driver.page_source, "html.parser")
        price_text = None
        
        # 1. Try JSON-LD (Best for Meesho/Snapdeal)
        price_text = get_smart_price(soup)

        # 2. Try XPath Text Search (If JSON fails)
        if not price_text:
            try:
                # Look for ANY header or span containing '₹'
                # This bypasses class name changes completely
                xpath_query = "//*[contains(text(), '₹') and string-length(text()) < 15]"
                elements = driver.find_elements(By.XPATH, xpath_query)
                
                for el in elements:
                    txt = el.text.strip()
                    # Filter out garbage, keep realistic prices
                    if re.match(r"₹\s?[\d,]+", txt):
                        price_text = txt
                        break
            except:
                pass

        # 3. Platform Specific Fallbacks
        if not price_text:
            if "snapdeal" in url:
                # Snapdeal often uses 'Rs.' instead of '₹'
                try:
                    el = driver.find_element(By.CLASS_NAME, "payBlkBig")
                    price_text = el.text
                except:
                    try:
                        el = driver.find_element(By.CLASS_NAME, "pdp-final-price")
                        price_text = el.text
                    except: pass

            elif "meesho" in url:
                 # Meesho Fallback
                 try:
                     el = driver.find_element(By.XPATH, "//h4[contains(text(), '₹')]")
                     price_text = el.text
                 except: pass

        # --- DEBUG: TAKE SCREENSHOT IF FAILED ---
        if not price_text:
            print(f"      [DEBUG] Failed to find price for {product_name}. Taking screenshot...")
            safe_name = re.sub(r'\W+', '_', product_name)[:15]
            driver.save_screenshot(f"error_{safe_name}.png")
            
            # Check Page Title for Blocking
            if "Access Denied" in driver.title or "Robot" in driver.title:
                return "Blocked by Website"

        return clean_price(price_text) if price_text else "Out of Stock / Error"

    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return "Error"

def main():
    print("Bot: Starting Driver...")
    driver = get_driver()
    fetch_time = datetime.now().strftime("%d %b %Y %I:%M %p")

    
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
            brand = str(row.iloc[0])
            product = str(row.iloc[1])
            row_data.append(brand) 
            row_data.append(product) 
            row_data.append(str(row.iloc[2])) 
            
            for col_idx in range(3, len(headers)):
                cell_value = row.iloc[col_idx]
                col_name = headers[col_idx]
                
                if isinstance(cell_value, str) and "http" in cell_value:
                    print(f"   -> Scraping {product} on {col_name}...")
                    price = get_price(driver, cell_value, product)
                    print(f"      [Result]: {price}")
                    row_data.append(price)
                else:
                    row_data.append("Not Available")
            row_data.append(fetch_time)
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