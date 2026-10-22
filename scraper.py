import os
import io
import re
import time
import glob
import zipfile
from urllib.parse import urlparse
import pandas as pd
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import Select
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from playwright.sync_api import sync_playwright
import sys

# Force UTF-8 encoding for Windows console to support emojis
sys.stdout.reconfigure(encoding='utf-8')
load_dotenv()
USERNAME = os.getenv("POLYMER_USERNAME")
PASSWORD = os.getenv("POLYMER_PASSWORD")

PROJECT_DIR = os.getcwd()
DOWNLOAD_DIR = os.path.join(PROJECT_DIR, "polymer_data")
ARCHIVE_DIR = os.path.join(PROJECT_DIR, "archive_data")

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)
if not os.path.exists(ARCHIVE_DIR):
    os.makedirs(ARCHIVE_DIR)

# Global ledger to track successful extraction
SUCCESSFUL_DOWNLOADS = []


def setup_driver():
    chrome_options = webdriver.ChromeOptions()

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0
    }
    chrome_options.add_experimental_option("prefs", prefs)

    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

    # Prevent renderer crashes ("Unable to receive message from renderer")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    # Bypass SSL warnings on en.pcklimited.in so the LATEX table loads
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.set_capability("acceptInsecureCerts", True)

    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    driver.execute_cdp_cmd('Page.setDownloadBehavior', {
        'behavior': 'allow',
        'downloadPath': DOWNLOAD_DIR
    })

    return driver


def clean_environment(download_dir, archive_dir):
    """CHECKPOINT 1: Archive Slate. Zips all files for backup, but keeps rolling ledgers in the workspace."""
    print("\n--- Pre-Flight Check: Archiving Workspace ---")
    stale_files = glob.glob(os.path.join(download_dir, "*.*"))

    # --- CRITICAL FIX: Files that must be backed up, but NOT deleted ---
    rolling_ledgers = ['LATEX_last_1_year_trend.xlsx']

    # Grab everything to create a complete backup snapshot
    files_to_zip = [f for f in stale_files if not (f.endswith('.png') or f.endswith('.log'))]

    if not files_to_zip:
        print("Workspace is already clean.\n")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_filename = os.path.join(archive_dir, f"polymer_data_archive_{timestamp}.zip")

    try:
        # 1. Back up EVERYTHING (including the Latex ledger) into the ZIP
        with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for f in files_to_zip:
                zipf.write(f, os.path.basename(f))

        # 2. Delete the temporary files, but KEEP the rolling ledgers active
        deleted_count = 0
        for f in files_to_zip:
            if os.path.basename(f) not in rolling_ledgers:
                os.remove(f)
                deleted_count += 1

        print(f"Backed up {len(files_to_zip)} files to {os.path.basename(zip_filename)}.")
        print(f"Cleared {deleted_count} stale files (Protected ledgers kept safe in workspace).\n")

    except PermissionError as e:
        print(f"CRITICAL ERROR: Cannot archive {os.path.basename(e.filename)}. The file is locked.")
        sys.exit(1)
    except Exception as e:
        print(f"CRITICAL ERROR during archiving: {e}")
        sys.exit(1)


def download_rubber_board_india(driver, download_dir):
    print("\n========================================")
    print("Downloading Rubber Board of India (Latex) Prices...")

    url = "http://rubberboard.org.in/public"
    target_path = os.path.join(download_dir, "LATEX_last_1_year_trend.xlsx")

    for attempt in range(3):
        try:
            driver.get(url)

            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            except Exception:
                pass

            time.sleep(3)
            tables = pd.read_html(io.StringIO(driver.page_source))

            target_table = None
            for df in tables:
                if 'latex' in df.to_string().lower():
                    target_table = df
                    break

            if target_table is None:
                raise Exception("Could not find the price table on Rubber Board website.")

            latex_price_100kg = None
            for index, row in target_table.iterrows():
                row_text = str(row.values).lower()
                if 'latex' in row_text:
                    numbers = re.findall(r'\d+(?:\.\d+)?', str(row.values))
                    valid_nums = [float(n) for n in numbers if float(n) > 100]
                    if valid_nums:
                        latex_price_100kg = valid_nums[0]
                        break

            if not latex_price_100kg:
                raise Exception("Found table, but could not parse the Latex price digits.")

            # --- CRITICAL FIX: CONVERT INR/100KG TO INR/KG ---
            # Divide by 100 and round to 2 decimal places to match the PCK format perfectly
            latex_price_kg = round(latex_price_100kg / 100, 2)

            today_date = datetime.now().strftime("%d %b %Y")
            new_data = pd.DataFrame([{
                "Date": today_date,
                "Full Advance Payment / Wet Kg Without Barrels (Supply in Tanker)": latex_price_kg
            }])

            if os.path.exists(target_path):
                try:
                    existing_df = pd.read_excel(target_path)
                    existing_df = existing_df[existing_df['Date'].astype(str) != today_date]
                    final_df = pd.concat([new_data, existing_df], ignore_index=True)
                except Exception as e:
                    print(f"  ℹ️ Could not merge with old PCK file ({e}). Starting fresh ledger.")
                    final_df = new_data
            else:
                final_df = new_data

            final_df.to_excel(target_path, index=False)

            if validate_file(target_path):
                print(f"✅ Success: Downloaded Rubber Board (Today's Latex Price: ₹{latex_price_kg}/KG)")
                SUCCESSFUL_DOWNLOADS.append("LATEX")
                return

        except Exception as e:
            print(f"⚠️ Rubber Board Attempt {attempt + 1} failed: {e}")
            time.sleep(3 * (attempt + 1))

    print("❌ FAILED: Rubber Board extraction.")


def validate_file(filepath):
    """CHECKPOINT 2: File Size Validation. Catches empty 0KB ghost files."""
    if not os.path.exists(filepath):
        return False
    if os.path.getsize(filepath) < 1024:
        os.remove(filepath)
        return False
    return True


def login(driver):
    """Logs into the PolymerUpdate portal."""
    print("Navigating to login page...")
    driver.get("https://www.polymerupdate.com/Account/Login")

    try:
        cookie_xpath = "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]"
        cookie_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, cookie_xpath)))
        cookie_btn.click()
        time.sleep(2)
    except Exception:
        pass

    try:
        user_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Security_Authentication_UserName")))
        pass_field = driver.find_element(By.ID, "currentpassword")
        user_field.send_keys(USERNAME)
        pass_field.send_keys(PASSWORD)

        login_btn = driver.find_element(By.XPATH, "//button[text()='Submit']")
        driver.execute_script("arguments[0].scrollIntoView(true);", login_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", login_btn)

        WebDriverWait(driver, 15).until(EC.url_changes("https://www.polymerupdate.com/Account/Login"))
        print("Login successful.")
    except Exception as e:
        print(f"CRITICAL: Login failed: {e}")
        sys.exit(1)


def download_data(driver, product_configs, download_dir):
    """Playwright Downloader with CHECKPOINT 3: Exponential Backoff."""
    print("\n--- Transferring Login Session to Playwright ---")
    selenium_cookies = driver.get_cookies()
    playwright_cookies = [{"name": c["name"], "value": c["value"], "domain": c["domain"], "path": c["path"]} for c in
                          selenium_cookies]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        context.add_cookies(playwright_cookies)
        page = context.new_page()

        for name, cfg in product_configs.items():
            success = False
            target_filename = f"{name}_last_1_year_trend.xlsx"
            target_path = os.path.join(download_dir, target_filename)

            for attempt in range(3):
                try:
                    url = cfg if isinstance(cfg, str) else cfg["url"]
                    guid = None if isinstance(cfg, str) else cfg.get("guid")
                    target_text = name if isinstance(cfg, str) else cfg.get("text", name)
                    period_val = "365" if isinstance(cfg, str) else cfg.get("period", "365")
                    is_direct = False if isinstance(cfg, str) else cfg.get("is_direct", False)
                    partial_match = False if isinstance(cfg, str) else cfg.get("partial_match", False)

                    print(f"\n[Playwright] {name} | Attempt {attempt + 1}/3 | {url}")
                    page.goto(url, wait_until="domcontentloaded", timeout=45000)
                    time.sleep(2)

                    if page.locator("text='Unfortunately, something went wrong on the page'").count() > 0:
                        raise Exception("PolymerUpdate server returned 500 error.")

                    tab_id = urlparse(url).fragment
                    if tab_id:
                        tab_trigger = page.locator(
                            f"a[href*='#{tab_id}'], button[data-bs-target*='{tab_id}'], [id='{tab_id}']").first
                        if tab_trigger.count() > 0 and tab_trigger.is_visible():
                            tab_trigger.click()
                            time.sleep(1.5)

                    page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(1.5)

                    if os.path.exists(target_path):
                        os.remove(target_path)

                    download_buttons = page.get_by_role("button", name="Download Historical Data")
                    if download_buttons.count() == 0:
                        download_buttons = page.locator("button, a, input[type='submit']").filter(
                            has_text="Download Historical Data")

                    target_container = None
                    target_btn = None
                    target_select = None
                    selected_option_val = None

                    for i in range(download_buttons.count()):
                        btn = download_buttons.nth(i)
                        if not btn.is_visible(): continue

                        container = btn.locator(
                            "xpath=./ancestor::div[contains(@class, 'card') or contains(@class, 'box') or contains(@class, 'table') or contains(@class, 'tab-pane') or .//select][1]")
                        if container.count() == 0: container = btn.locator("xpath=./..")

                        if is_direct:
                            target_container, target_btn = container, btn
                            break

                        select_elements = container.locator(
                            "select[id*='ProductList'], select[name*='GUID'], select[name*='Product']")
                        if select_elements.count() > 0:
                            candidate_select = select_elements.first
                            for opt in candidate_select.locator("option").all():
                                val = opt.get_attribute("value") or ""
                                txt = re.sub(r'\s+', ' ', opt.inner_text()).strip()

                                if guid and val.lower() == guid.lower():
                                    selected_option_val = val
                                    break
                                if not guid and target_text.strip().lower() == txt.lower():
                                    selected_option_val = val
                                    break
                                if not guid and partial_match and target_text.strip().lower() in txt.lower():
                                    selected_option_val = val
                                    break

                            if selected_option_val:
                                target_container, target_btn, target_select = container, btn, candidate_select
                                break
                        else:
                            target_container, target_btn = container, btn
                            break

                    if not target_btn: raise Exception("No visible download button matched criteria.")

                    target_btn.scroll_into_view_if_needed()
                    time.sleep(0.5)

                    if target_select and selected_option_val:
                        target_select.select_option(value=selected_option_val)
                        target_select.dispatch_event("change")
                        time.sleep(1)

                    period_loc = target_container.locator("select[id*='PeriodList']")
                    if period_loc.count() > 0:
                        period_select = period_loc.first
                        period_select.select_option(value=str(period_val))
                        period_select.dispatch_event("change")
                        time.sleep(1)

                    with page.expect_download(timeout=45000) as download_info:
                        target_btn.click()

                    download_info.value.save_as(target_path)

                    if validate_file(target_path):
                        print(f"✅ Success: Saved '{target_filename}'")
                        SUCCESSFUL_DOWNLOADS.append(name)
                        success = True
                        break
                    else:
                        raise Exception("Downloaded file failed size validation.")

                except Exception as e:
                    print(f"⚠️ Error on attempt {attempt + 1}: {e}")
                    time.sleep(3 * (attempt + 1))  # FIXED: Restored asterisk

            if not success:
                print(f"❌ FAILED: Could not process {name} after 3 attempts.")

        context.close()
        browser.close()


def download_cny_inr_history(download_dir):
    """Bypasses Cloudflare by using the open-source Frankfurter Exchange Rate API."""
    import requests
    print("\n========================================")
    print("Downloading CNY-INR History (Frankfurter API)...")
    target_path = os.path.join(download_dir, "CNY-INR_last_1_year_trend.xlsx")

    for attempt in range(3):
        try:
            # Query the last 365 days of data
            start_date = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")
            url = f"https://api.frankfurter.dev/v1/{start_date}..?base=CNY&symbols=INR"

            resp = requests.get(url, timeout=15)
            resp.raise_for_status()
            data = resp.json()

            if 'rates' in data and data['rates']:
                records = []
                for date_str, rate_data in data['rates'].items():
                    if 'INR' in rate_data:
                        records.append({'Date': date_str, 'Close': rate_data['INR']})

                df = pd.DataFrame(records)
                df['Date'] = pd.to_datetime(df['Date'])

                # --- CRITICAL FIX 1: Extend range to TODAY and eliminate all NaNs ---
                df.set_index('Date', inplace=True)
                full_date_range = pd.date_range(start=df.index.min(), end=datetime.now(), freq='D')
                df = df.reindex(full_date_range)

                # Forward-fill weekends, backward-fill missing start dates
                df['Close'] = df['Close'].ffill().bfill()

                df.reset_index(inplace=True)
                df.rename(columns={'index': 'Date'}, inplace=True)

                # Sort descending (newest first) to perfectly match the old scraper
                df.sort_values(by='Date', ascending=False, inplace=True)

                # --- CRITICAL FIX 2: Force DD/MM/YYYY string format ---
                df['Date'] = df['Date'].dt.strftime('%d/%m/%Y')

                if os.path.exists(target_path):
                    os.remove(target_path)

                df.to_excel(target_path, index=False)

                if validate_file(target_path):
                    print(f"✅ Success: Downloaded CNY-INR ({len(df)} contiguous records)")
                    SUCCESSFUL_DOWNLOADS.append("CNY-INR")
                    return
            raise Exception("API returned empty data.")
        except Exception as e:
            print(f"⚠️ CNY-INR Attempt {attempt + 1} failed: {e}")
            time.sleep(3 * (attempt + 1))

    print("❌ FAILED: CNY-INR extraction.")

def download_rbi_data(driver, download_dir, currency="USD"):
    print(f"\nDownloading RBI {currency}-INR...")
    target_label = "EURO" if currency == "EUR" else "USD"
    checkbox_id = f"chk{target_label}"

    # Restored to 5 years
    new_filepath = os.path.join(download_dir, f"{currency}-INR_last_5_year_trend.xls")

    for attempt in range(3):
        try:
            driver.get("https://www.rbi.org.in/scripts/referenceratearchive.aspx")
            time.sleep(4)

            # --- STEP 1: Select Currency ---
            try:
                target_checkbox = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, checkbox_id)))
                if not target_checkbox.is_selected():
                    driver.execute_script("arguments[0].click();", target_checkbox)
            except Exception:
                target_label_el = driver.find_element(By.XPATH, f"//label[contains(text(), '{target_label}')]")
                driver.execute_script("arguments[0].click();", target_label_el)

            # --- STEP 2: Set Dates (Restored Original 5-Year Range) ---
            today = datetime.now()
            from_date_str = (today - timedelta(days=365 * 5)).strftime("%d/%m/%Y")  # FIXED: Restored asterisk
            to_date_str = today.strftime("%d/%m/%Y")

            try:
                from_input = driver.find_element(By.ID, "txtFromDate")
                to_input = driver.find_element(By.ID, "txtToDate")
                driver.execute_script("arguments[0].removeAttribute('readonly');", from_input)
                driver.execute_script("arguments[0].removeAttribute('readonly');", to_input)
                driver.execute_script(f"arguments[0].value = '{from_date_str}';", from_input)
                driver.execute_script(f"arguments[0].value = '{to_date_str}';", to_input)
            except Exception as e:
                raise Exception(f"Failed to inject dates: {e}")

            # --- STEP 3: Click Submit ---
            try:
                submit_xpath = "//input[@type='submit' or @value='Go' or @value='Search']"
                go_btn = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, submit_xpath)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", go_btn)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", go_btn)
            except Exception as e:
                raise Exception(f"Failed to click Submit button: {e}")

            # --- STEP 3.5: Handle RBI Submit Alerts ---
            try:
                alert = WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert_text = alert.text
                alert.accept()
                raise Exception(f"RBI Website blocked request with popup: '{alert_text}'")
            except Exception as e:
                if "RBI Website blocked" in str(e):
                    raise e

            time.sleep(5)

            # --- STEP 4: Click Excel Download ---
            try:
                existing_files = set(glob.glob(os.path.join(download_dir, "*")))
                excel_xpath = "//*[contains(translate(text(), 'excel', 'EXCEL'), 'EXCEL') or contains(translate(@src, 'xls', 'XLS'), 'XLS')]"
                excel_btn = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, excel_xpath)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", excel_btn)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", excel_btn)
            except Exception as e:
                raise Exception(f"Failed to find Excel button: {e}")

            # --- STEP 4.5: Handle Excel Download Confirmation Popup ---
            try:
                download_alert = WebDriverWait(driver, 3).until(EC.alert_is_present())
                download_alert.accept()
            except Exception:
                pass

                # --- STEP 5: Wait & Validate ---
            start_time = time.time()
            while time.time() - start_time < 30:
                new_files = set(glob.glob(os.path.join(download_dir, "*"))) - existing_files
                valid_new_files = [f for f in new_files if not f.endswith('.crdownload') and not f.endswith('.tmp')]

                if valid_new_files:
                    time.sleep(1.5)

                    if os.path.exists(new_filepath):
                        os.remove(new_filepath)

                    try:
                        os.rename(valid_new_files[0], new_filepath)
                    except (PermissionError, OSError):
                        time.sleep(2)
                        os.rename(valid_new_files[0], new_filepath)

                    if validate_file(new_filepath):
                        print(f"✅ Success: RBI {currency}-INR")
                        SUCCESSFUL_DOWNLOADS.append(f"{currency}-INR")
                        return
                    else:
                        raise Exception("Downloaded file failed size validation (0KB Ghost File).")
                time.sleep(1)

            raise Exception("Download timed out waiting for the file to drop into the folder.")

        except Exception as e:
            print(f"⚠️ RBI {currency} Attempt {attempt + 1} failed: {e}")
            time.sleep(3 * (attempt + 1))  # FIXED: Restored asterisk

    print(f"❌ FAILED: RBI {currency} extraction.")


def download_brent_yfinance(download_dir):
    import yfinance as yf
    print("\nDownloading Brent Crude (Yahoo Finance)...")
    target_path = os.path.join(download_dir, "Brent_last_1_year_trend.xlsx")

    for attempt in range(3):
        try:
            df = yf.Ticker("BZ=F").history(period="1y")
            if not df.empty:
                df = df.reset_index()[['Date', 'Close']].rename(columns={'Close': 'Price'})
                df['Date'] = df['Date'].dt.tz_localize(None)
                df.to_excel(target_path, index=False)

                if validate_file(target_path):
                    print(f"✅ Success: Brent Crude")
                    SUCCESSFUL_DOWNLOADS.append("Brent")
                    return
            raise Exception("Empty dataset returned.")
        except Exception as e:
            print(f"⚠️ Brent Attempt {attempt + 1} failed: {e}")
            time.sleep(3 * (attempt + 1))  # FIXED: Restored asterisk
    print("❌ FAILED: Brent Crude extraction.")


def download_html_table_data(driver, url, name, download_dir):
    print(f"\nDownloading HTML Table: {name}...")
    target_path = os.path.join(download_dir, f"{name}_last_1_year_trend.xlsx")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    for attempt in range(3):
        try:
            tables = []

            # --- LAYER 1: Try Selenium (with Auto-Recovery if session died after RBI) ---
            try:
                try:
                    _ = driver.current_url
                except Exception:
                    print("  🔄 Browser session was closed. Restarting Chrome driver...")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    driver = setup_driver()
                    globals()['driver'] = driver

                driver.get(url)

                try:
                    WebDriverWait(driver, 12).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
                except Exception:
                    pass

                time.sleep(2)
                tables = pd.read_html(io.StringIO(driver.page_source))

                if tables and "westmetall.com" in url:
                    target_table = max(tables, key=lambda df: df.shape[0] * df.shape[1])
                    prev_year = str(datetime.now().year - 1)
                    try:
                        prev_link = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{prev_year}')]"))
                        )
                        driver.execute_script("arguments[0].click();", prev_link)
                        time.sleep(4)
                        prev_tables = pd.read_html(io.StringIO(driver.page_source))
                        if prev_tables:
                            target_table = pd.concat(
                                [target_table, max(prev_tables, key=lambda df: df.shape[0] * df.shape[1])],
                                ignore_index=True
                            )
                    except Exception:
                        pass
                elif tables:
                    target_table = max(tables, key=lambda df: df.shape[0] * df.shape[1])

            except Exception as sel_err:
                print(f"  ℹ️ Selenium fallback triggered for {name} ({sel_err}). Trying direct HTTP...")

            # --- LAYER 2: Direct HTTP Fallback (Immune to Chrome Crashes & SSL Blocks) ---
            if not tables:
                import requests
                import urllib3
                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

                resp = requests.get(url, headers=headers, verify=False, timeout=20)
                resp.raise_for_status()
                tables = pd.read_html(io.StringIO(resp.text))
                if not tables:
                    raise Exception("No tables found in DOM.")

                target_table = max(tables, key=lambda df: df.shape[0] * df.shape[1])

                if "westmetall.com" in url:
                    prev_year = str(datetime.now().year - 1)
                    try:
                        prev_url = f"{url}&year={prev_year}"
                        prev_resp = requests.get(prev_url, headers=headers, verify=False, timeout=20)
                        prev_tables = pd.read_html(io.StringIO(prev_resp.text))
                        if prev_tables:
                            target_table = pd.concat(
                                [target_table, max(prev_tables, key=lambda df: df.shape[0] * df.shape[1])],
                                ignore_index=True
                            )
                    except Exception:
                        pass

            if os.path.exists(target_path):
                os.remove(target_path)

            target_table.to_excel(target_path, index=False)

            if validate_file(target_path):
                print(f"✅ Success: {name}")
                SUCCESSFUL_DOWNLOADS.append(name)
                return
            raise Exception("Output file failed validation.")

        except Exception as e:
            print(f"⚠️ {name} Attempt {attempt + 1} failed: {e}")
            time.sleep(3 * (attempt + 1))  # FIXED: Restored asterisk

    print(f"❌ FAILED: {name} extraction.")


if __name__ == "__main__":
    PRODUCT_CONFIGS = {
        "PVC_Suspension": {"url": "https://polymerupdate.com/Prices/SouthAsia/Polymers/PVC#polymers",
                           "partial_match": True, "text": "Suspension Delhi"},
        "EVA": {"url": "https://polymerupdate.com/Prices/SouthAsia/OpenMarket/India/Delhi", "partial_match": True,
                "text": "EVA"},
        "PVC": {"url": "https://polymerupdate.com/Prices/SouthAsia/OpenMarket/India/Delhi", "partial_match": True,
                "text": "PVC"},
        "Naphtha": {"url": "https://polymerupdate.com/Prices/FEA/Naphtha", "partial_match": True,
                    "text": "Naphtha CFR Far East Asia"},
        "DOP": {"url": "https://polymerupdate.com/Prices/SEA/Intermediates/DOP", "partial_match": True,
                "text": "DOP CFR South East Asia"},
        "ACETONE": {"url": "https://www.polymerupdate.com/Prices/SouthAsia/ChemicalBulk/Bulk", "partial_match": True,
                    "text": "Acetone"},
        "PTA": {"url": "https://polymerupdate.com/Prices/SEA/Polymers/PET#feedstock", "partial_match": True,
                "text": "PTA CFR South East Asia"},
        "MEG": {"url": "https://polymerupdate.com/Prices/SEA/Polymers/PET#feedstock", "partial_match": True,
                "text": "MEG CFR South East Asia"},
        "PP": {"url": "https://polymerupdate.com/Prices/SouthAsia/Polymers/PP#polymers", "partial_match": True,
               "text": "BOPP CIF South Korea"},
        "LDPE": {"url": "https://www.polymerupdate.com/Prices/SouthAsia/OpenMarket/India/Delhi", "partial_match": True,
                 "text": "LDPE Film"},
        "LLDPE": {"url": "https://www.polymerupdate.com/Prices/SouthAsia/OpenMarket/India/Delhi", "partial_match": True,
                  "text": "LLDPE Film"}
    }

    html_table_urls = {
        "BRASS": "https://www.westmetall.com/en/markdaten.php?action=table&field=MB_MS_58_1",
        "ZINC": "https://www.westmetall.com/en/markdaten.php?action=table&field=LME_Zn_cash",
        # "LATEX": "https://en.pcklimited.in/Latex_isi",
    }

    EXPECTED_PRODUCTS = list(PRODUCT_CONFIGS.keys()) + list(html_table_urls.keys()) + ["USD-INR", "EUR-INR", "CNY-INR",
                                                                                       "Brent"]

    clean_environment(DOWNLOAD_DIR, ARCHIVE_DIR)

    driver = setup_driver()
    try:
        login(driver)
        download_data(driver, PRODUCT_CONFIGS, DOWNLOAD_DIR)

        download_rbi_data(driver, DOWNLOAD_DIR, currency="USD")
        download_rbi_data(driver, DOWNLOAD_DIR, currency="EUR")

        for name, url in html_table_urls.items():
            download_html_table_data(driver, url, name, DOWNLOAD_DIR)

        # --- NEW CALL ---
        download_rubber_board_india(driver, DOWNLOAD_DIR)

        # REWRITTEN CALL: Uses the new Yahoo Finance function
        download_cny_inr_history(DOWNLOAD_DIR)
        download_brent_yfinance(DOWNLOAD_DIR)

        time.sleep(5)

    finally:
        driver.quit()

        missing_products = [p for p in EXPECTED_PRODUCTS if p not in SUCCESSFUL_DOWNLOADS]
        log_path = os.path.join(PROJECT_DIR, "missing_products.log")

        if missing_products:
            print(f"\n⚠️ WARNING: The following products failed extraction: {missing_products}")
            with open(log_path, "w") as f:
                f.write(", ".join(missing_products))
        else:
            if os.path.exists(log_path):
                os.remove(log_path)
            print("\n✅ 100% Data Extraction Successful.")