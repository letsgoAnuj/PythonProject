import os
import io
import time
import glob
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
# Load credentials from a secure .env file
load_dotenv()
USERNAME = os.getenv("POLYMER_USERNAME")
PASSWORD = os.getenv("POLYMER_PASSWORD")

# Define where you want the Excel files to be saved
DOWNLOAD_DIR = os.path.join(os.getcwd(), "polymer_data")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# Configure Chrome to automatically save downloads to our folder
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)


# chrome_options.add_argument("--headless") # Uncomment to run invisibly

def setup_driver():
    """Initializes the Chrome WebDriver."""
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)


def login(driver):
    """Logs into the PolymerUpdate portal."""
    print("Navigating to login page...")
    driver.get("https://www.polymerupdate.com/Account/Login")

    # 1. HANDLE COOKIE BANNER FIRST
    try:
        print("Checking for cookie banner...")
        cookie_xpath = "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]"
        cookie_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, cookie_xpath))
        )
        cookie_btn.click()
        print("Cookie banner dismissed.")
        time.sleep(2)  # Give the page 2 seconds to settle after the banner closes
    except Exception:
        print("No cookie banner detected, proceeding...")
        # Take a picture of the invisible browser to see what went wrong
        screenshot_path = os.path.join(DOWNLOAD_DIR, "error_screenshot_RBI.png")
        driver.save_screenshot(screenshot_path)
        print(f"Saved crash screenshot to {screenshot_path}")

    # 2. LOCATE AND FILL FIELDS AFTER THE BANNER IS GONE
    try:
        # UPDATE THESE IDs IF NECESSARY (e.g., By.NAME, "user_password", etc.)
        user_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Security_Authentication_UserName"))
        )
        pass_field = driver.find_element(By.ID, "currentpassword")

        user_field.send_keys(USERNAME)
        pass_field.send_keys(PASSWORD)

        # 3. LOCATE AND CLICK THE SUBMIT BUTTON
        login_btn = driver.find_element(By.XPATH, "//button[text()='Submit']")

        # Scroll and force click to avoid "interactable" errors
        driver.execute_script("arguments[0].scrollIntoView(true);", login_btn)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", login_btn)

        # 4. WAIT FOR DASHBOARD TO LOAD
        WebDriverWait(driver, 15).until(
            EC.url_changes("https://www.polymerupdate.com/Account/Login")
        )
        print("Login successful.")

    except Exception as e:
        print(f"Login failed: {e}")
        # Take a picture of the invisible browser to see what went wrong
        screenshot_path = os.path.join(DOWNLOAD_DIR, "error_screenshot_RBI.png")
        driver.save_screenshot(screenshot_path)
        print(f"Saved crash screenshot to {screenshot_path}")
        driver.quit()
        exit()


# ==============================================================================
# CONFIGURATION REPOSITORY
# Add or update links and selectors here as you inspect each page.
# Each entry supports:
#   - 'url': The target URL (with #hash if applicable)
#   - 'guid': The exact <option value="..."> (preferred). Set to None if using text.
#   - 'text': Visible text to match if GUID is not yet known.
#   - 'period': Value for PeriodList (defaults to "365" for 1-year history)
# ==============================================================================


import os
import time
import re
from urllib.parse import urlparse
from playwright.sync_api import sync_playwright


def download_data(driver, product_configs, download_dir):
    """
    Hybrid Playwright Downloader.
    Resolves:
      1. ASP.NET single-form scope collapse via tight card-level ancestor isolation.
      2. Substring collisions via exact option text/GUID matching.
      3. Unfired AJAX triggers via explicit change event dispatch.
      4. Inactive tab masking and server-side 500 failure detection.
      5. Hidden HTML characters (non-breaking spaces) via aggressive Regex stripping.
      6. Direct Page downloads (bypassing dropdowns entirely).
    """
    print("\n--- Transferring Login Session to Playwright ---")

    # 1. Extract cookies from the authenticated Selenium session
    selenium_cookies = driver.get_cookies()
    playwright_cookies = [
        {
            "name": cookie["name"],
            "value": cookie["value"],
            "domain": cookie["domain"],
            "path": cookie["path"]
        }
        for cookie in selenium_cookies
    ]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        context.add_cookies(playwright_cookies)
        page = context.new_page()

        for name, cfg in product_configs.items():
            if isinstance(cfg, str):
                url = cfg
                guid = None
                target_text = name
                period_val = "365"
                is_direct = False
                partial_match = False
            else:
                url = cfg["url"]
                guid = cfg.get("guid")
                target_text = cfg.get("text", name)
                period_val = cfg.get("period", "365")
                is_direct = cfg.get("is_direct", False)
                partial_match = cfg.get("partial_match", False)

            print(f"\n========================================")
            print(f"[Playwright] Navigating to {name}: {url}")
            print(f"========================================")

            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
                time.sleep(2)

                # Pre-check: Detect server-side crash pages (like Brent 500 error)
                if page.locator("text='Unfortunately, something went wrong on the page'").count() > 0:
                    raise Exception("PolymerUpdate server returned an internal error page.")

                # Handle internal tab selection (#polymers, #feedstock, etc.)
                tab_id = urlparse(url).fragment
                if tab_id:
                    tab_trigger = page.locator(
                        f"a[href*='#{tab_id}'], button[data-bs-target*='{tab_id}'], [id='{tab_id}']").first
                    if tab_trigger.count() > 0 and tab_trigger.is_visible():
                        tab_trigger.click()
                        time.sleep(1.5)

                # Force scroll to load lazy-rendered bottom tables
                page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1.5)

                target_filename = f"{name}_last_1_year_trend.xlsx"
                target_path = os.path.join(download_dir, target_filename)

                if os.path.exists(target_path):
                    os.remove(target_path)

                # ========================================================
                # ISOLATE TARGET CARD & AVOID GLOBAL ASP.NET FORM TRAP
                # ========================================================
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
                    if not btn.is_visible():
                        continue

                    # Search upwards for the closest local card/box container
                    container = btn.locator(
                        "xpath=./ancestor::div[contains(@class, 'card') or contains(@class, 'box') or contains(@class, 'table') or contains(@class, 'tab-pane') or .//select][1]"
                    )
                    if container.count() == 0:
                        container = btn.locator("xpath=./..")

                    # DIRECT MATCH: If it's a direct page (like Brent), bypass dropdown checks entirely
                    if is_direct:
                        target_container = container
                        target_btn = btn
                        break

                    select_elements = container.locator(
                        "select[id*='ProductList'], select[name*='GUID'], select[name*='Product']")

                    if select_elements.count() > 0:
                        candidate_select = select_elements.first
                        option_handles = candidate_select.locator("option").all()
                        found = False

                        for opt in option_handles:
                            val = opt.get_attribute("value") or ""

                            # AGGRESSIVE CLEANING: Strip non-breaking spaces (\xa0), tabs, and trailing spaces
                            raw_txt = opt.inner_text()
                            txt = re.sub(r'\s+', ' ', raw_txt).strip()

                            # 1. Exact GUID Match (If explicitly provided)
                            if guid and val.lower() == guid.lower():
                                found = True
                                selected_option_val = val
                                break

                            # 2. Exact Text Match (Bulletproofed against hidden spaces)
                            if not guid and target_text.strip().lower() == txt.lower():
                                found = True
                                selected_option_val = val
                                print(f"  -> Text match accepted for '{txt}' (Extracted live GUID: {val})")
                                break

                            # 3. Partial Text Match (For grades like "EVA 18%")
                            if not guid and partial_match and target_text.strip().lower() in txt.lower():
                                found = True
                                selected_option_val = val
                                print(f"  -> Partial match accepted: '{txt}' (Extracted live GUID: {val})")
                                break

                        if found:
                            target_container = container
                            target_btn = btn
                            target_select = candidate_select
                            break
                    else:
                        # Fallback for pages without product dropdowns
                        target_container = container
                        target_btn = btn
                        break

                if not target_btn:
                    raise Exception(f"No visible download container matched product criteria for '{target_text}'.")

                # Scroll into view
                target_btn.scroll_into_view_if_needed()
                time.sleep(0.5)

                # ========================================================
                # SELECT OPTIONS & DISPATCH AJAX EVENTS
                # ========================================================
                if target_select and selected_option_val:
                    target_select.select_option(value=selected_option_val)
                    target_select.dispatch_event("change")
                    print(f"✓ Selected product option (Value: {selected_option_val}) and dispatched change event.")
                    time.sleep(1)

                period_loc = target_container.locator("select[id*='PeriodList']")
                if period_loc.count() > 0:
                    period_select = period_loc.first
                    period_select.select_option(value=str(period_val))
                    period_select.dispatch_event("change")
                    print(f"✓ Time Period set to {period_val} days.")
                    time.sleep(1)

                # ========================================================
                # EXECUTE DOWNLOAD & DIRECT SAVE
                # ========================================================
                with page.expect_download(timeout=45000) as download_info:
                    target_btn.click()

                download = download_info.value
                download.save_as(target_path)
                print(f"✅ Success: Downloaded and saved as '{target_filename}'")

            except Exception as e:
                print(f"❌ Failed to process {name}: {e}")
                screenshot_path = os.path.join(download_dir, f"error_{name}_playwright.png")
                page.screenshot(path=screenshot_path)
                print(f"Saved crash screenshot to {screenshot_path}")

        print("\n--- Playwright downloads complete. Returning control to Selenium ---")
        context.close()
        browser.close()


def download_cny_inr_history(download_dir):
    """
    Scrapes the full current year's CNY-INR exchange rate history table
    from exchangerates.org.uk by navigating directly to the year's archive page.
    """
    import pandas as pd
    import io
    from datetime import datetime

    print("\n========================================")
    print("[Playwright] Navigating to CNY-INR History")
    print("========================================")

    current_year = str(datetime.now().year)

    # 1. BYPASS DROPDOWN: Navigate directly to the dedicated yearly archive URL
    url = f"https://www.exchangerates.org.uk/CNY-INR-spot-exchange-rates-history-{current_year}.html"

    target_filename = "CNY-INR_last_1_year_trend.xlsx"
    target_path = os.path.join(download_dir, target_filename)

    try:
        if os.path.exists(target_path):
            os.remove(target_path)

        with sync_playwright() as p:
            # Launch browser with a standard User-Agent to prevent getting blocked
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

            print(f"Fetching full data directly for year: {current_year}...")
            page.goto(url, wait_until="domcontentloaded", timeout=45000)

            # Wait explicitly for the first table to appear in the DOM
            page.wait_for_selector("table", timeout=15000)

            # Extract the raw HTML of the fully updated page
            html_content = page.content()

            # 2. EXTRACT ALL HTML TABLES
            tables = pd.read_html(io.StringIO(html_content))

            valid_tables = []

            # 3. FILTER AND STITCH ALL MONTHLY TABLES TOGETHER
            for df in tables:
                # Only grab tables that actually contain exchange rate data headers
                if 'Date' in df.columns and 'Close' in df.columns:
                    valid_tables.append(df)

            if valid_tables:
                # Combine all the individual month tables into one master DataFrame
                history_df = pd.concat(valid_tables, ignore_index=True)

                # Clean up any completely empty rows or columns
                history_df.dropna(how='all', axis=1, inplace=True)
                history_df.dropna(how='all', axis=0, inplace=True)

                # Save it exactly like your other files
                history_df.to_excel(target_path, index=False)
                print(f"✅ Success: Downloaded {len(history_df)} days of data and saved as '{target_filename}'")
            else:
                raise Exception(f"No exchange rate data tables found for the year {current_year}.")

            browser.close()

    except Exception as e:
        print(f"❌ Failed to download CNY-INR data: {e}")


def download_rbi_data(driver, download_dir, currency="USD"):
    """
    Downloads historical rate from RBI for a specific currency.
    Usage: download_rbi_data(driver, DOWNLOAD_DIR, currency="USD")
           download_rbi_data(driver, DOWNLOAD_DIR, currency="EUR")
    """
    from selenium.common.exceptions import TimeoutException

    # RBI uses 'EURO' for its HTML IDs and Labels, but we want 'EUR' for our filename
    target_label = "EURO" if currency == "EUR" else "USD"
    checkbox_id = f"chk{target_label}"

    print(f"\nNavigating to RBI {currency}-INR...")
    driver.get("https://www.rbi.org.in/scripts/referenceratearchive.aspx")
    time.sleep(4)

    try:
        # STEP 1: Select Checkbox dynamically based on currency parameter
        try:
            # Optionally uncheck USD if we are fetching EUR (since USD is often checked by default)
            if currency == "EUR":
                try:
                    default_usd = driver.find_element(By.ID, "chkUSD")
                    if default_usd.is_selected():
                        driver.execute_script("arguments[0].click();", default_usd)
                except:
                    pass

            target_checkbox = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, checkbox_id))
            )
            if not target_checkbox.is_selected():
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_checkbox)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", target_checkbox)
                print(f"Selected {currency} currency.")
            else:
                print(f"{currency} currency already selected by default.")
        except Exception:
            print(f"Trying alternative {target_label} selector (label click)...")
            try:
                target_label_el = driver.find_element(By.XPATH, f"//label[contains(text(), '{target_label}')]")
                driver.execute_script("arguments[0].click();", target_label_el)
                print(f"Selected {currency} currency via label.")
            except Exception as e:
                print(f"Could not find {target_label} checkbox, proceeding... {e}")

        time.sleep(1)

        # STEP 2: Handle dates dynamically (5 Year Range)
        print("Applying 5-year date range...")
        today = datetime.now()
        five_years_ago = today - timedelta(days=365 * 5)

        # Format dates as DD/MM/YYYY for the RBI system
        from_date_str = five_years_ago.strftime("%d/%m/%Y")
        to_date_str = today.strftime("%d/%m/%Y")

        try:
            from_input = driver.find_element(By.ID, "txtFromDate")
            to_input = driver.find_element(By.ID, "txtToDate")
            driver.execute_script(f"arguments[0].value = '{from_date_str}';", from_input)
            driver.execute_script(f"arguments[0].value = '{to_date_str}';", to_input)
            print(f"Date range set: {from_date_str} to {to_date_str}")
        except Exception as e:
            print(f"Warning: Could not set dates automatically. {e}")

        # STEP 3: Click 'Go' / 'Submit' to refresh the table
        try:
            print("Looking for a 'Search' or 'Go' button to load the data...")
            submit_xpath = "//input[@type='submit' or @value='Go' or @value='Search']"
            go_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, submit_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", go_btn)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", go_btn)
            time.sleep(5)
        except Exception:
            print(
                "Note: No explicit 'Go' or 'Search' button found. The site may auto-fetch. Proceeding directly to Excel...")

        # Track existing files before we trigger the actual download
        existing_files = set(glob.glob(os.path.join(download_dir, "*")))

        # STEP 4: Find and click the actual Excel download button
        print("Looking for Excel download button...")
        try:
            excel_xpath = "//*[contains(translate(text(), 'excel', 'EXCEL'), 'EXCEL') or contains(translate(@title, 'excel', 'EXCEL'), 'EXCEL') or contains(translate(@alt, 'excel', 'EXCEL'), 'EXCEL') or contains(translate(@src, 'excel', 'EXCEL'), 'EXCEL') or contains(translate(@src, 'xls', 'XLS'), 'XLS')]"
            excel_btn = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, excel_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", excel_btn)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", excel_btn)
            print("Triggered Excel download.")

            # --- NEW: STEP 4.5 HANDLE BROWSER POP-UPS ---
            try:
                # Wait up to 3 seconds for a JS confirmation pop-up to appear
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                print(f"Pop-up detected: '{alert.text}'. Accepting...")
                alert.accept()  # This clicks "OK" or "Yes" on the pop-up
            except TimeoutException:
                # If no alert appears after 3 seconds, just ignore and proceed
                pass
            # --------------------------------------------

        except Exception as e:
            print(f"Warning: Could not find Excel button. {e}")
            screenshot_path = os.path.join(DOWNLOAD_DIR, "error_screenshot_RBI.png")
            driver.save_screenshot(screenshot_path)
            print(f"Saved crash screenshot to {screenshot_path}")

        # STEP 5: Wait for download and rename
        downloaded_file = None
        timeout = 45
        start_time = time.time()

        print(f"Waiting for {currency} download to complete...")
        while time.time() - start_time < timeout:
            current_files = set(glob.glob(os.path.join(download_dir, "*")))
            new_files = current_files - existing_files
            valid_new_files = [f for f in new_files if not f.endswith('.crdownload') and not f.endswith('.tmp')]

            if valid_new_files:
                downloaded_file = valid_new_files[0]
                # Cleanup duplicates if the server pushed multiple files simultaneously
                if len(valid_new_files) > 1:
                    for extra_file in valid_new_files[1:]:
                        try:
                            os.remove(extra_file)
                        except:
                            pass
                break
            time.sleep(1)

        if downloaded_file:
            # Dynamically name the file based on the currency passed in
            new_filename = f"{currency}-INR_last_5_year_trend.xls"
            new_filepath = os.path.join(download_dir, new_filename)
            if os.path.exists(new_filepath):
                os.remove(new_filepath)
            os.rename(downloaded_file, new_filepath)
            print(f"Success! File saved as: {new_filename}")
        else:
            print(f"Error: {currency} Download timed out.")

    except Exception as e:
        print(f"Failed to process RBI {currency}. Error: {e}")
        screenshot_path = os.path.join(DOWNLOAD_DIR, "error_screenshot_RBI.png")
        driver.save_screenshot(screenshot_path)
        print(f"Saved crash screenshot to {screenshot_path}")
    print("-" * 40)

def download_html_table_data(driver, url, name, download_dir):
    """Scrapes raw HTML tables for sites without direct export buttons."""
    print(f"Navigating to {name}...")
    driver.get(url)
    time.sleep(3)  # Give the page a moment to load

    try:
        # 1. Grab the raw HTML for the current page (Current Year)
        page_source = driver.page_source
        tables = pd.read_html(io.StringIO(page_source))

        if not tables:
            print(f"No table found on {url}")
            return

        # Pick the largest table on the page
        target_table = max(tables, key=lambda df: df.shape[0] * df.shape[1])

        # --- NEW: SPECIAL PAGINATION LOGIC FOR WESTMETALL ---
        if "westmetall.com" in url:
            prev_year = str(datetime.now().year - 1)  # Automatically calculates previous year
            print(f"Westmetall site detected. Fetching {prev_year} data for {name} to complete the 1-year trend...")
            try:
                # Look for the link that contains the previous year's text
                prev_year_link = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, f"//a[contains(text(), '{prev_year}')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", prev_year_link)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", prev_year_link)
                time.sleep(4)  # Wait for the previous year's table to load

                # Scrape the previous year's table
                page_source_prev = driver.page_source
                tables_prev = pd.read_html(io.StringIO(page_source_prev))
                target_table_prev = max(tables_prev, key=lambda df: df.shape[0] * df.shape[1])

                # Combine the current year and previous year tables into one DataFrame
                target_table = pd.concat([target_table, target_table_prev], ignore_index=True)
                print(f"Successfully appended {prev_year} data.")

            except Exception as e:
                print(f"Warning: Could not fetch {prev_year} data for BRASS_ZINC. Error: {e}")
        # ----------------------------------------------------

        # 2. Define the output file name
        new_filename = f"{name}_last_1_year_trend.xlsx"
        out_path = os.path.join(download_dir, new_filename)

        # 3. Delete existing file if it exists to avoid permission errors
        if os.path.exists(out_path):
            os.remove(out_path)

        # 4. Save the combined table to Excel
        target_table.to_excel(out_path, index=False)
        print(f"Success! Table saved as: {new_filename}")

    except Exception as e:
        print(f"Failed to scrape table for {name}. Error: {e}")


if __name__ == "__main__":
    # Dictionary of your target URLs
    PRODUCT_CONFIGS = {
        "PVC_Suspension": {
            "url": "https://polymerupdate.com/Prices/SouthAsia/Polymers/PVC#polymers",
            "guid": None,
            "partial_match": True,
            "text": "Suspension Delhi",
            "period": "365"
        },
        "Brent": {
            "url": "https://polymerupdate.com/Prices/SouthAsia/Crude/DatedBrent",
            "guid": None,
            "is_direct": True,
            "text": "Brent",
            "period": "365"
        },
        "EVA": {
            "url": "https://polymerupdate.com/Prices/SouthAsia/OpenMarket/India/Delhi",
            "guid": None,
            "partial_match": True,
            "text": "EVA",
            "period": "365"
        },
        "PVC": {
            "url": "https://polymerupdate.com/Prices/SouthAsia/OpenMarket/India/Delhi",
            "guid": None,
            "partial_match": True,
            "text": "PVC",
            "period": "365"
        },
        "Naphtha": {
            "url": "https://polymerupdate.com/Prices/FEA/Naphtha",
            "guid": None,
            "partial_match": True,
            "text": "Naphtha CFR Far East Asia",
            "period": "365"
        },
        "DOP": {
            "url": "https://polymerupdate.com/Prices/SEA/Intermediates/DOP",
            "guid": None,
            "partial_match": True,
            "text": "DOP CFR South East Asia",
            "period": "365"
        },
        "ACETONE": {
            "url": "https://www.polymerupdate.com/Prices/SouthAsia/ChemicalBulk/Bulk",
            "guid": None,
            "partial_match": True,
            "text": "Acetone",
            "period": "365"
        },
        "PTA": {
            "url": "https://polymerupdate.com/Prices/SEA/Polymers/PET#feedstock",
            "guid": None,
            "partial_match": True,
            "text": "PTA CFR South East Asia",
            "period": "365"
        },
        "MEG": {
            "url": "https://polymerupdate.com/Prices/SEA/Polymers/PET#feedstock",
            "guid": None,
            "partial_match": True,
            "text": "MEG CFR South East Asia",
            "period": "365"
        },
        "PP": {
            "url": "https://polymerupdate.com/Prices/SouthAsia/Polymers/PP#polymers",
            "guid": None,
            "partial_match": True,
            "text": "BOPP CIF South Korea",
            "period": "365"
        }
    }

    # HTML Table Scraping URLs
    html_table_urls = {
        "BRASS": "https://www.westmetall.com/en/markdaten.php?action=table&field=MB_MS_58_1",
        "ZINC": "https://www.westmetall.com/en/markdaten.php?action=table&field=LME_Zn_cash",
        "LATEX": "https://en.pcklimited.in/Latex_isi"
    }

    driver = setup_driver()
    try:
        # --- 1. RUN POLYMERUPDATE DATA ---
        login(driver)
        download_data(driver, PRODUCT_CONFIGS, DOWNLOAD_DIR)

        # --- 2. RUN RBI DATA ---
        download_rbi_data(driver, DOWNLOAD_DIR, currency="USD")
        download_rbi_data(driver, DOWNLOAD_DIR, currency="EUR")

        # --- 3. RUN HTML TABLE SCRAPING (Westmetall & Latex) ---
        for name, url in html_table_urls.items():
            download_html_table_data(driver, url, name, DOWNLOAD_DIR)

        # --- 4. RUN CNY-INR HISTORY (New Addition) ---
        download_cny_inr_history(DOWNLOAD_DIR)

        print("All downloads initiated. Waiting for active downloads to finish...")
        time.sleep(10)  # Final wait to ensure last file finishes saving

    finally:
        driver.quit()
        print(f"Scraping session closed. Files saved in: {DOWNLOAD_DIR}")