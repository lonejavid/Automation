"""
HME CLOUD AUTOMATION
Automatically logs into HMECloud and downloads Raw Car Data reports

USAGE:
    python3 -m automation.hmecloud

Or import and use:
    from automation.hmecloud import download_all_stores
    download_all_stores()
"""

import os
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from pathlib import Path

# ========== CONFIGURATION ==========
BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data"
DOWNLOADS_FOLDER = str((DATA_DIR / "downloads").resolve())
HME_STORES_FOLDER = str((DATA_DIR / "hme_stores").resolve())

HME_CLOUD_URL = "https://hmecloud.com/"
USERNAME = "doudit@hotmail.com"
PASSWORD = "Kfcguy123!@#"

# Store list from HMECloud
STORES = [
    "(Ungrouped) 2 Vlissengen Road – KFC",
    "(Ungrouped) 5 Mandela - KFC",
    "(Ungrouped) Movie Towne - KFC",
    "(Ungrouped) Giftland Mall - KFC",
    "(Ungrouped) Sheriff Street - KFC",
    "(Ungrouped) Providence - KFC"
]

# Default date: yesterday
DEFAULT_DATE = datetime.now() - timedelta(days=1)
# ===================================


def setup_chrome_driver(download_path=None):
    """Setup Chrome driver with download preferences"""
    
    if download_path is None:
        download_path = DOWNLOADS_FOLDER
    
    # Ensure download folder exists
    os.makedirs(download_path, exist_ok=True)
    
    # Chrome options
    options = webdriver.ChromeOptions()
    
    # Set download preferences
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    # Comment this line to see browser window (helpful for debugging)
    # options.add_argument("--headless")
    
    # Additional options for stability
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    
    driver = webdriver.Chrome(options=options)
    return driver


def login_to_hmecloud(driver, username=USERNAME, password=PASSWORD):
    """Login to HMECloud - Two-step process"""
    
    print("\n" + "="*80)
    print("🌐 LOGGING INTO HMECLOUD")
    print("="*80)
    
    try:
        # Navigate to HMECloud
        print(f"   Opening {HME_CLOUD_URL}...")
        driver.get(HME_CLOUD_URL)
        
        # Wait for login page to load
        wait = WebDriverWait(driver, 20)
        
        # STEP 1: Enter username
        print(f"   Step 1: Entering username: {username}")
        
        # Try different possible selectors for username field
        username_field = None
        selectors = [
            (By.NAME, "username"),
            (By.ID, "username"),
            (By.NAME, "email"),
            (By.ID, "email"),
            (By.CSS_SELECTOR, "input[type='text']"),
            (By.CSS_SELECTOR, "input[placeholder*='username' i]"),
            (By.CSS_SELECTOR, "input[placeholder*='email' i]"),
        ]
        
        for selector_type, selector_value in selectors:
            try:
                username_field = wait.until(
                    EC.presence_of_element_located((selector_type, selector_value))
                )
                print(f"   Found username field using: {selector_type}={selector_value}")
                break
            except:
                continue
        
        if not username_field:
            print("   ❌ Could not find username field")
            return False
        
        username_field.clear()
        username_field.send_keys(username)
        print(f"   ✅ Username entered: {username}")
        time.sleep(1)
        
        # Click Continue button
        print(f"   Clicking 'Continue' button...")
        continue_button = None
        
        # Try to find Continue button
        button_selectors = [
            (By.XPATH, "//button[contains(text(), 'Continue')]"),
            (By.XPATH, "//button[contains(text(), 'continue')]"),
            (By.CSS_SELECTOR, "button[type='submit']"),
            (By.CSS_SELECTOR, "button"),
        ]
        
        for selector_type, selector_value in button_selectors:
            try:
                continue_button = driver.find_element(selector_type, selector_value)
                print(f"   Found button using: {selector_type}={selector_value}")
                break
            except:
                continue
        
        if continue_button:
            continue_button.click()
            print(f"   ✅ Clicked 'Continue'")
            time.sleep(2)
        else:
            print("   ⚠️  Continue button not found, trying to proceed anyway...")
        
        # STEP 2: Enter password
        print(f"   Step 2: Entering password...")
        
        # Wait for password field to appear
        password_field = None
        password_selectors = [
            (By.NAME, "password"),
            (By.ID, "password"),
            (By.CSS_SELECTOR, "input[type='password']"),
        ]
        
        for selector_type, selector_value in password_selectors:
            try:
                password_field = wait.until(
                    EC.presence_of_element_located((selector_type, selector_value))
                )
                print(f"   Found password field using: {selector_type}={selector_value}")
                break
            except:
                continue
        
        if not password_field:
            print("   ❌ Could not find password field")
            return False
        
        password_field.clear()
        password_field.send_keys(password)
        print(f"   ✅ Password entered")
        time.sleep(1)
        
        # Click login/submit button
        print(f"   Clicking login button...")
        login_button = None
        
        login_selectors = [
            (By.XPATH, "//button[contains(text(), 'Login')]"),
            (By.XPATH, "//button[contains(text(), 'Sign in')]"),
            (By.XPATH, "//button[contains(text(), 'Submit')]"),
            (By.CSS_SELECTOR, "button[type='submit']"),
            (By.CSS_SELECTOR, "button"),
        ]
        
        for selector_type, selector_value in login_selectors:
            try:
                login_button = driver.find_element(selector_type, selector_value)
                print(f"   Found login button using: {selector_type}={selector_value}")
                break
            except:
                continue
        
        if login_button:
            login_button.click()
            print(f"   ✅ Clicked login button")
        
        # Wait for dashboard/home page to load
        time.sleep(5)
        
        print("   ✅ Login successful!")
        return True
        
    except TimeoutException:
        print("   ❌ Login timeout - page took too long to load")
        return False
    except NoSuchElementException as e:
        print(f"   ❌ Could not find login element: {e}")
        return False
    except Exception as e:
        print(f"   ❌ Login error: {e}")
        import traceback
        traceback.print_exc()
        return False


def navigate_to_reports(driver):
    """Navigate to Raw Car Data Report"""
    
    print("\n" + "="*80)
    print("📊 NAVIGATING TO REPORTS")
    print("="*80)
    
    try:
        wait = WebDriverWait(driver, 20)
        
        # Click on Reports menu in top navigation
        print("   Clicking 'REPORTS' in top menu...")
        
        # Try different selectors for Reports menu
        reports_clicked = False
        reports_selectors = [
            (By.LINK_TEXT, "REPORTS"),
            (By.LINK_TEXT, "Reports"),
            (By.PARTIAL_LINK_TEXT, "REPORT"),
            (By.XPATH, "//a[contains(text(), 'REPORTS')]"),
            (By.XPATH, "//a[contains(text(), 'Reports')]"),
        ]
        
        for selector_type, selector_value in reports_selectors:
            try:
                reports_menu = wait.until(
                    EC.element_to_be_clickable((selector_type, selector_value))
                )
                reports_menu.click()
                print(f"   ✅ Clicked 'REPORTS' menu")
                reports_clicked = True
                break
            except:
                continue
        
        if not reports_clicked:
            print("   ❌ Could not find REPORTS menu")
            return False
        
        # Wait for Reports Overview page to load
        time.sleep(3)
        print("   Waiting for Reports page to load...")
        
        # Click on Raw Car Data Report in the left sidebar
        print("   Clicking 'Raw Car Data Report'...")
        
        raw_car_clicked = False
        raw_car_selectors = [
            (By.LINK_TEXT, "Raw Car Data Report"),
            (By.PARTIAL_LINK_TEXT, "Raw Car Data"),
            (By.XPATH, "//a[contains(text(), 'Raw Car Data Report')]"),
            (By.XPATH, "//a[contains(text(), 'Raw Car Data')]"),
        ]
        
        for selector_type, selector_value in raw_car_selectors:
            try:
                raw_car_data = wait.until(
                    EC.element_to_be_clickable((selector_type, selector_value))
                )
                raw_car_data.click()
                print(f"   ✅ Clicked 'Raw Car Data Report'")
                raw_car_clicked = True
                break
            except:
                continue
        
        if not raw_car_clicked:
            print("   ❌ Could not find Raw Car Data Report link")
            return False
        
        # Wait for Raw Car Data Report page to load
        time.sleep(3)
        print("   ✅ Navigated to Raw Car Data Report page")
        return True
        
    except TimeoutException:
        print("   ❌ Navigation timeout")
        return False
    except Exception as e:
        print(f"   ❌ Navigation error: {e}")
        import traceback
        traceback.print_exc()
        return False


def select_store_and_date(driver, store_name, report_date=None):
    """Select store from dropdown and date from calendar picker"""
    
    if report_date is None:
        report_date = DEFAULT_DATE
    
    print(f"\n   📋 Selecting store and date...")
    print(f"      Store: {store_name}")
    print(f"      Date: {report_date.strftime('%Y-%m-%d')}")
    
    try:
        wait = WebDriverWait(driver, 20)
        
        # IMPORTANT: Switch to NESTED iframes (PowerBI uses double-iframe structure)
        print(f"      Looking for iframe...")
        time.sleep(3)  # Wait for page to load
        
        try:
            # Find the first PowerBI iframe
            iframe1 = wait.until(
                EC.presence_of_element_located((By.TAG_NAME, "iframe"))
            )
            print(f"      ✅ Found first iframe, switching to it...")
            driver.switch_to.frame(iframe1)
            print(f"      ⏳ Waiting for first iframe to load...")
            time.sleep(5)  # Wait longer for first iframe to load
            
            # Now find the NESTED iframe inside
            print(f"      Looking for nested iframe...")
            iframe2 = wait.until(
                EC.presence_of_element_located((By.TAG_NAME, "iframe"))
            )
            print(f"      ✅ Found nested iframe, switching to it...")
            driver.switch_to.frame(iframe2)
            print(f"      ⏳ Waiting for PowerBI report to fully load...")
            time.sleep(10)  # Wait even longer for PowerBI report to fully load
            print(f"      ✅ Switched to nested iframe successfully")
        except Exception as e:
            print(f"      ⚠️  Error with iframes: {e}")
            # Continue anyway
        
        # STEP 1: Select store from dropdown
        print(f"      Finding store dropdown...")
        
        # Wait a bit more for PowerBI content to fully load
        time.sleep(2)
        
        # Find the store input field (Fluent UI Combobox)
        dropdown_found = False
        
        try:
            # Find the store input by ID
            store_input = wait.until(
                EC.presence_of_element_located((By.ID, "P_STORE_ID-input"))
            )
            print(f"      ✅ Found store input field")
            
            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView(true);", store_input)
            time.sleep(0.5)
            
            # Click to open dropdown
            store_input.click()
            print(f"      ✅ Clicked store dropdown")
            dropdown_found = True
            time.sleep(1.5)
            
        except Exception as e:
            print(f"      ❌ Could not find store dropdown: {e}")
            return False
        
        # Type the store name into the input field (Fluent UI autocomplete)
        print(f"      Typing store name: '{store_name}'...")
        
        try:
            # Clear the field first
            store_input.clear()
            time.sleep(0.5)
            
            # Type the store name
            store_input.send_keys(store_name)
            print(f"      ✅ Typed store name")
            time.sleep(2)  # Wait for autocomplete to show
            
            # Press Enter or click on the matching option
            store_input.send_keys(Keys.ENTER)
            print(f"      ✅ Store selected: {store_name}")
            time.sleep(1)
            
        except Exception as e:
            print(f"      ❌ Could not select store: {e}")
            return False
        
        # STEP 2: Select date from calendar picker
        print(f"      Finding date picker...")
        
        try:
            # Find the calendar icon button (the button with calendar SVG)
            # It's in a span with role="button" next to the date input
            calendar_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "span[role='button'][aria-labelledby='datePicker-input8']"))
            )
            print(f"      ✅ Found calendar icon button")
            
            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView(true);", calendar_button)
            time.sleep(0.5)
            
            # Click on the calendar icon to open calendar
            calendar_button.click()
            print(f"      ✅ Clicked calendar icon to open calendar")
            time.sleep(2)  # Wait for calendar to appear
            
            # Format date (yesterday = one day before current date)
            formatted_date = report_date.strftime("%m/%d/%Y")
            target_day = report_date.day
            
            print(f"      Looking for day {target_day} in calendar...")
            
            # Try to click the day in the calendar (Fluent UI calendar)
            day_clicked = False
            day_selectors = [
                # Fluent UI uses button elements in table cells
                (By.XPATH, f"//button[@data-is-focusable='true' and normalize-space(text())='{target_day}']"),
                (By.XPATH, f"//td[@role='gridcell']//button[normalize-space(text())='{target_day}']"),
                (By.XPATH, f"//button[contains(@class, 'dayButton') and text()='{target_day}']"),
                (By.XPATH, f"//button[text()='{target_day}' and not(contains(@class, 'disabled'))]"),
            ]
            
            for selector_type, selector_value in day_selectors:
                try:
                    # Wait a bit and try to find the day button
                    day_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((selector_type, selector_value))
                    )
                    day_button.click()
                    print(f"      ✅ Clicked day {target_day} from calendar")
                    print(f"      ✅ Selected date: {formatted_date}")
                    day_clicked = True
                    break
                except:
                    continue
            
            if not day_clicked:
                print(f"      ⚠️  Could not click calendar day {target_day}, trying to type date...")
                # Find the date input field to type
                try:
                    date_input = driver.find_element(By.ID, "datePicker-input8")
                    date_input.clear()
                    time.sleep(0.5)
                    date_input.send_keys(formatted_date)
                    print(f"      ✅ Date typed: {formatted_date}")
                except Exception as e:
                    print(f"      ❌ Could not type date: {e}")
                    return False
            
            # Wait for other fields to auto-populate
            print(f"      ⏳ Waiting for other fields to auto-populate...")
            time.sleep(5)  # Wait for Start Time, Stop Time, etc. to fill automatically
            print(f"      ✅ Other fields should be populated now")
            
        except Exception as e:
            print(f"      ❌ Could not select date: {e}")
            return False
        
        print(f"      ✅ Store and date selected successfully!")
        
        # Wait for spinner overlay to clear after date selection
        print(f"      ⏳ Checking for spinner overlay after date selection...")
        time.sleep(2)
        
        try:
            # Wait for spinner to disappear (up to 20 seconds)
            WebDriverWait(driver, 20).until_not(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-testid='spinner']"))
            )
            print(f"      ✅ Spinner cleared, button is now clickable")
        except:
            print(f"      ⚠️  No spinner found or already gone")
        
        # STEP 3: Click "View Report" button
        print(f"      Waiting before clicking 'View Report'...")
        time.sleep(2)  # Small wait to ensure form is ready
        print(f"      Clicking 'View Report' button...")
        
        try:
            # Find the View Report button by data-testid
            view_report_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-testid='parameter-pane-submit-action']"))
            )
            
            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView(true);", view_report_button)
            time.sleep(0.5)
            
            # Click the button (FIRST TIME)
            view_report_button.click()
            print(f"      ✅ Clicked 'View Report' button (1st time)")
            
            # Wait for spinner to appear after first click
            print(f"      ⏳ Waiting for spinner overlay to appear after first click...")
            time.sleep(3)
            
            # Wait for spinner to disappear (this might take a while)
            print(f"      ⏳ Waiting for spinner overlay to disappear...")
            spinner_gone = False
            for attempt in range(60):  # Check for up to 60 seconds
                try:
                    spinners = driver.find_elements(By.CSS_SELECTOR, "div[data-testid='spinner']")
                    if len(spinners) == 0:
                        print(f"      ✅ Spinner disappeared after {attempt*1} seconds")
                        spinner_gone = True
                        break
                except:
                    break
                time.sleep(1)
                if attempt % 10 == 0 and attempt > 0:
                    print(f"         Still waiting... ({attempt}s)")
            
            if not spinner_gone:
                print(f"      ⚠️  Spinner timeout, proceeding anyway...")
            
            # Additional 25 second wait as requested
            print(f"      ⏳ Waiting additional 25 seconds...")
            for i in range(25, 0, -5):
                print(f"         {i} seconds...")
                time.sleep(5)
            
            print(f"      ✅ Ready for second click")
            
            # STEP 4: Click "View Report" button (SECOND TIME)
            print(f"      Clicking 'View Report' button (2nd time)...")
            
            # Find the button again
            view_report_button_2 = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-testid='parameter-pane-submit-action']"))
            )
            
            # Scroll into view
            driver.execute_script("arguments[0].scrollIntoView(true);", view_report_button_2)
            time.sleep(0.5)
            
            # Click the button (SECOND TIME)
            view_report_button_2.click()
            print(f"      ✅ Clicked 'View Report' button (2nd time)")
            
            # Wait for "Loading report..." spinner to appear
            print(f"      ⏳ Waiting for 'Loading report...' spinner...")
            time.sleep(3)
            
            # Wait for spinner to disappear (report data loads)
            print(f"      ⏳ Waiting for spinner to disappear and data to load...")
            max_spinner_wait = 120  # 2 minutes max
            spinner_wait = 0
            
            while spinner_wait < max_spinner_wait:
                try:
                    # Check if spinner is still present
                    spinner = driver.find_elements(By.XPATH, "//*[contains(text(), 'Loading report')]")
                    if len(spinner) == 0:
                        print(f"      ✅ Spinner disappeared - data is loaded!")
                        break
                except:
                    break
                time.sleep(5)
                spinner_wait += 5
                if spinner_wait % 15 == 0:
                    print(f"         Still loading... ({spinner_wait}s elapsed)")
            
            # Wait 15 seconds after data loads (as requested)
            print(f"      ⏳ Waiting 15 seconds for data to fully render...")
            time.sleep(15)
            print(f"      ✅ Data should be fully displayed now!")
            
            # STEP 5: Click "Export" button (stays in iframe)
            print(f"      Clicking 'Export' button in toolbar...")
            
            export_selectors = [
                (By.CSS_SELECTOR, "button[data-testid='toolbar-export-dropdown']"),
                (By.XPATH, "//button[@data-testid='toolbar-export-dropdown']"),
                (By.XPATH, "//button[@title='Export']"),
                (By.XPATH, "//button[contains(@class, 'ms-CommandBarItem')]//span[text()='Export']/ancestor::button"),
                (By.XPATH, "//span[@id and text()='Export']/ancestor::button"),
            ]
            
            export_button = None
            for selector_type, selector_value in export_selectors:
                try:
                    export_button = wait.until(
                        EC.element_to_be_clickable((selector_type, selector_value))
                    )
                    break
                except Exception:
                    continue
            
            if export_button is None:
                print(f"      ❌ Could not locate Export dropdown button")
                return False
            
            export_opened = False
            menu_container = None
            for attempt in range(1, 6):
                print(f"      🔄 Attempt {attempt}/5 to open Export dropdown...")
                try:
                    ActionChains(driver).move_to_element(export_button).pause(0.2).click().perform()
                    time.sleep(1.0)
                except Exception as e:
                    print(f"         ⚠️ ActionChains click failed: {e}. Trying JavaScript click")
                    try:
                        driver.execute_script("arguments[0].click();", export_button)
                        time.sleep(1.0)
                    except Exception as e2:
                        print(f"         ⚠️ JavaScript click also failed: {e2}")
                
                # Check for menu inside current iframe
                try:
                    menu_container = driver.find_element(By.XPATH, "//div[@role='menu']")
                    export_opened = True
                    print(f"         ✅ Export menu detected in iframe")
                    break
                except Exception:
                    pass
                
                # Try keyboard activation
                try:
                    export_button.send_keys(Keys.ENTER)
                    time.sleep(1.0)
                    menu_container = driver.find_element(By.XPATH, "//div[@role='menu']")
                    export_opened = True
                    print(f"         ✅ Export menu opened via ENTER")
                    break
                except Exception:
                    pass
                
                try:
                    export_button.send_keys(Keys.SPACE)
                    time.sleep(1.0)
                    menu_container = driver.find_element(By.XPATH, "//div[@role='menu']")
                    export_opened = True
                    print(f"         ✅ Export menu opened via SPACE")
                    break
                except Exception:
                    pass
            
            if not export_opened or menu_container is None:
                print(f"      ❌ Export dropdown did not open after multiple attempts")
                return False
            
            # STEP 6: Click "Microsoft Excel (.xlsx)" from dropdown
            print(f"      Selecting 'Microsoft Excel (.xlsx)' from dropdown...")
            excel_option = None
            for attempt in range(1, 4):
                try:
                    excel_option = menu_container.find_element(
                        By.XPATH,
                        "./descendant::button[@role='menuitem']//span[contains(text(),'Microsoft Excel (.xlsx)')]"
                    )
                    if excel_option:
                        break
                except Exception:
                    pass
                time.sleep(1)
            
            if excel_option is None:
                print(f"      ❌ Could not find Microsoft Excel option in dropdown")
                return False
            
            try:
                driver.execute_script("arguments[0].click();", excel_option)
                print(f"      ✅ Clicked 'Microsoft Excel (.xlsx)' from dropdown")
            except Exception as e:
                print(f"      ❌ Failed to click Microsoft Excel option: {e}")
                return False
            
            print(f"      ⏳ Waiting for download to complete...")
            time.sleep(10)
            print(f"      ✅ Download complete!")
            print(f"      📥 Excel file saved to downloads folder")
            
            # Run DT macro on downloaded file AUTOMATICALLY
            # This ensures every downloaded file gets the macro applied
            print(f"\n      " + "="*70)
            print(f"      🔄 AUTOMATICALLY RUNNING DT MACRO ON DOWNLOADED FILE")
            print(f"      " + "="*70)
            try:
                from .run_macro import process_downloaded_file
                # Wait for file to finish downloading, then run macro
                macro_success = process_downloaded_file(wait_for_download=True)
                if macro_success:
                    print(f"\n      ✅✅✅ FILE DOWNLOADED AND CONVERTED SUCCESSFULLY! ✅✅✅")
                else:
                    print(f"\n      ⚠️  WARNING: File downloaded but macro execution failed")
                    print(f"      💡 The file is in the downloads folder but needs manual macro execution")
                    print(f"      📁 Check the downloads folder for the file")
            except Exception as e:
                print(f"\n      ❌ ERROR: Could not run macro automatically: {e}")
                import traceback
                traceback.print_exc()
                print(f"      📥 File was downloaded successfully but macro must be run manually")
                print(f"      📁 File location: {DOWNLOADS_FOLDER}")
            
        except Exception as e:
            print(f"      ❌ Could not click View Report button: {e}")
            # Switch back to default content
            try:
                driver.switch_to.default_content()
            except:
                pass
            return False
        
        # Switch back to default content
        try:
            driver.switch_to.default_content()
            print(f"      Switched back to main content")
        except:
            pass
        
        return True
        
    except TimeoutException:
        print(f"      ❌ Timeout selecting store/date")
        # Switch back to default content
        try:
            driver.switch_to.default_content()
        except:
            pass
        return False
    except Exception as e:
        print(f"      ❌ Error selecting store/date: {e}")
        import traceback
        traceback.print_exc()
        # Switch back to default content
        try:
            driver.switch_to.default_content()
        except:
            pass
        return False


def download_store_report(driver, store_name, report_date=None):
    """Download report for a specific store and date"""
    
    if report_date is None:
        report_date = DEFAULT_DATE
    
    print(f"\n   📥 Downloading: {store_name}")
    print(f"      Date: {report_date.strftime('%Y-%m-%d')}")
    
    downloads_dir = Path(DOWNLOADS_FOLDER)
    if downloads_dir.exists():
        existing_files = sorted(downloads_dir.glob("*.xlsx"))
        if existing_files:
            print("   📁 Existing export found in downloads folder")
            try:
                from .run_macro import process_downloaded_file
                print("   🔄 Skipping new download and converting existing file...")
                if process_downloaded_file(downloads_dir):
                    print("   ✅ Existing file processed successfully")
                    return True
                else:
                    print("   ⚠️  Existing file conversion failed; continuing with fresh download")
            except Exception as e:
                print(f"   ⚠️  Could not process existing file: {e}")
                print("   🔁 Proceeding with fresh download")
    
    try:
        wait = WebDriverWait(driver, 60)
        
        # Select store and date
        if not select_store_and_date(driver, store_name, report_date):
            print(f"      ❌ Failed to select store/date for {store_name}")
            return False
        
        # Click "View Report" button
        print(f"      Clicking 'View Report'...")
        view_report_clicked = False
        
        view_report_selectors = [
            (By.XPATH, "//button[contains(text(), 'View report')]"),
            (By.XPATH, "//button[contains(text(), 'View Report')]"),
            (By.CSS_SELECTOR, "button.view-report"),
            (By.XPATH, "//button[contains(@class, 'view')]"),
        ]
        
        for selector_type, selector_value in view_report_selectors:
            try:
                view_report_btn = wait.until(
                    EC.element_to_be_clickable((selector_type, selector_value))
                )
                view_report_btn.click()
                print(f"      ✅ Clicked 'View Report'")
                view_report_clicked = True
                time.sleep(5)  # Wait for report to load
                break
            except:
                continue
        
        if not view_report_clicked:
            print(f"      ❌ Could not click 'View Report' button")
            return False
        
        # Click "Export" button
        print(f"      Clicking 'Export'...")
        export_btn = wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Export"))
        )
        export_btn.click()
        time.sleep(1)
        
        # Click "Microsoft Excel" option
        print(f"      Selecting 'Microsoft Excel'...")
        excel_option = wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Microsoft Excel"))
        )
        excel_option.click()
        
        # Wait for download to complete
        time.sleep(5)
        
        print(f"      ✅ Downloaded: {store_name}")
        return True
        
    except TimeoutException:
        print(f"      ❌ Timeout downloading {store_name}")
        return False
    except Exception as e:
        print(f"      ❌ Error downloading {store_name}: {e}")
        import traceback
        traceback.print_exc()
        return False


def download_all_stores(stores=None, report_date=None, download_path=None):
    """Download reports for all stores"""
    
    if stores is None:
        stores = STORES
    
    if report_date is None:
        report_date = DEFAULT_DATE
    
    if download_path is None:
        download_path = DOWNLOADS_FOLDER
    
    print("\n" + "="*80)
    print("🍗 KFC GUYANA - HME CLOUD AUTOMATION")
    print("="*80)
    print(f"Target date: {report_date.strftime('%B %d, %Y')}")
    print(f"Stores to download: {len(stores)}")
    print(f"Download folder: {download_path}")
    print("="*80)
    
    downloads_dir = Path(DOWNLOADS_FOLDER)
    existing_files = list(downloads_dir.glob("*.xlsx"))
    if existing_files:
        print("\n   📁 Existing export detected in downloads folder")
        try:
            from .run_macro import process_downloaded_file
            print("   🔄 Skipping new download and converting existing file...")
            if process_downloaded_file(downloads_dir):
                print("   ✅ Existing file processed successfully")
                return True
            else:
                print("   ⚠️  Existing file conversion failed; proceeding to download a fresh copy")
        except Exception as e:
            print(f"   ⚠️  Could not process existing file: {e}")
            print("   🔁 Proceeding with fresh download")
    
    # Setup Chrome driver
    driver = setup_chrome_driver(download_path)
    
    try:
        # Login
        if not login_to_hmecloud(driver):
            print("\n❌ Login failed! Exiting...")
            return False
        
        # Navigate to reports
        if not navigate_to_reports(driver):
            print("\n❌ Could not navigate to reports! Exiting...")
            return False
        
        # Download each store
        print("\n" + "="*80)
        print("📥 DOWNLOADING REPORTS")
        print("="*80)
        
        successful_downloads = 0
        failed_downloads = []
        
        for i, store in enumerate(stores, 1):
            print(f"\n[{i}/{len(stores)}]")
            
            if download_store_report(driver, store, report_date):
                successful_downloads += 1
            else:
                failed_downloads.append(store)
            
            # Small delay between downloads
            time.sleep(2)
        
        # Summary
        print("\n" + "="*80)
        print("✅ DOWNLOAD COMPLETE!")
        print("="*80)
        print(f"Successful: {successful_downloads}/{len(stores)}")
        
        if failed_downloads:
            print(f"\n❌ Failed downloads:")
            for store in failed_downloads:
                print(f"   - {store}")
        
        print(f"\nFiles saved to: {download_path}")
        print("="*80)
        
        # Keep browser open for a few seconds
        print("\nClosing browser in 5 seconds...")
        time.sleep(5)
        
        return successful_downloads == len(stores)
        
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        return False
        
    finally:
        # Close browser
        driver.quit()
        print("Browser closed.")


def download_single_store(store_name, report_date=None, download_path=None):
    """Download report for a single store"""
    
    if report_date is None:
        report_date = DEFAULT_DATE
    
    if download_path is None:
        download_path = DOWNLOADS_FOLDER
    
    print("\n" + "="*80)
    print("🍗 KFC GUYANA - SINGLE STORE DOWNLOAD")
    print("="*80)
    print(f"Store: {store_name}")
    print(f"Date: {report_date.strftime('%B %d, %Y')}")
    print(f"Download folder: {download_path}")
    print("="*80)
    
    driver = setup_chrome_driver(download_path)
    
    try:
        # Login
        if not login_to_hmecloud(driver):
            print("\n❌ Login failed! Exiting...")
            return False
        
        # Navigate to reports
        if not navigate_to_reports(driver):
            print("\n❌ Could not navigate to reports! Exiting...")
            return False
        
        # Download the store
        success = download_store_report(driver, store_name, report_date)
        
        if success:
            print("\n" + "="*80)
            print("✅ DOWNLOAD COMPLETE!")
            print("="*80)
            print(f"File saved to: {download_path}")
        
        # Keep browser open for a few seconds
        print("\nClosing browser in 5 seconds...")
        time.sleep(5)
        
        return success
        
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        return False
        
    finally:
        driver.quit()
        print("Browser closed.")


def interactive_download():
    """Interactive mode - let user choose store and date"""
    
    print("\n" + "="*80)
    print("🍗 KFC GUYANA - INTERACTIVE DOWNLOAD")
    print("="*80)
    
    # Choose download type
    print("\nDownload Options:")
    print("1. Download ALL stores (recommended)")
    print("2. Download SINGLE store")
    
    choice = input("\nEnter your choice (1 or 2): ").strip()
    
    # Choose date
    print("\nDate Options:")
    print("1. Yesterday (default)")
    print("2. Custom date")
    
    date_choice = input("Enter your choice (1 or 2): ").strip()
    
    if date_choice == "2":
        date_str = input("Enter date (YYYY-MM-DD): ").strip()
        try:
            report_date = datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            print("Invalid date format. Using yesterday.")
            report_date = DEFAULT_DATE
    else:
        report_date = DEFAULT_DATE
    
    # Execute based on choice
    if choice == "2":
        # Single store
        print("\nAvailable Stores:")
        for i, store in enumerate(STORES, 1):
            print(f"{i}. {store}")
        
        store_idx = input(f"\nSelect store (1-{len(STORES)}): ").strip()
        try:
            store_idx = int(store_idx) - 1
            if 0 <= store_idx < len(STORES):
                download_single_store(STORES[store_idx], report_date)
            else:
                print("Invalid store number!")
        except ValueError:
            print("Invalid input!")
    else:
        # All stores
        download_all_stores(report_date=report_date)


if __name__ == "__main__":
    # Run interactive download
    interactive_download()

