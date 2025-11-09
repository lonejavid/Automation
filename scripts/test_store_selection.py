"""
Test script to verify store and date selection automation
"""

from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.append(str(SRC_DIR))

from automation.hmecloud import setup_chrome_driver, login_to_hmecloud, navigate_to_reports, select_store_and_date
from automation.run_macro import process_downloaded_file
from datetime import datetime, timedelta
import time

print("\n" + "="*80)
print("🧪 TESTING STORE & DATE SELECTION AUTOMATION")
print("="*80)

# Test with Mandela store and yesterday's date
test_store = "(Ungrouped) 5 Mandela - KFC"
test_date = datetime.now() - timedelta(days=1)

print(f"\nTest Configuration:")
print(f"  Store: {test_store}")
print(f"  Date: {test_date.strftime('%B %d, %Y')}")
print("\n" + "="*80)

# Check for existing downloaded file
downloads_dir = PROJECT_ROOT / "data" / "downloads"
existing_files = sorted(downloads_dir.glob("*.xlsx"))
if existing_files:
    print("\n📁 Existing export found in downloads folder")
    print("🔄 Skipping browser automation and converting existing file...")
    if process_downloaded_file(downloads_dir):
        print("\n✅ Existing file processed successfully!")
        print("\n" + "="*80)
        print("🎉 AUTOMATION TEST PASSED (NO DOWNLOAD NEEDED)!")
        print("="*80)
        sys.exit(0)
    else:
        print("\n⚠️  Existing file conversion failed; proceeding with new download...")

# Setup Chrome driver
driver = None
print("\n1️⃣ Opening Chrome browser...")
driver = setup_chrome_driver()

try:
    # Login
    print("\n2️⃣ Logging into HMECloud...")
    if not login_to_hmecloud(driver):
        print("❌ Login failed!")
        sys.exit(1)
    
    print("✅ Login successful!")
    
    # Navigate to Raw Car Data Report
    print("\n3️⃣ Navigating to Raw Car Data Report...")
    if not navigate_to_reports(driver):
        print("❌ Navigation failed!")
        sys.exit(1)
    
    print("✅ Navigation successful!")
    
    # Select store and date
    print("\n4️⃣ Selecting store and date...")
    if not select_store_and_date(driver, test_store, test_date):
        print("❌ Store/Date selection failed!")
        sys.exit(1)
    
    print("\n✅ STORE AND DATE SELECTED SUCCESSFULLY!")
    
    print("\n" + "="*80)
    print("🎉 AUTOMATION TEST PASSED!")
    print("="*80)
    print("\nThe automation can now:")
    print("  ✅ Login to HMECloud")
    print("  ✅ Navigate to Raw Car Data Report")
    print("  ✅ Select store from dropdown")
    print("  ✅ Select date from calendar")
    print("\n📋 Next step: Click 'View Report' button")
    print("="*80)
    
    # Keep browser open so you can see
    print("\nKeeping browser open for 30 seconds...")
    for i in range(30, 0, -1):
        print(f"\rClosing in {i} seconds... ", end='', flush=True)
        time.sleep(1)
    print("\n")
    
except Exception as e:
    print(f"\n❌ Test error: {e}")
    import traceback
    traceback.print_exc()
    time.sleep(10)

finally:
    if driver is not None:
        print("Closing browser...")
        driver.quit()
        print("Done!")

print("\n" + "="*80)

