"""
Test script to verify store and date selection automation
"""

from hmecloud_automation import setup_chrome_driver, login_to_hmecloud, navigate_to_reports, select_store_and_date
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

# Setup Chrome driver
print("\n1️⃣ Opening Chrome browser...")
driver = setup_chrome_driver()

try:
    # Login
    print("\n2️⃣ Logging into HMECloud...")
    if not login_to_hmecloud(driver):
        print("❌ Login failed!")
        exit(1)
    
    print("✅ Login successful!")
    
    # Navigate to Raw Car Data Report
    print("\n3️⃣ Navigating to Raw Car Data Report...")
    if not navigate_to_reports(driver):
        print("❌ Navigation failed!")
        exit(1)
    
    print("✅ Navigation successful!")
    
    # Select store and date
    print("\n4️⃣ Selecting store and date...")
    if not select_store_and_date(driver, test_store, test_date):
        print("❌ Store/Date selection failed!")
        exit(1)
    
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
    # Close browser
    print("Closing browser...")
    driver.quit()
    print("Done!")

print("\n" + "="*80)

