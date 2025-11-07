"""
Quick test to demonstrate the HMECloud automation
"""

from hmecloud_automation import setup_chrome_driver, login_to_hmecloud, navigate_to_reports
import time

print("\n" + "="*80)
print("🍗 KFC GUYANA - HMECLOUD AUTOMATION DEMO")
print("="*80)
print("\nWatch the browser as it:")
print("  1. Opens Chrome")
print("  2. Goes to HMECloud.com")
print("  3. Enters username automatically")
print("  4. Clicks Continue")
print("  5. Enters password automatically")
print("  6. Logs in")
print("  7. Clicks REPORTS")
print("  8. Clicks Raw Car Data Report")
print("\n" + "="*80)

# Setup Chrome driver
print("\n🌐 Opening Chrome browser...")
driver = setup_chrome_driver()

try:
    # Login automatically
    login_success = login_to_hmecloud(driver)
    
    if login_success:
        print("\n✅ LOGIN SUCCESSFUL!")
        
        # Navigate to reports
        nav_success = navigate_to_reports(driver)
        
        if nav_success:
            print("\n✅ NAVIGATION SUCCESSFUL!")
            print("\n🎉 The automation is now at the Raw Car Data Report page!")
            print("\n" + "="*80)
            print("BROWSER WILL STAY OPEN FOR 30 SECONDS")
            print("You can see the Raw Car Data Report page")
            print("="*80)
            
            # Keep browser open so you can see
            for i in range(30, 0, -1):
                print(f"\rClosing in {i} seconds... ", end='', flush=True)
                time.sleep(1)
            print("\n")
        else:
            print("\n❌ Navigation failed")
            time.sleep(10)
    else:
        print("\n❌ Login failed")
        time.sleep(10)
    
except Exception as e:
    print(f"\n❌ Error: {e}")
    import traceback
    traceback.print_exc()
    time.sleep(10)

finally:
    # Close browser
    print("\nClosing browser...")
    driver.quit()
    print("✅ Done!")

print("\n" + "="*80)
print("🎉 AUTOMATION DEMO COMPLETE!")
print("="*80)

