"""
FULL END-TO-END AUTOMATION
Combines HMECloud download + data processing

USAGE:
    python3 full_automation.py

This script does EVERYTHING:
1. Logs into HMECloud automatically
2. Downloads all store reports
3. Transforms the data
4. Updates the Drive-Thru template
5. Done!
"""

import os
import sys
from datetime import datetime, timedelta

# Import our automation modules
from hmecloud_automation import download_all_stores
from complete_automation import main as run_complete_automation

# ========== CONFIGURATION ==========
DEFAULT_DATE = datetime.now() - timedelta(days=1)
DOWNLOADS_FOLDER = "/Users/sanctum/Desktop/Automation/downloads"
# ===================================


def full_automation(report_date=None, skip_download=False):
    """
    Run complete end-to-end automation
    
    Args:
        report_date: Date to download reports for (default: yesterday)
        skip_download: If True, skip HMECloud download (use existing files)
    """
    
    if report_date is None:
        report_date = DEFAULT_DATE
    
    print("\n" + "="*80)
    print("🍗 KFC GUYANA - FULL END-TO-END AUTOMATION")
    print("="*80)
    print(f"Target date: {report_date.strftime('%B %d, %Y')}")
    print(f"Skip download: {skip_download}")
    print("="*80)
    
    # PHASE 1: Download from HMECloud
    if not skip_download:
        print("\n" + "="*80)
        print("PHASE 1: DOWNLOADING FROM HMECLOUD")
        print("="*80)
        
        download_success = download_all_stores(report_date=report_date)
        
        if not download_success:
            print("\n⚠️  Warning: Some downloads may have failed")
            response = input("Continue with processing? (y/n): ").strip().lower()
            if response != 'y':
                print("\n❌ Automation cancelled by user")
                return False
    else:
        print("\n⏭️  Skipping HMECloud download (using existing files)")
    
    # PHASE 2: Process data and update template
    print("\n" + "="*80)
    print("PHASE 2: PROCESSING DATA & UPDATING TEMPLATE")
    print("="*80)
    
    processing_success = run_complete_automation()
    
    if not processing_success:
        print("\n❌ Data processing failed!")
        return False
    
    # FINAL SUCCESS
    print("\n" + "="*80)
    print("🎉🎉🎉 FULL AUTOMATION COMPLETE! 🎉🎉🎉")
    print("="*80)
    print("\n✅ Everything is done!")
    print("\n📋 Next steps:")
    print("   1. Open the Drive-Thru template in Excel")
    print("   2. Click 'Refresh All' to update pivot tables")
    print("   3. Verify data looks correct")
    print("\n" + "="*80)
    
    return True


def interactive_mode():
    """Interactive mode with user prompts"""
    
    print("\n" + "="*80)
    print("🍗 KFC GUYANA - FULL AUTOMATION (INTERACTIVE MODE)")
    print("="*80)
    
    # Ask about download
    print("\n📥 HMECloud Download:")
    print("1. Download fresh data from HMECloud (recommended)")
    print("2. Skip download (use existing files in downloads/)")
    
    download_choice = input("\nEnter your choice (1 or 2): ").strip()
    skip_download = (download_choice == "2")
    
    # Ask about date
    print("\n📅 Date Selection:")
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
    
    # Confirm
    print("\n" + "="*80)
    print("CONFIRMATION")
    print("="*80)
    print(f"Download from HMECloud: {'Yes' if not skip_download else 'No'}")
    print(f"Date: {report_date.strftime('%B %d, %Y')}")
    
    confirm = input("\nProceed with automation? (y/n): ").strip().lower()
    
    if confirm != 'y':
        print("\n❌ Automation cancelled by user")
        return False
    
    # Run automation
    return full_automation(report_date=report_date, skip_download=skip_download)


if __name__ == "__main__":
    # Check for command line arguments
    if len(sys.argv) > 1:
        if sys.argv[1] == "--skip-download":
            # Run with skip download
            full_automation(skip_download=True)
        elif sys.argv[1] == "--help" or sys.argv[1] == "-h":
            print("\nFULL AUTOMATION SCRIPT - USAGE:")
            print("\n  python3 full_automation.py              # Interactive mode")
            print("  python3 full_automation.py --skip-download # Use existing files")
            print("  python3 full_automation.py --help          # Show this help")
            print()
        else:
            print(f"Unknown option: {sys.argv[1]}")
            print("Use --help for usage information")
    else:
        # Run in interactive mode
        interactive_mode()

