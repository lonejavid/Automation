"""
COMPLETE DRIVE-THRU AUTOMATION
Automates Steps 6-10 from the procedure

USAGE:
1. Manually download Raw Car Data Excel files from HMECloud
2. Put them in: downloads/ folder
3. Run: python3 complete_automation.py
4. Done! Updated Drive Thru template ready!
"""

import os
import glob
from datetime import datetime, timedelta
from pathlib import Path
from .transform_data import transform_raw_car_data
from .template_operations import (
    create_backup,
    paste_to_template,
    concatenate_formulas,
    refresh_pivot_tables,
    update_dates,
    save_template
)

# ========== CONFIGURATION ==========
BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data"
DOWNLOADS_FOLDER = str((DATA_DIR / "downloads").resolve())
TEMPLATE_PATH = str((DATA_DIR / "templates" / "Drive Thru Optimization - KFC Guyana  (16-10)-copy.xlsx").resolve())
TARGET_SHEET = "AllStores"  # Or "Raw Data" - will auto-detect

# Columns with formulas (yellow headers) - UPDATE these column numbers
FORMULA_COLUMNS = [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]  # Columns L through V

# Date update configuration - UPDATE cell references as needed
DATE_CONFIGS = {
    "Consol Wkly time trnd": "A1",
    "Consol Wkly Txns Trnd": "A1",
    "Summary - Stores": "A1",
    "Wkly Time trend": "A1",
    "Wkly Txns Trend": "A1",
    "Day review - Txns time": "A1"
}

TARGET_DATE = datetime.now() - timedelta(days=1)  # Yesterday
# ===================================

def main():
    """Main automation workflow"""
    
    print("="*80)
    print("🍗 KFC GUYANA - COMPLETE DRIVE-THRU AUTOMATION")
    print("="*80)
    print(f"Target date: {TARGET_DATE.strftime('%B %d, %Y')}")
    print("="*80)
    
    # STEP 1: Find downloaded files
    print("\n" + "="*80)
    print("STEP 1: Finding downloaded Excel files")
    print("="*80)
    
    raw_files = glob.glob(os.path.join(DOWNLOADS_FOLDER, "*.xlsx"))
    raw_files = [f for f in raw_files if not f.endswith('_transformed.xlsx')]
    
    print(f"Found {len(raw_files)} raw Excel files:")
    for f in raw_files:
        print(f"   - {os.path.basename(f)}")
    
    if len(raw_files) == 0:
        print("\n❌ No files found in downloads/ folder!")
        print(f"   Please download Raw Car Data files to: {DOWNLOADS_FOLDER}")
        return False
    
    # STEP 2: Transform each file
    print("\n" + "="*80)
    print("STEP 2: Transforming data (Ctrl+D replacement)")
    print("="*80)
    
    transformed_dataframes = []
    
    for raw_file in raw_files:
        try:
            df = transform_raw_car_data(raw_file)
            transformed_dataframes.append(df)
        except Exception as e:
            print(f"   ❌ Error transforming {os.path.basename(raw_file)}: {e}")
    
    if len(transformed_dataframes) == 0:
        print("\n❌ No data transformed successfully!")
        return False
    
    print(f"\n✅ Transformed {len(transformed_dataframes)} files")
    total_rows = sum(len(df) for df in transformed_dataframes)
    print(f"   Total rows: {total_rows}")
    
    # STEP 3: Create backup and load template
    print("\n" + "="*80)
    print("STEP 3: Preparing template")
    print("="*80)
    
    backup_path = create_backup(TEMPLATE_PATH)
    
    # STEP 4: Paste data into template
    print("\n" + "="*80)
    print("STEP 4: Pasting data into template")
    print("="*80)
    
    wb = paste_to_template(transformed_dataframes, TEMPLATE_PATH, TARGET_SHEET)
    
    # STEP 5: Concatenate formulas
    print("\n" + "="*80)
    print("STEP 5: Concatenating formulas")
    print("="*80)
    
    # Find the target sheet name (might have changed)
    target_ws_name = None
    for name in [TARGET_SHEET, 'AllStores', 'Allstores', 'Raw Data']:
        if name in wb.sheetnames:
            target_ws_name = name
            break
    
    if target_ws_name and len(FORMULA_COLUMNS) > 0:
        ws = wb[target_ws_name]
        last_row = ws.max_row
        first_new_row = last_row - total_rows + 1
        
        wb = concatenate_formulas(
            wb,
            target_ws_name,
            first_new_row,
            last_row,
            FORMULA_COLUMNS
        )
    else:
        print("   ⚠️  Skipping formula concatenation (configure FORMULA_COLUMNS)")
    
    # STEP 6: Refresh pivot tables
    print("\n" + "="*80)
    print("STEP 6: Refreshing pivot tables")
    print("="*80)
    
    wb = refresh_pivot_tables(wb)
    
    # STEP 7: Update dates
    print("\n" + "="*80)
    print("STEP 7: Updating dates in sheets")
    print("="*80)
    
    wb = update_dates(wb, TARGET_DATE, DATE_CONFIGS)
    
    # STEP 8: Save final template
    print("\n" + "="*80)
    print("STEP 8: Saving final template")
    print("="*80)
    
    save_template(wb, TEMPLATE_PATH)
    
    # FINAL SUMMARY
    print("\n" + "="*80)
    print("✅✅✅ AUTOMATION COMPLETE! ✅✅✅")
    print("="*80)
    print(f"\n📊 Summary:")
    print(f"   - Files processed: {len(raw_files)}")
    print(f"   - Total rows added: {total_rows}")
    print(f"   - Template updated: {TEMPLATE_PATH}")
    print(f"   - Backup saved: {backup_path}")
    print(f"   - Date set to: {TARGET_DATE.strftime('%Y-%m-%d')}")
    
    print(f"\n📋 Next steps:")
    print(f"   1. Open the template in Excel")
    print(f"   2. Click 'Refresh All' to update pivot tables")
    print(f"   3. Verify data looks correct")
    
    print("\n" + "="*80)
    
    return True


if __name__ == "__main__":
    success = main()
    
    if success:
        print("\n🎉 SUCCESS! Your Drive-Thru template is ready!")
    else:
        print("\n❌ Automation failed - check error messages above")



