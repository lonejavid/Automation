"""
Run VBA Macro on Downloaded Excel File
Executes the DT macro (Ctrl+D) on the downloaded Raw Car Data report
"""

import os
import glob
import sys
import subprocess
from pathlib import Path
from datetime import datetime

def open_excel_file(excel_file_path):
    """Open the Excel workbook in the system's default spreadsheet application."""
    try:
        if sys.platform.startswith("darwin"):  # macOS
            subprocess.Popen(["open", excel_file_path])
        elif os.name == "nt":  # Windows
            os.startfile(excel_file_path)  # type: ignore[attr-defined]
        else:  # Linux / other
            subprocess.Popen(["xdg-open", excel_file_path])
        print(f"   📂 Excel opened: {excel_file_path}")
    except Exception as e:
        print(f"   ⚠️  Unable to open Excel automatically: {e}")


def find_latest_downloaded_file(downloads_folder):
    """
    Find the most recently downloaded Excel file
    
    Args:
        downloads_folder: Path to downloads directory
    
    Returns:
        Path to latest Excel file or None
    """
    excel_files = glob.glob(os.path.join(downloads_folder, "*.xlsx"))
    
    if not excel_files:
        return None
    
    # Sort by modification time, most recent first
    latest_file = max(excel_files, key=os.path.getmtime)
    return latest_file


def run_dt_macro(excel_file_path):
    """
    Run the DT macro on the Excel file using xlwings
    
    Args:
        excel_file_path: Path to Excel file
    
    Returns:
        True if successful, False otherwise
    """
    try:
        import xlwings as xw
        
        print(f"\n📝 Opening Excel file: {os.path.basename(excel_file_path)}")
        
        # Open Excel file and keep the window visible for the user
        app = xw.App(visible=True)
        wb = app.books.open(excel_file_path)
        
        print("   Running DT macro (Ctrl+D equivalent)...")
        
        # Run the macro
        try:
            app.api.Run("DT")
            print("   ✅ Macro executed successfully")
            wb.save()
            # Keep workbook and Excel window open for the user
            print("   📂 Excel workbook left open for review")
            
            print("\n" + "="*80)
            print("✅ FILE CONVERTED SUCCESSFULLY!")
            print("="*80)
            print(f"   📁 File: {os.path.basename(excel_file_path)}")
            print(f"   📍 Location: {excel_file_path}")
            print("="*80)
            return True
        except Exception as macro_error:
            print(f"   ⚠️  Macro not found, replicating logic in Python...")
            try:
                wb.close()
            except:
                pass
            app.quit()
            return run_dt_macro_python_logic(excel_file_path)
        
    except ImportError:
        print("   ⚠️  xlwings not installed. Replicating macro logic in Python...")
        return run_dt_macro_python_logic(excel_file_path)
    except Exception as e:
        print(f"   ❌ Error running macro: {e}")
        try:
            wb.close()
        except:
            pass
        try:
            app.quit()
        except:
            pass
        return False


def run_dt_macro_python_logic(excel_file_path, wb=None, ws=None):
    """
    Replicate the DT macro logic using openpyxl
    This is a fallback if xlwings is not available or macro doesn't exist
    """
    should_close = False
    try:
        from openpyxl import load_workbook
        
        if wb is None:
            print(f"\n📝 Opening Excel file: {os.path.basename(excel_file_path)}")
            wb = load_workbook(excel_file_path)
            ws = wb.active
            should_close = True
        
        print("   Replicating DT macro logic...")
        
        # Step 1: Unmerge all cells
        merged_ranges = list(ws.merged_cells.ranges)
        for merged_range in merged_ranges:
            ws.unmerge_cells(str(merged_range))
        
        # Step 2: Add "Store Name" header at B7
        ws['B7'] = "Store Name"
        
        # Step 3: Copy store name from B4 to column B starting at B8
        store_name = ws['B4'].value
        if store_name:
            max_row = ws.max_row
            for row in range(8, min(max_row + 1, 198)):  # B8:B197
                ws[f'B{row}'] = store_name
        
        # Step 4: Delete rows 1-6
        ws.delete_rows(1, 6)
        
        # After deleting rows 1-6, what was B8 is now B2
        # Fill B2:B191 with store name if empty
        if store_name:
            max_row = ws.max_row
            for row in range(2, min(max_row + 1, 192)):
                if ws[f'B{row}'].value is None:
                    ws[f'B{row}'] = store_name
        
        # Step 5: Delete columns in reverse order to maintain indices
        if ws.max_column >= 11:
            ws.delete_cols(11, 1)  # Column K
        if ws.max_column >= 10:
            ws.delete_cols(10, 1)  # Column J
        if ws.max_column >= 9:
            ws.delete_cols(9, 1)   # Column I
        if ws.max_column >= 7:
            ws.delete_cols(7, 1)   # Column G
        if ws.max_column >= 6:
            ws.delete_cols(5, 2)   # Columns E-F
        if ws.max_column >= 4:
            ws.delete_cols(4, 1)   # Column D
        
        # Step 6: Set column widths
        try:
            ws.column_dimensions['I'].width = 1.57
        except:
            pass
        try:
            ws.column_dimensions['J'].width = 4
        except:
            pass
        ws.column_dimensions['B'].width = 33.29
        
        # Step 7: Fill blank cells in column A with value above
        max_row = ws.max_row
        for row in range(2, max_row + 1):
            cell_a = ws[f'A{row}']
            if cell_a.value is None or cell_a.value == '':
                cell_above = ws[f'A{row-1}']
                if cell_above.value:
                    cell_a.value = cell_above.value
        
        # Step 8: Delete columns F-G and J-K (final cleanup)
        if ws.max_column >= 7:
            ws.delete_cols(6, 2)
        if ws.max_column >= 11:
            ws.delete_cols(10, 2)
        
        # Save changes
        wb.save(excel_file_path)
        if should_close and wb:
            wb.close()
        
        print("   ✅ Macro logic replicated and applied")
        print("\n" + "="*80)
        print("✅ FILE CONVERTED SUCCESSFULLY!")
        print("="*80)
        print(f"   📁 File: {os.path.basename(excel_file_path)}")
        print(f"   📍 Location: {excel_file_path}")
        print("="*80)
        
        # Attempt to open the workbook in Excel for the user
        open_excel_file(excel_file_path)
        return True
        
    except Exception as e:
        print(f"   ❌ Error replicating macro: {e}")
        import traceback
        traceback.print_exc()
        if should_close and wb:
            try:
                wb.close()
            except:
                pass
        return False


def process_downloaded_file(downloads_folder=None):
    """
    Find latest downloaded file and run DT macro on it
    
    Args:
        downloads_folder: Path to downloads folder (defaults to data/downloads)
    
    Returns:
        True if successful, False otherwise
    """
    if downloads_folder is None:
        BASE_DIR = Path(__file__).resolve().parents[2]
        downloads_folder = BASE_DIR / "data" / "downloads"
    
    downloads_folder = str(downloads_folder)
    
    print("\n" + "="*80)
    print("🔄 RUNNING DT MACRO ON DOWNLOADED FILE")
    print("="*80)
    
    # Find latest file
    latest_file = find_latest_downloaded_file(downloads_folder)
    
    if not latest_file:
        print("   ❌ No Excel files found in downloads folder")
        return False
    
    print(f"   📄 Found file: {os.path.basename(latest_file)}")
    print(f"   📅 Modified: {datetime.fromtimestamp(os.path.getmtime(latest_file)).strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Run macro
    success = run_dt_macro(latest_file)
    
    if success:
        print("\n" + "="*80)
        print("✅ FILE CONVERTED SUCCESSFULLY!")
        print("="*80)
        print(f"   📁 File: {os.path.basename(latest_file)}")
        print(f"   📍 Location: {latest_file}")
        print("="*80)
        return True
    else:
        print("\n" + "="*80)
        print("❌ FILE CONVERSION FAILED")
        print("="*80)
        return False


if __name__ == "__main__":
    process_downloaded_file()

