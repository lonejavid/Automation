"""
Run VBA Macro on Downloaded Excel File
Executes the DT macro (Ctrl+D) on the downloaded Raw Car Data report
"""

import os
import glob
import sys
import subprocess
import time
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


def wait_for_file_download(file_path, max_wait=60):
    """
    Wait for a file to be fully downloaded (not still being written)
    
    Args:
        file_path: Path to the file
        max_wait: Maximum seconds to wait
    
    Returns:
        True if file is ready, False if timeout
    """
    start_time = time.time()
    last_size = -1
    
    while time.time() - start_time < max_wait:
        if os.path.exists(file_path):
            try:
                current_size = os.path.getsize(file_path)
                # If file size hasn't changed in 2 seconds, it's done downloading
                if current_size == last_size and current_size > 0:
                    time.sleep(1)  # One more second to be sure
                    return True
                last_size = current_size
            except (OSError, IOError):
                # File might be locked, wait a bit
                pass
        time.sleep(0.5)
    
    return os.path.exists(file_path) and os.path.getsize(file_path) > 0


def unblock_downloaded_file(file_path):
    """
    Unblock a downloaded file on Windows (removes Zone.Identifier)
    This prevents Excel from opening in Protected View
    
    Args:
        file_path: Path to the file
    """
    if sys.platform == "win32":
        try:
            zone_file = file_path + ":Zone.Identifier"
            if os.path.exists(zone_file):
                os.remove(zone_file)
                print(f"   🔓 Unblocked Windows file: {os.path.basename(file_path)}")
        except Exception as e:
            print(f"   ⚠️  Could not unblock file: {e}")


def find_latest_downloaded_file(downloads_folder, wait_for_download=True):
    """
    Find the most recently downloaded Excel file
    
    Args:
        downloads_folder: Path to downloads directory
        wait_for_download: If True, wait for file to finish downloading
    
    Returns:
        Path to latest Excel file or None
    """
    excel_files = glob.glob(os.path.join(downloads_folder, "*.xlsx"))
    
    if not excel_files:
        return None
    
    # Sort by modification time, most recent first
    latest_file = max(excel_files, key=os.path.getmtime)
    
    if wait_for_download:
        print(f"   ⏳ Waiting for file to finish downloading...")
        if wait_for_file_download(latest_file, max_wait=60):
            print(f"   ✅ File download complete")
        else:
            print(f"   ⚠️  File may still be downloading, proceeding anyway...")
    
    # Unblock file on Windows
    unblock_downloaded_file(latest_file)
    
    return latest_file


def ensure_excel_edit_mode(wb):
    """
    Ensure Excel workbook is in edit mode (not Protected View)
    
    Args:
        wb: xlwings workbook object
    """
    try:
        if wb.api.ProtectStructure or wb.api.ProtectWindows:
            print("   ⚠️  Workbook is protected, attempting to enable editing...")
            # Try to enable editing
            wb.api.Application.DisplayAlerts = False
            wb.api.Application.EnableEvents = True
    except:
        pass


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
        
        # Disable alerts to prevent popups
        app.api.DisplayAlerts = False
        app.api.ScreenUpdating = True
        
        # Open workbook
        try:
            wb = app.books.open(excel_file_path, read_only=False)
        except Exception as e:
            print(f"   ⚠️  Could not open file with xlwings: {e}")
            print("   🔄 Falling back to Python logic...")
            app.quit()
            return run_dt_macro_python_logic(excel_file_path)
        
        # Ensure edit mode
        ensure_excel_edit_mode(wb)
        
        print("   Running DT macro (Ctrl+D equivalent)...")
        
        # Try multiple methods to run the macro
        macro_success = False
        
        # Method 1: Try to run from add-in or workbook
        try:
            app.api.Run("DT")
            print("   ✅ Macro executed successfully via DT")
            macro_success = True
        except Exception as macro_error:
            # Method 2: Try with full path
            try:
                app.api.Run(f"'{os.path.basename(excel_file_path)}'!DT")
                print("   ✅ Macro executed successfully (with workbook name)")
                macro_success = True
            except:
                # Method 3: Try from Personal.xlsb or add-in
                try:
                    app.api.Run("Personal.xlsb!DT")
                    print("   ✅ Macro executed successfully from Personal.xlsb")
                    macro_success = True
                except:
                    pass
        
        if macro_success:
            try:
                wb.save()
                print("   💾 File saved successfully")
            except Exception as save_error:
                print(f"   ⚠️  Could not save automatically: {save_error}")
                print("   📝 File is open in Excel - please save manually if needed")
            
            # Keep workbook and Excel window open for the user
            print("   📂 Excel workbook left open for review")
            
            print("\n" + "="*80)
            print("✅ FILE CONVERTED SUCCESSFULLY!")
            print("="*80)
            print(f"   📁 File: {os.path.basename(excel_file_path)}")
            print(f"   📍 Location: {excel_file_path}")
            print("="*80)
            return True
        else:
            print(f"   ⚠️  Macro 'DT' not found in add-ins or workbook")
            print(f"   🔄 Replicating macro logic in Python...")
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
        print(f"   ❌ Error running macro with xlwings: {e}")
        import traceback
        traceback.print_exc()
        try:
            wb.close()
        except:
            pass
        try:
            app.quit()
        except:
            pass
        print("   🔄 Falling back to Python logic...")
        return run_dt_macro_python_logic(excel_file_path)


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


def process_downloaded_file(downloads_folder=None, wait_for_download=True):
    """
    Find latest downloaded file and run DT macro on it automatically
    
    This function is called automatically after each download to ensure
    the DT macro runs on every downloaded Excel file.
    
    Args:
        downloads_folder: Path to downloads folder (defaults to data/downloads)
        wait_for_download: If True, wait for file to finish downloading
    
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
    
    # Ensure downloads folder exists
    if not os.path.exists(downloads_folder):
        print(f"   ❌ Downloads folder does not exist: {downloads_folder}")
        return False
    
    # Find latest file (and wait for download to complete)
    latest_file = find_latest_downloaded_file(downloads_folder, wait_for_download=wait_for_download)
    
    if not latest_file:
        print("   ❌ No Excel files found in downloads folder")
        print(f"   📁 Folder: {downloads_folder}")
        return False
    
    print(f"   📄 Found file: {os.path.basename(latest_file)}")
    print(f"   📅 Modified: {datetime.fromtimestamp(os.path.getmtime(latest_file)).strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"   📍 Full path: {latest_file}")
    
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
        print("   ⚠️  The file was downloaded but the macro could not be executed.")
        print("   💡 You may need to run the macro manually in Excel.")
        return False


if __name__ == "__main__":
    process_downloaded_file()

