"""
Run VBA Macro on Downloaded Excel File
Executes the DT macro (Ctrl+D) on the downloaded Raw Car Data report
"""

import os
import glob
import sys
import subprocess
import shutil
import tempfile
from pathlib import Path
from datetime import datetime

ADDIN_MACRO_CANDIDATES = [
    "DT",
    "DTMacro.xlam!DT",
    "PERSONAL.XLSB!DT",
]

DT_MACRO_CODE = """Sub DT()
'
' DT Macro
'
' Keyboard Shortcut: Ctrl+d
'
    Cells.Select
    Selection.UnMerge
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "Store Name"
    Range("B4").Select
    Selection.Copy
    Range("B8").Select
    ActiveSheet.Paste
    Rows("1:6").Select
    Range("A6").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B197")
    Range("B2:B197").Select
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").ColumnWidth = 1.57
    Columns("I:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").ColumnWidth = 4
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").ColumnWidth = 19.86
    Columns("B:B").ColumnWidth = 33.29
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A197").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A2:A197").Select
    Range("A197").Activate
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("F:G").Delete Shift:=xlToLeft
    Columns("J:K").Delete Shift:=xlToLeft
End Sub"""


def unblock_downloaded_file(excel_file_path):
    """Remove the 'Mark of the Web' so Excel opens directly in edit mode (Windows)."""
    if os.name != "nt":
        return

    try:
        # Remove alternate data stream that triggers Protected View
        safe_path = excel_file_path.replace('"', '`"')
        cmd = [
            "powershell",
            "-NoProfile",
            "-Command",
            f'If (Test-Path -Path "{safe_path}") {{ Unblock-File -Path "{safe_path}" -ErrorAction SilentlyContinue }}'
        ]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
    except Exception as e:
        print(f"   ⚠️  Unable to unblock file automatically: {e}")

    try:
        os.chmod(excel_file_path, 0o666)
    except Exception:
        pass


def ensure_excel_edit_mode(app):
    """Set Excel automation security to allow macros and enable editing if opened in Protected View."""
    try:
        app.display_alerts = False
    except Exception:
        pass

    try:
        # msoAutomationSecurityLow = 1 allows macros to run
        app.api.AutomationSecurity = 1
    except Exception:
        pass

    try:
        pv_windows = app.api.ProtectedViewWindows
        if pv_windows.Count > 0:
            print("   ⚠️  Excel opened in Protected View. Enabling editing...")
            wb_com = pv_windows(1).Open()
            return wb_com
    except Exception:
        pass

    return None


def ensure_dt_macro_present(wb):
    """Ensure the DT macro exists in the workbook before running."""
    try:
        wb.macro("DT")
        return True
    except Exception:
        pass

    try:
        vb_project = wb.api.VBProject
    except Exception as e:
        print("   ⚠️  Excel security prevented access to the VBA project.")
        print("       Enable 'Trust access to the VBA project object model' in Excel Options to run the VBA macro.")
        return False

    try:
        module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        module.CodeModule.AddFromString(DT_MACRO_CODE)
        print("   ℹ️  DT macro injected into workbook for this session.")
        return True
    except Exception as e:
        print(f"   ⚠️  Unable to inject DT macro automatically: {e}")
        return False


def run_macro_via_addin(app, wb) -> bool:
    """Attempt to execute the DT macro via a loaded add-in (preferred path)."""
    try:
        wb.activate()
    except Exception:
        pass

    for macro_name in ADDIN_MACRO_CANDIDATES:
        try:
            app.api.Run(macro_name)
            print(f"   ✅ Macro executed successfully via {macro_name}")
            return True
        except Exception:
            continue

    return False


def run_macro_from_workbook(wb) -> bool:
    """Execute the DT macro stored inside the active workbook."""
    try:
        wb.activate()
    except Exception:
        pass

    try:
        dt_macro = wb.macro("DT")
    except Exception:
        return False

    dt_macro()
    print("   ✅ Macro executed successfully (workbook module)")
    return True


def save_workbook_gracefully(wb, excel_file_path: str, app):
    """Try to save the workbook without interrupting the automation.

    Returns the workbook (possibly reopened) so the caller can keep Excel visible.
    """
    if wb is None:
        return None

    try:
        wb.save()
        return wb
    except Exception as save_error:
        message = str(save_error)
        if "OSERROR: -50" in message:
            print("   ⚠️  Excel could not save automatically (macOS OSERROR -50).")
        else:
            print(f"   ⚠️  Unable to save workbook automatically: {save_error}")

        print("   ℹ️  Attempting to save a copy and restore the workbook automatically...")
        temp_dir = Path(tempfile.mkdtemp(prefix="dtmacro_save_"))
        temp_copy = temp_dir / Path(excel_file_path).name

        try:
            wb.impl.save(path=str(temp_copy), password=None)
            wb.close(False)
            shutil.copy2(temp_copy, excel_file_path)
            print("   ✅ Saved changes by copying workbook back to original file.")
            reopened = app.books.open(excel_file_path, update_links=False, read_only=False)
            return reopened
        except Exception as copy_error:
            print(f"   ❌ Could not preserve changes automatically: {copy_error}")
            print("       Please save manually in Excel (⌘+S) before closing.")
            return wb

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
    excel_files = [
        path
        for path in glob.glob(os.path.join(downloads_folder, "*.xlsx"))
        if not Path(path).name.startswith("~$")
    ]
    
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
    app = None
    wb = None
    try:
        import xlwings as xw
        
        print(f"\n📝 Opening Excel file: {os.path.basename(excel_file_path)}")
        unblock_downloaded_file(excel_file_path)
        
        # Open Excel file and keep the window visible for the user
        app = xw.App(visible=True)
        app.screen_updating = True
        
        wb_com_from_protected_view = ensure_excel_edit_mode(app)
        if wb_com_from_protected_view is not None:
            wb = xw.Book(wb_com_from_protected_view)
        else:
            try:
                wb = app.books.open(excel_file_path, update_links=False, read_only=False)
            except Exception as open_error:
                print(f"   ⚠️  Excel refused to open the file directly ({open_error}). Retrying...")
                wb_com_from_protected_view = ensure_excel_edit_mode(app)
                if wb_com_from_protected_view is not None:
                    wb = xw.Book(wb_com_from_protected_view)
                else:
                    raise
            else:
                # After opening, check whether Excel still placed the workbook in Protected View
                pv_retry = ensure_excel_edit_mode(app)
                if pv_retry is not None:
                    wb = xw.Book(pv_retry)
        
        # Preferred path: use the permanent add-in if it is installed.
        if run_macro_via_addin(app, wb):
            wb = save_workbook_gracefully(wb, excel_file_path, app)
            print("   📂 Excel workbook left open for review")
            print("\n" + "=" * 80)
            print("✅ FILE CONVERTED SUCCESSFULLY!")
            print("=" * 80)
            print(f"   📁 File: {os.path.basename(excel_file_path)}")
            print(f"   📍 Location: {excel_file_path}")
            print("=" * 80)
            return True

        print("   ℹ️  DT add-in not detected. Injecting macro into this workbook…")
        if ensure_dt_macro_present(wb) and run_macro_from_workbook(wb):
            wb = save_workbook_gracefully(wb, excel_file_path, app)
            print("   📂 Excel workbook left open for review")
            print("\n" + "=" * 80)
            print("✅ FILE CONVERTED SUCCESSFULLY!")
            print("=" * 80)
            print(f"   📁 File: {os.path.basename(excel_file_path)}")
            print(f"   📍 Location: {excel_file_path}")
            print("=" * 80)
            return True

        print("   ⚠️  Could not run DT macro via Excel. Falling back to Python transformation.")
        try:
            if wb:
                wb.close()
        except Exception:
            pass
        try:
            if app:
                app.quit()
        except Exception:
            pass
        return run_dt_macro_python_logic(excel_file_path)
        
    except ImportError:
        print("   ⚠️  xlwings not installed. Replicating macro logic in Python...")
        return run_dt_macro_python_logic(excel_file_path)
    except Exception as e:
        print(f"   ❌ Error running macro: {e}")
        import traceback
        traceback.print_exc()
        try:
            if wb:
                wb.close()
        except Exception:
            pass
        try:
            if app:
                app.quit()
        except Exception:
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

