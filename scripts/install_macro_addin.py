#!/usr/bin/env python3
"""Install the DT macro as an Excel add-in on macOS."""

from __future__ import annotations

import shutil
import sys
import tempfile
import os
from pathlib import Path

ADDIN_NAME = "DTMacro.xlam"
MACRO_SOURCE = Path(__file__).resolve().parent / ADDIN_NAME

# Determine platform-specific add-in folder and instructions
if sys.platform == "darwin":
    TARGET_FOLDER = Path.home() / "Library/Group Containers/UBF8T346G9.Office/User Content/Add-ins"
    ENABLE_INSTRUCTIONS = (
        "After installation, open Excel and enable the add-in via Tools → Excel Add-ins…\n"
        "Check the box next to \"DTMacro\" (or click Browse… if it is not listed)."
    )
elif os.name == "nt":
    TARGET_FOLDER = Path.home() / "AppData/Roaming/Microsoft/AddIns"
    ENABLE_INSTRUCTIONS = (
        "After installation, open Excel and enable the add-in via File → Options → Add-ins → Go…\n"
        "Click Browse…, pick DTMacro.xlam, then make sure \"DTMacro\" is checked."
    )
else:
    TARGET_FOLDER = Path.cwd() / "excel_addins"
    ENABLE_INSTRUCTIONS = (
        "After installation, copy the add-in file to your platform-specific Excel add-in folder\n"
        "and enable it through the Excel Add-ins dialog."
    )

TARGET_PATH = TARGET_FOLDER / ADDIN_NAME

INFO_BANNER = "=" * 80

MACRO_README = f"""
{INFO_BANNER}
📦 DT Macro Add-in Installer
{INFO_BANNER}

This script creates a permanent Excel add-in containing the DT macro and places it in:
{TARGET_PATH}

{ENABLE_INSTRUCTIONS}
"""


def require_xlwings() -> None:
    try:
        import xlwings  # noqa: F401
    except ImportError:  # pragma: no cover - guidance only
        print("❌ xlwings is required to install the add-in.")
        print("   Install it with: pip install xlwings")
        sys.exit(1)


def read_macro_source() -> str:
    if not MACRO_SOURCE.exists():
        print(f"❌ Macro source file missing: {MACRO_SOURCE}")
        print("   Make sure scripts/DTMacro.xlam is part of the repository.")
        sys.exit(1)
    return MACRO_SOURCE.read_text(encoding="utf-8")


def create_addin_file(macro_code: str, destination: Path) -> None:
    """Create an .xlam add-in at *destination* using the provided macro code."""
    import xlwings as xw

    print("→ Launching Excel to compile add-in…")
    app = xw.App(visible=False)  # keep hidden during setup
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.add()
        try:
            vb_module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        except Exception:
            print("❌ Excel denied access to the VBA project.")
            print("   Enable 'Trust access to the VBA project object model' in Excel → Preferences")
            print("   (Security > Macro Security). Then run this installer again.")
            wb.close()
            raise SystemExit(1)

        vb_module.Name = "DTMacroModule"
        vb_module.CodeModule.AddFromString(macro_code)

        print(f"→ Saving add-in to {destination}")
        destination.parent.mkdir(parents=True, exist_ok=True)

        try:
            wb.api.SaveAs(str(destination), FileFormat=55)  # 55 => xlam add-in
        except Exception as save_error:
            print(f"❌ Could not save add-in: {save_error}")
            wb.close()
            raise SystemExit(1)
    finally:
        try:
            wb.close()
        except Exception:
            pass
        try:
            app.quit()
        except Exception:
            pass


def install_addin() -> None:
    print(MACRO_README)
    require_xlwings()

    macro_code = read_macro_source()

    if TARGET_PATH.exists():
        print(f"ℹ️  Existing add-in found at: {TARGET_PATH}")
        backup_dir = Path(tempfile.mkdtemp(prefix="dtmacro_backup_"))
        backup_path = backup_dir / ADDIN_NAME
        print(f"→ Backing up current add-in to: {backup_path}")
        shutil.copy2(TARGET_PATH, backup_path)

    temp_dir = Path(tempfile.mkdtemp(prefix="dtmacro_builder_"))
    temp_addin = temp_dir / ADDIN_NAME

    create_addin_file(macro_code, temp_addin)

    print("→ Copying add-in into Excel add-in folder…")
    TARGET_FOLDER.mkdir(parents=True, exist_ok=True)
    shutil.copy2(temp_addin, TARGET_PATH)

    print(MACRO_README)
    print("✅ Installation complete!")
    print("Next steps:")
    if sys.platform == "darwin":
        print("  1. Open Excel")
        print("  2. Tools → Excel Add-ins…")
        print("  3. Check 'DTMacro' (or Browse… and select it)")
        print("  4. Click OK, then test Tools → Macro → DT")
    elif os.name == "nt":
        print("  1. Open Excel")
        print("  2. File → Options → Add-ins → Go…")
        print("  3. Browse to DTMacro.xlam if needed, check 'DTMacro'")
        print("  4. Click OK, then test Developer → Macros → DT")
    else:
        print("  1. Copy DTMacro.xlam to your Excel add-in folder")
        print("  2. Enable it via the Excel Add-ins dialog")
    print("\nIf Excel was already running, restart it so the add-in loads.")


def is_installed() -> bool:
    return TARGET_PATH.exists()


def main() -> None:
    if "--check" in sys.argv:
        if is_installed():
            print(f"✅ DTMacro add-in found at {TARGET_PATH}")
        else:
            print("❌ DTMacro add-in not found. Run this script without --check to install it.")
        return

    install_addin()


if __name__ == "__main__":
    main()
