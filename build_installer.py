"""
Build Windows Desktop App with Installer
Creates executable and installer for Windows users
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent
BUILD_DIR = PROJECT_ROOT / "build"
DIST_DIR = PROJECT_ROOT / "dist"
SPEC_FILE = PROJECT_ROOT / "kfc_automation.spec"

def print_header(text):
    """Print formatted header"""
    print("\n" + "="*80)
    print(f"  {text}")
    print("="*80)

def check_dependencies():
    """Check if required tools are installed"""
    print_header("Checking Dependencies")
    
    try:
        import PyInstaller
        print("✅ PyInstaller installed")
    except ImportError:
        print("❌ PyInstaller not found")
        print("   Install with: pip install pyinstaller")
        return False
    
    return True

def clean_build_dirs():
    """Clean previous build directories"""
    print_header("Cleaning Build Directories")
    
    dirs_to_clean = [BUILD_DIR, DIST_DIR]
    
    for dir_path in dirs_to_clean:
        if dir_path.exists():
            try:
                shutil.rmtree(dir_path)
                print(f"   🧹 Cleaned: {dir_path.name}")
            except Exception as e:
                print(f"   ⚠️  Could not clean {dir_path.name}: {e}")
    
    BUILD_DIR.mkdir(exist_ok=True)
    DIST_DIR.mkdir(exist_ok=True)
    print("   ✅ Build directories ready")

def create_spec_file():
    """Create PyInstaller spec file"""
    print_header("Creating PyInstaller Spec File")
    
    spec_content = """# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['desktop_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('data/templates', 'data/templates'),
        ('src', 'src'),
        ('scripts', 'scripts'),
    ],
    hiddenimports=[
        'selenium',
        'xlwings',
        'pandas',
        'openpyxl',
        'streamlit',
        'webdriver_manager',
        'automation',
        'automation.hmecloud',
        'automation.complete_automation',
        'automation.transform_data',
        'automation.template_operations',
        'automation.run_macro',
        'tkinter',
        'tkinter.ttk',
        'tkinter.scrolledtext',
        'tkinter.messagebox',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='KFC_DriveThru_Automation',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
"""
    
    SPEC_FILE.write_text(spec_content)
    print(f"   ✅ Created spec file: {SPEC_FILE.name}")
    return SPEC_FILE

def build_executable(spec_path):
    """Build the executable using PyInstaller"""
    print_header("Building Executable")
    print("   This may take 5-10 minutes...")
    print("   Please wait...")
    
    # Try different ways to run PyInstaller
    pyinstaller_commands = [
        [sys.executable, '-m', 'PyInstaller', '--clean', '--noconfirm', str(spec_path)],
        ['pyinstaller', '--clean', '--noconfirm', str(spec_path)],
    ]
    
    for cmd in pyinstaller_commands:
        try:
            result = subprocess.run(
                cmd,
                cwd=PROJECT_ROOT,
                check=True,
                capture_output=True,
                text=True
            )
            print("   ✅ Build successful!")
            return True
        except subprocess.CalledProcessError as e:
            if cmd == pyinstaller_commands[-1]:  # Last attempt
                print(f"   ❌ Build failed!")
                if e.stderr:
                    print(f"   Error: {e.stderr}")
                if e.stdout:
                    print(f"   Output: {e.stdout}")
                return False
            continue  # Try next command
        except FileNotFoundError:
            if cmd == pyinstaller_commands[-1]:  # Last attempt
                print("   ❌ PyInstaller not found")
                print("   Make sure PyInstaller is installed: pip install pyinstaller")
                return False
            continue  # Try next command
    
    return False

def create_installer_script():
    """Create NSIS installer script"""
    print_header("Creating Installer Script")
    
    nsi_content = """; KFC Drive-Thru Automation Installer
; NSIS Installer Script

!define APP_NAME "KFC Drive-Thru Automation"
!define APP_VERSION "1.0.0"
!define APP_PUBLISHER "KFC Guyana"
!define APP_EXE "KFC_DriveThru_Automation.exe"
!define APP_DIR "KFC_Automation"

Name "${APP_NAME}"
OutFile "KFC_DriveThru_Automation_Setup.exe"
InstallDir "$PROGRAMFILES\\${APP_DIR}"
RequestExecutionLevel admin

; UI
!include "MUI2.nsh"

!define MUI_ICON "${NSISDIR}\\Contrib\\Graphics\\Icons\\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\\Contrib\\Graphics\\Icons\\modern-uninstall.ico"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE "English"

Section "Install"
    SetOutPath "$INSTDIR"
    
    ; Copy all files from dist
    File /r "dist\\KFC_DriveThru_Automation\\*"
    
    ; Create uninstaller
    WriteUninstaller "$INSTDIR\\Uninstall.exe"
    
    ; Registry entries
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APP_NAME}" \\
        "DisplayName" "${APP_NAME}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APP_NAME}" \\
        "UninstallString" "$INSTDIR\\Uninstall.exe"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APP_NAME}" \\
        "Publisher" "${APP_PUBLISHER}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APP_NAME}" \\
        "DisplayVersion" "${APP_VERSION}"
    
    ; Create desktop shortcut
    CreateShortcut "$DESKTOP\\${APP_NAME}.lnk" "$INSTDIR\\${APP_EXE}"
    
    ; Create start menu shortcut
    CreateDirectory "$SMPROGRAMS\\${APP_NAME}"
    CreateShortcut "$SMPROGRAMS\\${APP_NAME}\\${APP_NAME}.lnk" "$INSTDIR\\${APP_EXE}"
    CreateShortcut "$SMPROGRAMS\\${APP_NAME}\\Uninstall.lnk" "$INSTDIR\\Uninstall.exe"
    
    ; Create data directories
    CreateDirectory "$INSTDIR\\data\\downloads"
    CreateDirectory "$INSTDIR\\data\\templates"
    CreateDirectory "$INSTDIR\\logs"
SectionEnd

Section "Uninstall"
    ; Remove files
    RMDir /r "$INSTDIR"
    
    ; Remove shortcuts
    Delete "$DESKTOP\\${APP_NAME}.lnk"
    RMDir /r "$SMPROGRAMS\\${APP_NAME}"
    
    ; Remove registry entries
    DeleteRegKey HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APP_NAME}"
SectionEnd
"""
    
    nsi_path = PROJECT_ROOT / "installer.nsi"
    nsi_path.write_text(nsi_content)
    print(f"   ✅ Created installer script: {nsi_path.name}")
    return nsi_path

def main():
    """Main build process"""
    print_header("KFC DRIVE-THRU AUTOMATION - WINDOWS APP BUILDER")
    
    if not check_dependencies():
        print("\n❌ Please install missing dependencies and try again.")
        return False
    
    clean_build_dirs()
    spec_path = create_spec_file()
    
    if not build_executable(spec_path):
        print("\n❌ Build failed. Check errors above.")
        return False
    
    create_installer_script()
    
    print_header("BUILD COMPLETE!")
    
    exe_path = DIST_DIR / "KFC_DriveThru_Automation.exe"
    exe_folder = DIST_DIR / "KFC_DriveThru_Automation"
    
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024*1024)
        print(f"\n✅ Executable created: {exe_path}")
        print(f"   Size: {size_mb:.1f} MB")
    elif exe_folder.exists():
        print(f"\n✅ Executable folder created: {exe_folder}")
        exe_file = exe_folder / "KFC_DriveThru_Automation.exe"
        if exe_file.exists():
            size_mb = exe_file.stat().st_size / (1024*1024)
            print(f"   Executable: {exe_file}")
            print(f"   Size: {size_mb:.1f} MB")
    
    print("\n📋 Next Steps:")
    print("   1. Test the executable first")
    print("   2. Create installer (if NSIS is installed):")
    print("      - Right-click installer.nsi → Compile NSIS Script")
    print("   3. Share with users:")
    print("      - Share KFC_DriveThru_Automation_Setup.exe (if created)")
    print("      - OR share the .exe file directly")
    
    return True

if __name__ == "__main__":
    success = main()
    input("\nPress Enter to exit...")
    sys.exit(0 if success else 1)

