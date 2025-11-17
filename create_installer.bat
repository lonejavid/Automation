@echo off
REM Create a simple installer batch script
REM This will create a setup folder with the app and installer

echo ========================================
echo Creating Installer Package
echo ========================================

REM Create installer directory
if exist "Installer" rmdir /s /q "Installer"
mkdir "Installer"

REM Copy the executable
copy "dist\KFC_DriveThru_Automation.exe" "Installer\KFC_DriveThru_Automation.exe"

REM Create a simple installer batch file
echo @echo off > "Installer\Install.bat"
echo echo Installing KFC Drive-Thru Automation... >> "Installer\Install.bat"
echo echo. >> "Installer\Install.bat"
echo set INSTALL_DIR=%%ProgramFiles%%\KFC_Automation >> "Installer\Install.bat"
echo if not exist "%%INSTALL_DIR%%" mkdir "%%INSTALL_DIR%%" >> "Installer\Install.bat"
echo copy "KFC_DriveThru_Automation.exe" "%%INSTALL_DIR%%\" >> "Installer\Install.bat"
echo echo. >> "Installer\Install.bat"
echo echo Creating desktop shortcut... >> "Installer\Install.bat"
echo powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%%USERPROFILE%%\Desktop\KFC Drive-Thru Automation.lnk'); $Shortcut.TargetPath = '%%INSTALL_DIR%%\KFC_DriveThru_Automation.exe'; $Shortcut.Save()" >> "Installer\Install.bat"
echo echo. >> "Installer\Install.bat"
echo echo Installation complete! >> "Installer\Install.bat"
echo echo The app has been installed to: %%INSTALL_DIR%% >> "Installer\Install.bat"
echo echo A shortcut has been created on your desktop. >> "Installer\Install.bat"
echo pause >> "Installer\Install.bat"

REM Create README
echo KFC Drive-Thru Automation - Installation Instructions > "Installer\README.txt"
echo. >> "Installer\README.txt"
echo INSTALLATION: >> "Installer\README.txt"
echo 1. Run Install.bat as Administrator >> "Installer\README.txt"
echo 2. Or simply double-click KFC_DriveThru_Automation.exe to run directly >> "Installer\README.txt"
echo. >> "Installer\README.txt"
echo USAGE: >> "Installer\README.txt"
echo - Double-click the desktop shortcut >> "Installer\README.txt"
echo - Click "Start Automation" button >> "Installer\README.txt"
echo - Wait for completion >> "Installer\README.txt"
echo. >> "Installer\README.txt"
echo REQUIREMENTS: >> "Installer\README.txt"
echo - Windows 10 or later >> "Installer\README.txt"
echo - Microsoft Excel installed >> "Installer\README.txt"
echo - Internet connection >> "Installer\README.txt"

echo.
echo ========================================
echo Installer package created in "Installer" folder
echo ========================================
echo.
echo Files created:
echo   - Installer\KFC_DriveThru_Automation.exe (Main application)
echo   - Installer\Install.bat (Installation script)
echo   - Installer\README.txt (Instructions)
echo.
echo To share with users:
echo   1. Zip the "Installer" folder
echo   2. Share the zip file
echo   3. Users extract and run Install.bat
echo.
pause

