@echo off 
echo Installing KFC Drive-Thru Automation... 
echo. 
set INSTALL_DIR=%ProgramFiles%\KFC_Automation 
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%" 
copy "KFC_DriveThru_Automation.exe" "%INSTALL_DIR%\" 
echo. 
echo Creating desktop shortcut... 
powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%USERPROFILE%\Desktop\KFC Drive-Thru Automation.lnk'); $Shortcut.TargetPath = '%INSTALL_DIR%\KFC_DriveThru_Automation.exe'; $Shortcut.Save()" 
echo. 
echo Installation complete! 
echo The app has been installed to: %INSTALL_DIR% 
echo A shortcut has been created on your desktop. 
pause 
