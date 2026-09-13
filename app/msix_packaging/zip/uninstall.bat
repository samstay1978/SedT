@echo off
REM OfficeSensitiveEncryptor - uninstaller
REM Removes: desktop shortcut, context menus, install folder
setlocal
set "DEST=%ProgramFiles%\OfficeSensitiveEncryptor"

net session >nul 2>&1
if errorlevel 1 (
    echo Administrator rights are required. Requesting elevation...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

if not exist "%DEST%\OfficeSensitiveEncryptor.exe" (
    echo Not installed at "%DEST%".
    pause
    exit /b 1
)

echo Closing the app if it is running...
taskkill /f /im OfficeSensitiveEncryptor.exe >nul 2>&1

echo Removing the desktop shortcut and context menus...
"%DEST%\OfficeSensitiveEncryptor.exe" --uninstall

echo Removing the installation folder...
cd /d "%TEMP%"
(goto) 2>nul & rd /s /q "%DEST%"
