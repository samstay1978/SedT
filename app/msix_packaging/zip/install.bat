@echo off
REM OfficeSensitiveEncryptor - ZIP installer
REM Default install location: %ProgramFiles%\OfficeSensitiveEncryptor
REM (MSIX cannot choose its install folder; this installer can.)
setlocal
set "DEST=%ProgramFiles%\OfficeSensitiveEncryptor"

net session >nul 2>&1
if errorlevel 1 (
    echo Administrator rights are required. Requesting elevation...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo Installing to "%DEST%" ...
if not exist "%DEST%" mkdir "%DEST%"
copy /y "%~dp0OfficeSensitiveEncryptor.exe" "%DEST%\" >nul
if errorlevel 1 goto :copy_fail
if exist "%DEST%\word" rd /s /q "%DEST%\word"
xcopy /e /i /y "%~dp0word" "%DEST%\word" >nul
copy /y "%~dp0uninstall.bat" "%DEST%\" >nul

echo First launch (creates the desktop shortcut and context menus)...
explorer.exe "%DEST%\OfficeSensitiveEncryptor.exe"

echo.
echo Done. Installed to: %DEST%
echo The user manual (README.txt) and license are released there
echo automatically on first launch.
echo To uninstall: run uninstall.bat in that folder.
pause
endlocal
exit /b 0

:copy_fail
echo.
echo ERROR: copying files failed.
echo If the app is running, close it first and retry.
pause
exit /b 1
