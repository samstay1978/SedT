@echo off
REM ============================================================
REM  Office Sensitive Word Encryptor - One-Click Build Pipeline (Windows)
REM
REM  Pipeline: source code -> exe -> manual/license txt -> MSIX
REM    [1] Check Python
REM    [2] Install build/runtime dependencies
REM    [3] Run logic self-test (test_encryptor.py)
REM    [4] Build exe with PyInstaller -> msix_packaging\app\
REM    [5] Run the exe once (--gen-files) so it creates the two
REM        txt files (user manual + license) next to itself
REM    [6] Pack the "app" folder into an MSIX
REM
REM  Inputs : office_sensitive_encryptor.py
REM           test_encryptor.py
REM           office_sensitive_encryptor.spec
REM  Outputs: msix_packaging\app\OfficeSensitiveEncryptor.exe  (+ 2 txt)
REM           msix_packaging\dist_msix\OfficeSensitiveEncryptor.msix
REM
REM  Modes:
REM    build.bat           -> pack MSIX without signing
REM                           (Microsoft Store submission mode)
REM    build.bat sign      -> pack and sign the MSIX
REM                           (side-loading; needs mycert.pfx in
REM                            msix_packaging, edit CERT_PWD there)
REM
REM  Requirements: Windows 10/11 with Python 3.9+ installed
REM ============================================================
cd /d "%~dp0"

set APP_DIR=msix_packaging\app
set EXE=%APP_DIR%\OfficeSensitiveEncryptor.exe
set MSIX=msix_packaging\dist_msix\OfficeSensitiveEncryptor.msix

echo [1/6] Checking Python environment...
where python >nul 2>nul
if errorlevel 1 goto :err_python

echo [2/6] Installing build and runtime dependencies...
python -m pip install --upgrade pyinstaller python-docx openpyxl python-pptx cryptography
if errorlevel 1 goto :err_deps

echo [3/6] Running logic self-test (10 regression tests)...
python test_encryptor.py
if errorlevel 1 goto :err_test

echo [4/6] Building with PyInstaller (about 1-3 minutes)...
mkdir "%APP_DIR%" 2>nul
del /q "%APP_DIR%\*.txt" 2>nul
python -m PyInstaller office_sensitive_encryptor.spec --noconfirm --distpath "%APP_DIR%"
if errorlevel 1 goto :err_build
if not exist "%EXE%" goto :err_exe

echo [5/6] Running the exe once to generate manual and license files...
start "" /wait "%EXE%" --gen-files
if not exist "%APP_DIR%\*.txt" goto :err_txt

echo [6/6] Packing MSIX from the app folder...
if /i "%~1"=="sign" goto :pack_sign
call msix_packaging\build_msix.bat nosign auto
if errorlevel 1 goto :err_msix
goto :finish

:pack_sign
call msix_packaging\build_msix.bat sign auto
if errorlevel 1 goto :err_msix
goto :finish

:err_python
echo.
echo ERROR: Python not found.
echo Please install Python 3.9+ from the official site
echo   https://www.python.org/downloads/
echo and make sure "Add python.exe to PATH" is checked during installation.
pause
exit /b 1

:err_deps
echo.
echo ERROR: Failed to install dependencies.
echo Please check your network and retry.
pause
exit /b 1

:err_test
echo.
echo ERROR: Logic self-test failed. Build aborted.
pause
exit /b 1

:err_build
echo.
echo ERROR: PyInstaller build failed.
echo Please take a screenshot of the error above and report it.
pause
exit /b 1

:err_exe
echo.
echo ERROR: exe not found at %EXE%
pause
exit /b 1

:err_txt
echo.
echo ERROR: manual/license txt files were not generated in %APP_DIR%
pause
exit /b 1

:err_msix
echo.
echo ERROR: MSIX packing failed. See messages above.
pause
exit /b 1

:finish
echo.
echo ============================================================
echo  Pipeline complete.
echo    exe : %cd%\%EXE%
echo    txt : manual + license files next to the exe
echo    msix: %cd%\%MSIX%
echo.
echo  - Store submission: submit the .msix to Partner Center
echo    (the Store re-signs it - no certificate needed)
echo  - Side-loading:    run build_msix.bat manually with mycert.pfx
echo ============================================================
pause
