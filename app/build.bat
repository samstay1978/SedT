@echo off
REM ============================================================
REM  Office Sensitive Word Encryptor - One-Click Build Pipeline (Windows)
REM
REM  Pipeline: source code -> exe -> word assets -> manual/license txt -> MSIX
REM    [1] Check Python
REM    [2] Install build/runtime dependencies
REM    [3] Run logic self-test (test_encryptor.py)
REM    [4] Build exe with PyInstaller -> msix_packaging\app\
REM    [5] Copy word\ (demo vocabularies + sample documents) next to exe
REM    [6] Run the exe once (--gen-files) so it creates the two
REM        txt files (user manual + license) next to itself
REM    [7] Pack the "app" folder into an MSIX
REM
REM  Inputs : office_sensitive_encryptor.py
REM           test_encryptor.py
REM           office_sensitive_encryptor.spec
REM           word\  (4 domain vocabularies + 4 sample .docx files)
REM  Outputs: msix_packaging\app\OfficeSensitiveEncryptor.exe (+ word\ + 2 txt)
REM           msix_packaging\dist_msix\OfficeSensitiveEncryptor.msix
REM
REM  Modes:
REM    build.bat              -> pack MSIX without signing
REM                             (Microsoft Store submission mode)
REM    build.bat sign         -> pack and sign the MSIX
REM                             (side-loading; needs mycert.pfx in
REM                              msix_packaging)
REM    build.bat PASSWORD <p> -> pass the certificate password for
REM                             signing (forwarded as CERT_PWD), e.g.
REM                             build.bat sign PASSWORD mysecret
REM
REM  Requirements: Windows 10/11 with Python 3.9+ installed
REM ============================================================
cd /d "%~dp0"

set APP_DIR=msix_packaging\app
set EXE=%APP_DIR%\OfficeSensitiveEncryptor.exe
set MSIX=msix_packaging\dist_msix\OfficeSensitiveEncryptor.msix

echo [1/8] Checking Python environment...
where python >nul 2>nul
if errorlevel 1 goto :err_python

echo [2/8] Installing build and runtime dependencies...
python -m pip install --upgrade pyinstaller python-docx openpyxl python-pptx cryptography
if errorlevel 1 goto :err_deps

echo [3/8] Running logic self-test (14 regression tests)...
python test_encryptor.py
if errorlevel 1 goto :err_test

echo [4/8] Building with PyInstaller (about 1-3 minutes)...
mkdir "%APP_DIR%" 2>nul
del /q "%APP_DIR%\*.txt" 2>nul
python -m PyInstaller office_sensitive_encryptor.spec --noconfirm --distpath "%APP_DIR%"
if errorlevel 1 goto :err_build
if not exist "%EXE%" goto :err_exe

echo [5/8] Copying demo vocabularies and sample documents (word dir)...
xcopy "%~dp0word" "%APP_DIR%\word\" /e /i /y >nul
if errorlevel 1 goto :err_word
if not exist "%APP_DIR%\word\vocab_legal.json" goto :err_word

echo [6/8] Running the exe once to generate manual and license files...
start "" /wait "%EXE%" --gen-files
if not exist "%APP_DIR%\*.txt" goto :err_txt

echo [7/8] Packing MSIX from the app folder...
set MSIX_ARGS=nosign auto
:parse_pwd
if "%~1"=="" goto :msix_go
if /i "%~1"=="sign" set MSIX_ARGS=sign auto
if /i "%~1"=="PASSWORD" (
    set MSIX_ARGS=%MSIX_ARGS% PASSWORD "%~2"
    shift
)
shift
goto :parse_pwd
:msix_go
call msix_packaging\build_msix.bat %MSIX_ARGS%
if errorlevel 1 goto :err_msix

echo [8/8] Building ZIP distribution (installs to Program Files)...
if not exist "%~dp0dist_zip" mkdir "%~dp0dist_zip"
set "ZIP_STAGE=%~dp0dist_zip\stage"
if exist "%ZIP_STAGE%" rd /s /q "%ZIP_STAGE%"
mkdir "%ZIP_STAGE%"
copy /y "%APP_DIR%\OfficeSensitiveEncryptor.exe" "%ZIP_STAGE%\" >nul
if errorlevel 1 goto :err_zip
xcopy /e /i /y "%APP_DIR%\word" "%ZIP_STAGE%\word" >nul
if errorlevel 1 goto :err_zip
copy /y "%~dp0msix_packaging\zip\install.bat" "%ZIP_STAGE%\" >nul
if errorlevel 1 goto :err_zip
copy /y "%~dp0msix_packaging\zip\uninstall.bat" "%ZIP_STAGE%\" >nul
for /f "delims=" %%v in ('powershell -NoProfile -Command "$m=[xml](Get-Content -Raw '%~dp0msix_packaging\app\AppxManifest.xml'); $m.Package.Identity.Version"') do set "ZIPVER=%%v"
if "%ZIPVER%"=="" set ZIPVER=2.4.0
powershell -NoProfile -Command "Compress-Archive -Path '%ZIP_STAGE%\*' -DestinationPath '%~dp0dist_zip\OfficeSensitiveEncryptor_%ZIPVER%.zip' -Force"
if errorlevel 1 goto :err_zip
rd /s /q "%ZIP_STAGE%"
echo      ZIP created: dist_zip\OfficeSensitiveEncryptor_%ZIPVER%.zip
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

:err_word
echo.
echo ERROR: word directory with demo vocabularies not found next to build.bat.
echo Expected: %~dp0word\vocab_legal.json (and 3 more vocabularies + 4 sample docx)
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

:err_zip
echo.
echo ERROR: ZIP distribution creation failed.
pause
exit /b 1

:finish
echo.
echo ============================================================
echo  Pipeline complete.
echo    exe : %cd%\%EXE%
echo    word: demo vocabularies + sample documents next to the exe
echo    txt : manual + license files next to the exe
echo    msix: %cd%\%MSIX%
echo    zip : %cd%\dist_zip\OfficeSensitiveEncryptor_%ZIPVER%.zip
echo.
echo  - Store submission: submit the .msix to Partner Center
echo    (the Store re-signs it - no certificate needed)
echo  - Side-loading:    build.bat sign PASSWORD ^<pfx-password^>
echo                     (needs mycert.pfx in msix_packaging)
echo  - Program Files:   unzip dist_zip\OfficeSensitiveEncryptor_*.zip,
echo                     run install.bat as admin (installs to
echo                     C:\Program Files\OfficeSensitiveEncryptor;
echo                     uninstall.bat in that folder removes it)
echo ============================================================
pause
