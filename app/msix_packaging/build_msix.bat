@echo off

setlocal

REM ============================================================

REM  MSIX packer - pack the "app" folder into an MSIX

REM

REM  Modes:

REM    build_msix.bat              -> pack and sign (side-loading)

REM    build_msix.bat nosign       -> pack only, NO signing

REM       (Microsoft Store submissions: the Store re-signs your

REM        package automatically, so no cert is needed)

REM    build_msix.bat [mode] auto  -> non-interactive (no pause);

REM       used by build.bat as the last step of the pipeline

REM

REM  The "app" folder must contain:

REM    OfficeSensitiveEncryptor.exe, AppxManifest.xml, Assets\, word\

REM    (build.bat produces all of these automatically)

REM

REM  Output: dist_msix\OfficeSensitiveEncryptor.msix

REM  Prereq: Windows SDK installed (provides makeappx.exe / signtool.exe)

REM          AppxManifest.xml Publisher must match the certificate Subject.

REM ============================================================

cd /d "%~dp0"



set APP_DIR=%~dp0app

set OUT_DIR=%~dp0dist_msix

set OUT=%OUT_DIR%\OfficeSensitiveEncryptor.msix

set CERT=%~dp0mycert.pfx

set CERT_PWD=YOUR_PFX_PASSWORD



REM auto mode (called from build.bat) -> suppress pause prompts

set PAUSE_CMD=pause

if /i "%~2"=="auto" set PAUSE_CMD=



if not exist "%APP_DIR%\AppxManifest.xml" goto :err_manifest

if not exist "%APP_DIR%\OfficeSensitiveEncryptor.exe" goto :err_exe

if not exist "%APP_DIR%\word" goto :err_word

if /i "%~1"=="nosign" goto :locate_tools

if not exist "%CERT%" goto :err_cert



:locate_tools

REM --- locate makeappx / signtool (PATH first, then Windows Kits) ---

set MAKEAPPX=makeappx

where makeappx >nul 2>nul || (

    for /f "delims=" %%i in ('dir /s /b "%ProgramFiles(x86)%\Windows Kits\10\bin\makeappx.exe" 2^>nul') do set MAKEAPPX=%%i

)

if "%MAKEAPPX%"=="makeappx" (

    for /f "delims=" %%i in ('dir /s /b "%ProgramFiles%\Windows Kits\10\bin\makeappx.exe" 2^>nul') do set MAKEAPPX=%%i

)

if "%MAKEAPPX%"=="makeappx" goto :err_makeappx

set SIGNTOOL=signtool

where signtool >nul 2>nul || (

    for /f "delims=" %%i in ('dir /s /b "%ProgramFiles(x86)%\Windows Kits\10\bin\signtool.exe" 2^>nul') do set SIGNTOOL=%%i

)



mkdir "%OUT_DIR%" 2>nul



echo [1/2] Packing directory into MSIX...

"%MAKEAPPX%" pack /d "%APP_DIR%" /p "%OUT%" /l

if errorlevel 1 goto :err_pack



if /i "%~1"=="nosign" goto :nosign_done

echo [2/2] Signing package...

"%SIGNTOOL%" sign /f "%CERT%" /p "%CERT_PWD%" /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 "%OUT%"

if errorlevel 1 goto :err_sign

goto :finish



:nosign_done

echo [2/2] Skipping signing (Store submission mode).

echo        The Microsoft Store will re-sign this package automatically.

goto :finish



:err_manifest

echo.

echo ERROR: missing %APP_DIR%\AppxManifest.xml

%PAUSE_CMD%

exit /b 1



:err_exe

echo.

echo ERROR: copy OfficeSensitiveEncryptor.exe into the app\ folder first.

%PAUSE_CMD%

exit /b 1



:err_word

echo.

echo ERROR: missing %APP_DIR%\word (demo vocabularies and sample documents).

echo build.bat copies it automatically; or copy the word\ folder manually.

%PAUSE_CMD%

exit /b 1



:err_cert

echo.

echo ERROR: certificate not found: %CERT%

echo Create a test one with:

echo   $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=Sam Li" -CertStoreLocation Cert:\CurrentUser\My

echo   $pwd  = ConvertTo-SecureString -String "YOUR_PFX_PASSWORD" -Force -AsPlainText

echo   Export-PfxCertificate -Cert $cert -FilePath mycert.pfx -Password $pwd

%PAUSE_CMD%

exit /b 1



:err_makeappx

echo.

echo ERROR: makeappx.exe not found.

echo.

echo Install the Windows SDK, then retry:

echo   Option 1 (easiest): Microsoft Store, search "Windows SDK",

echo        click Install (free official tool, about 1 GB).

echo   Option 2: download from

echo        https://developer.microsoft.com/windows/downloads/windows-sdk/

echo   IMPORTANT: during setup, make sure the "Windows SDK for

echo        Desktop C++ x86 Apps" component is selected - that is

echo        where makeappx.exe lives.

echo.

echo Verify after install (should print a path):

echo   dir /s /b "%ProgramFiles(x86)%\Windows Kits\10\bin\makeappx.exe"

echo.

echo Or skip the search: edit this script and set MAKEAPPX to the

echo full path, e.g.:

echo   set MAKEAPPX=C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\makeappx.exe

%PAUSE_CMD%

exit /b 1



:err_pack

echo.

echo ERROR: makeappx failed.

%PAUSE_CMD%

exit /b 1



:err_sign

echo.

echo ERROR: signing failed.

%PAUSE_CMD%

exit /b 1



:finish

echo.

echo ============================================================

echo  Done: %OUT%

echo  - nosign mode: submit this .msix to Partner Center (Store signs it)

echo  - signed mode: install via  Add-AppxPackage -Path "%OUT%"

echo  Install (admin PowerShell):

echo    Add-AppxPackage -Path "%OUT%"

echo  If the cert is self-signed, trust it first or enable

echo  Developer Mode: Settings ^> Privacy and security ^> For developers.

echo.

echo  If install fails with 0x800B010A (untrusted publisher cert):

echo    Export-Certificate -Cert (Get-ChildItem Cert:\CurrentUser\My ^| Where-Object {$_.Subject -eq "CN=Sam Li"}) -FilePath mycert.cer

echo    Import-Certificate -FilePath .\mycert.cer -CertStoreLocation Cert:\LocalMachine\Root

echo    (run as Administrator, then install again)

echo ============================================================

%PAUSE_CMD%

