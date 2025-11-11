@echo off

echo ========================================
echo  OpenGov Office Sync - Uninstaller
echo ========================================
echo.
echo This will uninstall the Excel add-in (any version)...
echo.

REM Close Excel if it's running
echo [1/4] Closing Excel if running...
taskkill /F /IM EXCEL.EXE 2>&1
if %ERRORLEVEL% EQU 0 (
    echo    ^> Excel was closed
) else (
    echo    ^> Excel was not running
)
timeout /t 2 /nobreak >nul

REM Remove OpenGov Excel add-ins by known registry keys
echo.
echo [2/3] Removing OpenGov EXCEL add-ins...
echo.

REM Remove local version
echo Checking for local version...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "opengov-excel-addin-local" /f >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo    ^> Removed local version ✓
) else (
    echo    ^> Local version not found
)

REM Remove production version
echo Checking for production version...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "opengov-excel-addin-prod" /f >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo    ^> Removed production version ✓
) else (
    echo    ^> Production version not found
)

REM Also check Office 15.0
reg delete "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" /v "opengov-excel-addin-local" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" /v "opengov-excel-addin-prod" /f >nul 2>&1

echo Done

echo.
echo [3/3] Cleaning up files and cache...

REM Clean up temp files
if exist "%TEMP%\opengov-excel-addin" (
    rd /s /q "%TEMP%\opengov-excel-addin" 2>nul
    echo    ^> Temp files removed
)

REM Clear Office cache
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
    rd /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul
    mkdir "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" 2>nul
    echo    ^> Office cache cleared
)

echo.
echo ========================================
echo  Done!
echo ========================================
echo.
echo All OpenGov EXCEL add-ins have been removed.
echo Word add-ins were left untouched.
echo.
echo Open Excel to verify it's gone:
echo   Insert ^> My Add-ins ^> Developer Add-ins
echo.
echo If it's still there, restart Excel or reboot.
echo.

pause

