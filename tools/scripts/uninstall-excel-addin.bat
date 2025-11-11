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

REM Find and remove EXCEL-ONLY OpenGov add-ins
echo.
echo [2/3] Finding and removing OpenGov EXCEL add-ins...
echo.

REM Check Office 16.0
for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" 2^>nul ^| findstr /i /r "REG_SZ.*opengov"') do (
    set "MANIFEST_PATH=%%b"
    setlocal enabledelayedexpansion
    if exist "!MANIFEST_PATH!" (
        findstr /i /c:"Host Name=\""Workbook\"" "!MANIFEST_PATH!" >nul 2>&1
        if !ERRORLEVEL! EQU 0 (
            REM This is an Excel add-in, safe to remove
            for /f "tokens=1" %%c in ('reg query "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" 2^>nul ^| findstr /i "opengov" ^| findstr /i "!MANIFEST_PATH!"') do (
                echo Found Excel add-in: %%c
                reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%%c" /f >nul 2>&1
                echo    ^> Removed ✓
            )
        ) else (
            echo Skipping non-Excel add-in: !MANIFEST_PATH!
        )
    )
    endlocal
)

REM Check Office 15.0
for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" 2^>nul ^| findstr /i /r "REG_SZ.*opengov"') do (
    set "MANIFEST_PATH=%%b"
    setlocal enabledelayedexpansion
    if exist "!MANIFEST_PATH!" (
        findstr /i /c:"Host Name=\""Workbook\"" "!MANIFEST_PATH!" >nul 2>&1
        if !ERRORLEVEL! EQU 0 (
            for /f "tokens=1" %%c in ('reg query "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" 2^>nul ^| findstr /i "opengov" ^| findstr /i "!MANIFEST_PATH!"') do (
                echo Found Excel add-in: %%c
                reg delete "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" /v "%%c" /f >nul 2>&1
                echo    ^> Removed ✓
            )
        )
    )
    endlocal
)

echo Done checking registry

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

