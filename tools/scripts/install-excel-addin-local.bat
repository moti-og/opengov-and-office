@echo off

echo ========================================
echo  OpenGov Office Sync - LOCAL TESTING
echo ========================================
echo.
echo This will install the Excel add-in for local development...
echo.

REM Close Excel if it's running
echo Closing Excel if running...
taskkill /F /IM EXCEL.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

REM Remove production version if present
echo Checking for production version...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "opengov-excel-addin-prod" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" /v "opengov-excel-addin-prod" /f >nul 2>&1
echo    ^> Production version removed (if present)

REM Download manifest from local server
echo Downloading manifest from localhost:3001...
set TEMP_DIR=%TEMP%\opengov-excel-addin
if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"

REM Use PowerShell to download the manifest
powershell -NoProfile -Command "try { Invoke-WebRequest -Uri 'http://localhost:3001/manifest-local.xml' -OutFile '%TEMP_DIR%\manifest.xml' -UseBasicParsing; exit 0 } catch { Write-Host 'ERROR: Could not download manifest from server'; Write-Host 'Make sure the server is running: npm run server'; exit 1 }"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Failed to download manifest!
    echo.
    echo Make sure the backend server is running:
    echo   npm run server
    echo.
    echo Then try running this installer again.
    pause
    exit /b 1
)
echo    ^> Manifest downloaded ✓

REM Register manifest via Developer registry key
echo Registering add-in (localhost:3001)...
set MANIFEST_PATH=%TEMP_DIR%\manifest.xml
set ADDIN_ID=opengov-excel-addin-local

REM Office 2016/2019/2021/365 (16.0)
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

REM Office 2013 (15.0) - fallback
reg add "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" /v "%ADDIN_ID%" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

echo.
echo ========================================
echo  Installation Complete!
echo ========================================
echo.
echo Verifying installation...
reg query "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "opengov-excel-addin-local" >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo    ^> Add-in registered successfully ✓
) else (
    echo    ^> Warning: Add-in registration not found!
    echo    ^> This might be normal if you have Office 2013 instead of 2016+
)
echo.
echo IMPORTANT: Make sure your local servers are running!
echo   - Backend: http://localhost:3001
echo   - Frontend: http://localhost:3000
echo.
echo Run this command in another terminal:
echo   npm run server
echo.
echo Opening Excel...
echo.

REM Open Excel
start excel.exe
timeout /t 2 /nobreak >nul

echo.
echo ========================================
echo  Next Steps: Activate the Add-in
echo ========================================
echo.
echo Excel is now open.
echo.
echo TO ACTIVATE THE ADD-IN (first time only):
echo   1. In Excel, click the "Insert" tab
echo   2. Click "Get Add-ins" or "My Add-ins"
echo   3. Click "Developer Add-ins" at the top
echo   4. Click "OpenGov Office Sync (Local)"
echo.
echo The add-in panel will appear on the right side.
echo After this first activation, it will remember your choice.
echo.
echo You can now open: http://localhost:3000
echo to see real-time sync between Excel and the web!
echo.

echo Press any key to close this window...
pause >nul
exit /b 0

