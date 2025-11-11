@echo off

echo ========================================
echo  OpenGov Office Sync - Excel Add-in
echo ========================================
echo.
echo This will install the Excel add-in...
echo.

REM Close Excel if it's running
echo Closing Excel if running...
taskkill /F /IM EXCEL.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

REM Remove local version if present
echo Checking for local version...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "opengov-excel-addin-local" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" /v "opengov-excel-addin-local" /f >nul 2>&1
echo    ^> Local version removed (if present)

REM Download manifest to temp location
echo Downloading manifest...
set TEMP_DIR=%TEMP%\opengov-excel-addin
if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"
powershell -Command "Invoke-WebRequest -Uri 'https://opengov-and-office.onrender.com/manifest.xml' -OutFile '%TEMP_DIR%\manifest.xml'"

REM Register manifest via Developer registry key
echo Registering add-in...
set MANIFEST_PATH=%TEMP_DIR%\manifest.xml
set ADDIN_ID=opengov-excel-addin-prod

REM Office 2016/2019/2021/365 (16.0)
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ADDIN_ID%" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

REM Office 2013 (15.0) - fallback
reg add "HKCU\Software\Microsoft\Office\15.0\WEF\Developer" /v "%ADDIN_ID%" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul 2>&1

echo.
echo ========================================
echo  Installation Complete!
echo ========================================
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
echo   4. Click "OpenGov Office Sync"
echo.
echo The add-in panel will appear on the right side.
echo After this first activation, it will remember your choice.
echo.
echo You can now open: https://opengov-and-office.onrender.com
echo to see real-time sync between Excel and the web!
echo.

echo Press any key to close this window...
pause >nul
exit /b 0

