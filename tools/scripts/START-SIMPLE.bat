@echo off
echo ========================================
echo   OpenGov Office Sync - SIMPLE MODE
echo ========================================
echo.
echo Killing any running servers...
taskkill /F /IM node.exe >nul 2>&1
timeout /t 2 /nobreak >nul

echo.
echo Starting Backend Server (Port 3001)...
start "Backend" cmd /k "cd /d "%~dp0" && npm run server"

timeout /t 4 /nobreak >nul

echo Starting Frontend Server (Port 3000)...
start "Frontend" cmd /k "cd /d "%~dp0" && npm start"

echo.
echo ========================================
echo   Servers Starting!
echo ========================================
echo.
echo IMPORTANT - Wait 30 seconds, then:
echo   1. Open browser: http://localhost:3000
echo   2. Hard refresh (Ctrl+Shift+R)
echo   3. Open Excel ^& reload add-in
echo.
echo New Features:
echo   - Simple HTML table (no Luckysheet!)
echo   - Editable cells - click to edit
echo   - 2 second sync delay
echo   - Rate limiting active
echo.
pause

