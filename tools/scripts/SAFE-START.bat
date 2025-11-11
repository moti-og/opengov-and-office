@echo off
echo ========================================
echo SAFE START - Killing all processes
echo ========================================
echo.

echo Killing all Node.js processes...
taskkill /F /IM node.exe >nul 2>&1

echo.
echo Waiting for processes to stop...
timeout /t 3 /nobreak >nul

echo.
echo ========================================
echo Starting servers with safe settings
echo ========================================
echo.

echo Starting Backend Server (Port 3001)...
start "Backend Server" cmd /k "cd /d "%~dp0" && npm run server"

timeout /t 5 /nobreak >nul

echo Starting Frontend Server (Port 3000)...
start "Frontend Server" cmd /k "cd /d "%~dp0" && npm start"

echo.
echo ========================================
echo Servers are starting!
echo ========================================
echo.
echo IMPORTANT:
echo 1. Wait for both servers to fully start (30 seconds)
echo 2. Hard refresh web browser (Ctrl+Shift+R)
echo 3. Close and reopen Excel add-in task pane
echo.
echo Rate limiting is now active:
echo - Minimum 2 seconds between syncs
echo - 3 second debounce on changes
echo - Maximum 5 reconnect attempts
echo.
pause

