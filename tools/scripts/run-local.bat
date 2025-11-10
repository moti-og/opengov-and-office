@echo off
REM OpenGov Office - Server Management Script
REM Usage: run-local.bat [start|stop|restart|status]

setlocal

set ACTION=%1

if "%ACTION%"=="" (
    echo.
    echo ========================================
    echo   OpenGov Office - Server Manager
    echo ========================================
    echo.
    echo Choose an action:
    echo   1. Start servers
    echo   2. Stop servers
    echo   3. Restart servers
    echo   4. Check status
    echo   5. Exit
    echo.
    set /p CHOICE="Enter choice (1-5): "
    
    if "%CHOICE%"=="1" set ACTION=start
    if "%CHOICE%"=="2" set ACTION=stop
    if "%CHOICE%"=="3" set ACTION=restart
    if "%CHOICE%"=="4" set ACTION=status
    if "%CHOICE%"=="5" exit /b 0
    
    if "%ACTION%"=="" (
        echo Invalid choice!
        pause
        exit /b 1
    )
)

if /I "%ACTION%"=="start" goto :start
if /I "%ACTION%"=="stop" goto :stop
if /I "%ACTION%"=="restart" goto :restart
if /I "%ACTION%"=="status" goto :status

echo Unknown action: %ACTION%
echo Usage: servers.bat [start^|stop^|restart^|status]
exit /b 1

:start
echo.
echo ========================================
echo   Starting OpenGov Office Servers
echo ========================================
echo.

REM Start backend server (MongoDB + Express)
echo [1/2] Starting backend server (port 3001)...
start "OpenGov Backend" cmd /k "cd /d %~dp0..\..\server && node index.js"
timeout /t 2 /nobreak >nul

REM Start Excel add-in (webpack dev server)
echo [2/2] Starting Excel add-in (port 3000)...
echo.
echo This will:
echo   - Build the add-in with webpack
echo   - Start dev server with HTTPS
echo   - Sideload manifest into Excel
echo   - Open Excel automatically
echo.
start "OpenGov Add-in" cmd /k "cd /d %~dp0..\.. && npm start"

echo.
echo ========================================
echo   Servers Started!
echo ========================================
echo.
echo Backend:  http://localhost:3001
echo Add-in:   https://localhost:3000
echo.
echo Excel will open automatically with the add-in loaded.
echo.
goto :end

:stop
echo.
echo ========================================
echo   Stopping OpenGov Office Servers
echo ========================================
echo.

REM Stop Node processes
echo Stopping Node.js processes...
taskkill /FI "WINDOWTITLE eq OpenGov*" /F >nul 2>&1

REM Also try to stop Excel debugging
cd /d %~dp0..\..
call npm stop >nul 2>&1

echo.
echo Servers stopped.
echo.
pause
goto :end

:restart
echo.
echo Restarting servers...
call :stop
timeout /t 2 /nobreak >nul
call :start
goto :end

:status
echo.
echo ========================================
echo   Server Status
echo ========================================
echo.

REM Check if backend is running
netstat -ano | findstr :3001 >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Backend Server [port 3001]:  RUNNING
) else (
    echo Backend Server [port 3001]:  STOPPED
)

REM Check if add-in dev server is running
netstat -ano | findstr :3000 >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Add-in Dev Server [port 3000]:  RUNNING
) else (
    echo Add-in Dev Server [port 3000]:  STOPPED
)

echo.

REM Check Node processes
tasklist /FI "IMAGENAME eq node.exe" 2>NUL | find /I /N "node.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo Node.js processes:
    tasklist /FI "IMAGENAME eq node.exe" | findstr node.exe
) else (
    echo No Node.js processes running
)

echo.
pause
goto :end

:end
endlocal

