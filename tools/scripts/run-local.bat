@echo off
REM OpenGov Office - Server Management Script
REM Usage: run-local.bat [start|stop|restart|status]
REM Double-click to start servers automatically

setlocal

set ACTION=%1

REM Default to start if no argument provided
if "%ACTION%"=="" set ACTION=start

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

REM Kill any processes on our ports first (clean start)
echo [0/2] Cleaning up ports 3000 and 3001...
powershell -NoProfile -Command "$proc = Get-NetTCPConnection -LocalPort 3001 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty OwningProcess; if ($proc) { Write-Host '  - Killing process on port 3001'; Stop-Process -Id $proc -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 1 }" >nul 2>&1
powershell -NoProfile -Command "$proc = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty OwningProcess; if ($proc) { Write-Host '  - Killing process on port 3000'; Stop-Process -Id $proc -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 1 }" >nul 2>&1
echo   - Ports cleared
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

REM Kill ONLY processes on our ports (3000, 3001)
echo Stopping Node.js processes on ports 3000 and 3001...

REM Kill port 3001 (backend)
powershell -NoProfile -Command "$proc = Get-NetTCPConnection -LocalPort 3001 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty OwningProcess; if ($proc) { Write-Host '  - Killed backend server on port 3001 (PID:' $proc ')'; Stop-Process -Id $proc -Force -ErrorAction SilentlyContinue } else { Write-Host '  - No process found on port 3001' }"

REM Kill port 3000 (add-in dev server)
powershell -NoProfile -Command "$proc = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty OwningProcess; if ($proc) { Write-Host '  - Killed add-in dev server on port 3000 (PID:' $proc ')'; Stop-Process -Id $proc -Force -ErrorAction SilentlyContinue } else { Write-Host '  - No process found on port 3000' }"

REM Also try to stop Excel debugging
cd /d %~dp0..\..
call npm stop >nul 2>&1

echo.
echo Servers stopped.
echo.
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
goto :end

:end
endlocal

