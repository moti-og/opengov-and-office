@echo off
echo Stopping all Node processes...
taskkill /F /IM node.exe >nul 2>&1

echo.
echo Clearing MongoDB...
mongosh opengov-office --eval "db.documents.deleteMany({})" --quiet

echo.
echo Starting servers...
start "Backend Server" cmd /k "npm run server"
timeout /t 3 /nobreak >nul
start "Frontend Server" cmd /k "npm start"

echo.
echo Done! Servers are starting in new windows.
echo Press any key to close this window...
pause >nul

