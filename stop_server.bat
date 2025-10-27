@echo off
echo ========================================
echo Flask Server Safe Shutdown Script
echo ========================================
echo.

echo Checking for Flask server on port 5000...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :5000 ^| findstr LISTENING') do (
    echo Found process: %%a
    echo Attempting to stop gracefully...
    taskkill /PID %%a /T
    timeout /t 2 /nobreak >nul

    rem Check if still running
    tasklist /FI "PID eq %%a" 2>nul | find "%%a" >nul
    if errorlevel 1 (
        echo Process stopped successfully!
    ) else (
        echo Process still running, forcing shutdown...
        taskkill /PID %%a /F /T
    )
)

echo.
echo Checking for any remaining Python processes...
tasklist | findstr python.exe >nul
if errorlevel 1 (
    echo No Python processes found.
) else (
    echo Found Python processes:
    tasklist | findstr python.exe
    echo.
    set /p answer="Do you want to stop all Python processes? (y/n): "
    if /i "%answer%"=="y" (
        taskkill /IM python.exe /F /T
        echo All Python processes stopped.
    )
)

echo.
echo ========================================
echo Done!
echo ========================================
pause
