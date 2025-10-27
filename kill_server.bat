@echo off
chcp 65001 >nul
echo ========================================
echo Stopping Flask Server
echo ========================================

for /f "tokens=5" %%a in ('netstat -ano ^| findstr :5000 ^| findstr LISTENING') do (
    echo Stopping process %%a on port 5000...
    taskkill /PID %%a /F /T
)

echo.
echo Done! Server stopped.
echo ========================================
