@echo off
echo 개발서버 시작 중 (포트 5002)...
cd /d "C:\Users\kiko9\Desktop\itglobal"
set PORT=5002
set FLASK_ENV=development
set FLASK_DEBUG=1
echo 포트: %PORT%
python -m dashboard.app