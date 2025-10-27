@echo off
echo 개발서버 시작 중...
cd /d "C:\Users\kiko9\Desktop\itglobal"
set PORT=5001
set FLASK_ENV=development
set FLASK_DEBUG=1
echo PORT is %PORT%
python -m dashboard.app
pause