@echo off
chcp 65001 >nul
echo 🚀 IT Global 개발 환경 시작
echo ==============================
cd /d "%~dp0"
python dev-start.py
pause