@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
title ITG Dashboard - 전체 서비스 시작

REM ============================================================
REM  ITG Dashboard 통합 시작 스크립트
REM  순서: 관리자 권한 → Docker Desktop → Redis 컨테이너 → ITGFlask → 헬스 체크
REM  이 순서를 지켜야 Flask 가 Redis fallback 모드로 빠지지 않음.
REM  2026-07-10 작성 — 사고 재발 방지용.
REM ============================================================

REM ---- 1. 관리자 권한 확인 (net start/stop 필요) ----
net session >nul 2>&1
if !errorlevel! neq 0 (
    echo.
    echo [알림] 관리자 권한이 필요합니다. 자동으로 관리자 창을 엽니다...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

cls
echo ============================================================
echo   ITG Dashboard - 전체 서비스 시작 ^(관리자 모드^)
echo ============================================================
echo.

REM ---- 2. Docker Desktop 실행 ----
echo [1/5] Docker Desktop 실행 여부 확인...
tasklist /FI "IMAGENAME eq Docker Desktop.exe" 2>nul | find /I "Docker Desktop.exe" >nul
if !errorlevel! neq 0 (
    echo   Docker Desktop 실행 중이 아님. 시작합니다...
    start "" "C:\Program Files\Docker\Docker\Docker Desktop.exe"
) else (
    echo   Docker Desktop 이미 실행 중.
)

REM ---- 3. Docker daemon 준비 대기 (최대 90초, 초기 부팅은 오래 걸림) ----
echo.
echo [2/5] Docker daemon 준비 대기 ^(최대 90초^)...
set /a wait=0
:docker_wait
docker info >nul 2>&1
if !errorlevel! equ 0 (
    echo   Docker daemon 준비 완료 ^(대기 !wait!초^).
    goto docker_ready
)
set /a wait+=3
if !wait! geq 90 (
    echo   [오류] Docker daemon 90초 대기 초과.
    echo         Docker Desktop 창을 직접 확인하시고 다시 실행하세요.
    goto :error
)
timeout /t 3 /nobreak >nul
goto docker_wait
:docker_ready

REM ---- 4. Redis 컨테이너 시작 ----
echo.
echo [3/5] Redis 컨테이너 시작...
docker start redis >nul 2>&1
if !errorlevel! equ 0 (
    echo   Redis 컨테이너 실행됨.
) else (
    echo   [경고] 기존 redis 컨테이너 없음. 새로 생성...
    docker run -d --name redis -p 6379:6379 --restart unless-stopped redis:alpine
    if !errorlevel! neq 0 (
        echo   [오류] Redis 컨테이너 생성 실패.
        goto :error
    )
)

REM ---- 5. Redis PING 응답 대기 (최대 30초) ----
echo.
echo [4/5] Redis PING 응답 대기...
set /a wait=0
:redis_wait
docker exec redis redis-cli ping 2>nul | findstr /I PONG >nul
if !errorlevel! equ 0 (
    echo   Redis PONG 응답 확인 ^(대기 !wait!초^).
    goto redis_ready
)
set /a wait+=2
if !wait! geq 30 (
    echo   [오류] Redis 30초 응답 없음.
    goto :error
)
timeout /t 2 /nobreak >nul
goto redis_wait
:redis_ready

REM ---- 6. ITGFlask 서비스 재시작 (fallback 모드 잔재 제거) ----
echo.
echo [5/5] ITGFlask 서비스 재시작...
sc query ITGFlask | findstr /I RUNNING >nul
if !errorlevel! equ 0 (
    echo   실행 중인 서비스 중지...
    net stop ITGFlask >nul 2>&1
    timeout /t 3 /nobreak >nul
)
echo   서비스 시작...
net start ITGFlask
if !errorlevel! neq 0 (
    echo   [오류] ITGFlask 서비스 시작 실패. NSSM 로그 확인 필요.
    echo         %~dp0dashboard\logs\service_stderr.log
    goto :error
)

REM ---- 7. 헬스 체크 (최대 60초) ----
echo.
echo ============================================================
echo   헬스 체크 대기 ^(최대 60초^)...
echo ============================================================
set /a wait=0
set "HEALTHFILE=%TEMP%\itg_health_%RANDOM%.txt"
:health_wait
timeout /t 3 /nobreak >nul
C:\Windows\System32\curl.exe -sS -o nul -w "%%{http_code}" http://localhost:5000/api/health > "!HEALTHFILE!" 2>nul
set /p code=<"!HEALTHFILE!"
if "!code!"=="200" (
    del "!HEALTHFILE!" 2>nul
    echo.
    echo ============================================================
    echo    모든 서비스 정상 동작 중
    echo ============================================================
    echo.
    echo   웹 대시보드: http://localhost:5000
    echo   상세 상태  : http://localhost:5000/api/health
    echo.
    goto :end
)
set /a wait+=3
if !wait! geq 60 (
    del "!HEALTHFILE!" 2>nul
    echo.
    echo   [경고] 헬스 체크 60초 대기 초과. 마지막 응답 상태 코드: !code!
    echo         전체 헬스 응답:
    C:\Windows\System32\curl.exe -s http://localhost:5000/api/health
    echo.
    goto :error
)
echo   대기 중... ^(!wait!초, 마지막 상태=!code!^)
goto health_wait

:error
echo.
echo ============================================================
echo   ⚠ 실행 중 오류 발생
echo ============================================================
echo   로그 위치:
echo     - Flask: %~dp0dashboard\logs\itglobal_dashboard.log
echo     - 에러  : %~dp0dashboard\logs\itglobal_dashboard_error.log
echo     - NSSM  : %~dp0dashboard\logs\service_stderr.log
echo.
pause
exit /b 1

:end
echo   ^(창을 닫으려면 아무 키나 누르세요^)
pause >nul
exit /b 0
