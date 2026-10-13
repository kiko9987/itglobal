# ITGFlask 워치독 스케줄 작업 등록 — 반드시 "관리자 권한" PowerShell 에서 1회 실행
#
# 사용:  cd 프로젝트폴더  ->  .\install_watchdog.ps1
#
# 하는 일: watchdog.ps1 을 1분마다 SYSTEM 계정으로 실행하는 작업 'ITGFlask-Watchdog' 등록.
#   /health 3회 연속 무응답 시 서비스 자동 재시작 + 관리자 슬랙 통보.
# 제거:  Unregister-ScheduledTask -TaskName 'ITGFlask-Watchdog' -Confirm:$false

$ErrorActionPreference = 'Stop'

# 관리자 권한 확인
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "⚠️ 관리자 권한이 아닙니다. '관리자 권한으로 실행'한 PowerShell 에서 다시 실행하세요." -ForegroundColor Red
    return
}

$TaskName = 'ITGFlask-Watchdog'
$script   = Join-Path $PSScriptRoot 'watchdog.ps1'
if (-not (Test-Path $script)) {
    Write-Host "❌ watchdog.ps1 을 찾을 수 없습니다: $script" -ForegroundColor Red
    return
}

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $script)

# 1분마다 무기한 반복
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes 1) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Principal $principal -Settings $settings -Force `
    -Description 'ITGFlask /health 워치독: 3회 연속 무응답(wedge) 시 자동 재시작 + 슬랙 통보' | Out-Null

Write-Host "✅ '$TaskName' 등록 완료 — 1분마다 SYSTEM 계정으로 /health·Redis 감시." -ForegroundColor Green

# Redis 복구용 사용자 세션 작업 (2026-10-13) — Docker Desktop 은 사용자 세션 앱이라 SYSTEM 이 직접 못 다룸.
#   트리거 없음: 워치독이 Redis 3분 무응답 시 Start-ScheduledTask 로 시작.
#   실행 계정 = 이 관리자 PowerShell 을 연 사용자(도커를 쓰는 로그인 계정), 로그인 중일 때만 실행.
$DockerTask   = 'ITGFlask-DockerRecover'
$dockerScript = Join-Path $PSScriptRoot 'docker_recover.ps1'
if (Test-Path $dockerScript) {
    $dAction = New-ScheduledTaskAction -Execute 'powershell.exe' `
        -Argument ('-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}"' -f $dockerScript)
    $dPrincipal = New-ScheduledTaskPrincipal -UserId ("{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME) `
        -LogonType Interactive -RunLevel Limited
    $dSettings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    Register-ScheduledTask -TaskName $DockerTask -Action $dAction -Principal $dPrincipal `
        -Settings $dSettings -Force `
        -Description 'Redis 무응답 시 워치독이 시작: docker start redis → docker desktop restart' | Out-Null
    Write-Host "✅ '$DockerTask' 등록 완료 — Redis 꺼지면 워치독이 Docker Desktop 자동 재시작." -ForegroundColor Green
}
else {
    Write-Host "⚠️ docker_recover.ps1 없음 — Redis 자동 복구 작업은 건너뜀" -ForegroundColor Yellow
}
Write-Host "   로그: dashboard\logs\watchdog.log" -ForegroundColor Cyan
Write-Host "   즉시 1회 실행 테스트: Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Cyan
Write-Host "   제거: Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false" -ForegroundColor DarkGray
