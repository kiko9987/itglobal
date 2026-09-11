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

Write-Host "✅ '$TaskName' 등록 완료 — 1분마다 SYSTEM 계정으로 /health 감시." -ForegroundColor Green
Write-Host "   로그: dashboard\logs\watchdog.log" -ForegroundColor Cyan
Write-Host "   즉시 1회 실행 테스트: Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Cyan
Write-Host "   제거: Unregister-ScheduledTask -TaskName '$TaskName' -Confirm:`$false" -ForegroundColor DarkGray
