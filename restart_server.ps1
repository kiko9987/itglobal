# ITGFlask 확실한 재시작 스크립트  (반드시 "관리자 권한" PowerShell 에서)
#
# 사용:  cd 프로젝트폴더 후  ->  .\restart_server.ps1     ( 앞에 .\ 필수 )
#
# 배경: nssm 서비스가 포트 5000 python 을 못 죽이는 경우가 있어 Restart-Service 만으론
#   옛 코드가 계속 서비스될 수 있다. 또 python 이 LocalSystem 이라 PowerShell Stop-Process 는
#   권한 거부로 실패한다(taskkill 은 서비스 프로세스도 종료 가능). 이 스크립트는:
#     1) Restart-Service 시도  2) 그래도 옛 프로세스가 남으면 taskkill 로 트리째 종료 후 재시작
#     3) 새 프로세스가 뜰 때까지 대기.

$ErrorActionPreference = 'SilentlyContinue'
Write-Host "== ITGFlask 재시작 ==" -ForegroundColor Cyan

$before = (Get-NetTCPConnection -LocalPort 5000 -State Listen | Select-Object -First 1).OwningProcess
Write-Host "현재 포트5000 PID: $before"

Write-Host "1) Restart-Service..."
Restart-Service ITGFlask -Force
Start-Sleep -Seconds 3

$after = (Get-NetTCPConnection -LocalPort 5000 -State Listen | Select-Object -First 1).OwningProcess
if ($after -and $after -eq $before) {
    Write-Host "2) 잔여 프로세스가 남음 -> taskkill 로 트리째 강제 종료..." -ForegroundColor Yellow
    $parent = (Get-CimInstance Win32_Process -Filter "ProcessId=$after").ParentProcessId
    cmd /c "taskkill /F /T /PID $after"   | Out-Null
    if ($parent) { cmd /c "taskkill /F /T /PID $parent" | Out-Null }
    Start-Sleep -Seconds 2
    Start-Service ITGFlask
}

Write-Host "3) 기동 대기 (30~60초)..." -ForegroundColor Cyan
for ($i = 0; $i -lt 40; $i++) {
    Start-Sleep -Seconds 3
    $c = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($c) {
        $p = Get-CimInstance Win32_Process -Filter "ProcessId=$($c.OwningProcess)"
        if ((Get-Date $p.CreationDate) -gt (Get-Date).AddMinutes(-5)) {
            Write-Host "`n✅ 새 프로세스 기동 완료 (PID $($c.OwningProcess), 시작 $((Get-Date $p.CreationDate -f 'HH:mm:ss')))" -ForegroundColor Green
            return
        }
    }
    Write-Host "   대기 $($i + 1)..."
}
Write-Host "`n⚠️ 시간 내 새 프로세스 확인 안 됨 — dashboard\logs 확인 필요" -ForegroundColor Yellow
