# Docker Desktop / Redis 복구 — 사용자 세션 작업 'ITGFlask-DockerRecover' 로 실행
#
# 워치독(watchdog.ps1, SYSTEM)이 Redis 3분 연속 무응답을 감지하면 이 작업을 시작한다.
# Docker Desktop 은 로그인한 사용자 세션에서 도는 앱이라 SYSTEM 에서는 직접 재시작할 수 없다.
# (2026-10-13 WSL 자동 업데이트로 Docker VM 이 죽어 Redis 3.5시간 다운 — 사람이 재시작할 때까지 방치됨)
#
# 순서: ① 엔진은 살아 있고 컨테이너만 멈췄으면 docker start redis
#       ② 그래도 안 되면 docker desktop restart → 최대 5분 동안 redis 기동 확인
# 설치: install_watchdog.ps1 이 함께 등록 (트리거 없음 — 워치독이 필요할 때만 시작)

$ErrorActionPreference = 'SilentlyContinue'
$LogPath = Join-Path $PSScriptRoot 'dashboard\logs\watchdog.log'

function Write-WLog($msg) {
    try {
        $line = ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg)
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    } catch {}
}

function Test-Redis {
    try {
        $c = New-Object System.Net.Sockets.TcpClient
        $iar = $c.BeginConnect('127.0.0.1', 6379, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne(3000)) { $c.Close(); return $false }
        $c.EndConnect($iar)
        $st = $c.GetStream(); $st.ReadTimeout = 3000; $st.WriteTimeout = 3000
        $b = [Text.Encoding]::ASCII.GetBytes("PING`r`n"); $st.Write($b, 0, $b.Length)
        $buf = New-Object byte[] 64; $n = $st.Read($buf, 0, 64)
        $c.Close()
        $resp = [Text.Encoding]::ASCII.GetString($buf, 0, $n)
        return ($resp -like '+PONG*' -or $resp -like '-NOAUTH*')
    } catch { return $false }
}

Write-WLog 'DOCKER-RECOVER: 시작'
if (Test-Redis) { Write-WLog 'DOCKER-RECOVER: 이미 Redis 응답 — 조치 없음'; return }

# ① 컨테이너만 멈춘 경우
docker start redis 2>$null | Out-Null
Start-Sleep -Seconds 5
if (Test-Redis) { Write-WLog 'DOCKER-RECOVER: docker start redis 로 복구'; return }

# ② Docker Desktop 재시작 (VM 이 죽은 경우)
Write-WLog 'DOCKER-RECOVER: docker desktop restart 실행'
docker desktop restart 2>&1 | Out-Null
$deadline = (Get-Date).AddMinutes(5)
while ((Get-Date) -lt $deadline) {
    if (Test-Redis) { Write-WLog 'DOCKER-RECOVER: Docker Desktop 재시작으로 복구'; return }
    docker start redis 2>$null | Out-Null
    Start-Sleep -Seconds 10
}
Write-WLog 'DOCKER-RECOVER: 5분 안에 복구 안 됨 — 수동 조치 필요'
