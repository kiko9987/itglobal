# ITGFlask 헬스 워치독 — /health 연속 무응답(wedge) 시 자동 재시작 + 슬랙 통보
#
# 배경: NSSM 은 프로세스가 '죽으면' 재시작하지만, 프로세스는 살아있는데 서빙 불능인
#   'wedge'(2026-08-19 Redis 순단→Waitress 소진→2시간 마비, memory project_redis_blip_app_wedge)
#   는 감지하지 못한다. 이 워치독이 그 공백을 메운다.
#
# 동작: 1분마다(스케줄 작업) /health probe →
#   - 연속 $FailThreshold 회 무응답 + 재시작 쿨다운 경과 시 → 서비스 재시작
#   - 부팅/재시작 직후 grace(프로세스 uptime < $StartupGraceSec)면 실패 무시(기동 중)
#   - 재시작·복구·복구실패(에스컬레이션)를 관리자 슬랙 DM 으로 통보
#   - 쿨다운으로 재시작 폭주(crash-loop) 방지
#
#
# Redis 감시 (2026-10-13 추가): WSL 자동 업데이트로 Docker VM 이 죽어 Redis 가 3.5시간 꺼졌는데
#   /health 는 200(임시 저장소 모드)이라 워치독이 못 잡았다. 그래서 Redis 를 직접 PING 한다.
#   - Redis 3분 연속 무응답 → 사용자 세션 작업 'ITGFlask-DockerRecover'(docker_recover.ps1) 시작
#     (Docker Desktop 은 사용자 세션 앱이라 SYSTEM 에서 직접 못 다룸) + 슬랙 통보
#   - Redis 는 정상인데 앱이 임시 저장소(fallback_mode)에 갇혀 있으면 → 앱 재시작
#
# 설치: 관리자 PowerShell 에서  .\install_watchdog.ps1  1회 실행 (SYSTEM 계정 1분 반복 작업 등록)
# 점검: .\watchdog.ps1 -SelfTest  → 재시작 없이 슬랙 알림 경로만 점검

param([switch]$SelfTest)

$ErrorActionPreference = 'SilentlyContinue'

# ── 설정 ──────────────────────────────────────────────────────
$HealthUrl          = 'http://localhost:5000/health'
$TimeoutSec         = 10       # /health 응답 대기 (정상 ~85ms, 첫히트 warmup 여유 포함)
$FailThreshold      = 3        # 연속 실패 N회 → 재시작 (약 3분 무응답)
$RestartCooldownSec = 600      # 재시작 간 최소 간격 10분 (crash-loop 방지)
$StartupGraceSec    = 120      # 프로세스 기동 후 2분은 실패 무시 (부팅/재시작 중)
$EscalateThrottleSec= 1800     # '복구 실패' 알림 최소 간격 30분
$ServiceName        = 'ITGFlask'
$ApiHealthUrl       = 'http://localhost:5000/api/health'   # services.cache = redis / fallback_mode
$RedisHost          = '127.0.0.1'
$RedisPort          = 6379
$RedisFailThreshold = 3        # Redis 연속 무응답 N회(약 3분) → 도커 복구 작업 시작
$RedisRecoverCooldownSec = 600 # 도커 복구 시도 간 최소 간격 10분
$RedisSlackThrottleSec   = 1800  # 'Redis 꺼짐' 반복 알림 최소 간격 30분
$FallbackThreshold  = 2        # Redis 정상인데 앱이 임시 저장소 모드 N회 연속 → 앱 재시작
$DockerTaskName     = 'ITGFlask-DockerRecover'
$LogPath            = Join-Path $PSScriptRoot 'dashboard\logs\watchdog.log'
$StatePath          = Join-Path $PSScriptRoot 'dashboard\logs\watchdog_state.json'

# ── 유틸 ──────────────────────────────────────────────────────
function Now-Unix { [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() }

function Write-WLog($msg) {
    try {
        $line = ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg)
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    } catch {}
}

function Get-EnvVal($name) {
    $envPath = Join-Path $PSScriptRoot '.env'
    if (-not (Test-Path $envPath)) { return $null }
    $m = Select-String -Path $envPath -Pattern ('^\s*' + [regex]::Escape($name) + '\s*=') | Select-Object -First 1
    if (-not $m) { return $null }
    return (($m.Line -split '=', 2)[1]).Trim()
}

function Send-Slack($text) {
    try {
        $token = Get-EnvVal 'SLACK_BOT_TOKEN'
        $chan  = Get-EnvVal 'ERROR_ALERT_CHANNEL'
        if (-not $chan) { $chan = Get-EnvVal 'SLACK_ADMIN_CHANNEL' }
        if (-not $token -or -not $chan) { Write-WLog 'SLACK skip: 토큰/채널 미설정'; return }
        $body = (@{ channel = $chan; text = $text } | ConvertTo-Json -Compress)
        $bytes = [Text.Encoding]::UTF8.GetBytes($body)
        Invoke-RestMethod -Uri 'https://slack.com/api/chat.postMessage' -Method Post -TimeoutSec 8 `
            -Headers @{ Authorization = "Bearer $token" } `
            -ContentType 'application/json; charset=utf-8' -Body $bytes | Out-Null
    } catch { Write-WLog ('SLACK 실패: ' + $_.Exception.Message) }
}

function Load-State {
    if (Test-Path $StatePath) {
        try { return (Get-Content $StatePath -Raw | ConvertFrom-Json) } catch {}
    }
    return [pscustomobject]@{ consecutiveFailures = 0; lastRestartUnix = 0; lastEscalateUnix = 0; incidentOpen = $false }
}

function Save-State($s) {
    try { ($s | ConvertTo-Json -Compress) | Set-Content -Path $StatePath -Encoding UTF8 } catch {}
}

function Set-StateProp($s, $name, $value) {
    # 옛 상태 파일엔 새 항목이 없음 → 없으면 추가
    if ($s.PSObject.Properties[$name]) { $s.$name = $value }
    else { $s | Add-Member -NotePropertyName $name -NotePropertyValue $value }
}

function Test-Redis {
    # Redis 에 직접 PING. +PONG(또는 비밀번호 요구 -NOAUTH)이면 살아 있음.
    try {
        $c = New-Object System.Net.Sockets.TcpClient
        $iar = $c.BeginConnect($RedisHost, $RedisPort, $null, $null)
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

function Get-AppCacheMode {
    # /api/health 는 정상이 아니면 503 이라 curl.exe 로 본문만 읽는다
    try {
        $raw = & curl.exe -s --max-time 10 $ApiHealthUrl
        if (-not $raw) { return $null }
        return (($raw | ConvertFrom-Json).services.cache)
    } catch { return $null }
}

function Test-Health {
    try {
        [System.Net.WebRequest]::DefaultWebProxy = $null   # 프록시 자동탐지 지연 제거
        $r = Invoke-WebRequest -Uri $HealthUrl -TimeoutSec $TimeoutSec -UseBasicParsing
        return ($r.StatusCode -eq 200)
    } catch { return $false }
}

function Get-AppUptimeSec {
    $c = Get-NetTCPConnection -LocalPort 5000 -State Listen -EA SilentlyContinue | Select-Object -First 1
    if (-not $c) { return $null }   # 리스너 없음 = 진짜 다운 (grace 아님)
    $p = Get-CimInstance Win32_Process -Filter ("ProcessId=" + $c.OwningProcess) -EA SilentlyContinue
    if (-not $p) { return $null }
    return ([DateTimeOffset]::UtcNow - [DateTimeOffset]$p.CreationDate).TotalSeconds
}

function Invoke-Restart {
    # restart_server.ps1 과 동일 전략: Restart-Service → 잔여 프로세스 taskkill 폴백
    $before = (Get-NetTCPConnection -LocalPort 5000 -State Listen -EA SilentlyContinue | Select-Object -First 1).OwningProcess
    Restart-Service $ServiceName -Force -EA SilentlyContinue
    Start-Sleep -Seconds 3
    $after = (Get-NetTCPConnection -LocalPort 5000 -State Listen -EA SilentlyContinue | Select-Object -First 1).OwningProcess
    if ($after -and $after -eq $before) {
        $parent = (Get-CimInstance Win32_Process -Filter ("ProcessId=" + $after) -EA SilentlyContinue).ParentProcessId
        cmd /c "taskkill /F /T /PID $after" | Out-Null
        if ($parent) { cmd /c "taskkill /F /T /PID $parent" | Out-Null }
        Start-Sleep -Seconds 2
        Start-Service $ServiceName -EA SilentlyContinue
    }
}

# ── 셀프테스트: 다운타임/재시작 없이 알림 경로만 점검 ─────────
if ($SelfTest) {
    Write-WLog 'SELFTEST: 슬랙 알림 경로 점검'
    Send-Slack (':test_tube: *[워치독 셀프테스트]* ITGFlask 워치독 알림 경로 정상 — 이 메시지가 보이면 자동복구 통보가 작동합니다. (실제 재시작 아님)')
    Write-Output 'SelfTest 완료 — 관리자 슬랙 DM 확인'
    return
}

# ── 메인 ──────────────────────────────────────────────────────
$now   = Now-Unix
$state = Load-State

# ── Redis 감시 (앱 /health 와 별개) ───────────────────────────
if (Test-Redis) {
    if ($state.redisIncidentOpen) {
        Write-WLog 'REDIS RECOVERED: PING 응답 재개'
        Send-Slack ':white_check_mark: *[자동복구] Redis 다시 응답* — 서버가 임시 저장소 모드면 곧 자동 재시작합니다.'
        Set-StateProp $state 'redisIncidentOpen' $false
    }
    Set-StateProp $state 'redisFailures' 0
    # Redis 는 정상인데 앱이 임시 저장소 모드에 갇힘(Redis 가 늦게 뜬 경우) → 앱 재시작해야 풀림
    if ((Get-AppCacheMode) -eq 'fallback_mode') {
        $fb = [int]$state.fallbackStreak + 1
        Set-StateProp $state 'fallbackStreak' $fb
        $up = Get-AppUptimeSec
        Write-WLog ('FALLBACK: Redis 정상인데 앱이 임시 저장소 모드 (연속 {0}회)' -f $fb)
        if ($fb -ge $FallbackThreshold -and $up -ne $null -and $up -ge $StartupGraceSec -and
                ($now - [int]$state.lastRestartUnix) -ge $RestartCooldownSec) {
            Write-WLog 'RESTART 실행 (임시 저장소 모드 해제)'
            $state.lastRestartUnix = $now
            Set-StateProp $state 'fallbackStreak' 0
            Save-State $state
            Invoke-Restart
            Send-Slack ':arrows_counterclockwise: *[자동복구] ITGFlask 재시작* — Redis 는 정상인데 서버가 임시 저장소 모드에 갇혀 있어 재시작함.'
        }
    }
    else { Set-StateProp $state 'fallbackStreak' 0 }
}
else {
    $rf = [int]$state.redisFailures + 1
    Set-StateProp $state 'redisFailures' $rf
    Write-WLog ('REDIS FAIL: PING 무응답 (연속 {0}회)' -f $rf)
    if ($rf -ge $RedisFailThreshold -and ($now - [int]$state.lastRedisRecoverUnix) -ge $RedisRecoverCooldownSec) {
        Set-StateProp $state 'lastRedisRecoverUnix' $now
        Set-StateProp $state 'redisIncidentOpen' $true
        Save-State $state
        $started = $false
        try { Start-ScheduledTask -TaskName $DockerTaskName -ErrorAction Stop; $started = $true } catch {}
        Write-WLog ('REDIS RECOVER: 도커 복구 작업 {0}' -f $(if ($started) { '시작' } else { '없음/실패' }))
        if (($now - [int]$state.lastRedisSlackUnix) -ge $RedisSlackThrottleSec) {
            Set-StateProp $state 'lastRedisSlackUnix' $now
            if ($started) {
                Send-Slack ((':rotating_light: *[Redis 꺼짐]* {0}분째 응답 없음 → 도커 자동 재시작 시도 중.' -f $rf) + "`n" +
                    '_복구되면 다시 알립니다. 그동안 방문 사진·상담 저장이 늦어질 수 있습니다._')
            }
            else {
                Send-Slack ((':rotating_light: *[Redis 꺼짐]* {0}분째 응답 없음.' -f $rf) + "`n" +
                    '*도커 자동 재시작 작업이 없어 수동 조치 필요* — 이 PC 에서 Docker Desktop 재시작.')
            }
        }
    }
}
Save-State $state

$ok    = Test-Health

if ($ok) {
    if ($state.incidentOpen) {
        Send-Slack ':white_check_mark: *[자동복구] ITGFlask 정상 복구 확인* — /health 응답 재개.'
        Write-WLog 'RECOVERED: /health 200, incident 종료'
    }
    $state.consecutiveFailures = 0
    $state.incidentOpen = $false
    Save-State $state
    return
}

# 실패
$state.consecutiveFailures = [int]$state.consecutiveFailures + 1
$uptime = Get-AppUptimeSec
$inGrace = ($uptime -ne $null -and $uptime -lt $StartupGraceSec)
Write-WLog ('FAIL: /health 무응답 (연속 {0}회, uptime={1})' -f $state.consecutiveFailures, ($(if ($uptime -eq $null) { 'no-listener' } else { [math]::Round($uptime) })))

if ($inGrace) {
    Write-WLog '기동 grace — 조치 보류'
    Save-State $state
    return
}

if ([int]$state.consecutiveFailures -ge $FailThreshold) {
    $sinceRestart = $now - [int]$state.lastRestartUnix
    if ($sinceRestart -ge $RestartCooldownSec) {
        Write-WLog ('RESTART 실행 (연속 {0}회 무응답)' -f $state.consecutiveFailures)
        $state.incidentOpen = $true
        Save-State $state
        Invoke-Restart
        $state.lastRestartUnix = $now
        $state.consecutiveFailures = 0
        Save-State $state
        Send-Slack (':arrows_counterclockwise: *[자동복구] ITGFlask 재시작 실행*' + "`n" +
            ('/health {0}회 연속 무응답(wedge) 감지 → 서비스 자동 재시작함.' -f $FailThreshold) + "`n" +
            '_수동 조치 불필요. dashboard\logs 확인 권장._')
    }
    else {
        # 쿨다운 내 지속 실패 = 직전 재시작으로도 복구 안 됨 → 에스컬레이션(throttle)
        if (($now - [int]$state.lastEscalateUnix) -ge $EscalateThrottleSec) {
            $state.lastEscalateUnix = $now
            $state.incidentOpen = $true
            Save-State $state
            Send-Slack (':rotating_light: *[자동복구 실패] ITGFlask* — 재시작 후에도 /health 무응답 지속.' + "`n" +
                '**수동 개입 필요** (Redis/시트 API/프로세스 상태 확인).')
            Write-WLog 'ESCALATE: 쿨다운 내 지속 실패 → 수동개입 알림'
        }
        else {
            Write-WLog ('쿨다운 대기 중(재시작 {0}s 전) — 조치 보류' -f $sinceRestart)
            Save-State $state
        }
    }
}
else {
    Save-State $state
}
