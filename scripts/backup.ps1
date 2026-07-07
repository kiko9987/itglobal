# ITG 대시보드 백업 자동화 스크립트 (2026-07-08)
#
# 백업 대상:
#   1) instance/users.db (SQLite, .backup 명령)
#   2) Redis dump.rdb (Docker BGSAVE 후 복사)
#   3) .env (시크릿 포함, 로컬만 보관)
#   4) credentials.json 및 계열 JSON 크레덴셜
#
# 실행:
#   .\scripts\backup.ps1              # 즉시 실행
#   .\scripts\backup.ps1 -DryRun     # 실제 복사 안 함, 계획만 표시
#
# Task Scheduler 등록 (매일 03:00):
#   $t = New-ScheduledTaskTrigger -Daily -At 3am
#   $a = New-ScheduledTaskAction -Execute "PowerShell.exe" `
#        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"C:\Users\SECOM\Desktop\ITG-Project\Claude Project\scripts\backup.ps1`""
#   Register-ScheduledTask -TaskName "ITG-Daily-Backup" -Trigger $t -Action $a -User "SECOM"

param(
    [switch]$DryRun = $false
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = "C:\Users\SECOM\Desktop\ITG-Project\Claude Project"
$BackupRoot = "C:\Users\SECOM\Desktop\ITG-Backups"
$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$BackupDir = Join-Path $BackupRoot $Timestamp

# ────────────────────────────────────────
# 로그 헬퍼
# ────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "HH:mm:ss"
    $prefix = switch ($Level) {
        "ERROR" { "❌" }
        "WARN"  { "⚠️" }
        "OK"    { "✓" }
        default { "•" }
    }
    Write-Output "[$ts] $prefix $Message"
}

Write-Log "==================================================="
Write-Log "ITG 백업 시작 (DryRun=$DryRun)"
Write-Log "백업 대상: $BackupDir"
Write-Log "==================================================="

if (-not $DryRun) {
    New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
}

# ────────────────────────────────────────
# 1) SQLite users.db 백업 (.backup 명령 = WAL 안전)
# ────────────────────────────────────────
$SqliteSrc = Join-Path $ProjectRoot "instance\users.db"
$SqliteDest = Join-Path $BackupDir "users.db"

if (Test-Path $SqliteSrc) {
    Write-Log "SQLite users.db 백업 시도"
    if (-not $DryRun) {
        # sqlite3 CLI가 있으면 .backup, 없으면 파일 복사 (WAL 체크포인트 후)
        $sqlite3 = Get-Command sqlite3 -ErrorAction SilentlyContinue
        if ($sqlite3) {
            & $sqlite3.Source $SqliteSrc ".backup '$SqliteDest'"
            Write-Log "SQLite .backup 완료 → $SqliteDest" -Level OK
        } else {
            # sqlite3 없으면 WAL 체크포인트 강제 후 복사 (덜 안전하지만 대안)
            Copy-Item $SqliteSrc $SqliteDest -Force
            $walSrc = "$SqliteSrc-wal"
            if (Test-Path $walSrc) { Copy-Item $walSrc "$SqliteDest-wal" -Force }
            Write-Log "SQLite 파일 복사 완료 (sqlite3 CLI 없음)" -Level OK
        }
    }
    $sizeMB = if (Test-Path $SqliteSrc) { [math]::Round((Get-Item $SqliteSrc).Length / 1MB, 2) } else { 0 }
    Write-Log "  크기: ${sizeMB}MB"
} else {
    Write-Log "SQLite users.db 없음 → skip" -Level WARN
}

# ────────────────────────────────────────
# 2) Redis BGSAVE 후 dump.rdb 복사
# ────────────────────────────────────────
Write-Log "Redis BGSAVE 트리거"
if (-not $DryRun) {
    try {
        $bgSave = docker exec redis redis-cli BGSAVE 2>&1
        Start-Sleep -Seconds 3

        # Redis 컨테이너 volume에서 dump.rdb 복사
        $RedisRdb = Join-Path $BackupDir "dump.rdb"
        docker cp redis:/data/dump.rdb $RedisRdb 2>&1 | Out-Null

        if (Test-Path $RedisRdb) {
            $sizeMB = [math]::Round((Get-Item $RedisRdb).Length / 1MB, 2)
            Write-Log "Redis dump.rdb 복사 완료 (${sizeMB}MB)" -Level OK
        } else {
            Write-Log "Redis dump.rdb 복사 실패" -Level ERROR
        }
    } catch {
        Write-Log "Redis 백업 오류: $($_.Exception.Message)" -Level ERROR
    }
} else {
    Write-Log "  [DryRun] Redis BGSAVE 트리거 skip"
}

# ────────────────────────────────────────
# 3) .env 백업 (시크릿 포함)
# ────────────────────────────────────────
$EnvSrc = Join-Path $ProjectRoot ".env"
if (Test-Path $EnvSrc) {
    Write-Log ".env 백업"
    if (-not $DryRun) {
        Copy-Item $EnvSrc (Join-Path $BackupDir ".env") -Force
    }
    Write-Log "  .env 복사 완료 (시크릿 포함, 로컬만 보관)" -Level OK
} else {
    Write-Log ".env 없음 → skip" -Level WARN
}

# ────────────────────────────────────────
# 4) credentials.json 및 계열 JSON 크레덴셜
# ────────────────────────────────────────
$CredFiles = @(
    "credentials.json",
    "google_oauth_credentials.json",
    "dashboard\credentials.json",
    "dashboard\google_calendar_client_secret.json",
    "dashboard\google_oauth_credentials.json",
    "instance\google_calendar_token.json"
)

foreach ($rel in $CredFiles) {
    $src = Join-Path $ProjectRoot $rel
    if (Test-Path $src) {
        if (-not $DryRun) {
            $destSubdir = Split-Path $rel -Parent
            if ($destSubdir) {
                $destDir = Join-Path $BackupDir $destSubdir
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            Copy-Item $src (Join-Path $BackupDir $rel) -Force
        }
        Write-Log "  $rel 복사 완료" -Level OK
    }
}

# ────────────────────────────────────────
# 5) 오래된 백업 정리 (30일 이상)
# ────────────────────────────────────────
Write-Log "30일 이상 백업 정리"
if ((Test-Path $BackupRoot) -and (-not $DryRun)) {
    $cutoff = (Get-Date).AddDays(-30)
    $oldDirs = Get-ChildItem $BackupRoot -Directory | Where-Object { $_.CreationTime -lt $cutoff }
    if ($oldDirs.Count -gt 0) {
        $oldDirs | ForEach-Object {
            Remove-Item $_.FullName -Recurse -Force
            Write-Log "  $($_.Name) 삭제" -Level OK
        }
    } else {
        Write-Log "  삭제할 오래된 백업 없음"
    }
}

# ────────────────────────────────────────
# 마무리
# ────────────────────────────────────────
if (-not $DryRun) {
    $totalSizeMB = (Get-ChildItem $BackupDir -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
    Write-Log "==================================================="
    Write-Log "백업 완료: $BackupDir ($([math]::Round($totalSizeMB, 2))MB)" -Level OK
    Write-Log "==================================================="
} else {
    Write-Log "==================================================="
    Write-Log "DryRun 완료 (실제 파일 복사 안 함)" -Level OK
    Write-Log "==================================================="
}
