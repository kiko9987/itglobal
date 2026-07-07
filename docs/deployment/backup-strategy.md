# 백업 전략

## 자동 백업 스크립트

`scripts/backup.ps1` — Windows PowerShell 스크립트로 다음 백업:

1. **SQLite `instance/users.db`** — `.backup` 명령 (WAL 안전), sqlite3 CLI 없으면 파일 복사
2. **Redis `dump.rdb`** — Docker `BGSAVE` 후 `docker cp`로 로컬 저장
3. **`.env`** — 시크릿 포함, 로컬만 보관
4. **credentials JSON** — Google OAuth·서비스 계정·캘린더 토큰

**보관 위치**: `C:\Users\SECOM\Desktop\ITG-Backups\{yyyy-MM-dd_HHmm}\`

**보관 기간**: 30일 (자동 정리)

## 실행 방법

### 즉시 실행 (검증)
```powershell
cd "C:\Users\SECOM\Desktop\ITG-Project\Claude Project"
.\scripts\backup.ps1 -DryRun  # 계획만 표시, 실제 복사 안 함
.\scripts\backup.ps1           # 실제 백업
```

### Windows Task Scheduler 자동화 (매일 03:00)

```powershell
$trigger = New-ScheduledTaskTrigger -Daily -At 3am
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"C:\Users\SECOM\Desktop\ITG-Project\Claude Project\scripts\backup.ps1`""
$principal = New-ScheduledTaskPrincipal -UserId "SECOM" -RunLevel Highest

Register-ScheduledTask -TaskName "ITG-Daily-Backup" `
    -Trigger $trigger `
    -Action $action `
    -Principal $principal `
    -Description "ITG 대시보드 매일 03시 자동 백업"
```

### 검증

Task Scheduler 실행 상태:
```powershell
Get-ScheduledTask -TaskName "ITG-Daily-Backup" | Format-List
Get-ScheduledTaskInfo -TaskName "ITG-Daily-Backup"  # 마지막 실행·다음 실행
```

수동 트리거:
```powershell
Start-ScheduledTask -TaskName "ITG-Daily-Backup"
```

## 복구 절차

### SQLite users.db 복구
```powershell
Stop-Service ITGFlask
Copy-Item "C:\Users\SECOM\Desktop\ITG-Backups\{날짜}\users.db" `
    "C:\Users\SECOM\Desktop\ITG-Project\Claude Project\instance\users.db" -Force
Start-Service ITGFlask
```

### Redis 복구
```powershell
docker stop redis
docker cp "C:\Users\SECOM\Desktop\ITG-Backups\{날짜}\dump.rdb" redis:/data/dump.rdb
docker start redis
```

### `.env` 복구
```powershell
Copy-Item "C:\Users\SECOM\Desktop\ITG-Backups\{날짜}\.env" `
    "C:\Users\SECOM\Desktop\ITG-Project\Claude Project\.env" -Force
schtasks /Run /TN "Restart-ITGFlask"
```

## 참고

- **Google Sheets 자체 백업**: 시트는 Google Drive에서 자체 버전 관리 (변경 이력 30일).
  추가 export 필요 시 스케줄러에 `export_sheet_to_csv` job 추가 검토 (지금은 미포함).
- **BitLocker 권장**: 백업 볼륨은 BitLocker 등으로 암호화 (시크릿·PII 보호).
- **외부 백업**: 하드웨어 장애 대비 월 1회 외부 저장소(USB·클라우드) 수동 복사 권장.
