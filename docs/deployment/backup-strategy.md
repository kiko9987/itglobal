# 백업 전략

## 자동 백업 (2026-07-09~)

`scripts/backup_daily.py` — Python 스크립트로 다음 백업 수행:

| 대상 | 경로 | 방식 |
|---|---|---|
| SQLite `instance/users.db` | `backup/users_db/{YYYYMMDD}.db` | `shutil.copy2` (WAL 은 SQLite 가 안전하게 처리) |
| Redis `dump.rdb` | `backup/redis/{YYYYMMDD}.rdb` | Docker 컨테이너에 `BGSAVE` 후 `docker cp` |
| 시크릿 (`.env`, `credentials.json`, `google_calendar_token.json`) | `backup/secrets/{YYYYMMDD}/*` | `shutil.copy2` — Windows ACL 그대로 유지 |
| 로그 자동 정리 | `dashboard/logs/*.log*` | 30일 이상된 파일 자동 삭제 |

**보관 위치**: 프로젝트 루트 `backup/` (git ignore됨).
**보관 기간**: 30일 (자동 정리 — 스크립트가 매일 실행 후 오래된 파일 삭제).

## 실행 방법

### 자동 실행 (권장 — 이미 활성)

Flask 부팅 시 APScheduler 크론 등록 (매일 03:15 KST):
```python
# dashboard/services/sync_scheduler.py
scheduler.add_job(_safe_daily_backup, 'cron', hour=3, minute=15, id='_safe_daily_backup')
```

확인:
```powershell
# 스케줄러 로그에서 등록 확인
Select-String -Path dashboard\logs\service_stdout.log -Pattern "_safe_daily_backup" | Select-Object -Last 5

# 실제 실행 결과 (다음날 오전에 확인)
Get-ChildItem backup\users_db\ | Sort-Object LastWriteTime -Descending | Select-Object -First 5
```

### 수동 실행 (검증·즉시 백업)
```powershell
cd "C:\Users\SECOM\Desktop\ITG-Project\Claude Project"
.\.venv\Scripts\python.exe -X utf8 scripts\backup_daily.py
```

정상 출력 예시:
```
[BACKUP] 시작 @ 2026-07-09T21:34:21
  users_db: OK: 20260709.db (208 KB), 정리 0건
  redis: OK: 20260709.rdb (2566 KB, docker), 정리 0건
  secrets: OK: 복사 3건 (.env, credentials.json, google_calendar_token.json), 정리 0개 폴더
  log_cleanup: OK: 0개 파일 삭제 (0.0 MB 확보)
```

## 복구 절차

### SQLite users.db 복구
매니저·권한 데이터 손실 시.
```powershell
# 1. 서비스 정지
Stop-Service ITGFlask  # 관리자 권한 필요

# 2. 백업 복원 (원하는 날짜 선택)
Copy-Item "backup\users_db\20260709.db" "instance\users.db" -Force

# 3. 서비스 재시작
schtasks /Run /TN "Restart-ITGFlask"

# 4. 확인
curl https://pm.itg-aircon.com/api/health
```
**RTO**: ~2분. **손실 범위**: 마지막 백업 이후 신규 사용자 등록·권한 변경.

### Redis 복구
캐시·세션·큐·프로젝트 락 데이터 손실 시.
```powershell
# 1. 컨테이너 정지
docker stop redis

# 2. 백업 rdb 를 컨테이너 안에 복사
docker cp "backup\redis\20260709.rdb" redis:/data/dump.rdb

# 3. 재시작
docker start redis
docker exec redis redis-cli PING   # PONG 확인
docker exec redis redis-cli DBSIZE # 정상 시 수천 건

# 4. Flask 도 재시작 (기존 캐시 무효)
schtasks /Run /TN "Restart-ITGFlask"
```
**RTO**: ~5분. **손실 범위**: 마지막 백업 이후 진행 중이던 sheet write-behind 큐 op·모든 매니저 세션(재로그인 필요).

### 시크릿 복구
`.env` 편집 실수·credentials 만료·삭제 등.
```powershell
# 1. 백업에서 복사
Copy-Item "backup\secrets\20260709\.env" ".env" -Force
Copy-Item "backup\secrets\20260709\credentials.json" "credentials.json" -Force
Copy-Item "backup\secrets\20260709\google_calendar_token.json" "instance\google_calendar_token.json" -Force

# 2. 서비스 재시작
schtasks /Run /TN "Restart-ITGFlask"
```
**RTO**: ~1분. **주의**: 슬랙 토큰 재발급 후에는 백업 복원하면 새 토큰이 무효화됨 — 최신 값 재입력 필수.

## 재해 시나리오별 RPO/RTO

| 시나리오 | RPO (허용 손실) | RTO (복구 시간) | 절차 |
|---|---|---|---|
| users.db 손상 | 최대 24시간 | 2분 | ↑ SQLite 복구 |
| Redis 데이터 손상 | 최대 24시간 (진행 중 큐 op 유실 가능) | 5분 | ↑ Redis 복구 |
| `.env` 편집 실수 | 최대 24시간 | 1분 | ↑ 시크릿 복구 |
| Flask 코드 회귀 | 0 (git 되돌리기) | 10분 | `incident-response.md` §3.1 |
| 세콤 PC 하드웨어 장애 | ~24시간 (최근 backup/) | 4시간 (옛 PC 롤백) | `memory/project_secom_migration_2026-07-07.md` §롤백 절차 |
| Google Sheets 시트 파괴 | 30일 (Google 자체 버전 관리) | 30분 (매니저 안내 포함) | `incident-response.md` §2.5 |

## Non-destructive 백업 검증 (권장 — 매주)

서비스 중단 없이 백업 파일이 유효한지 확인.

```powershell
.\.venv\Scripts\python.exe -c "
import sqlite3, hashlib
from pathlib import Path

# SQLite: 해시 비교 + integrity check
orig = Path('instance/users.db').read_bytes()
back = Path('backup/users_db/{어제날짜}.db').read_bytes()
print('해시 일치:', hashlib.sha256(orig).hexdigest() == hashlib.sha256(back).hexdigest())
conn = sqlite3.connect('backup/users_db/{어제날짜}.db')
print('integrity:', conn.execute('PRAGMA integrity_check').fetchone()[0])
for t in ['users', 'constructors', 'audit_logs']:
    n = conn.execute(f'SELECT COUNT(*) FROM {t}').fetchone()[0]
    print(f'  {t}: {n}행')
"

# Redis: rdb 파일 무결성
docker run --rm -v "${PWD}\backup\redis:/data" redis:alpine `
  sh -c "redis-check-rdb /data/{어제날짜}.rdb | tail -10"
# '\o/ RDB looks OK! \o/' 출력이면 정상
```

**2026-07-09 리허설 결과** (실제 검증됨):
- users.db 212KB, 해시 일치, integrity OK, 22 users / 24 constructors / 296 audit_logs
- Redis dump.rdb 2.5MB, 11,556 keys, checksum OK

## 파괴적 리허설 (분기 1회 권장)

실제 복구가 되는지 검증. 새 백업이 만들어진 다음날 오전에 5분 이내 완료 가능.

### users.db 복구 리허설
```powershell
# 1. 현재 DB 를 임시로 옮김
Move-Item instance\users.db instance\users.db.rehearsal-backup

# 2. 어제 백업으로 복구
Copy-Item "backup\users_db\{어제날짜}.db" instance\users.db

# 3. Flask 재시작 (관리자 계정으로 로그인 되는지 확인)
schtasks /Run /TN "Restart-ITGFlask"

# 4. 브라우저에서 로그인 확인 후 원래대로 복원
Stop-Service ITGFlask
Remove-Item instance\users.db
Move-Item instance\users.db.rehearsal-backup instance\users.db
Start-Service ITGFlask
```

### Redis 복구 리허설
운영 시간 외(저녁·주말)에만 시도. 세션 강제 종료됨.
```powershell
# 1. 현재 dump 를 리허설용 별도 저장
docker exec redis redis-cli --rdb /tmp/rehearsal-current.rdb
docker cp redis:/tmp/rehearsal-current.rdb backup\redis\_rehearsal-current.rdb

# 2. 어제 백업으로 되돌림
docker stop redis
docker cp "backup\redis\{어제날짜}.rdb" redis:/data/dump.rdb
docker start redis
docker exec redis redis-cli DBSIZE  # 정상 확인

# 3. 원래대로 복원
docker stop redis
docker cp backup\redis\_rehearsal-current.rdb redis:/data/dump.rdb
docker start redis

# 4. Flask 재시작 (모든 매니저 재로그인 필요)
schtasks /Run /TN "Restart-ITGFlask"
```

## 참고

- **Google Sheets 자체 백업**: Google 이 자체 버전 관리 (30일 보관). 시트 열기 → 파일 → 버전 기록.
- **BitLocker 권장**: `backup/` 폴더는 시크릿·PII 포함. C 드라이브에 BitLocker 활성 유지.
- **외부 백업**: 하드웨어 장애 대비 월 1회 `backup/` 를 외부 USB·클라우드에 수동 복사 권장.
- **관련 문서**: [incident-response.md](./incident-response.md) — 이슈 유형별 대응 runbook.
