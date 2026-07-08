# 실서비스 개시 후 긴급 대응 (Incident Response Runbook)

**대상**: 관리자(kiko@itg-aircon.com) — 매니저 20명 실사용 중 이슈 발생 시 즉시 참조.
**전제 스택**: Flask + Waitress + Caddy(HTTPS) + Redis(Docker) + Google Sheets API.
**서버 PC**: SECOM (192.168.0.42) — `C:\Users\SECOM\Desktop\ITG-Project\Claude Project`

---

## 섹션 1. 이슈 감지 (첫 신호 & 즉시 확인)

| 신호 | 출처 | 즉시 확인 명령 |
|---|---|---|
| Sentry 이슈 알림 | Slack DM/이메일 | Sentry 대시보드에서 스택트레이스 → 커밋 특정 |
| `/api/health` 503 | `curl https://pm.itg-aircon.com/api/health` | 응답 JSON의 6개 서비스 (redis/db/filesystem/cache/scheduler/google_sheets) 중 실패 항목 |
| 매니저 슬랙 문의 | Slack 채널 | "어떤 화면 / 어떤 조작 / 몇 시경" 3가지부터 확보 |
| Waitress queue 폭주 | `service_stdout.log` "queue depth" 로그 | `Get-Content dashboard\logs\service_stdout.log -Tail 100 -Wait` |
| 서비스 자체 응답 없음 | 브라우저 접속 실패 | `Get-Service ITGFlask, Caddy` + `docker ps --filter name=redis` |

즉시 참조 3개 명령:
```powershell
curl -s https://pm.itg-aircon.com/api/health | ConvertFrom-Json
Get-Service ITGFlask, Caddy; docker ps --filter name=redis
Get-Content dashboard\logs\service_stderr.log -Tail 50
```

---

## 섹션 2. 이슈 유형별 대응 (Runbook)

### 2.1 서비스 자체 다운 (HTTPS 접속 실패)

```powershell
# 1. 상태 확인
Get-Service ITGFlask, Caddy

# 2. 최근 로그 (에러 위주)
Get-Content dashboard\logs\service_stderr.log -Tail 50
Get-Content dashboard\logs\service_stdout.log -Tail 50

# 3. 재시작 (Scheduled Task로 권한 없이 실행 가능)
schtasks /Run /TN "Restart-ITGFlask"
Start-Sleep 10
curl https://pm.itg-aircon.com/api/health
```

3회 재시도해도 실패 → **섹션 4의 "점검 중" 공지 발송 + 섹션 3 롤백 진행**.
Caddy만 죽었으면 `Restart-Service Caddy` (관리자 권한 필요).

### 2.2 Redis 다운

세션·캐시·Pub/Sub·큐가 Redis 의존 → 매니저 재로그인 필요.

```powershell
docker ps --filter name=redis
docker start redis                       # 컨테이너 stop된 경우
docker exec redis redis-cli PING         # PONG 확인
docker exec redis redis-cli DBSIZE       # 정상 시 4000+
```

Docker Desktop 자체가 죽었으면 → Docker Desktop 앱 실행 → `docker start redis`.
Redis 데이터가 손상됐다면 `docs/deployment/backup-strategy.md` 복구 절차.

### 2.3 Google Sheets API 오류 (429 / 500)

쿼터: read 1500/min (프로젝트), read/write 300/min (사용자). 승인된 상태.

```powershell
# 사용률 확인 (Cloud Console)
Start-Process "https://console.cloud.google.com/apis/api/sheets.googleapis.com/quotas?project=smooth-unison-470801-p5"

# 429는 tenacity로 자동 재시도 — 로그에서 재시도 성공 여부 확인
Select-String -Path dashboard\logs\service_stdout.log -Pattern "429|RateLimit" | Select-Object -Last 20
```

임시 조치 (지속 429 시): `.env`에서 `CACHE_TTL=1200` → `2400`으로 확장 → `schtasks /Run /TN "Restart-ITGFlask"`.
500 계열은 Google 측 장애 가능성 — https://status.cloud.google.com 확인.

### 2.4 취소·재개·편집 오류

```powershell
# 1. 최근 커밋에서 원인 후보
git log --oneline -20

# 2. Sentry에서 스택트레이스 → 파일:라인 특정
# 3. 프로젝트별 락이 남아있으면 강제 해제 (해당 lead_no)
docker exec redis redis-cli DEL "consult_submit_lock:{lead_no}"
docker exec redis redis-cli DEL "visit_action_lock:{lead_no}:{action}"
docker exec redis redis-cli DEL "channeltalk_lead_lock:{chat_id}"
```

특정 커밋이 원인으로 확정되면 → 섹션 3.1 revert.

### 2.5 데이터 손상 (시트 값 이상)

Google Sheets는 자체 버전 관리(30일) → 시트에서 직접 이전 버전 복원.
- 시트 열기 → 파일 → 버전 기록 → 특정 시점 복원 (해당 행만 복사·붙여넣기 권장)
- 컬럼 시프트 잔재 관련 데이터 파괴(과거 `efabf2a` 사례) 재발 시: `git log --oneline --all --grep="컬럼\|column"` 로 유사 패턴 확인
- 로컬 SQLite / Redis / `.env` 복구는 `docs/deployment/backup-strategy.md` 참조

---

## 섹션 3. 즉시 롤백 절차

### 3.1 마지막 커밋 되돌리기 (가장 흔한 케이스)

```powershell
git log --oneline -10                    # 문제 커밋 SHA 확인
git revert <sha> --no-edit               # 되돌리기 커밋 생성
git push origin main
# 프론트 변경이 포함됐다면
cd dashboard; npm run build; cd ..
schtasks /Run /TN "Restart-ITGFlask"
curl https://pm.itg-aircon.com/api/health
```

### 3.2 이전 안정 버전으로 롤백 (여러 커밋이 얽혔을 때)

```powershell
# 안정 커밋 특정 (예: 실서비스 개시 시점 태그 또는 SHA)
git log --oneline -30
git checkout -b hotfix-rollback <안정_sha>
cd dashboard; npm install; npm run build; cd ..
schtasks /Run /TN "Restart-ITGFlask"
# 정상 확인 후 main에 병합
```

### 3.3 세콤 PC 자체가 문제 (하드웨어·OS)

옛 KiKO PC로 롤백 → `memory/project_secom_migration_2026-07-07.md` "🔁 롤백 필요 시 절차" 참조 (공유기 포트포워딩 원복 → 옛 PC 서비스 재활성화).

### 3.4 서비스 완전 중단 (원인 조사 시간 필요)

```powershell
Stop-Service ITGFlask       # 관리자 권한
# 섹션 4의 "서비스 이슈 발생" 공지 즉시 발송
# 원인 분석 후 재개
Start-Service ITGFlask
```

---

## 섹션 4. 매니저 공지 템플릿

Slack 채널(운영 공지 채널 또는 매니저 DM). 격식 없이 짧게.

**잠시 점검 중 (5분 이내 복구 예상)**
> [ITG 대시보드] 방금 잠깐 점검 들어갑니다. 5분 안에 복구 예정이고, 지금 작성 중인 내용은 잠시 후 다시 저장 부탁드립니다.

**서비스 이슈 발생 (원인 조사 중)**
> [ITG 대시보드] 지금 서비스 접속·저장이 원활하지 않습니다. 원인 확인 중이고, 복구 시 이 채널로 재공지드립니다. 급한 건은 슬랙으로 공유 부탁드려요.

**서비스 복구 완료**
> [ITG 대시보드] 정상 복구 완료했습니다. 방금 전 작업(취소/편집 등) 결과 한 번씩만 확인 부탁드립니다. 반영 안 된 게 있으면 알려주세요.

**데이터 확인 요청 (편집 이력 확인)**
> [ITG 대시보드] {프로젝트 코드} 관련해서 오늘 {시간대}에 편집하신 분 계신가요? 시트 값 확인이 필요합니다. 편집자 확인되면 원본 복구 진행하겠습니다.

---

## 섹션 5. 첫 3일 모니터링 체크리스트

### 매일 아침 09:00
- [ ] `curl https://pm.itg-aircon.com/api/health` → 200 응답, 6개 서비스 모두 healthy
- [ ] Sentry 새 이슈 (전날 이후) 검토
- [ ] Waitress 스레드 사용률 정상 (< 30 pool, connection < 200)
- [ ] `docker exec redis redis-cli DBSIZE` 정상 증가율 (전날 대비 ±수십 건)
- [ ] Google Sheets API 사용률 (Cloud Console) < 20%
- [ ] 매니저 슬랙 문의 확인
- [ ] `Select-String -Path dashboard\logs\service_stdout.log -Pattern "ERROR|CRITICAL" | Select-Object -Last 20`

### 매일 저녁 18:00
- [ ] Task Scheduler `ITG-Daily-Backup` 실행 이력 확인 (`Get-ScheduledTaskInfo -TaskName "ITG-Daily-Backup"`)
- [ ] 시트 데이터 sample 검증 (당일 편집된 프로젝트 3~5건 UI ↔ 시트 값 일치 확인)
- [ ] `dashboard\logs\` 디스크 사용량 확인 (100MB 초과 시 로테이션 검토)

### 주간 (금요일 저녁)
- [ ] Redis AOF/RDB 파일 크기 및 백업 아카이브 존재 확인 (`C:\Users\SECOM\Desktop\ITG-Backups\`)
- [ ] Caddy 인증서 만료일 (90일 주기) — `SLACK_PUBLIC_HOST` 도메인 자동 체크가 09시 매일 알림

---

## 섹션 6. 긴급 연락처 · 리소스

| 리소스 | URL / 값 |
|---|---|
| Google Cloud Console (프로젝트) | https://console.cloud.google.com/home/dashboard?project=smooth-unison-470801-p5 |
| Sheets API 쿼터 | https://console.cloud.google.com/apis/api/sheets.googleapis.com/quotas?project=smooth-unison-470801-p5 |
| Sentry 프로젝트 | https://sentry.io (DSN은 `.env`의 `SENTRY_DSN`) |
| Slack 앱 관리 (봇 4개) | https://api.slack.com/apps |
| Slack webhook 단일 진입점 | https://pm.itg-aircon.com/slack/events |
| GitHub 저장소 | https://github.com/kiko9987/itglobal |
| Let's Encrypt 상태 | https://letsencrypt.status.io |
| Google Cloud 상태 | https://status.cloud.google.com |
| 관리자 Slack DM | `SLACK_ADMIN_CHANNEL=U04UL2ZLJAX` (kiko) |
| 관리자 이메일 | kiko@itg-aircon.com |

### 관련 문서
- `OPERATIONS.md` — 시스템 구성, 서비스 상태 명령, 자주 발생 문제
- `docs/deployment/backup-strategy.md` — 백업 스크립트 및 복구 절차
- `memory/project_secom_migration_2026-07-07.md` — 세콤 PC 이관 스냅샷 및 옛 PC 롤백 절차
- `memory/project_folder_open_protocol.md` — `itgfolder://` 프로토콜 문제 시
