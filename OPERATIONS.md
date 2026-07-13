# ITG-Aircon Dashboard 운영 가이드

회사 서버 PC 세팅 + 원격 유지보수 참고용.

## 1. 시스템 구성

| 컴포넌트 | 종류 | 비고 |
|---|---|---|
| Flask 백엔드 | NSSM 윈도우 서비스 (`ITGFlask`, **LocalSystem 계정**) | 포트 5000 — 모든 슬랙 webhook/UI 진입점 |
| Caddy 리버스 프록시 | NSSM 윈도우 서비스 | 포트 443(HTTPS) → 내부 5000 |
| Vite 빌드 산출물 | `dashboard/static/dist/` | 프론트엔드 JS 변경 시 `npm run build` 필요 |
| Redis | Docker 컨테이너 (`redis:alpine`) | 포트 6379, 볼륨 `redis-data`, AOF+RDB 영속화. Pub/Sub · 캐시 · **Sheet write-behind 큐** · 프로젝트 락 |
| Sheet write-behind 큐 워커 | Flask 프로세스 내부 데몬 스레드 | Google Sheets API 지연 UX 분리. `/admin/queue-status` 로 모니터링 |
| 일 백업 스케줄러 | APScheduler cron (매일 03:15) | `scripts/backup_daily.py` — users.db + Redis dump.rdb + 시크릿, 30일 유지 + 30일 이상 로그 자동 삭제 |
| 슬랙 API 실패 자동 감지 | 로깅 필터 (`_SlackErrorSnooper`) | 5분 이내 3건 초과 시 관리자 슬랙 DM. `/admin/queue-status` 카드로 실시간 확인 |
| 데이터 정합성 감사 | 관리자 페이지 `/admin/data-integrity` | write-behind 워커 실패 조용한 누적 감지. 캐시 vs 시트 실제 diff |
| 라우터 포트포워딩 | 외부 443 → 내부 PC:443 | DDNS 도메인: `pm.itg-aircon.com` |
| Google Sheets | 외부 API | 서비스 계정 credentials.json |
| Google Drive | 외부 API + 로컬 동기화 | 시설별 폴더 자동 생성 |
| Slack 봇 4개 | 외부 (Bolt for Python) | main / visit / project / payment |
| Gmail | 외부 (OAuth) | 홈페이지 메일 자동 처리 |
| 채널톡 | 외부 (Developer API) | 채팅 인입 ↔ 슬랙 양방향 |
| Kakao Local API | 외부 (REST) | 주소 검증 |

### 슬랙 webhook 진입점
- 단일 도메인: `https://pm.itg-aircon.com/slack/events`
- 4개 봇 + 슬래시 명령 모두 이 URL 사용

### Flask 서비스 계정 = LocalSystem (중요)
NSSM으로 등록된 `ITGFlask`는 **LocalSystem 계정**으로 실행됩니다. 이로 인해:
- 사용자 세션의 매핑 드라이브(`G:` 등 Google Drive Desktop)가 **보이지 않음**
- 사용자 데스크톱에 창을 띄울 수 없음 (`subprocess.run(['explorer', ...])` 등 무효)
- 이 제약 때문에 문서 폴더 열기는 서버 subprocess가 아닌 **클라이언트 커스텀 URL 프로토콜** (`itgfolder://`)로 구현됨 — [docs/employee-deployment/itgfolder-install.md](docs/employee-deployment/itgfolder-install.md) 참고

## 2. 핵심 파일 / 위치

| 항목 | 경로 |
|---|---|
| 프로젝트 루트 | `C:\Users\KiKO\Desktop\ITG-Project\Claude Project` |
| 환경 변수 | `.env` (루트) — 모든 토큰, 시트 ID, webhook URL |
| Google 서비스 계정 | `credentials.json` (루트) — 시트/드라이브 접근 |
| Gmail OAuth 토큰 | `token.json` (루트) |
| Caddy 설정 | `Caddyfile` (루트) |
| Redis 자동 백업 | `backup/redis/{YYYYMMDD}.rdb` (03:15 크론) |
| SQLite 자동 백업 | `backup/users_db/{YYYYMMDD}.db` (03:15 크론) |
| 백업 스크립트 | `scripts/backup_daily.py` |
| Flask 진입점 | `app.py` |
| 슬랙 봇 핸들러 | `dashboard/blueprints/slack_bot.py` (3500줄) |
| 슬랙 공통 유틸 | `dashboard/blueprints/slack_helpers.py` (모달 state 추출, 이니셜 매핑, 시간 표시) |
| 채널톡 핸들러 | `dashboard/blueprints/channeltalk.py` |
| 채널톡 공통 유틸 | `dashboard/blueprints/channeltalk_helpers.py` (시간 포맷, 스팸 감지) |
| 폴링 스케줄러 | `dashboard/services/sync_scheduler.py` (관리자 알림 + 인증서 체크 포함) |
| 리드 서비스 | `dashboard/services/lead_service.py` — 시트 CRUD (16열, A~P) |
| 리드 sync | `dashboard/services/lead_sync.py` — 당근/전화 워크플로 폴링 |
| 폴더 API | `dashboard/blueprints/folders.py` + `dashboard/utils/google_drive.py` |
| Vite 프론트엔드 소스 | `dashboard/src/` |
| Vite 빌드 산출물 | `dashboard/static/dist/` (git에 커밋됨) |
| 직원 배포 문서 | `docs/employee-deployment/` |
| pytest 테스트 | `tests/unit/test_slack_channeltalk_core.py` |
| 가상환경 | `.venv/` |

## 3. 회사 서버 PC 마이그레이션 체크리스트

### 사전 준비
- [ ] 서버 PC에 고정 사설 IP 설정 (또는 DHCP 예약)
- [ ] Windows 자동 로그인 + 절전 모드 OFF
- [ ] 방화벽: 5000/443 포트 개방 (또는 Flask/Caddy 추가)

### 소프트웨어 설치
- [ ] Python 3.9+ (현재 PC 버전과 동일)
- [ ] Node.js 18+ (Vite 프론트엔드 빌드용)
- [ ] Git for Windows
- [ ] Docker Desktop (WSL2 백엔드)
- [ ] NSSM (https://nssm.cc)
- [ ] Caddy (https://caddyserver.com)

### 코드/데이터 이전
- [ ] `git clone <리포지토리>` — 코드 가져오기
- [ ] `python -m venv .venv` + `pip install -r requirements.txt`
- [ ] `cd dashboard && npm install` — Vite 의존성
- [ ] `npm run build` — 프론트엔드 빌드 (필수: `static/dist/`가 있어야 UI 로드됨)
- [ ] `.env` 복사 (수동, 시크릿 포함)
- [ ] `credentials.json`, `token.json` 복사
- [ ] `Caddyfile` 복사
- [ ] Redis 데이터 마이그레이션:
  ```powershell
  # 현재 PC에서
  docker exec redis redis-cli BGSAVE
  docker cp redis:/data/dump.rdb backup.rdb

  # 서버 PC에서
  docker volume create redis-data
  docker run -d --name redis --restart always -p 6379:6379 -v redis-data:/data redis:alpine redis-server --save 900 1
  docker cp backup.rdb redis:/data/dump.rdb
  docker restart redis
  docker exec redis redis-cli CONFIG SET appendonly yes
  docker exec redis redis-cli BGREWRITEAOF
  # 컨테이너 재생성 시 --appendonly yes 추가
  ```

### 서비스 등록 (NSSM)
- [ ] Flask:
  ```powershell
  nssm install ITGFlask "C:\path\to\.venv\Scripts\python.exe" "C:\path\to\app.py"
  nssm set ITGFlask AppDirectory "C:\path\to\project"
  nssm set ITGFlask Start SERVICE_AUTO_START
  nssm start ITGFlask
  ```
- [ ] Caddy:
  ```powershell
  nssm install Caddy "C:\path\to\caddy.exe" "run --config Caddyfile"
  nssm set Caddy AppDirectory "C:\path\to\project"
  nssm start Caddy
  ```

### 네트워크
- [ ] 라우터 포트포워딩 외부 443 → 새 서버 PC:443으로 변경
- [ ] DDNS 클라이언트 (`pm.itg-aircon.com`)가 서버 PC에서 실행되는지 확인
  - 또는 라우터 자체 DDNS 사용
- [ ] Caddy 자동 인증서 갱신 동작 확인 (Let's Encrypt)

### 검증
- [ ] `https://pm.itg-aircon.com/healthz` 외부에서 접속 확인
- [ ] 슬랙 봇 4개 모두 응답 (인입 카드, 모달, 슬래시 명령)
- [ ] APScheduler 작업 로그 확인 (`logs/`)
- [ ] 첫 인입 lead 전체 플로우 검증 (홈페이지 → 시트 → 슬랙 → 모달 제출 → 리스트 등록)
- [ ] 관리 사이트에서 프로젝트 상세 열기 → 문서 폴더 링크 렌더 확인 (UI 로드)
- [ ] `.venv\Scripts\python.exe -m pytest tests/unit/test_slack_channeltalk_core.py -v --no-cov` 통과

### 직원 PC 배포 (병렬 진행)
- [ ] 각 직원 PC에 Google Drive Desktop 설치 및 G: 마운트 확인
- [ ] `itgfolder://` 프로토콜 등록 — [docs/employee-deployment/itgfolder-install.md](docs/employee-deployment/itgfolder-install.md) 참고
- [ ] Slack Desktop 앱 설치 및 로그인

## 4. 원격 유지보수 환경

### 옵션 A: RDP (가장 단순)
- 서버 PC에 RDP 활성화
- 내 PC에서 `mstsc` 또는 클라이언트로 접속
- **사내망 한정**이면 그대로 OK, 외부 접속이면 VPN 권장

### 옵션 B: SSH + VS Code Remote (개발 친화적)
- 서버 PC에 OpenSSH Server 활성화
- 내 PC VS Code에서 Remote-SSH 확장으로 접속
- 코드 편집/디버깅이 RDP보다 가벼움

### 옵션 C: Git 푸시 기반 (가장 안전)
- 코드는 GitHub에서 관리
- 서버 PC에 작은 deploy 스크립트:
  ```powershell
  git pull
  pip install -r requirements.txt
  Restart-Service ITGFlask
  ```
- 변경 시 내 PC에서 git push → 서버 PC에서 deploy 스크립트 실행

**추천 조합**: A(RDP) + C(Git) — 일상 배포는 Git, 트러블슈팅 시 RDP

## 5. 운영 주의점

### 보안
- [ ] `.env`, `credentials.json`, `token.json`은 절대 Git 커밋 X (`.gitignore` 확인)
- [ ] 슬랙 봇 토큰 노출 시 즉시 회전 — Slack 앱 관리 페이지
- [ ] 카카오 API 키 마찬가지
- [ ] 방화벽 인바운드: 443만 허용 (5000 직접 노출 X)
- [ ] Windows 자동 업데이트는 작업 시간 외(새벽)로 설정 — 재부팅 후 서비스 자동 시작 보장
- [ ] 정기적으로 노출된 토큰 점검 (현재 task #23 pending)

### 백업
- [x] Redis: AOF + RDB 자동 (Docker 볼륨 영속화)
- [x] **자동 일 백업 (2026-07-09~)** — 매일 새벽 03:15 APScheduler 크론
  - `scripts/backup_daily.py` 실행 → `backup/users_db/{YYYYMMDD}.db` + `backup/redis/{YYYYMMDD}.rdb`
  - 30일 이상 지난 파일 자동 삭제
  - 수동 실행: `python scripts/backup_daily.py`
- [ ] Google 시트는 외부 SaaS — 별도 백업 불필요 (Google이 관리)
- [ ] 로그 디렉토리(`logs/`) 주기 정리 — 디스크 차오름 방지

### 모니터링
- 일일 점검:
  - Flask 서비스 상태: `Get-Service ITGFlask`
  - Caddy 상태: `Get-Service Caddy`
  - Redis 상태: `docker ps --filter name=redis`
  - 슬랙 #수금_관리, #방문_일정 채널에 정상 메시지 발송됐는지 확인

### 폭주 안전망 (구현됨)
- Redis 다운 시 모든 sync 작업 자동 skip (`_redis_healthy()` circuit breaker)
- 한 sync에 신규 lead 5건 초과 시 발송 skip (`KARROT_MAX_NEW_PER_SYNC`)
- SSL 에러로 누락된 슬랙 알림은 5분마다 자동 재시도 (`pending_slack_notify` Redis 큐)
- 채널톡 스팸 자동 감지 — 마케팅 키워드 + URL 조합 판정 시 미응답 알림 큐 skip
  (`_is_spam_message` in `channeltalk_helpers.py`)

### Sheet write-behind 아키텍처 (2026-07-08~)
Google Sheets API 지연을 사용자 UX에서 완전 분리하기 위한 비동기 쓰기 큐.

**동작 원리**
1. 매니저가 편집·취소·재개·계산서 요청 등 mutation 실행
2. 백엔드: 캐시(Redis + 메모리)만 즉시 업데이트하고 `sheet_write_queue`에 op 등록 → **<300ms 응답**
3. 백그라운드 워커가 Redis List(`sheet_write_queue`)를 LPOP 폴링하며 Google Sheets에 실제 반영

**핵심 파일**
| 위치 | 역할 |
|---|---|
| `dashboard/services/sheet_write_queue.py` | 큐 코어 (enqueue/worker/데드레터/depth 알림/rollup) |
| `dashboard/blueprints/projects.py` | `sheet_batch_write`, `sheet_bg_color`, `project_update_sheet`, `project_cancel_sheet`, `project_resume_sheet` 핸들러 등록 |
| `dashboard/services/project_slack_actions.py` | 슬랙 [편집]/[취소]/[재개] mutation write-behind |
| `dashboard/services/as_service.py` | A/S 상태 업데이트 write-behind |

**Redis 키**
| 키 | 종류 | 용도 |
|---|---|---|
| `sheet_write_queue` | List | 대기 op |
| `sheet_write_processing` | List | 처리 중 op (크래시 복구용) |
| `sheet_write_failed` | List | 3회 실패 → 데드레터 |
| `sheet_write_hourly_processed` | String | rollup 카운터 |

**안전장치**
- **3회 재시도** — 지수 백오프 (2s → 4s → 8s), 이후 데드레터
- **데드레터 슬랙 DM** — 관리자(`SLACK_ADMIN_CHANNEL`)에게 즉시 알림
- **큐 depth 알림** — pending > 100건 시 관리자 슬랙 DM (5분 쿨다운)
- **처리량 rollup** — 매 시간 처리 건수 로그 (`[QUEUE][ROLLUP]`)
- **크래시 복구** — 처리 중 op 는 `sheet_write_processing` 에 남아 부팅 시 복구 가능

**관리자 모니터링**
- 페이지: `/admin/queue-status` (5초 자동 갱신)
  - pending / processing / failed 카운트
  - 최근 실패 op 10건 + 인라인 재시도 버튼
- API: `GET /api/admin/sheet-write-queue`, `POST /api/admin/sheet-write-queue/retry/<op_id>`

**주의**
- Google Sheets는 이제 관리자 전용 (매니저는 대시보드만 사용). 시트를 직접 수정하면 캐시와 불일치 발생 가능.
- 워커 종료 시 처리 중 op가 유실될 수 있으나 캐시에는 이미 반영돼 있어 사용자 관점 영향 없음. 다음 prefetch 사이클에서 시트-캐시 불일치 감지 시 관리자가 큐 상태 페이지에서 확인 가능.

### 동시성 보호 (구현됨)
3중 락으로 race condition 방지:
| Redis 키 | TTL | 보호 대상 |
|---|---|---|
| `channeltalk_lead_lock:{chat_id}` | 60초 | 채널톡 매니저 응답 + 슬랙 모달 동시 처리 |
| `consult_submit_lock:{lead_no}` | 30초 | 두 매니저가 같은 lead 모달 동시 제출 (데이터 손실 방지) |
| `visit_action_lock:{lead_no}:{action}` | 5초 | 방문 카드 [방문일 수정] / [방문 취소] 동시 클릭 |
| `_APPEND_LEAD_LOCK` (process-wide threading.Lock) | - | 신규 lead_no 발번 race (빈 행 방지) |

### 운영 알림 (구현됨)
Redis 다운 / sync 실패 / SSL 인증서 만료 30일 이내 시 자동 알림:
- 슬랙 DM (`SLACK_ADMIN_CHANNEL=U04UL2ZLJAX`)
- 이메일 (`ADMIN_NOTIFY_EMAIL=kiko@itg-aircon.com`) — Gmail API send
- 같은 알림은 30분 cooldown (스팸 방지)
- SSL 인증서: 매일 09시 `SLACK_PUBLIC_HOST` 도메인 체크

**Gmail send scope 요구**: Google Workspace Admin 콘솔 도메인 위임에 `gmail.send` scope 추가 필요.

## 6. 자주 발생하는 문제

### Redis 다운
**증상**: 슬랙 메시지 폭주, 또는 `_redis_healthy()` 로그 다수
**조치**:
```powershell
docker ps --filter name=redis
docker start redis  # 또는 Docker Desktop 재시작
docker exec redis redis-cli PING  # PONG 확인
```

### Flask 응답 없음
**증상**: 슬랙 봇 무응답, healthz 안 뜸
**조치**:
```powershell
Get-Service ITGFlask
Restart-Service ITGFlask  # 관리자 권한 필요
```
원인 추적: `logs/dashboard.log`

### Caddy 인증서 갱신 실패
**증상**: HTTPS 접속 시 인증서 만료 경고
**조치**:
- 라우터 포트포워딩 443 정상인지
- Let's Encrypt rate limit 걸렸으면 staging 모드로 임시 회피
- `caddy reload --config Caddyfile`

### 시트 → 슬랙 알림 누락
**증상**: 인입은 됐는데 슬랙 카드 안 옴
**조치**:
- Redis `pending_slack_notify:*` 키 확인 — 5분마다 자동 재시도
- 즉시 재시도: `docker exec redis redis-cli KEYS "pending_slack_notify:*"`

### 슬랙 메시지 중복
**증상**: 같은 lead가 슬랙에 여러 번 옴
**원인**: Redis dedup 키 사라짐 (컨테이너 데이터 손실)
**조치**: 이미 영속화 적용됨 — 재발 시 AOF 파일 상태 확인

### 문서 폴더 링크 클릭해도 탐색기가 안 열림 (사용자 문의)
**증상**: 프로젝트 상세의 폴더 ID 링크 클릭 시 반응 없음, 또는 브라우저 오류
**원인 후보** (순서대로 확인):
1. 해당 직원 PC에 `itgfolder://` 프로토콜이 등록되지 않음 — [docs/employee-deployment/itgfolder-install.md](docs/employee-deployment/itgfolder-install.md) 재설치
2. `C:\ITG\open-itg-folder.vbs` 파일 없음
3. Google Drive Desktop 미실행 또는 G: 미마운트
4. 브라우저 팝업 차단 — 첫 클릭 시 "itgfolder 열기" 허용 필요

**주의**: 서버(Flask/NSSM/LocalSystem)는 이 흐름에 관여하지 않음. 서버 로그에는 흔적 없음. 모두 클라이언트 이슈로 좁혀서 진단.

### 신규 프로젝트 등록 시 API 400 (범위 관련)
**증상**: 프로젝트 시트 API 에러 `Requested writing within range ... but tried writing to column [XX]`
**원인**: `_build_row_values`가 반환하는 리스트 길이와 `append_row`의 하드코딩 범위 불일치
**조치**: 컬럼 추가/이동 시 `dashboard/utils/google_sheets.py:797` 의 `A{next_row}:AP{next_row}`도 함께 갱신. 현재 프로젝트 시트는 A~AP(42열, AO=Lead No, AP=_version).

## 7. 환경 변수 핵심 목록 (`.env`)

```bash
# Google Sheets — 프로젝트 시트 + 리드 시트 분리
GOOGLE_SHEET_ID=...                   # 공사 현황 시트 ID
GOOGLE_SHEET_NAME=공사 현황
ONLINE_LEADS_SHEET_ID=...             # 리드 관리 시트 ID (별개 문서)
ONLINE_LEADS_SHEET_NAME=리드 관리     # 탭 이름 (개명됨: 옛 "고객 리드 관리")
KARROT_AUTO_SHEET_ID=...              # 당근 자동 등록 시트
KARROT_AUTO_SHEET_TAB=...
GOOGLE_APPLICATION_CREDENTIALS=credentials.json
GOOGLE_CREDENTIALS_FILE=credentials.json
GOOGLE_CALENDAR_ID=primary
# Google Cloud Vision API — 사업자등록증 OCR 로 법인명·상호 자동 추출
# 사용 전 GCP Console → API & Services → Cloud Vision API → Enable
# 무료 티어 매월 1,000장, 이후 $1.5/1000장. 서비스 계정에 Vision API 접근 권한 필요.

# Slack 봇 6개 (각각 토큰 + signing secret 필요)
SLACK_BOT_TOKEN=...              # 메인 (온라인 문의 알림봇)
SLACK_SIGNING_SECRET=...
SLACK_VISIT_BOT_TOKEN=...        # 방문 일정 알림봇
SLACK_VISIT_SIGNING_SECRET=...
SLACK_PROJECT_BOT_TOKEN=...      # 공사 현황 알림봇
SLACK_PROJECT_SIGNING_SECRET=...
SLACK_PAYMENT_BOT_TOKEN=...      # 수금 관리 알림봇
SLACK_AS_BOT_TOKEN=...           # A/S 관리 알림봇 (선택)
SLACK_AS_SIGNING_SECRET=...

# Slack webhook 워크플로 (Slack List 조작)
SLACK_LIST_WEBHOOK_URL=...           # 방문 List 등록
SLACK_LIST_UPDATE_WEBHOOK_URL=...    # 방문 List 업데이트
SLACK_VISIT_CANCEL_WEBHOOK_URL=...
SLACK_VISIT_MODIFY_WEBHOOK_URL=...
SLACK_VISIT_RESTORE_WEBHOOK_URL=...
SLACK_VISIT_COMPLETE_WEBHOOK_URL=... # 방문 완료 (List 삭제)

# Slack 채널
SLACK_LEAD_CHANNEL=...          # 온라인 문의
SLACK_VISIT_CHANNEL=...         # 방문 일정
SLACK_PROJECT_CHANNEL=...       # 공사 확정
SLACK_PAYMENT_CHANNEL=...       # 수금 관리
SLACK_CHANNELTALK_CHANNEL=...
SLACK_INVOICE_CHANNEL_ID=...    # 세금계산서 발행 요청 카드 발송처 (#영업_관리)

# 세금계산서 관리 알림 봇 (invoice_bot) — 별도 Bolt App
SLACK_INVOICE_BOT_TOKEN=xoxb-... # Bot User OAuth Token
SLACK_INVOICE_SIGNING_SECRET=... # App Credentials → Signing Secret
# Endpoint: https://pm.itg-aircon.com/slack/invoice-events (Events + Interactivity 같은 URL)
# Bot events: message.channels, message.groups
# 스코프: chat:write, chat:write.public, channels:join, files:read,
#         channels:history, groups:history, groups:read, users:read, reactions:write

# 스케줄러 플래그
PHONE_WORKFLOW_SYNC_ENABLED=false  # 슬랙 워크플로 도입 전엔 false 유지 (수동 시트 입력 시 자동 카드 발송 방지)

# 외부 API
KAKAO_REST_API_KEY=...
CHANNELTALK_ACCESS_KEY=...
CHANNELTALK_ACCESS_SECRET=...
CHANNELTALK_OPERATOR_ID=...        # 봇 답변 시 표기할 운영자 ID
CHANNELTALK_REMIND_AFTER_MIN=30    # 미응답 알림 딜레이 (분)
HOMEPAGE_MAIL_USER=...             # Gmail 계정 (홈페이지 문의 처리용)
HOMEPAGE_PROCESSED_LABEL=처리완료   # 처리된 메일에 붙는 라벨
HOMEPAGE_MAIL_DAYS_BACK=3

# 폴링 주기 (선택)
KARROT_SYNC_INTERVAL_MIN=2
HOMEPAGE_MAIL_SYNC_INTERVAL_MIN=2
KARROT_MAX_NEW_PER_SYNC=5

# Google Drive 시설 폴더 (방문 사진 자동 저장 + itgfolder://)
GOOGLE_DRIVE_VISIT_FOLDER_ID=...              # 부모 폴더 ID
GOOGLE_DRIVE_WINDOWS_BASE_PATH=G:\공유 드라이브\...  # Drive Desktop 미러 경로

# 활성화 플래그
SLACK_BOT_ENABLED=true
PHONE_WORKFLOW_SYNC_ENABLED=true   # 전화 워크플로 sync 폴러

# Redis
REDIS_HOST=localhost
REDIS_PORT=6379
REDIS_DB=0

# 운영 알림 (Redis 다운, sync 실패, SSL 인증서 만료)
SLACK_ADMIN_CHANNEL=U04UL2ZLJAX     # 관리자 슬랙 User ID
SLACK_PUBLIC_HOST=pm.itg-aircon.com  # SSL 인증서 체크 도메인
ADMIN_NOTIFY_EMAIL=kiko@itg-aircon.com  # 이메일 알림 수신처
```

### 시트 컬럼 참고
| 시트 | 범위 | 최근 주요 변경 |
|---|---|---|
| **프로젝트 시트** (공사 현황) | A~AP (42열) | 2026-07 AO=Lead No 신설, _version이 AN→AP 이동 |
| **리드 시트** (리드 관리) | A~P (16열) | 2026-07 P열 = 폴더 ID (방문 사진 폴더 자동 저장) |

프로젝트 시트 컬럼 변경 시 반드시 `dashboard/utils/google_sheets.py:797`의 `A:AP` 범위와 `_build_row_values`의 리스트 크기(현재 42)를 함께 갱신.

## 8. 백업/복구 절차

### 전체 시스템 백업
```powershell
$date = Get-Date -Format yyyyMMdd
$dst = "backups\system_$date"
New-Item -ItemType Directory -Force $dst

# Redis
docker exec redis redis-cli BGSAVE
Start-Sleep 3
docker cp redis:/data/dump.rdb "$dst\redis_dump.rdb"

# 감사 로그 SQLite
Copy-Item dashboard_db.sqlite "$dst\dashboard_db.sqlite"

# 설정 파일
Copy-Item .env, credentials.json, token.json, Caddyfile $dst
```

### 복구
1. 위 백업 디렉토리에서 `.env`, `credentials.json`, `token.json`, `Caddyfile` 복원
2. Redis 데이터: `docker cp dump.rdb redis:/data/` + 재시작

## 9. 유용한 명령 모음

```powershell
# 서비스 상태
Get-Service ITGFlask, Caddy

# 서비스 재시작 (관리자)
Restart-Service ITGFlask

# Redis 상태
docker ps --filter name=redis
docker exec redis redis-cli INFO replication
docker exec redis redis-cli DBSIZE

# 로그 실시간 확인
Get-Content logs\dashboard.log -Tail 50 -Wait

# 슬랙 메시지 일괄 삭제 (특정 채널)
# 슬랙에서: /청소 (직접 입력)

# 환경 변수 확인 (Python REPL)
python -c "import os; from dotenv import load_dotenv; load_dotenv(); print(os.getenv('GOOGLE_SHEET_ID'))"

# 핵심 함수 회귀 검증 (코드 변경 후 권장)
.venv\Scripts\python.exe -m pytest tests/unit/test_slack_channeltalk_core.py -v --no-cov

# 프론트엔드 리빌드 (JS/CSS 변경 후 필수)
cd dashboard
npm run build
cd ..
Restart-Service ITGFlask

# 프로젝트 시트 append_row 범위 확인 (컬럼 추가 시)
Select-String -Path dashboard\utils\google_sheets.py -Pattern "A{next_row}"
```

## 10. 배포 · 문서

### Git pre-commit hook (자동 회귀 감지)
커밋할 때마다 `test_core_flows.py` 순수 함수 테스트 자동 실행. 실패하면
커밋 차단. 실행 시간 ~1초. Python 파일 변경 없는 커밋은 skip.

**최초 1회 설정** (새 clone 후):
```powershell
git config core.hooksPath .githooks
```

- 훅 스크립트: `.githooks/pre-commit`
- 커밋 시 자동 실행됨, 별도 명령 불필요
- 급하게 우회 필요 시: `git commit --no-verify` (권장 X)

### 직원 대상 배포
- **문서 폴더 프로토콜**: [docs/employee-deployment/itgfolder-install.md](docs/employee-deployment/itgfolder-install.md)
  - 각 직원 PC에 `.reg` + `.vbs` 설치 필요
  - Google Drive Desktop 사전 설치 요구
  - 미설치 시 서비스가 폴더 링크 클릭에 응답 X (서버 로그에 아무 흔적 없음)

### 관련 문서
- [README.md](README.md) — 시스템 개요 및 기능
- [SETUP.md](SETUP.md) — 개발 환경 셋업
- [GOOGLE_OAUTH_SETUP.md](GOOGLE_OAUTH_SETUP.md) — Google OAuth 상세
- [dashboard/프로젝트_진행현황_및_다음단계_계획.md](dashboard/프로젝트_진행현황_및_다음단계_계획.md) — Phase 이력 및 로드맵
- [dashboard/다음_세션_시작_가이드.md](dashboard/다음_세션_시작_가이드.md) — 새 세션 시작 시 체크리스트

## 11. 연락처 / 외부 자료

- 슬랙 앱 관리: https://api.slack.com/apps
- Google Cloud Console: https://console.cloud.google.com
- 카카오 디벨로퍼스: https://developers.kakao.com
- 채널톡 개발자: https://developers.channel.io
- DDNS (도메인): pm.itg-aircon.com

---
**마지막 업데이트**: 2026-07-03 (Phase 5-G 완료 반영)
