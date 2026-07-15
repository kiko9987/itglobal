# ITG-Aircon 관리 시스템

냉난방기 설치/A·S/수금 전체 라이프사이클을 관리하는 통합 대시보드 + 자동화 봇.

**핵심**: 리드 인입 → 상담 → 방문 → 견적 → 공사 확정 → 시공 → 수금 → 사후관리 전 과정이 Google Sheets를 단일 진실 원천으로 삼고, Flask 대시보드 + 4개 Slack 봇 + 채널톡·Gmail 인입 통합으로 운영.

## 🌟 주요 기능

### 📊 관리 대시보드 (pm.itg-aircon.com)
- **프로젝트 관리**: 시트-스타일 인라인 편집, 실시간 필터·검색·정렬
- **리드 관리**: 온라인/전화/거래처 리드 통합 관리, 프로젝트 자동 연동
- **통계·차트**: 월별 매출, 담당자별 성과, 미수금 현황
- **감사 로그**: 모든 편집 이력 추적 (사용자·시각·이전/이후 값)
- **권한 관리**: Admin/Editor/Viewer 3단계 롤

### 🤖 Slack 봇 (6개 분리 앱)
| 봇 | 담당 |
|---|---|
| **온라인 문의 알림봇** | 홈페이지/카카오톡/채널톡/당근/전화 문의 인입 → 카드 발송 → 상담 모달 |
| **방문 일정 알림봇** | 방문 예약 카드 + 방문일 수정/완료/취소 액션 + 현장 사진 자동 저장 |
| **공사 현황 알림봇** | 리드 → 프로젝트 등록 모달 (`/공사확정`) + 세금계산서 발행 요청 모달 |
| **수금 관리 알림봇** | 입금 메모 자동 감지 → 미수금 카드 → 확정 액션 |
| **A/S 관리 알림봇** | `/as` 슬래시 + 3단계 모달 (요청/접수/완료) — 시공자 pre-fill |
| **세금계산서 관리 알림봇** | `#영업_관리` 카드 발송 + 스레드 이미지/PDF 자동 감지 → 자동 완료 |

**진입점**: 각 봇마다 별도 endpoint (`/slack/events`, `/slack/project-events`, `/slack/visit-events`, `/slack/as-events`, `/slack/invoice-events`).

### 🔄 자동 인입 파이프라인
| 채널 | 흐름 |
|---|---|
| 홈페이지 문의 | Gmail API → 파싱 → 시트 등록 → Slack 카드 |
| 카카오톡/채널톡 | 채널톡 Developer API → 시트 등록 → Slack 카드 (매니저 답변 forward) |
| 전화 문의 | Slack 워크플로 → 시트 등록 → sync 폴링 → Slack 카드 |
| 당근마켓 | 당근 시트 → 자동 sync → 메인 시트 → Slack 카드 |
| 거래처/기타 방문 | Slack `/방문` 슬래시 명령 → 방문 List 등록 |

### 📁 Google Drive 자동 폴더 파이프라인
- 방문 카드 thread에 현장 사진 첨부 → 자동으로 프로젝트 폴더 생성
- 폴더명 규칙: `(담당자이니셜) 주소 YY.MM.DD`
- 폴더 ID가 리드 시트 P열에 저장 → 신규 프로젝트 등록 시 자동 채움
- 대시보드에서 폴더 링크 클릭 시 각자 자기 PC 탐색기로 열림 ([itgfolder://](docs/employee-deployment/itgfolder-install.md) 프로토콜)
- **대량 배치 지원** — Slack UI 한 번당 10장 상한이라 여러 답글로 나눠 올려도 같은 폴더에 이어짐. 4장 이상은 진행 답글 (`⏳ K/N`) 로 실시간 표시

### 💰 세금계산서 발행 요청/완료 흐름
- 공사 확정 카드 하단 `[💰 계산서 요청]` 버튼 → 프로젝트 정보 pre-fill 모달 → `#영업_관리` 채널에 **세금계산서 관리 알림봇** 아바타로 카드 발송
- Submit 시 이중 병렬 검증: (1) 사업자등록증 첨부, (2) 시트 S열 부가세 체크박스 채움. 미충족 시 modal errors 로 반려
- **자동 완료** — 회계가 스레드에 이미지/PDF 첨부하면 봇이 감지 → 헤더 `🔔 요청` → `✅ 완료`, 첨부 상태 `⬜ 미첨부` → `✅ 첨부됨` 자동 갱신 + 처리자 이니셜 표시. 별도 버튼 클릭 불필요
- 첨부 파일 삭제 감지 시 스레드에 재첨부 안내

### 📝 사업자등록증 OCR (Vision API)
- 공사확정 카드 스레드에 사업자등록증 첨부 → Drive 저장 + Google Cloud Vision API 로 법인명/상호 자동 추출
- 정규식으로 `법 인 명 (단 체 명) : XXX` / `상 호 : XXX` 매치 → 스레드에 `📝 OCR 결과 — 사업자명 추정: ...` 안내
- 매니저는 관리 페이지에서 확인 후 프로젝트 사업자명 수정 (오탐 대비 자동 시트 반영 안 함)
- 무료 티어 매월 1,000장, 이후 $1.5/1000장

### 🔔 안정성 안전망
- **고아 리드 감지** — 시트 등록됐지만 슬랙 카드 없는 리드를 5분 주기 스캔·자동 재발송 (Flask 재시작 등 유실 대비)
- **재문의 알림** — 채팅 리드 진행 중 상태에서 재문의 오면 채널에 top-level 알림 카드 추가 (스레드 리플라이는 슬랙 알림 안 뜨는 문제 회피)
- **프로젝트 편집 알림** — 대시보드에서 편집 시 원본 공사 확정 카드가 최신 스냅샷으로 갱신되고, 스레드에 `[프로젝트코드 데이터 수정 알림]` 히스토리 답글
- **Pending 큐 재발송** — 슬랙 SSL 에러 등으로 누락된 카드·방문 알림 5분 주기 자동 복구
- **Sheet write-behind 큐** — 모든 mutation을 Redis 지속 큐에 등록 후 백그라운드 워커가 Google Sheets에 반영. 응답 시간 <300ms, API 지연을 UX와 완전 분리. 3회 재시도 후 데드레터 + 관리자 슬랙 DM.
- **자동 일 백업** — 매일 새벽 03:15 users.db + Redis dump.rdb를 `backup/` 로 스냅샷 (30일 유지)
- **관리자 큐 상태 페이지** — `/admin/queue-status` 실시간 (5초 갱신) pending/processing/failed 카운트 + 실패 op 재시도
- **방문 사진 배치 처리 안정성** — 스레드 대량 첨부 (30~60장) 대응: 파일 수 비례 Redis lock TTL (상한 10분), Drive 429/5xx 지수 백오프 3회, 진행 답글 5장 단위 갱신, 부분 실패는 성공/실패/스킵 카운트 분리 표기

### 🔍 주소 정규화
- Kakao Local API로 도로명/지번 검증
- 시설명·층·호 정보 자동 부착 (verified 주소 뒤에 부가정보)
- 재문의 자동 감지 (같은 주소 다른 리드 매칭)

## 🏗️ 시스템 구성

| 컴포넌트 | 배포 | 비고 |
|---|---|---|
| Flask 백엔드 | Windows NSSM 서비스 (`ITGFlask`) | 포트 5000 |
| Caddy 리버스 프록시 | Windows NSSM 서비스 | 포트 443 HTTPS → 5000 |
| Redis | Docker 컨테이너 | 캐시·락·pending 큐·**Sheet write-behind 큐** |
| Google Sheets | 외부 API | 서비스 계정 인증 |
| Google Drive | 외부 API + Desktop 앱 | 폴더 자동 생성 · 로컬 미러 |
| Slack | Bolt for Python | 4개 봇 앱 |
| 채널톡 | Developer API | 실시간 webhook |
| Gmail | OAuth (홈페이지 메일 처리) | |

**상세 운영 가이드**: [OPERATIONS.md](OPERATIONS.md)

## 📁 프로젝트 구조

```
Claude Project/
├── app.py                        # Flask 진입점
├── dashboard/
│   ├── blueprints/               # Flask 라우트 (블루프린트별 분리)
│   │   ├── projects.py           # 프로젝트 CRUD + 자동 코드 생성
│   │   ├── leads.py              # 리드 관리
│   │   ├── slack_bot.py          # Slack 4봇 통합 핸들러
│   │   ├── slack_helpers.py      # Slack 공통 유틸
│   │   ├── channeltalk.py        # 채널톡 webhook
│   │   ├── channeltalk_helpers.py
│   │   ├── folders.py            # Google Drive 폴더 API
│   │   ├── auth.py, admin.py, users.py
│   │   └── ...
│   ├── services/                 # 비즈니스 로직
│   │   ├── project_service.py    # 프로젝트 시트 로드/캐시
│   │   ├── lead_service.py       # 리드 시트 로드/CRUD
│   │   ├── lead_sync.py          # 당근/전화 워크플로 자동 동기화
│   │   ├── lead_helpers.py       # 전화번호 정규화, 키워드 매핑
│   │   ├── address_resolver.py   # Kakao Local API
│   │   ├── homepage_mail_sync.py # Gmail 파싱
│   │   ├── payment_sync.py       # 수금 메모 감지
│   │   ├── project_slack_notifier.py
│   │   ├── channeltalk_api.py, channeltalk_threads.py
│   │   ├── calendar_service.py, calendar_sync_scheduler.py
│   │   └── sync_scheduler.py     # apscheduler 폴링 관리
│   ├── utils/
│   │   ├── google_sheets.py, google_drive.py
│   │   ├── redis_client.py, smart_cache_manager.py
│   │   └── ...
│   ├── templates/, static/, src/ # UI (Vite 빌드)
│   └── logs/
├── docs/                         # 문서
│   └── employee-deployment/
├── tests/                        # pytest
├── credentials.json              # Google 서비스 계정 (git-ignored)
├── token.json                    # Gmail OAuth 토큰 (git-ignored)
├── .env                          # 환경 변수 (git-ignored)
├── Caddyfile
├── docker-compose.yml            # Redis 컨테이너
└── OPERATIONS.md                 # 운영 가이드
```

## 🚀 시작하기

### 개발 환경 (로컬)
[SETUP.md](SETUP.md) 참고.

### 프로덕션 배포 (회사 서버 PC)
[OPERATIONS.md](OPERATIONS.md) 참고 — NSSM 서비스 등록, Caddy HTTPS, Redis 백업 포함.

### 직원용 폴더 프로토콜 설치
[docs/employee-deployment/itgfolder-install.md](docs/employee-deployment/itgfolder-install.md) 참고 — 각 직원 PC에서 프로젝트 문서 폴더가 탐색기로 열리도록 설정.

## 🔒 보안 · 인증

- Google OAuth (도메인 제한: `@itg-aircon.com`)
- 세션 4시간 타임아웃
- 편집 락 5분 자동 해제
- 모든 시트 편집 감사 로그
- CSRF 토큰 (Flask-WTF)
- `.env`, `credentials.json`, `token.json` git-ignored

## 🧪 테스트

```bash
pytest tests/
```

주요 테스트: `tests/unit/test_slack_channeltalk_core.py` (Slack/채널톡 핵심 파서·모달 로직).

## 📚 관련 문서

- [SETUP.md](SETUP.md) — 개발 환경 셋업
- [OPERATIONS.md](OPERATIONS.md) — 프로덕션 운영 · 서버 마이그레이션 · 백업
- [GOOGLE_OAUTH_SETUP.md](GOOGLE_OAUTH_SETUP.md) — Google OAuth 상세
- [docs/employee-deployment/](docs/employee-deployment/) — 직원 대상 설치 가이드
- [dashboard/프로젝트_진행현황_및_다음단계_계획.md](dashboard/프로젝트_진행현황_및_다음단계_계획.md) — 진행 이력

## 📝 라이선스

내부 사용 전용.
