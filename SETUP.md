# 개발 환경 셋업 가이드

로컬 개발 시작용. 프로덕션 운영/서버 마이그레이션은 [OPERATIONS.md](OPERATIONS.md) 참고.

## 사전 요구사항

- Python 3.9+
- Node.js 18+ (프론트엔드 빌드용)
- Docker Desktop (Redis 컨테이너용)
- Google Cloud Console 프로젝트 (Sheets/Drive/Calendar/OAuth API 활성화)
- Slack 워크스페이스 관리자 권한 (봇 등록용, 개발 참여 시)

## 1. 저장소 복제 및 가상환경

```powershell
git clone <repo-url>
cd "Claude Project"

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
```

## 2. 프론트엔드 의존성 (Vite)

```powershell
cd dashboard
npm install
cd ..
```

## 3. 필수 인증 파일

프로젝트 루트에 배치:

| 파일 | 용도 | 획득 방법 |
|---|---|---|
| `credentials.json` | Google 서비스 계정 (Sheets/Drive 접근) | Google Cloud Console → IAM → 서비스 계정 → JSON 키 |
| `token.json` | Gmail OAuth 토큰 | 최초 실행 시 자동 생성 (아래 설명) |
| `google_calendar_client_secret.json` | Calendar OAuth | `dashboard/` 폴더에 배치 |

모두 `.gitignore` 처리됨. **절대 커밋 금지.**

## 4. 환경 변수 (.env)

```powershell
Copy-Item .env.example .env
notepad .env
```

**최소 채워야 하는 값** (개발용):

```env
# Flask
FLASK_ENV=development
SECRET_KEY=<랜덤 64자>

# Google Sheets (프로젝트 관리 + 리드 관리)
GOOGLE_SHEET_ID=<공사 현황 시트 ID>
GOOGLE_SHEET_NAME=공사 현황의 사본
ONLINE_LEADS_SHEET_ID=<리드 관리 시트 ID>
ONLINE_LEADS_SHEET_NAME=리드 관리

# Google OAuth (로그인)
GOOGLE_OAUTH_ALLOWED_DOMAIN=itg-aircon.com
GOOGLE_OAUTH_REDIRECT_URIS=http://localhost:5000/auth/callback

# Redis
REDIS_HOST=localhost
REDIS_PORT=6379

# 관리자 (감사 로그·알림 대상)
ADMIN_EMAILS=your@email.com
```

**슬랙/채널톡/카카오 통합 등을 개발하려면** 추가 필요 (자세한 목록은 [OPERATIONS.md 3장](OPERATIONS.md) 참조):
- `SLACK_BOT_TOKEN`, `SLACK_SIGNING_SECRET` (main bot)
- `SLACK_VISIT_BOT_TOKEN`, `SLACK_VISIT_SIGNING_SECRET` (방문봇)
- `SLACK_PROJECT_BOT_TOKEN`, `SLACK_PAYMENT_BOT_TOKEN` (공사확정/수금봇)
- `SLACK_LIST_WEBHOOK_URL`, `SLACK_VISIT_COMPLETE_WEBHOOK_URL` (워크플로 트리거)
- `CHANNELTALK_ACCESS_KEY`, `CHANNELTALK_ACCESS_SECRET`
- `KAKAO_REST_API_KEY`
- `GOOGLE_DRIVE_VISIT_FOLDER_ID`, `GOOGLE_DRIVE_WINDOWS_BASE_PATH`

## 5. Redis 실행 (Docker)

```powershell
docker run -d --name itg-redis -p 6379:6379 -v redis-data:/data redis:alpine
```

또는 `docker-compose.yml` 사용:
```powershell
docker-compose up -d redis
```

## 6. 최초 실행

프론트엔드 빌드:
```powershell
cd dashboard
npm run build
cd ..
```

Flask 실행:
```powershell
python app.py
```

또는:
```powershell
python -m flask --app app run --port 5000 --debug
```

브라우저: http://localhost:5000

## 7. 개발 워크플로

### 프론트엔드 개발 중일 때
```powershell
cd dashboard
npm run dev
```
Vite 개발 서버가 별도 포트에 뜸 (HMR 지원).

### 프론트엔드 프로덕션 빌드
```powershell
cd dashboard
npm run build:prod
```
- lint → build → manifest verify

### 테스트 실행
```powershell
pytest tests/
```

## 트러블슈팅

- **`ModuleNotFoundError`**: 가상환경 활성화 확인 (`.venv\Scripts\activate`)
- **Google Sheets 403**: 서비스 계정을 해당 시트에 편집자 권한으로 공유했는지 확인
- **Redis 연결 실패**: `docker ps`로 컨테이너 실행 확인
- **Slack 서명 검증 실패**: `.env`의 `SLACK_*_SIGNING_SECRET`이 Slack 앱 설정과 일치하는지 확인
- **포트 5000 충돌**: `.env`의 `PORT=` 변경 또는 기존 프로세스 종료

## 참고 문서

- [OPERATIONS.md](OPERATIONS.md) — 프로덕션 배포 · NSSM 서비스 · 백업 · 인증서 · Slack 웹훅 라우팅
- [GOOGLE_OAUTH_SETUP.md](GOOGLE_OAUTH_SETUP.md) — Google OAuth 상세 설정
- [docs/employee-deployment/itgfolder-install.md](docs/employee-deployment/itgfolder-install.md) — 직원 대상 폴더 프로토콜 배포
- [README.md](README.md) — 시스템 개요 및 기능 목록
