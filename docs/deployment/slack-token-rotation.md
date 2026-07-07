# Slack 토큰 회전 가이드 (관리자 액션)

## 왜 필요한가

과거 세션에서 노출 이력이 있었던 슬랙 봇 토큰들을 정기적으로 회전(재발급)해야 함. `[[project-system-audit-checklist]]` Task #23의 후속.

회전 대상 (`.env` 기준):
- `SLACK_BOT_TOKEN` — 메인 봇
- `SLACK_VISIT_BOT_TOKEN` — 방문 관리 봇
- `SLACK_PRICE_BOT_TOKEN` — 견적 봇
- `SLACK_PAYMENT_BOT_TOKEN` — 수금 관리 봇
- (신규 봇 추가 시 함께)

Signing secret은 봇 앱 자체 회전 시에만 바꿈. 토큰만 회전이 일반적.

## 회전 절차 (봇 1개당 3분)

### 1. Slack API 사이트에서 새 토큰 발급

1. `api.slack.com/apps` 접속
2. 대상 봇 앱 선택 (예: "ITG 봇")
3. **좌측 → OAuth & Permissions**
4. **Bot User OAuth Token** 섹션 → **Reissue token** 클릭
5. 새 토큰(`xoxb-...`) 복사

**주의**: Reissue 하면 옛 토큰은 즉시 무효화됨. 아래 서버 반영 전에는 봇 응답 안 함. 재발급 → 서버 반영을 30초 이내 진행.

### 2. `.env`에 새 토큰 반영

```powershell
# 편집기로 .env 열고
notepad "C:\Users\SECOM\Desktop\ITG-Project\Claude Project\.env"

# 해당 라인 새 토큰으로 교체
# 예: SLACK_BOT_TOKEN=xoxb-옛값 → SLACK_BOT_TOKEN=xoxb-새값
```

### 3. 서비스 재시작

```powershell
schtasks /Run /TN "Restart-ITGFlask"
```

25초 후 서비스 정상 응답 (health check로 검증):
```powershell
Invoke-WebRequest -Uri "https://pm.itg-aircon.com/api/health" -UseBasicParsing
```

### 4. 검증

Slack 봇 채널에서 명령어 하나 실행 (예: `/전화 010-1234-5678`) → 정상 응답 확인.

## 여러 봇 한 번에 회전 시

1. 모든 봇 새 토큰 미리 발급 (아직 서버 반영 X)
2. `.env` 여러 라인 함께 수정
3. 서비스 한 번 재시작

## 이력 남기기

회전 후 이 파일 하단에 기록:

| 일시 | 회전 대상 | 담당 |
|---|---|---|
| 예: 2026-07-07 | SLACK_BOT_TOKEN, SLACK_VISIT_BOT_TOKEN | 고광일 |
