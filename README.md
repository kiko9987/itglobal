# 냉난방기 설치 공사 관리 시스템

구글 시트의 웹 GUI 버전 + 실시간 대시보드 + 자동 알림 시스템

**영업사원들을 위한 쉬운 데이터 입력과 관리자를 위한 통합 대시보드**

## 🌟 주요 기능

### 🖥️ **웹 기반 데이터 입력 시스템 (구글 시트 GUI)**
- **직관적인 입력 폼**: 영업사원들도 쉽게 사용할 수 있는 웹 인터페이스
- **자동 데이터 검증**: 필수 항목 체크 및 형식 검증
- **실시간 계산**: 부가세, 미수금, 마진율 자동 계산
- **프로젝트 코드 자동 생성**: 지역별 순차 번호 자동 부여
- **구글 시트 실시간 동기화**: 입력 즉시 구글 시트에 반영

### 📋 **프로젝트 관리 시스템**
- **전체 프로젝트 목록**: 필터링, 검색, 정렬 기능
- **상세 정보 보기**: 각 프로젝트의 완전한 정보 조회
- **수정 및 삭제**: 웹에서 직접 데이터 편집
- **CSV 내보내기**: 필터된 데이터 엑셀 다운로드
- **상태별 관리**: 대기/진행중/완료 상태 추적

### 📊 **실시간 대시보드**
- **전체 현황 요약**: 총 프로젝트 수, 매출액, 미수금, 회수율
- **월별 매출 분석**: 시각적 차트로 매출 추이 확인
- **지역별 성과**: 영업사원별 매출 및 성과 분석
- **브랜드별 분석**: 냉난방기 브랜드별 매출 및 마진율
- **미수금 현황**: 미수금 현황 및 기간별 분류
- **실시간 업데이트**: WebSocket으로 실시간 데이터 반영

### 📧 **스마트 알림 시스템**
- **빈 칸 알림**: 영업사원별 미입력 항목 자동 체크 및 이메일 발송
- **일일 보고서**: 관리자용 일일 현황 요약 보고서
- **슬랙 통합**: 팀 채널로 실시간 알림
- **스케줄 자동화**: 매일 오전 9시, 오후 6시 자동 실행
- **개인별 맞춤 알림**: 각 영업사원에게 해당하는 누락 항목만 전송

## 🚀 빠른 시작

### 1. 필수 요구사항
- Python 3.8+
- 구글 계정 및 Google Cloud Console 프로젝트
- Gmail 계정 (알림용)

### 2. 설치 및 설정

#### 라이브러리 설치
```bash
pip install -r requirements.txt
```

#### 구글 API 설정
1. [Google Cloud Console](https://console.cloud.google.com/)에서 프로젝트 생성
2. Google Sheets API 활성화
3. 서비스 계정 생성 후 JSON 키 다운로드
4. `credentials.json` 파일을 프로젝트 루트에 저장

#### 환경 변수 설정
`.env` 파일을 생성하고 다음 정보를 입력:

```env
# 구글 API 설정
GOOGLE_SHEET_ID=your_google_sheet_id_here
GOOGLE_SHEET_NAME=공사 현황

# Flask 설정
FLASK_SECRET_KEY=your_secret_key_here
DEBUG=True

# 이메일 설정 (Gmail SMTP)
EMAIL_HOST=smtp.gmail.com
EMAIL_PORT=587
EMAIL_USERNAME=your_email@gmail.com
EMAIL_PASSWORD=your_app_password

# 관리자 이메일 (쉼표로 구분)
ADMIN_EMAILS=admin1@company.com,admin2@company.com

# 영업사원 이메일 (JSON 형태)
SALES_EMAILS={"양곡":"yangkok@company.com","종로":"jongno@company.com"}

# 슬랙 웹훅 URL (선택사항)
SLACK_WEBHOOK_URL=https://hooks.slack.com/services/your/webhook/url

# 알림 설정
NOTIFICATION_INTERVAL_HOURS=24
MISSING_FIELDS_THRESHOLD=3
```

### 3. 실행

#### 통합 시스템 시작
```bash
python start_dashboard.py
```

시스템이 시작되면 다음 URL에서 이용 가능합니다:
- **메인 대시보드**: `http://localhost:5000`
- **프로젝트 관리**: `http://localhost:5000/projects`  
- **새 프로젝트 등록**: `http://localhost:5000/project/new`

#### 알림 시스템 시작 (별도 터미널)
```bash
python start_notifications.py
```

## 📁 프로젝트 구조

```
프로젝트/
├── dashboard/
│   ├── app.py                 # Flask 웹 애플리케이션
│   ├── templates/
│   │   └── dashboard.html     # 대시보드 HTML 템플릿
│   └── utils/
│       ├── google_sheets.py   # 구글 시트 연동
│       ├── data_analyzer.py   # 데이터 분석 모듈
│       └── notification_system.py # 알림 시스템
├── data/                      # 엑셀 데이터 파일 (폴백용)
├── start_dashboard.py         # 대시보드 시작 스크립트
├── start_notifications.py     # 알림 시스템 시작 스크립트
├── requirements.txt           # Python 패키지 의존성
├── .env                       # 환경 변수 설정
├── credentials.json          # 구글 API 자격증명
└── README.md                  # 프로젝트 문서
```

## 🔧 상세 설정

### 구글 시트 설정
1. 구글 시트에서 '공사 현황' 시트 생성
2. 필수 컬럼: 프로젝트 코드, 사업자, 담당자, 거래처, 현장 주소, 공사 내용, 공사 시작, 총액 2, 미수금 등
3. 시트 공유 설정에서 서비스 계정 이메일에 뷰어 권한 부여

### 이메일 설정
Gmail 사용 시:
1. Google 계정에서 2단계 인증 활성화
2. 앱 비밀번호 생성 후 `EMAIL_PASSWORD`에 입력

### 슬랙 설정 (선택사항)
1. 슬랙 워크스페이스에서 Incoming Webhook 생성
2. 웹훅 URL을 `SLACK_WEBHOOK_URL`에 입력

## 📈 대시보드 기능 상세

### 메인 대시보드
- **요약 카드**: 프로젝트 수, 총매출, 미수금, 회수율
- **월별 매출 차트**: 라인 차트로 매출 추이 시각화
- **지역별 매출**: 도넛 차트로 지역별 비중 표시
- **미수금 현황**: 담당자별 미수금 바 차트
- **브랜드 분석**: 파이 차트로 브랜드별 매출 비중

### 실시간 기능
- **WebSocket 연결**: 실시간 데이터 업데이트
- **자동 새로고침**: 10분마다 자동으로 데이터 갱신
- **연결 상태 표시**: 실시간 연결 상태 확인

## 🔔 알림 시스템 상세

### 빈 칸 알림
- **자동 검증**: 중요 필드의 누락 데이터 자동 감지
- **개인별 알림**: 각 영업사원에게 개별 이메일 발송
- **우선순위**: 누락 항목 수에 따른 긴급도 표시
- **프로젝트 목록**: 구체적인 누락 프로젝트 코드 제공

### 관리자 보고서
- **일일 요약**: 전체 현황 요약 보고서
- **미수금 현황**: 미수금 상세 현황
- **팀 성과**: 지역별 성과 요약

## 🛠️ 커스터마이징

### 새로운 차트 추가
1. `data_analyzer.py`에 분석 함수 추가
2. `app.py`에 API 엔드포인트 추가
3. `dashboard.html`에 차트 컴포넌트 추가

### 알림 규칙 변경
1. `.env`에서 `MISSING_FIELDS_THRESHOLD` 값 조정
2. `notification_system.py`에서 체크 로직 수정

### 새로운 영업사원 추가
1. `.env`의 `SALES_EMAILS`에 추가
2. 구글 시트에 새로운 담당자 지역 추가

## 🔍 문제 해결

### 일반적인 문제들

**1. 구글 시트 연결 실패**
- `credentials.json` 파일 확인
- 구글 시트 ID 정확성 확인
- 서비스 계정 권한 확인

**2. 이메일 발송 실패**
- Gmail 앱 비밀번호 확인
- 2단계 인증 활성화 확인
- 방화벽 설정 확인

**3. 데이터 로드 오류**
- 구글 시트 컬럼명 일치 확인
- 데이터 타입 확인
- 네트워크 연결 확인

### 로그 확인
- 대시보드 로그: `dashboard.log`
- 알림 시스템 로그: `notifications.log`

## 🔒 보안 고려사항

- `.env` 파일을 버전 관리에 포함하지 마세요
- `credentials.json` 파일을 안전하게 보관하세요
- 프로덕션 환경에서는 `DEBUG=False` 설정
- 정기적으로 앱 비밀번호 갱신

## 📞 지원

문제가 발생하면 다음을 확인해보세요:
1. 요구사항 체크 스크립트 실행
2. 로그 파일 확인
3. .env 설정 재확인

## 🔄 업데이트

새로운 기능이나 버그 수정을 위해 정기적으로 업데이트하세요:
```bash
git pull origin main
pip install -r requirements.txt
```

---

💡 **팁**: 처음 설정할 때는 대시보드를 먼저 실행해서 정상 동작을 확인한 후 알림 시스템을 시작하는 것을 추천합니다.