# Google OAuth 설정 가이드

구글 OAuth를 사용하여 @itg-aircon.com 도메인 계정만 접속할 수 있도록 설정하는 방법입니다.

## 1. Google Cloud Console 설정

### 1.1 프로젝트 생성 또는 선택
1. [Google Cloud Console](https://console.cloud.google.com/) 접속
2. 기존 프로젝트 선택하거나 새 프로젝트 생성

### 1.2 OAuth 2.0 클라이언트 ID 생성
1. **API 및 서비스** → **사용자 인증 정보** 메뉴 이동
2. **사용자 인증 정보 만들기** → **OAuth 클라이언트 ID** 선택
3. 애플리케이션 유형: **웹 애플리케이션** 선택
4. 이름: `IT Global Dashboard` (원하는 이름)
5. **승인된 리디렉션 URI** 추가:
   ```
   http://localhost:5000/auth/callback
   https://yourdomain.com/auth/callback  (운영 서버가 있는 경우)
   ```
6. **만들기** 클릭
7. **클라이언트 ID**와 **클라이언트 보안 비밀** 복사

### 1.3 OAuth 동의 화면 설정 (필수)
1. **OAuth 동의 화면** 메뉴 이동
2. 사용자 유형: **내부** 선택 (G Suite/Google Workspace 사용하는 경우)
3. 앱 정보 입력:
   - 앱 이름: `IT Global 프로젝트 관리`
   - 사용자 지원 이메일: 본인 이메일
   - 개발자 연락처 정보: 본인 이메일
4. **저장 후 계속** 클릭
5. 범위는 기본값 사용 (이메일, 프로필, openid)
6. 테스트 사용자에 @itg-aircon.com 계정들 추가

## 2. 애플리케이션 설정

### 방법 1: 환경 변수 사용 (권장)
`.env` 파일에 다음 추가:
```env
GOOGLE_OAUTH_CLIENT_ID=발급받은_클라이언트_ID.apps.googleusercontent.com
GOOGLE_OAUTH_CLIENT_SECRET=발급받은_클라이언트_보안_비밀
```

### 방법 2: JSON 파일 사용
1. `dashboard/google_oauth_credentials.json` 파일 생성
2. 다음 형식으로 작성:
```json
{
  "web": {
    "client_id": "발급받은_클라이언트_ID.apps.googleusercontent.com",
    "client_secret": "발급받은_클라이언트_보안_비밀",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "redirect_uris": [
      "http://localhost:5000/auth/callback",
      "https://yourdomain.com/auth/callback"
    ]
  }
}
```

## 3. 테스트

1. 서버 실행: `python dashboard/app.py`
2. 브라우저에서 `http://localhost:5000` 접속
3. **"회사 구글 계정으로 로그인"** 버튼 클릭
4. @itg-aircon.com 계정으로 로그인 시도

## 4. 운영 환경 설정

### 도메인 등록
운영 서버가 있는 경우:
1. Google Cloud Console에서 승인된 리디렉션 URI에 실제 도메인 추가
2. `google_oauth.py`의 `redirect_uri` 설정 확인

### 보안 강화
- 클라이언트 보안 비밀은 절대 공개 저장소에 올리지 마세요
- 운영 환경에서는 HTTPS 사용 필수
- OAuth 동의 화면을 "내부"로 설정하여 조직 내부만 접근 가능

## 5. 사용자 관리

- **신규 사용자**: @itg-aircon.com 계정으로 처음 로그인 시 자동으로 `viewer` 권한으로 등록
- **권한 변경**: 관리자가 사용자 관리 메뉴에서 `editor` 또는 `admin` 권한 부여
- **도메인 제한**: @itg-aircon.com이 아닌 계정은 자동으로 접근 차단

## 문제 해결

### 오류 메시지별 해결 방법
- `구글 로그인이 설정되지 않았습니다`: OAuth 설정 확인
- `@itg-aircon.com 도메인 계정만 접속 가능`: 회사 계정으로 로그인 필요
- `구글 로그인 중 오류가 발생했습니다`: 클라이언트 ID/시크릿 또는 리디렉션 URI 확인

### 로그 확인
서버 실행 시 터미널에서 OAuth 관련 로그 확인:
```
✅ 새 Google 사용자 등록: 홍길동 (hong@itg-aircon.com) - viewer 권한
Google OAuth 로그인 성공: hong@itg-aircon.com
```