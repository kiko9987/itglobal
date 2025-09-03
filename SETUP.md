# ITG Sheets Dashboard 설정 가이드

## 집에서 작업 시작하기

### 1. 프로젝트 클론
```bash
git clone https://github.com/[YOUR_USERNAME]/itg-sheets-dashboard.git
cd itg-sheets-dashboard
```

### 2. 가상환경 설정
```bash
# Python 가상환경 생성
python -m venv venv

# 가상환경 활성화
# Windows
venv\Scripts\activate
# macOS/Linux  
source venv/bin/activate

# 의존성 설치
pip install -r requirements.txt
```

### 3. 환경 설정 파일 생성
```bash
# .env.example을 .env로 복사
cp .env.example .env
```

### 4. 필수 파일 설정

#### A. Google API 자격증명 파일 (credentials.json)
- Google Cloud Console에서 서비스 계정 JSON 키 다운로드
- 프로젝트 루트에 `credentials.json`으로 저장
- ⚠️ 이 파일은 Git에 커밋되지 않음 (.gitignore에 포함)

#### B. .env 파일 수정
```bash
# 중요: 실제 값으로 변경 필요
GOOGLE_SHEET_ID=your-actual-google-sheet-id
SECRET_KEY=your-random-secret-key
# 기타 필요한 설정들...
```

### 5. 애플리케이션 실행
```bash
python start_dashboard.py
```

## 동기화 작업 흐름

### 작업 시작 전 (pull)
```bash
git pull origin main
```

### 작업 완료 후 (push)
```bash
git add .
git commit -m "작업 내용 설명"
git push origin main
```

## 브랜치 작업 (권장)
```bash
# 새 기능 개발 시
git checkout -b feature/새기능명
git add .
git commit -m "새 기능 추가"
git push origin feature/새기능명

# GitHub에서 Pull Request 생성
```

## 주의사항
- `.env` 파일과 `credentials.json`은 Git에 포함되지 않음
- 각 환경에서 이 파일들을 직접 설정해야 함
- 민감한 정보를 GitHub에 올리지 않도록 주의
- 작업 전후로 항상 `git pull`과 `git push` 실행

## 문제 해결
- 의존성 오류: `pip install -r requirements.txt` 재실행
- 권한 오류: Google API 서비스 계정 권한 확인
- 포트 충돌: `.env`에서 PORT 변경