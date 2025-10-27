# ITGlobal Dashboard - 개발 환경 설정 가이드

## 현재 작업 상태 (2024-09-22)

### 최근 완료된 작업
1. **배지 시스템 현대화**: 해시 기반 색상 할당으로 일관성 확보
2. **테이블 시스템 개선**: 레거시 칼럼 폭을 모던 CSS Grid에 적용
3. **헤더 정렬 조정**: 텍스트 중앙정렬 + 화살표 우측 배치 (진행중)
4. **DataTables 통합**: 정렬 기능과 모던 반응형 디자인 통합

### 현재 이슈
- 테이블 헤더 정렬이 완전히 해결되지 않음 (브라우저 캐시 문제 가능성)
- 화살표와 텍스트 위치 조정 필요

## 개발 환경 요구사항

### Python 환경
```bash
python 3.8+
pip install -r requirements.txt
```

### Node.js 환경
```bash
node 16+
npm install
```

### 필수 환경변수
```bash
# Google OAuth 설정
GOOGLE_CLIENT_ID=your_client_id
GOOGLE_CLIENT_SECRET=your_client_secret

# 기본 포트 설정
PORT=5000
FLASK_ENV=development
FLASK_DEBUG=1
```

## 개발 서버 실행

### 방법 1: 통합 스크립트
```bash
python dev-start.py
```

### 방법 2: 수동 실행
```bash
# Python 서버
python run_server.py

# Vite 개발 서버 (별도 터미널)
cd dashboard
npm run dev
```

### 방법 3: 배치 파일 (Windows)
```bash
dev-start.bat
```

## 빌드 및 배포

### 프론트엔드 빌드
```bash
cd dashboard
npm run build
```

### CSS/JS 변경 후 빌드 필수
- CSS 변경 시: `npm run build` 실행
- 브라우저 캐시 클리어: `Ctrl + F5`

## 프로젝트 구조

```
itglobal/
├── dashboard/
│   ├── src/
│   │   ├── css/
│   │   │   ├── components/
│   │   │   │   ├── project-table.css      # 테이블 스타일
│   │   │   │   └── performance-optimizations.css
│   │   │   └── design-system/
│   │   │       ├── badges.css            # 배지 스타일
│   │   │       └── tokens.css            # CSS 변수
│   │   └── js/
│   │       └── components/
│   │           ├── LegacyBadgeSystem.js  # 배지 로직
│   │           └── ProjectTable.js       # 테이블 로직
│   ├── templates/
│   │   ├── modern_project_list.html      # 메인 테이블 페이지
│   │   └── components/
│   ├── static/dist/                      # 빌드된 파일들
│   └── package.json
├── data/                                 # 데이터 파일들
└── run_server.py                         # 메인 서버
```

## 주요 기술 스택

### 백엔드
- Flask (Python)
- Google Sheets API
- Google Drive API
- SQLite

### 프론트엔드
- Vite (빌드 도구)
- Vanilla JavaScript (ES6+)
- CSS Grid + Flexbox
- DataTables.js
- Bootstrap 5

### 디자인 시스템
- CSS 변수 기반 토큰 시스템
- 컴포넌트 기반 CSS 아키텍처
- 반응형 Container Queries

## 다음 작업 계획

### 우선순위 1: 헤더 정렬 완료
1. 브라우저 캐시 문제 해결
2. 텍스트 중앙정렬 + 화살표 우측 배치 확정
3. 모든 브라우저에서 테스트

### 우선순위 2: 성능 최적화
1. 가상 스크롤링 구현 검토
2. 테이블 렌더링 성능 개선
3. 모바일 최적화

### 우선순위 3: 사용자 경험 개선
1. 로딩 상태 개선
2. 에러 처리 강화
3. 접근성 향상

## 문제 해결 가이드

### 빌드 문제
```bash
# 캐시 클리어
rm -rf node_modules/.vite
npm run build

# 의존성 재설치
rm -rf node_modules package-lock.json
npm install
```

### 서버 연결 문제
```bash
# 포트 충돌 확인
netstat -ano | findstr :5000

# 프로세스 종료
taskkill /PID [PID번호] /F
```

### CSS 변경이 반영되지 않는 경우
1. `npm run build` 실행
2. 브라우저에서 `Ctrl + F5` (하드 리프레시)
3. 개발자 도구에서 네트워크 탭의 캐시 비활성화

## Git 워크플로우

### 기본 작업 흐름
```bash
git pull origin main
# 작업 수행
git add .
git commit -m "작업 내용"
git push origin main
```

### 브랜치 전략
- `main`: 메인 개발 브랜치
- 필요시 기능별 브랜치 생성

## 연락처 및 참고자료

### 문서
- `LEGACY_VS_MODERN_VALIDATION_GUIDE.md`: 레거시/모던 비교 가이드
- `FINAL_VALIDATION_REPORT.md`: 최종 검증 보고서
- `critical_findings_summary.md`: 중요 이슈 요약

### 개발 도구
- `dev-start.py`: 통합 개발 서버 실행
- `css_feature_flag_test.py`: CSS 기능 테스트
- `verification_checklist.md`: 검증 체크리스트

---
*마지막 업데이트: 2024-09-22*
*작성자: Claude Code Assistant*