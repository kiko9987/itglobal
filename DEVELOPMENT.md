# 🚀 IT Global 개발 환경 가이드

## 빠른 시작

### 방법 1: 원클릭 실행 (권장)
```bash
# 배치 파일 실행 (Windows)
dev-start.bat

# 또는 Python 스크립트 직접 실행
python dev-start.py
```

### 방법 2: 수동 실행
```bash
# Flask 서버 (터미널 1)
python run_server.py

# Vite 개발서버 (터미널 2)
cd dashboard
npm run dev
```

## 🎯 자동화된 기능

### 포트 자동 관리
- **Flask**: 기본 5000번, 사용 중이면 자동으로 다음 포트 할당
- **Vite**: 기본 5173번, 사용 중이면 자동으로 다음 포트 할당
- **프록시 설정 자동 업데이트**: Vite 설정에서 Flask 포트 자동 연동

### 브라우저 자동 열기
- Chrome으로 개발 환경 자동 접속
- 캐시 무시 및 보안 설정 최적화

### 실시간 로그 모니터링
- Flask와 Vite 서버 로그 통합 출력
- 색상 구분으로 가독성 향상

## 🔥 핫리로드 개발

### Frontend 개발
- **JavaScript/CSS 파일 수정** → 즉시 브라우저 반영
- **컴포넌트 수정** → 상태 유지하며 업데이트
- **새로고침 불필요**

### Backend 개발
- **Python 파일 수정** → Flask 서버 자동 재시작
- **설정 파일 변경** → 서버 재시작 필요

## 📁 주요 파일 구조

```
itglobal/
├── dev-start.py          # 통합 개발 스크립트
├── dev-start.bat         # Windows 배치 파일
├── run_server.py         # Flask 서버 실행
├── dashboard/
│   ├── vite.config.js    # Vite 설정 (자동 포트 연동)
│   ├── package.json      # npm 스크립트
│   └── src/
│       ├── js/           # JavaScript 모듈
│       └── css/          # 스타일시트
└── DEVELOPMENT.md        # 이 파일
```

## 🛠 개발 팁

### 포트 충돌 해결
```bash
# 사용 중인 포트 확인
netstat -ano | findstr :5000

# 프로세스 종료
taskkill /PID <프로세스ID> /F
```

### 캐시 문제 해결
- **브라우저**: Ctrl+Shift+R (강력 새로고침)
- **Vite**: 서버 재시작
- **Flask**: Python 파일 수정 시 자동 재시작

### 디버깅
```bash
# Flask 디버그 모드 활성화
set FLASK_DEBUG=1
python run_server.py

# Vite 디버그 로그
npm run dev -- --debug
```

## ⚡ 성능 최적화

### 개발 환경
- **Vite HMR**: 0.1초 내 반영
- **Flask Auto-reload**: 1초 내 재시작
- **프록시 최적화**: API 요청 지연 최소화

### 빌드 환경
```bash
# 프로덕션 빌드
cd dashboard
npm run build

# 빌드 검증
npm run preview
```

## 🐛 문제 해결

### 공통 문제
1. **포트 충돌**: `dev-start.py`가 자동으로 다른 포트 할당
2. **404 에러**: Vite 프록시 설정 확인
3. **핫리로드 실패**: 브라우저 캐시 삭제

### 연락처
- **개발 문의**: [담당자 이메일]
- **버그 리포트**: Issues 탭 활용

## 📊 환경 상태 확인

### 서버 상태
```bash
# 실행 중인 서버 확인
python -c "
import socket
ports = [5000, 5173]
for port in ports:
    s = socket.socket()
    try:
        s.connect(('localhost', port))
        print(f'포트 {port}: 사용 중')
    except:
        print(f'포트 {port}: 사용 가능')
    finally:
        s.close()
"
```

### 의존성 확인
```bash
# Python 패키지
pip list

# Node.js 패키지
cd dashboard
npm list
```

---

**🎉 즐거운 개발되세요!**