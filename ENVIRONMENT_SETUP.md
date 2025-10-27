# 개발 환경 설정 가이드

이 문서는 회사와 집에서 동일한 개발 환경을 구축하기 위한 가이드입니다.

## 📋 목차
1. [로컬 전용 파일 (Git에 포함되지 않음)](#로컬-전용-파일)
2. [데이터베이스 설정](#데이터베이스-설정)
3. [환경 변수 설정](#환경-변수-설정)
4. [초기 설정 방법](#초기-설정-방법)
5. [환경 동기화 체크리스트](#환경-동기화-체크리스트)

---

## 🔒 로컬 전용 파일 (Git에 포함되지 않음)

다음 파일들은 `.gitignore`에 포함되어 Git에 커밋되지 않습니다:

### 1. 데이터베이스 파일
```
dashboard/data/dashboard.db          # 사용자, 감사 로그 등
instance/users.db                    # 레거시 사용자 DB
users.db                             # 레거시 사용자 DB
*.db, *.sqlite, *.sqlite3           # 모든 데이터베이스 파일
```

### 2. 인증 및 환경 설정
```
.env                                 # 환경 변수
credentials.json                     # Google API 자격증명
client_secret_*.json                # OAuth 클라이언트 시크릿
```

### 3. 기타
```
*.log                               # 로그 파일
logs/                               # 로그 디렉토리
__pycache__/                        # Python 캐시
node_modules/                       # Node 의존성
```

---

## 💾 데이터베이스 설정

### 데이터베이스 위치
- **경로**: `dashboard/data/dashboard.db`
- **타입**: SQLite3

### 주요 테이블

#### 1. users (사용자 관리)
```sql
CREATE TABLE users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT,
    permission_level TEXT DEFAULT 'viewer',  -- 'viewer', 'editor', 'admin', 'super_admin'
    is_active BOOLEAN DEFAULT 1,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    last_login TIMESTAMP,
    google_id TEXT UNIQUE
);
```

#### 2. audit_logs (감사 로그)
```sql
CREATE TABLE audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_email TEXT NOT NULL,
    user_name TEXT NOT NULL,
    user_role TEXT,
    action TEXT NOT NULL,
    project_code TEXT,
    field_name TEXT,
    old_value TEXT,
    new_value TEXT,
    ip_address TEXT,
    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

### 권한 레벨 설명

| 권한 | 설명 | 프론트엔드 표시 |
|------|------|----------------|
| `viewer` | 읽기 전용 | Viewer |
| `editor` | 편집 가능 | Editor |
| `admin` | 관리자 | Admin |
| `super_admin` | 슈퍼 관리자 (시스템 전체 관리) | Admin |

**참고**: `super_admin`은 백엔드에서만 구별되며, 프론트엔드에서는 `admin`과 동일하게 "Admin"으로 표시됩니다.

### 초기 사용자 설정 방법

#### 옵션 1: Python으로 직접 추가
```python
import sqlite3
from werkzeug.security import generate_password_hash

conn = sqlite3.connect('dashboard/data/dashboard.db')
cursor = conn.cursor()

cursor.execute('''
    INSERT INTO users (name, email, password_hash, permission_level, is_active)
    VALUES (?, ?, ?, ?, ?)
''', ('관리자', 'admin@example.com', generate_password_hash('password'), 'super_admin', 1))

conn.commit()
conn.close()
```

#### 옵션 2: Google OAuth 사용
1. Google 계정으로 로그인
2. 자동으로 `viewer` 권한으로 사용자 생성됨
3. 데이터베이스에서 직접 권한 변경:
```sql
UPDATE users SET permission_level = 'super_admin' WHERE email = 'your-email@example.com';
```

### 현재 데이터베이스 사용자 확인
```bash
cd dashboard
python -c "import sqlite3; conn = sqlite3.connect('data/dashboard.db'); cursor = conn.cursor(); cursor.execute('SELECT email, permission_level, is_active FROM users'); [print(f'{row[0]}: {row[1]} (활성: {row[2]})') for row in cursor.fetchall()]; conn.close()"
```

---

## 🔧 환경 변수 설정

### .env 파일 생성
프로젝트 루트에 `.env` 파일을 생성하세요:

```bash
# .env.example을 복사
cp .env.example .env
```

### 필수 환경 변수

```env
# Flask 설정
FLASK_APP=dashboard
FLASK_ENV=development
SECRET_KEY=your-secret-key-here

# 데이터베이스
DATABASE_URL=sqlite:///dashboard/data/dashboard.db

# Google OAuth (선택사항)
GOOGLE_CLIENT_ID=your-client-id
GOOGLE_CLIENT_SECRET=your-client-secret
GOOGLE_REDIRECT_URI=http://localhost:5000/auth/google/callback

# Google Sheets API
GOOGLE_SPREADSHEET_ID=your-spreadsheet-id
GOOGLE_SHEET_NAME=your-sheet-name
```

### Google OAuth 설정 (선택사항)

1. [Google Cloud Console](https://console.cloud.google.com/) 접속
2. 프로젝트 생성 또는 선택
3. "API 및 서비스" > "OAuth 동의 화면" 설정
4. "사용자 인증 정보" > "OAuth 2.0 클라이언트 ID" 생성
5. 승인된 리디렉션 URI 추가: `http://localhost:5000/auth/google/callback`
6. 클라이언트 ID와 시크릿을 `.env`에 추가
7. `credentials.json` 파일 다운로드하여 프로젝트 루트에 저장

---

## 🚀 초기 설정 방법

### 1. 저장소 클론 (첫 설정 시)
```bash
git clone <repository-url>
cd itglobal
```

### 2. Python 환경 설정
```bash
# 가상 환경 생성
python -m venv venv

# 가상 환경 활성화
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 의존성 설치
pip install -r requirements.txt
```

### 3. Node.js 환경 설정
```bash
cd dashboard
npm install
```

### 4. 환경 변수 설정
```bash
# .env.example을 복사하여 .env 생성
cp .env.example .env

# .env 파일을 편집하여 필요한 값 입력
```

### 5. 데이터베이스 초기화
```bash
# 데이터베이스 디렉토리 생성
mkdir -p dashboard/data

# Python으로 초기 사용자 생성 (선택사항)
python -c "
from dashboard.utils.user_database import get_user_database
db = get_user_database()
db.create_user('관리자', 'admin@example.com', 'password', 'super_admin')
print('초기 관리자 계정 생성 완료')
"
```

### 6. Google Sheets API 설정 (필요 시)
```bash
# credentials.json 파일을 프로젝트 루트에 배치
# .env에 GOOGLE_SPREADSHEET_ID와 GOOGLE_SHEET_NAME 설정
```

### 7. 서버 실행
```bash
# Flask 서버 (터미널 1)
python run_server.py

# Vite 개발 서버 (터미널 2)
cd dashboard
npm run dev
```

### 8. 접속
- Flask: http://localhost:5000
- Vite (핫 리로드): http://localhost:5173

---

## ✅ 환경 동기화 체크리스트

새로운 컴퓨터나 환경에서 설정할 때 다음을 확인하세요:

### Git으로 동기화되는 항목 ✅
- [x] 소스 코드 (Python, JavaScript, CSS)
- [x] 설정 파일 템플릿 (`.env.example`)
- [x] 문서 (README, 가이드)
- [x] 빌드 설정 (package.json, vite.config.js)

### Git으로 동기화되지 않는 항목 ❌ (수동 설정 필요)
- [ ] **데이터베이스 파일** (`dashboard.db`)
  - 각 환경에서 독립적으로 관리됨
  - 사용자, 로그 등은 각 환경마다 다름
- [ ] **환경 변수** (`.env`)
  - `.env.example`을 복사하여 `.env` 생성 후 값 입력
- [ ] **Google 인증 정보** (`credentials.json`)
  - Google Cloud Console에서 다운로드
- [ ] **node_modules**
  - `npm install`로 설치
- [ ] **Python 가상 환경** (`venv/`)
  - 각 환경에서 새로 생성

### 환경별 차이 확인 방법

#### 사용자 권한 확인
```bash
cd dashboard
python -c "import sqlite3; conn = sqlite3.connect('data/dashboard.db'); cursor = conn.cursor(); cursor.execute('SELECT email, permission_level FROM users'); print('현재 사용자:', cursor.fetchall()); conn.close()"
```

#### 감사 로그 수 확인
```bash
cd dashboard
python -c "import sqlite3; conn = sqlite3.connect('data/dashboard.db'); cursor = conn.cursor(); cursor.execute('SELECT COUNT(*) FROM audit_logs'); print(f'감사 로그 수: {cursor.fetchone()[0]}개'); conn.close()"
```

#### 환경 변수 확인
```bash
# .env 파일 존재 여부
ls -la .env

# .env 내용 확인 (민감한 정보 주의)
cat .env | grep -v "SECRET\|PASSWORD\|CLIENT"
```

---

## 🔄 회사 ↔ 집 환경 동기화

### 코드 동기화 (Git 사용)
```bash
# 회사에서 작업 후
git add .
git commit -m "feat: 새 기능 추가"
git push

# 집에서
git pull
npm install  # package.json이 변경된 경우
pip install -r requirements.txt  # requirements.txt가 변경된 경우
cd dashboard && npm run build  # 프론트엔드 빌드
```

### 데이터베이스 동기화 (필요 시)
**주의**: 데이터베이스는 일반적으로 환경마다 독립적으로 관리합니다.

동기화가 필요한 경우:
1. **수동 백업/복원**:
```bash
# 회사에서 백업
cp dashboard/data/dashboard.db dashboard/data/dashboard.backup.db
# Git에 커밋하지 말고 별도로 전송 (Dropbox, USB 등)

# 집에서 복원
cp dashboard/data/dashboard.backup.db dashboard/data/dashboard.db
```

2. **특정 데이터만 마이그레이션**:
```sql
-- 사용자 정보만 내보내기
.mode csv
.output users_export.csv
SELECT * FROM users;
.quit

-- 다른 환경에서 가져오기
.mode csv
.import users_export.csv users
```

---

## 🐛 문제 해결

### 1. "사용자 목록을 불러올 수 없습니다"
- **원인**: 로그인이 안 되어 있거나 권한 부족
- **해결**:
  1. Google OAuth로 다시 로그인
  2. 데이터베이스에서 사용자 권한 확인
  3. `permission_level`이 `admin` 또는 `super_admin`인지 확인

### 2. "시간이 다르게 표시됩니다"
- **원인**: 브라우저 캐시
- **해결**: Ctrl+Shift+R (강력 새로고침) 또는 `npm run build` 후 새로고침

### 3. "로그가 다릅니다"
- **원인**: 각 환경의 데이터베이스가 독립적임
- **해결**: 정상 동작입니다. 환경마다 다른 로그가 생성됩니다.

### 4. "Flask 서버가 시작되지 않습니다"
- **원인**: 포트 충돌 또는 환경 변수 누락
- **해결**:
  1. `.env` 파일 존재 확인
  2. 포트 5000이 사용 중인지 확인
  3. 가상 환경 활성화 확인

---

## 📚 추가 리소스

- [Flask 공식 문서](https://flask.palletsprojects.com/)
- [Vite 공식 문서](https://vitejs.dev/)
- [Google OAuth 가이드](https://developers.google.com/identity/protocols/oauth2)
- [SQLite 공식 문서](https://www.sqlite.org/docs.html)

---

**마지막 업데이트**: 2025-10-14
**작성자**: Claude Code
