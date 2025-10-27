# ITG 대시보드 완전 파이썬 재구축 계획

## 🔍 현재 시스템 문제점

### 심각한 복잡성
- **13,062줄** 단일 HTML 템플릿
- 4,000줄+ JavaScript 코드
- DataTable.js + Socket.IO + 복잡한 DOM 조작
- 캐시 동기화 문제
- 클라이언트-서버 상태 불일치

### 유지보수 불가능
- 하나의 파일에 모든 기능 집중
- 복잡한 의존성 관계
- 디버깅 어려움
- 새 기능 추가 시 부작용 위험

## 🎯 새로운 아키텍처: "Simple Python Web App"

### 기술 스택 선택

#### Option A: FastAPI + Jinja2 (추천)
```
FastAPI (최신, 빠름, 타입 힌트)
├── SQLite/PostgreSQL (데이터 영속성)
├── Jinja2 Templates (서버사이드 렌더링)
├── HTMX (최소한의 동적 기능)
└── Bootstrap 5 (스타일링)
```

#### Option B: Django (검증된 선택)
```
Django (풀스택, 검증됨)
├── Django ORM (데이터베이스)
├── Django Templates
├── Django Admin (관리 기능)
└── Bootstrap 5
```

#### Option C: Flask 개선 (현재 기반 활용)
```
현재 Flask 유지
├── SQLite 데이터베이스 추가
├── 템플릿 완전 재작성
├── JavaScript 90% 제거
└── 서버사이드 중심
```

### 🗄️ 데이터 아키텍처 개선

#### 현재: Google Sheets 중심
```
Google Sheets ← 모든 데이터
├── 캐시 복잡성
├── 동기화 문제
└── 성능 제약
```

#### 신규: 하이브리드 데이터베이스
```
SQLite/PostgreSQL (주 데이터베이스)
├── Projects (프로젝트 정보)
├── Users (사용자 관리)
├── AuditLogs (감사 로그)
├── Cache (캐시 테이블)
└── Settings (설정)

Google Sheets (백업/동기화)
├── 주기적 동기화
├── 백업 용도
└── 외부 접근용
```

## 🏗️ 구현 계획

### Phase 1: 데이터베이스 설계 (1일)

#### 테이블 구조
```sql
-- 프로젝트 테이블
CREATE TABLE projects (
    id INTEGER PRIMARY KEY,
    project_code VARCHAR(20) UNIQUE,
    site_name VARCHAR(200),
    site_address TEXT,
    business VARCHAR(100),
    manager VARCHAR(100),
    manager_email VARCHAR(100),
    status VARCHAR(50) DEFAULT 'active',  -- active, cancelled, completed
    total_amount_1 DECIMAL(15,2),
    total_amount_2 DECIMAL(15,2),
    contract_deposit DECIMAL(15,2),
    mid_payment DECIMAL(15,2),
    final_payment DECIMAL(15,2),
    collection_notes TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    is_cancelled BOOLEAN DEFAULT FALSE
);

-- 사용자 테이블
CREATE TABLE users (
    id INTEGER PRIMARY KEY,
    email VARCHAR(200) UNIQUE,
    name VARCHAR(100),
    role VARCHAR(50),  -- admin, editor, viewer
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 감사 로그 테이블
CREATE TABLE audit_logs (
    id INTEGER PRIMARY KEY,
    user_id INTEGER,
    project_code VARCHAR(20),
    action VARCHAR(100),
    field_name VARCHAR(100),
    old_value TEXT,
    new_value TEXT,
    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    ip_address VARCHAR(45),
    FOREIGN KEY (user_id) REFERENCES users(id)
);

-- 프로젝트 잠금 테이블
CREATE TABLE project_locks (
    id INTEGER PRIMARY KEY,
    project_code VARCHAR(20),
    user_id INTEGER,
    locked_fields TEXT,  -- JSON array
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    expires_at TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES users(id)
);
```

### Phase 2: 백엔드 API 구현 (2-3일)

#### FastAPI 구조
```python
app/
├── main.py                 # 애플리케이션 진입점
├── models/
│   ├── project.py         # 프로젝트 모델
│   ├── user.py           # 사용자 모델
│   └── audit.py          # 감사 로그 모델
├── routers/
│   ├── projects.py       # 프로젝트 관련 API
│   ├── auth.py          # 인증 관련 API
│   └── admin.py         # 관리 기능 API
├── services/
│   ├── project_service.py # 프로젝트 비즈니스 로직
│   ├── sheets_sync.py    # Google Sheets 동기화
│   └── cache_service.py  # 캐시 관리
├── templates/
│   ├── base.html        # 기본 템플릿
│   ├── projects/
│   │   ├── list.html    # 프로젝트 목록 (< 200줄)
│   │   ├── detail.html  # 프로젝트 상세 (< 150줄)
│   │   └── edit.html    # 프로젝트 편집 (< 100줄)
│   └── components/
│       ├── header.html  # 헤더 컴포넌트
│       └── footer.html  # 푸터 컴포넌트
└── static/
    ├── css/
    │   └── app.css      # 메인 스타일 (< 500줄)
    └── js/
        └── app.js       # 최소한의 JS (< 200줄)
```

#### 핵심 API 엔드포인트
```python
# 프로젝트 관리
GET  /projects              # 프로젝트 목록 (서버사이드 페이징)
GET  /projects/{code}       # 프로젝트 상세
POST /projects              # 프로젝트 생성
PUT  /projects/{code}       # 프로젝트 수정
POST /projects/{code}/cancel # 프로젝트 취소
POST /projects/{code}/resume # 프로젝트 재개

# 인증 및 사용자
GET  /auth/login           # 로그인 페이지
POST /auth/login           # 로그인 처리
GET  /auth/logout          # 로그아웃
GET  /users/profile        # 프로필 페이지

# 관리 기능
GET  /admin/cache          # 캐시 관리
POST /admin/sync-sheets    # Google Sheets 동기화
GET  /admin/audit-logs     # 감사 로그
```

### Phase 3: 프론트엔드 단순화 (2일)

#### 템플릿 구조 (총 1,000줄 이하)
```html
<!-- base.html (100줄) -->
<!DOCTYPE html>
<html>
<head>
    <title>ITG 대시보드</title>
    <link href="/static/css/app.css" rel="stylesheet">
</head>
<body>
    {% include 'components/header.html' %}
    <main>{% block content %}{% endblock %}</main>
    {% include 'components/footer.html' %}
    <script src="/static/js/app.js"></script>
</body>
</html>

<!-- projects/list.html (200줄) -->
{% extends 'base.html' %}
{% block content %}
<div class="container">
    <!-- 필터링 (서버사이드) -->
    <form method="GET" class="filters mb-4">
        <input name="search" value="{{ request.args.search }}">
        <select name="status">
            <option value="">전체</option>
            <option value="active">진행중</option>
            <option value="cancelled">취소됨</option>
        </select>
        <button type="submit">검색</button>
    </form>

    <!-- 프로젝트 테이블 -->
    <table class="table">
        <thead>
            <tr>
                <th>상태</th>
                <th>프로젝트 코드</th>
                <th>현장명</th>
                <th>액션</th>
            </tr>
        </thead>
        <tbody>
            {% for project in projects %}
            <tr class="{% if project.is_cancelled %}cancelled{% endif %}">
                <td>
                    {% if project.is_cancelled %}
                        <span class="badge bg-danger">취소됨</span>
                    {% else %}
                        <span class="badge bg-success">진행중</span>
                    {% endif %}
                </td>
                <td>{{ project.project_code }}</td>
                <td>{{ project.site_name }}</td>
                <td>
                    {% if project.is_cancelled %}
                        <button onclick="resumeProject('{{ project.project_code }}')">재개</button>
                    {% else %}
                        <a href="/projects/{{ project.project_code }}">상세</a>
                        <button onclick="cancelProject('{{ project.project_code }}')">취소</button>
                    {% endif %}
                </td>
            </tr>
            {% endfor %}
        </tbody>
    </table>

    <!-- 페이징 (서버사이드) -->
    {% if pagination.pages > 1 %}
    <nav>
        {% for page in pagination.iter_pages() %}
            <a href="?page={{ page }}">{{ page }}</a>
        {% endfor %}
    </nav>
    {% endif %}
</div>
{% endblock %}
```

#### 최소한의 JavaScript (200줄)
```javascript
// app.js - 전체 JavaScript 파일
class ITGDashboard {
    constructor() {
        this.initEventListeners();
    }

    initEventListeners() {
        // 폼 제출 시 로딩 표시
        document.querySelectorAll('form').forEach(form => {
            form.addEventListener('submit', this.showLoading);
        });

        // 확인 대화상자
        document.querySelectorAll('[data-confirm]').forEach(btn => {
            btn.addEventListener('click', this.confirmAction);
        });
    }

    showLoading(e) {
        const btn = e.target.querySelector('button[type="submit"]');
        if (btn) {
            btn.disabled = true;
            btn.innerHTML = '<i class="spinner"></i> 처리 중...';
        }
    }

    confirmAction(e) {
        const message = e.target.dataset.confirm;
        if (!confirm(message)) {
            e.preventDefault();
        }
    }

    // 프로젝트 취소/재개 (HTMX 또는 간단한 POST)
    async cancelProject(projectCode) {
        if (!confirm('정말로 취소하시겠습니까?')) return;

        try {
            const response = await fetch(`/projects/${projectCode}/cancel`, {
                method: 'POST',
                headers: {'Content-Type': 'application/json'}
            });

            if (response.ok) {
                location.reload(); // 서버사이드 렌더링으로 새로고침
            } else {
                alert('취소에 실패했습니다.');
            }
        } catch (error) {
            alert('오류가 발생했습니다.');
        }
    }

    async resumeProject(projectCode) {
        if (!confirm('프로젝트를 재개하시겠습니까?')) return;

        const response = await fetch(`/projects/${projectCode}/resume`, {
            method: 'POST',
            headers: {'Content-Type': 'application/json'}
        });

        if (response.ok) {
            location.reload();
        }
    }
}

// 초기화
document.addEventListener('DOMContentLoaded', () => {
    new ITGDashboard();
});

// 전역 함수 (템플릿에서 호출용)
function cancelProject(code) { window.dashboard.cancelProject(code); }
function resumeProject(code) { window.dashboard.resumeProject(code); }
```

### Phase 4: 데이터 마이그레이션 (1일)

#### Google Sheets → Database 마이그레이션
```python
# migration/migrate_data.py
import sqlite3
import pandas as pd
from google_sheets import GoogleSheetsManager

def migrate_from_sheets():
    """Google Sheets 데이터를 SQLite로 마이그레이션"""

    # 기존 데이터 읽기
    manager = GoogleSheetsManager()
    df = manager.get_sheet_data(sheet_id, sheet_range)

    # 데이터베이스 연결
    conn = sqlite3.connect('itg_dashboard.db')

    # 프로젝트 데이터 삽입
    for _, row in df.iterrows():
        project = {
            'project_code': row['프로젝트 코드'],
            'site_name': row['현장명'],
            'site_address': row['현장 주소'],
            'business': row['사업자'],
            'manager': row['담당자'],
            'is_cancelled': row['수금 관련 특이사항'] == '공사취소',
            'total_amount_1': parse_amount(row['총액1']),
            'total_amount_2': parse_amount(row['총액2']),
            # ... 기타 필드
        }

        insert_project(conn, project)

    conn.close()
    print(f"마이그레이션 완료: {len(df)}개 프로젝트")

def setup_sync_schedule():
    """Google Sheets와 정기 동기화 설정"""
    # 1시간마다 양방향 동기화
    pass
```

### Phase 5: 배포 및 테스트 (1일)

#### Docker 구성
```dockerfile
# Dockerfile
FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .
EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
```

#### 배포 스크립트
```bash
#!/bin/bash
# deploy.sh

# 백업
cp itg_dashboard.db itg_dashboard.db.backup

# 새 버전 배포
docker build -t itg-dashboard .
docker stop itg-dashboard-old || true
docker run -d --name itg-dashboard -p 8000:8000 itg-dashboard

# 헬스체크
curl -f http://localhost:8000/health || exit 1

echo "배포 완료"
```

## 🎯 예상 결과

### 코드 복잡성 대폭 감소
- **13,062줄** → **1,000줄** (92% 감소)
- JavaScript **4,000줄** → **200줄** (95% 감소)
- 단일 파일 → 모듈화된 구조

### 성능 및 안정성 향상
- ✅ 캐시 문제 완전 해결 (데이터베이스 기반)
- ✅ 실시간 동기화 문제 해결 (서버 중심)
- ✅ DOM 조작 복잡성 제거
- ✅ 상태 일관성 100% 보장

### 개발 효율성 향상
- ✅ 파이썬 단일 언어
- ✅ 타입 힌트 지원 (FastAPI)
- ✅ 자동 API 문서화
- ✅ 테스트 용이성

### 유지보수성 향상
- ✅ 모듈화된 구조
- ✅ 명확한 책임 분리
- ✅ 디버깅 용이성
- ✅ 새 기능 추가 안정성

## 📅 구현 일정

| 단계 | 기간 | 주요 작업 |
|------|------|-----------|
| Phase 1 | 1일 | 데이터베이스 설계 및 구축 |
| Phase 2 | 2-3일 | FastAPI 백엔드 구현 |
| Phase 3 | 2일 | 프론트엔드 단순화 |
| Phase 4 | 1일 | 데이터 마이그레이션 |
| Phase 5 | 1일 | 배포 및 테스트 |

**총 소요 기간: 7-8일**

## 🚀 다음 단계

1. **기술 스택 선택** (FastAPI vs Django vs Flask 개선)
2. **데이터베이스 스키마 확정**
3. **마이그레이션 전략 수립**
4. **단계별 구현 시작**

어떤 기술 스택을 선호하시나요?
- FastAPI (최신, 빠름, 타입 힌트)
- Django (검증됨, 풀스택)
- Flask 개선 (현재 코드 활용)