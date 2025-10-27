# IT Global Dashboard API 개발자 가이드

## 시작하기

이 가이드는 IT Global Dashboard API API를 사용하여 애플리케이션을 개발하는 방법을 설명합니다.

## 빠른 시작

### 1. 인증 설정

```bash
# Bearer 토큰 인증
curl -H "Authorization: Bearer YOUR_TOKEN" \
     -H "Content-Type: application/json" \
     {base_url}/api/v1/test
```

### 2. 기본 요청 예시

```python
import requests

# API 기본 URL
BASE_URL = "http://localhost:5000"

# 인증 헤더
headers = {
    "Authorization": "Bearer YOUR_TOKEN",
    "Content-Type": "application/json"
}

# GET 요청
response = requests.get(f"{BASE_URL}/api/v1/projects", headers=headers)
data = response.json()

print(f"성공: {data['success']}")
print(f"데이터: {data['data']}")
```

## 표준 응답 형식

모든 API 응답은 다음과 같은 표준 형식을 따릅니다:

```json
{
  "success": true,
  "data": "실제 응답 데이터",
  "error": null,
  "meta": {
    "timestamp": "2025-01-01T00:00:00Z",
    "version": "v1",
    "request_id": "uuid-string"
  }
}
```

### 성공 응답
- `success`: `true`
- `data`: 요청한 데이터 또는 작업 결과
- `error`: `null`

### 오류 응답
- `success`: `false`
- `data`: `null`
- `error`: 오류 상세 정보

## 오류 코드

| HTTP 코드 | 설명 | 해결 방법 |
|-----------|------|-----------|
| 400 | 잘못된 요청 | 요청 파라미터나 본문을 확인하세요 |
| 401 | 인증 필요 | 유효한 인증 토큰을 제공하세요 |
| 403 | 권한 없음 | 해당 리소스에 접근할 권한이 있는지 확인하세요 |
| 404 | 찾을 수 없음 | URL이나 리소스 ID를 확인하세요 |
| 500 | 서버 오류 | 서버 관리자에게 문의하세요 |

## 페이지네이션

목록 조회 API는 페이지네이션을 지원합니다:

```python
# 페이지네이션 파라미터
params = {
    "page": 1,      # 페이지 번호 (기본: 1)
    "limit": 20     # 페이지당 항목 수 (기본: 20, 최대: 100)
}

response = requests.get(f"{BASE_URL}/api/v1/projects", headers=headers, params=params)
```

## 코드 예시

### 프로젝트 생성
```python
project_data = {
    "name": "새 프로젝트",
    "description": "프로젝트 설명",
    "start_date": "2025-01-01",
    "budget": 10000.00
}

response = requests.post(
    f"{BASE_URL}/api/v1/projects",
    headers=headers,
    json=project_data
)

if response.json()['success']:
    print("프로젝트 생성 성공!")
else:
    print(f"오류: {response.json()['error']}")
```

### 프로젝트 목록 조회
```python
response = requests.get(f"{BASE_URL}/api/v1/projects", headers=headers)
projects = response.json()['data']

for project in projects:
    print(f"프로젝트: {project['name']} (ID: {project['id']})")
```

## SDK 및 클라이언트 라이브러리

### Python 클라이언트
```python
# pip install requests

class ITGlobalAPIClient:
    def __init__(self, base_url, token):
        self.base_url = base_url
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

    def get_projects(self, page=1, limit=20):
        response = requests.get(
            f"{self.base_url}/api/v1/projects",
            headers=self.headers,
            params={"page": page, "limit": limit}
        )
        return response.json()

    def create_project(self, project_data):
        response = requests.post(
            f"{self.base_url}/api/v1/projects",
            headers=self.headers,
            json=project_data
        )
        return response.json()

# 사용법
client = ITGlobalAPIClient("http://localhost:5000", "your_token")
projects = client.get_projects()
```

### JavaScript 클라이언트
```javascript
class ITGlobalAPIClient {
    constructor(baseUrl, token) {
        this.baseUrl = baseUrl;
        this.headers = {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        };
    }

    async getProjects(page = 1, limit = 20) {
        const response = await fetch(
            `${this.baseUrl}/api/v1/projects?page=${page}&limit=${limit}`,
            { headers: this.headers }
        );
        return await response.json();
    }

    async createProject(projectData) {
        const response = await fetch(`${this.baseUrl}/api/v1/projects`, {
            method: 'POST',
            headers: this.headers,
            body: JSON.stringify(projectData)
        });
        return await response.json();
    }
}

// 사용법
const client = new ITGlobalAPIClient('http://localhost:5000', 'your_token');
client.getProjects().then(data => console.log(data));
```

## 개발 환경 설정

### 로컬 개발
1. 개발 서버 시작: `http://localhost:5000`
2. API 문서: `http://localhost:5000/docs`
3. 테스트 엔드포인트: `http://localhost:5000/api/v1/test`

### 프로덕션 환경
- URL: `https://dashboard.itglobal.com`
- SSL 인증서 필요
- Rate Limiting 적용

## 문의 및 지원

- 개발팀 이메일: dev@itglobal.com
- API 문서: [Swagger UI](/docs)
- GitHub Issues: [프로젝트 저장소](https://github.com/itglobal/dashboard)

---

*이 가이드는 자동으로 생성되었습니다. 최종 업데이트: 2025-09-19 16:46:29*
