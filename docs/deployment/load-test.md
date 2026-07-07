# 20명 동시 사용 부하 테스트 가이드

실서비스 개시 전 20 매니저 동시 사용 시나리오를 시뮬레이션해 병목·오류 지점 확인.

## 도구: Locust

Python 기반 부하 테스트. 웹 UI로 실시간 관측.

### 설치 (venv에)

```powershell
& "C:\Users\SECOM\Desktop\ITG-Project\Claude Project\.venv\Scripts\pip.exe" install locust
```

## 테스트 스크립트

`docs/deployment/locustfile.py`에 저장:

```python
"""20 매니저 동시 사용 시뮬레이션.

시나리오:
- 각 유저가 로그인
- 프로젝트 리스트 조회 (weight 5)
- 특정 프로젝트 상세 조회 (weight 3)
- 편집 시도 (weight 2, 이따금)
- 취소·재개 (weight 1, 드물게)
"""
from locust import HttpUser, task, between
import random


class DashboardUser(HttpUser):
    wait_time = between(1, 5)  # 각 액션 사이 1~5초 대기 (실사용자 흉내)
    host = "https://pm.itg-aircon.com"

    def on_start(self):
        """세션당 한 번, 로그인"""
        # TODO: 실제 로그인 방식에 맞게 조정
        # 예: Google OAuth flow는 자동화 어려움 → 세션 쿠키 미리 획득 후 붙이기
        # self.client.cookies['session'] = 'PRE_ACQUIRED_SESSION_COOKIE'
        pass

    @task(5)
    def list_projects(self):
        """프로젝트 리스트"""
        self.client.get("/projects", name="/projects")

    @task(3)
    def api_get_projects(self):
        """프로젝트 API"""
        self.client.get("/api/projects/list", name="/api/projects/list")

    @task(2)
    def health_check(self):
        """헬스 체크 (모니터링 흉내)"""
        self.client.get("/api/health", name="/api/health")

    # 편집·취소·재개는 데이터 변경이라 부하 테스트에선 주의 (테스트 DB로만)
```

## 실행

### 로컬 부하 (같은 PC에서)

```powershell
cd "C:\Users\SECOM\Desktop\ITG-Project\Claude Project\docs\deployment"
locust -f locustfile.py --users 20 --spawn-rate 5
```

- 웹 UI: <http://localhost:8089>
- Users: 20 (동시 매니저 수)
- Spawn rate: 5/초 (4초 만에 20명 다 접속)

Start swarm 클릭 → 통계 관측.

### 관측 지표

| 지표 | 정상 범위 | 이상 신호 |
|---|---|---|
| Response Time (p50) | < 500ms | > 2s → Waitress 스레드 부족 or Google Sheets 지연 |
| Response Time (p99) | < 3s | > 10s → 캐시 미스로 시트 로드 |
| Failure rate | 0% | > 1% → 로그 확인 필요 |
| RPS | 5~15 req/s | 갑자기 떨어지면 크래시 or lock |

### 동시 서버 로그 관측

부하 테스트 중 별도 터미널:

```powershell
Get-Content "C:\Users\SECOM\Desktop\ITG-Project\Claude Project\dashboard\logs\service_stdout.log" -Wait -Tail 20
```

### Google Sheets API 사용률 관측

Cloud Console → APIs & Services → Google Sheets API → **Metrics** 페이지.
- 실시간 트래픽 그래프
- 429 오류 발생 여부

## 파괴적 시나리오 테스트

부하 테스트 별개로 다음도 확인:

### 1. Redis 다운 시나리오

```powershell
docker stop redis
Invoke-WebRequest -Uri "https://pm.itg-aircon.com/api/health" -UseBasicParsing
# 응답 status: 'critical', cache: 'fallback_mode' 확인
docker start redis
```

Waitress 단일 프로세스라 fallback 모드에서도 앱은 응답. 검증.

### 2. Google Sheets API 500 에러 처리

방화벽으로 `sheets.googleapis.com` 임시 차단 → 재시도 로직·에러 알림 확인.

### 3. 취소·재개 반복 크래시 재현 시도

같은 프로젝트에 반복적으로 취소·재개 30회+ → SIGSEGV 재발 여부.

## 결과 판단

- **전부 정상**: 실서비스 개시 가능
- **response time 초과**: Waitress `threads`, `connection_limit`  튜닝 or CACHE_TTL 상향
- **failure rate 초과**: 로그 traceback 분석 후 개별 fix
- **Google API 429**: 쿼터 상향이 아직 반영 안 됐거나 캐시 hit rate 낮음

## 참고

- Locust 문서: <https://docs.locust.io/>
- 지금 서버 스펙 (Waitress threads=30, connection_limit=200) 은 100 유저까지 감당 설계됨. 20명은 여유.
