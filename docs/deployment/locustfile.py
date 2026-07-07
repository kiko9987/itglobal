"""20 매니저 동시 사용 시뮬레이션.

실행:
  locust -f locustfile.py --users 20 --spawn-rate 5

주의:
- 편집/취소/재개는 데이터 변경이라 프로덕션 시트에 부하 걸면 안 됨.
  스테이징 환경 or 테스트 시트에서만 실행.
- 로그인은 Google OAuth라 자동화 어려움. 세션 쿠키 미리 획득해서 self.client.cookies에 세팅.
"""
from locust import HttpUser, task, between


class DashboardUser(HttpUser):
    wait_time = between(1, 5)
    host = "https://pm.itg-aircon.com"

    def on_start(self):
        # TODO: 실제 세션 쿠키를 여기 세팅
        # self.client.cookies['session'] = 'YOUR_SESSION_COOKIE'
        pass

    @task(5)
    def list_projects(self):
        self.client.get("/projects", name="/projects")

    @task(3)
    def api_get_projects(self):
        self.client.get("/api/projects/list", name="/api/projects/list")

    @task(2)
    def health_check(self):
        self.client.get("/api/health", name="/api/health")
