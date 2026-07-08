"""익명 접근 가능한 엔드포인트 부하 테스트.

로그인 세션 획득 없이 서버 반응성·스레드풀·SocketIO polling 처리 능력 실측.

실행:
    cd "C:\\Users\\SECOM\\Desktop\\ITG-Project\\Claude Project\\docs\\deployment"
    locust -f locustfile_anonymous.py --headless --users 20 --spawn-rate 5 --run-time 60s --host https://pm.itg-aircon.com
"""
from locust import HttpUser, task, between


class AnonymousUser(HttpUser):
    wait_time = between(1, 3)
    host = "https://pm.itg-aircon.com"

    @task(3)
    def health_check(self):
        """모니터링 도구 흉내"""
        self.client.get("/api/health", name="/api/health")

    @task(2)
    def landing_page(self):
        """로그인 페이지 렌더링"""
        self.client.get("/", name="/", allow_redirects=False)

    @task(2)
    def static_asset_manifest(self):
        """Vite manifest 조회 (dist 서빙 확인)"""
        self.client.get("/static/dist/.vite/manifest.json", name="/static/manifest")

    @task(1)
    def favicon(self):
        self.client.get("/favicon.ico", name="/favicon.ico")
