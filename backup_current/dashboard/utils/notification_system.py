import os, time, logging, requests
from typing import List, Dict
logger = logging.getLogger(__name__)

class NotificationSystem:
    def __init__(self):
        self.slack_webhook = os.getenv("SLACK_WEBHOOK_URL", "").strip()
        self.suppress_minutes = int(os.getenv("NOTIFY_SUPPRESS_MINUTES", "120"))
        self._recent: Dict[str, float] = {}

    def _should_send(self, key: str) -> bool:
        now = time.time()
        last = self._recent.get(key, 0)
        if now - last >= self.suppress_minutes * 60:
            self._recent[key] = now
            return True
        return False

    def _post_slack(self, text: str):
        if not self.slack_webhook:
            logger.info(f"[SLACK DRYRUN] {text}")
            return True
        try:
            resp = requests.post(self.slack_webhook, json={"text": text}, timeout=10)
            return (resp.status_code // 100) == 2
        except Exception:
            return False

    def notify_missing_fields(self, items: List[Dict]):
        for it in items:
            proj = it.get("프로젝트 코드") or "(미지정)"
            missing = ", ".join(it.get("_missing_fields", []))
            owner = it.get("담당자", "")
            key = f"missing:{proj}:{missing}"
            if self._should_send(key):
                text = f"⚠️ 누락 필드 감지\n• 프로젝트: {proj}\n• 누락: {missing}\n• 담당자: {owner}"
                self._post_slack(text)
