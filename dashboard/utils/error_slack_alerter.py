"""에러 로그 → 관리자 슬랙 큐레이션 알림 (error_slack_alerter).

목적
----
코드 어디서든 발생한 ERROR/CRITICAL 로그를, 우리가 실제로 보는 관리자 슬랙
DM(SLACK_ADMIN_CHANNEL)으로 '큐레이션'해서 전달한다. Sentry(이미 활성)는
아카이브·검색·그룹핑용이고, 이 알림은 "지금 봐야 함" 호출기 역할.

왜 '전량 전달'이 아니라 큐레이션인가
------------------------------------
측정(2026-09-11): 최근 2개월 ERROR 6,005건 중 대부분이 구글 API 타임아웃·
재시도 소진·engineio 소켓 잡음 등 만성 반복이었다. 전량 전달하면 사장님 DM
도배 → 알림 무력화. 그래서:

  * 완전 잡음(denylist): engineio/socketio/waitress·재시작 잔재 → 무시
  * 일시적/만성(transient): 매 건 무시, '급증(spike)' 시에만 1건 경보
    (구글 API 순단·wedge 사고 같은 걸 포착)
  * 신규(new) 시그니처: 즉시 경보 (전에 못 보던 게 터짐 = 최고 신호)
  * 반복(recurring) 일반 에러: 1시간 dedup
  * CRITICAL: 무조건 통과(120초 dedup)
  * 전역 rate cap(60초당 N건): 다수 시그니처 동시 폭주 방어

설계 원칙
--------
  * logging.Handler 로 root 에 부착. emit() 은 분류+큐잉만(로그 호출 차단 없음).
  * 실제 슬랙 전송은 백그라운드 워커 스레드(bounded queue).
  * 예외 완전 격리: 알림 기계가 앱을 절대 죽이지 않는다.
  * cooldown/seen 은 Redis(재시작 넘어 유지) + Redis 장애 시 in-memory 폴백.
    spike 윈도우·전역 rate 는 짧은 창이라 in-memory(프로세스 로컬)로 충분.

환경변수
--------
  ERROR_SLACK_ALERTS_ENABLED  (기본 '1'; '0'/'false'/'no' 면 비활성)
  ERROR_ALERT_CHANNEL         (발송 대상 override; 기본 SLACK_ADMIN_CHANNEL)
  ERROR_ALERT_SPIKE_THRESHOLD (기본 20; 5분 내 동일 유형 N건 이상이면 급증 경보)
  SLACK_BOT_TOKEN             (발송 토큰; 없으면 알림 비활성)
"""

import json
import logging
import os
import queue
import re
import threading
import time
import traceback as _traceback
import urllib.request
from collections import defaultdict, deque
from typing import Callable, Optional

logger = logging.getLogger(__name__)

# ── 임계값/상수 ────────────────────────────────────────────────
_MIN_LEVEL = logging.ERROR
SPIKE_WINDOW_SEC = 300                                    # 급증 판정 창 (5분)
SPIKE_THRESHOLD = int(os.getenv('ERROR_ALERT_SPIKE_THRESHOLD', '20') or 20)
SPIKE_COOLDOWN_SEC = 1800                                 # 급증 경보 후 동일 sig 30분 억제
RECURRING_COOLDOWN_SEC = 3600                            # 반복 일반 에러 sig 1시간 dedup
CRITICAL_COOLDOWN_SEC = 120                              # CRITICAL 은 120초 dedup(버스트 중복만 억제)
GLOBAL_RATE_MAX = 6                                      # 60초당 최대 경보 수
GLOBAL_RATE_WINDOW_SEC = 60
SEEN_TTL_SEC = 30 * 86400                                # 시그니처 '앎' 유지 30일
MIN_KNOWN_FOR_NEW = 5                                    # 이만큼 쌓이기 전엔 new 라벨 안 붙임(빈 Redis 오탐 방지)
QUEUE_MAX = 200

# 완전 무시할 잡음: logger 이름 접두 / 메시지 부분문자열
_DENY_LOGGER_PREFIXES = (
    'waitress', 'engineio', 'socketio', 'geventwebsocket',
    __name__, 'dashboard.utils.slack_health',
)
_DENY_MSG_SUBSTR = (
    'Invalid session', 'Session is disconnected',
    'I/O operation on closed file', 'Exception while serving /socket.io/',
    'write() before start_response',
)
# 일시적/만성 유형: 매 건 무시하되 급증할 때만 경보 (외부 API 순단 신호)
_TRANSIENT_SUBSTR = (
    'TimeoutError', 'read operation timed out', 'The read operation timed out',
    '최대 재시도 횟수', 'Timeout reading from socket', '연결 오류',
    'HttpError 500', 'HttpError 502', 'HttpError 503', 'HttpError 429',
    'API 사용량 한도', 'ServerNotFoundError', 'Connection aborted',
    'Connection reset', 'Read timed out', 'Max retries exceeded',
)

_CODE_RE = re.compile(r'[GRNgrn]\d{3,4}-[A-Za-z]{1,3}')   # 프로젝트 코드 정규화
_NUM_RE = re.compile(r'\d+')

_KIND_HEADER = {
    'new': '🆕 신규 에러',
    'recurring': '🔁 반복 에러',
    'spike': '📈 에러 급증',
    'critical': '⛔ CRITICAL',
}


def _safe_msg(record: logging.LogRecord) -> str:
    try:
        return record.getMessage()
    except Exception:
        return str(getattr(record, 'msg', ''))


def _signature(record: logging.LogRecord, msg: str) -> str:
    """동일 유형 판정용 시그니처. 프로젝트코드·숫자를 정규화해 변주를 하나로 묶음."""
    name = record.name or 'root'
    where = f"{getattr(record, 'module', '?')}:{getattr(record, 'lineno', '?')}"
    exc_type = ''
    if record.exc_info and record.exc_info[0]:
        exc_type = record.exc_info[0].__name__
    core = _CODE_RE.sub('<code>', msg)
    core = _NUM_RE.sub('#', core)[:60]
    return f"{name}|{where}|{exc_type}|{core}"


class _AlertGate:
    """분류·억제 로직(순수 결정). Redis 있으면 cooldown/seen 을 재시작 넘어 유지."""

    def __init__(self, *, now: Callable[[], float] = time.time,
                 redis_getter: Optional[Callable] = None):
        self._now = now
        self._redis_getter = redis_getter
        self._lock = threading.Lock()
        self._sig_window = defaultdict(deque)   # sig -> deque[ts]  (급증 판정)
        self._global = deque()                   # 전송된 경보 ts (rate cap)
        self._suppressed = 0                     # rate cap 으로 버린 수(다음 경보에 표기)
        # Redis 폴백용 in-memory
        self._mem_cooldown = {}                  # sig -> until_ts
        self._mem_seen = {}                      # sig -> last_ts

    # ── Redis 헬퍼 (원시 클라이언트 사용; 래퍼엔 해시 연산 없음) ──
    def _raw(self):
        if not self._redis_getter:
            return None
        try:
            rc = self._redis_getter()
            return getattr(rc, 'redis', None) or rc
        except Exception:
            return None

    def _cooldown_acquire(self, sig: str, ttl: int) -> bool:
        """지금 보낼 수 있으면 True(그리고 ttl 동안 잠금). 이미 쿨다운 중이면 False."""
        key = f'error_alert:cd:{abs(hash(sig)) % (10 ** 12)}'
        raw = self._raw()
        if raw is not None:
            try:
                got = raw.set(key, '1', nx=True, ex=ttl)
                return bool(got)
            except Exception:
                pass
        # in-memory 폴백
        now = self._now()
        with self._lock:
            until = self._mem_cooldown.get(sig, 0)
            if now < until:
                return False
            self._mem_cooldown[sig] = now + ttl
            return True

    def _seen_classify(self, sig: str) -> bool:
        """신규 시그니처면 True. 부수효과로 seen 기록. (일반 에러에만 사용)"""
        raw = self._raw()
        if raw is not None:
            try:
                hkey = 'error_alert:sigs'
                exists = raw.hexists(hkey, sig)
                cnt = raw.hlen(hkey)
                raw.hset(hkey, sig, f'{self._now():.0f}')
                raw.expire(hkey, SEEN_TTL_SEC)
                return (not exists) and (cnt >= MIN_KNOWN_FOR_NEW)
            except Exception:
                pass
        # in-memory 폴백
        with self._lock:
            exists = sig in self._mem_seen
            cnt = len(self._mem_seen)
            self._mem_seen[sig] = self._now()
        return (not exists) and (cnt >= MIN_KNOWN_FOR_NEW)

    def _mark_seen(self, sig: str) -> None:
        """spike/critical 도 seen 에 기록(나중에 'new' 오분류 방지). 실패 무시."""
        raw = self._raw()
        if raw is not None:
            try:
                raw.hset('error_alert:sigs', sig, f'{self._now():.0f}')
                raw.expire('error_alert:sigs', SEEN_TTL_SEC)
                return
            except Exception:
                pass
        with self._lock:
            self._mem_seen[sig] = self._now()

    def decide(self, record: logging.LogRecord) -> Optional[dict]:
        """전송 결정. 보낼 것 없으면 None, 보낼 거면 dict(kind/sig/window_count/...)."""
        if record.levelno < _MIN_LEVEL:
            return None
        msg = _safe_msg(record)

        name = (record.name or '').lower()
        if any(name.startswith(p.lower()) for p in _DENY_LOGGER_PREFIXES):
            return None
        if any(s in msg for s in _DENY_MSG_SUBSTR):
            return None

        now = self._now()
        sig = _signature(record, msg)
        is_critical = record.levelno >= logging.CRITICAL
        transient = any(s in msg for s in _TRANSIENT_SUBSTR)

        # 급증 판정용 윈도우 갱신
        with self._lock:
            w = self._sig_window[sig]
            w.append(now)
            cutoff = now - SPIKE_WINDOW_SEC
            while w and w[0] < cutoff:
                w.popleft()
            window_count = len(w)

        # 후보 종류 결정
        if is_critical:
            kind, ttl = 'critical', CRITICAL_COOLDOWN_SEC
        elif transient:
            if window_count < SPIKE_THRESHOLD:
                return None                      # 만성 잡음: 급증 아니면 무시
            kind, ttl = 'spike', SPIKE_COOLDOWN_SEC
        else:
            kind, ttl = 'normal', RECURRING_COOLDOWN_SEC

        # 쿨다운 확보(강한 dedup; 재시작 넘어 유지)
        if not self._cooldown_acquire(sig, ttl):
            return None

        # 일반 에러는 new/recurring 세분화
        if kind == 'normal':
            kind = 'new' if self._seen_classify(sig) else 'recurring'
        else:
            self._mark_seen(sig)

        # 전역 rate cap
        with self._lock:
            g = self._global
            gcut = now - GLOBAL_RATE_WINDOW_SEC
            while g and g[0] < gcut:
                g.popleft()
            if len(g) >= GLOBAL_RATE_MAX:
                self._suppressed += 1
                return None
            g.append(now)
            suppressed = self._suppressed
            self._suppressed = 0

        return {
            'sig': sig, 'kind': kind, 'window_count': window_count,
            'suppressed': suppressed, 'msg': msg,
        }


class SlackErrorAlertHandler(logging.Handler):
    """ERROR/CRITICAL 로그를 큐레이션해 관리자 슬랙으로 전송하는 로깅 핸들러."""

    def __init__(self, token: str, channel: str, gate: _AlertGate):
        super().__init__(level=_MIN_LEVEL)
        self._token = token
        self._channel = channel
        self._gate = gate
        self._q: "queue.Queue[str]" = queue.Queue(maxsize=QUEUE_MAX)
        self._local = threading.local()
        self._worker = threading.Thread(
            target=self._run, name='SlackErrorAlerter', daemon=True)
        self._worker.start()

    def emit(self, record: logging.LogRecord) -> None:
        # 재진입 가드: 이 핸들러(또는 워커) 안에서 나온 로그가 되돌아오지 않도록
        if getattr(self._local, 'in_emit', False):
            return
        try:
            self._local.in_emit = True
            decision = self._gate.decide(record)
            if not decision:
                return
            text = self._format(record, decision)
            try:
                self._q.put_nowait(text)
            except queue.Full:
                pass
        except Exception:
            pass  # 알림 기계는 절대 앱을 죽이지 않는다
        finally:
            self._local.in_emit = False

    def _format(self, record: logging.LogRecord, d: dict) -> str:
        head = _KIND_HEADER.get(d['kind'], '⛔ 시스템 에러')
        where = (f"{getattr(record, 'name', '?')} "
                 f"({getattr(record, 'module', '?')}:{getattr(record, 'lineno', '?')})")
        lines = [f"*{head}*", f"위치: `{where}`", f"메시지: {d['msg'][:400]}"]

        if record.exc_info and record.exc_info[0]:
            etype = record.exc_info[0].__name__
            eval_ = str(record.exc_info[1])[:200] if record.exc_info[1] else ''
            lines.append(f"예외: `{etype}: {eval_}`")

        ctx = []
        for attr in ('endpoint', 'method', 'user_email', 'project_code', 'operation'):
            v = getattr(record, attr, None)
            if v:
                ctx.append(f"{attr}={v}")
        if ctx:
            lines.append("컨텍스트: " + ", ".join(ctx))

        if d['kind'] == 'spike':
            lines.append(f"최근 5분 동일 유형 *{d['window_count']}건* — 외부 API 순단·과부하 의심")
        elif d['window_count'] > 1:
            lines.append(f"최근 5분 동일 유형 {d['window_count']}건")

        if d.get('suppressed'):
            lines.append(f"_(직전 rate cap 으로 {d['suppressed']}건 억제됨)_")

        if record.exc_info:
            try:
                tb = ''.join(_traceback.format_exception(*record.exc_info))
                tail = "\n".join(tb.strip().splitlines()[-8:])
                lines.append("```\n" + tail[:1200] + "\n```")
            except Exception:
                pass

        lines.append("_Sentry 전체기록 · 동일 유형 dedup_")
        return "\n".join(lines)

    def _run(self) -> None:
        while True:
            try:
                text = self._q.get()
            except Exception:
                continue
            try:
                self._post(text)
            except Exception:
                pass

    def _post(self, text: str) -> None:
        payload = json.dumps(
            {'channel': self._channel, 'text': text}).encode('utf-8')
        req = urllib.request.Request(
            'https://slack.com/api/chat.postMessage',
            data=payload,
            headers={
                'Authorization': f'Bearer {self._token}',
                'Content-Type': 'application/json; charset=utf-8',
            },
        )
        urllib.request.urlopen(req, timeout=5).read()


_installed = False
_lock = threading.Lock()


def install_error_slack_alerter() -> bool:
    """root 로거에 큐레이션 에러 알림 핸들러를 1회 부착. 성공 시 True."""
    global _installed
    with _lock:
        if _installed:
            return False

        if os.getenv('ERROR_SLACK_ALERTS_ENABLED', '1').strip().lower() in ('0', 'false', 'no'):
            logger.info('[ERROR_ALERT] 비활성화됨 (ERROR_SLACK_ALERTS_ENABLED=0)')
            return False
        if os.getenv('FLASK_ENV', '').strip().lower() == 'testing':
            return False

        token = os.getenv('SLACK_BOT_TOKEN', '').strip()
        channel = (os.getenv('ERROR_ALERT_CHANNEL', '').strip()
                   or os.getenv('SLACK_ADMIN_CHANNEL', '').strip())
        if not token or not channel:
            logger.warning('[ERROR_ALERT] SLACK_BOT_TOKEN/SLACK_ADMIN_CHANNEL 미설정 — 알림 skip')
            return False

        def _rget():
            from dashboard.utils.redis_client import get_redis_client
            return get_redis_client()

        gate = _AlertGate(redis_getter=_rget)
        handler = SlackErrorAlertHandler(token, channel, gate)
        logging.getLogger().addHandler(handler)
        _installed = True
        logger.info(f'[ERROR_ALERT] 큐레이션 에러 슬랙 알림 활성화 (대상={channel[:6]}…)')
        return True
