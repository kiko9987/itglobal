# -*- coding: utf-8 -*-
"""홈페이지 전화 버튼 클릭 측정 — 아임웹 스니펫 sendBeacon → Redis 기록.

목적: 전화 문의는 광고 전환(폼 제출)에 안 잡힘 → 모바일 광고 판단이 왜곡됨.
홈페이지의 tel: 링크 클릭 순간 유입경로(스니펫이 이미 아는 utm/네이버 키워드/gclid)를
같이 보내 '어느 광고에서 온 사람이 전화 버튼을 눌렀는지' 를 서버에 남긴다.

- POST /track/call-click  (body = JSON 문자열, sendBeacon 은 text/plain 으로 보냄 → CORS preflight 없음)
- 로그인 세션 없는 공개 엔드포인트: security_middleware 가 /track/ 는 CSRF 우회(rate limit 은 유지).
- 개인정보 없음: 유입·페이지 경로·버튼 종류·모바일 여부만. IP 는 저장하지 않고 중복 억제 키로만 해시 사용.
- 저장: zset call_click:index (score=epoch) + hash call_click:{id}, TTL 400일.
"""

import hashlib
import json
import re
import time
import uuid

from flask import Blueprint, jsonify, request

from dashboard.utils.logging_config import get_logger
from dashboard.utils.redis_client import get_redis_client

logger = get_logger(__name__)

call_track_bp = Blueprint('call_track', __name__, url_prefix='/track')

CALL_CLICK_INDEX = 'call_click:index'
_RECORD_TTL = 60 * 60 * 24 * 400   # 1년+ (1년 단위 분석 원칙)
_DEDUP_TTL = 60                    # 같은 기기가 1분 안에 여러 번 눌러도 1건
_MAX_BODY = 2048
_ALLOWED_ORIGINS = ('https://www.itg-aircon.com', 'https://itg-aircon.com')

_SAFE_RE = re.compile(r'[^0-9A-Za-z가-힣_\-./: ]')


def _clean(value, limit: int) -> str:
    """제어문자·태그 제거 + 길이 제한 (슬랙/리포트 노출 시 주입 방지)."""
    if value is None:
        return ''
    s = _SAFE_RE.sub('', str(value)).strip()
    return s[:limit]


def parse_click_payload(raw: bytes) -> dict | None:
    """비콘 body → 저장할 필드. 형식이 틀리면 None."""
    if not raw or len(raw) > _MAX_BODY:
        return None
    try:
        data = json.loads(raw.decode('utf-8'))
    except (ValueError, UnicodeDecodeError):
        return None
    if not isinstance(data, dict):
        return None
    return {
        'inflow': _clean(data.get('inflow'), 130),
        'gclid': _clean(data.get('gclid'), 120),
        'ref': _clean(data.get('ref'), 60),
        'page': _clean(data.get('page'), 80) or '/',
        'btn': _clean(data.get('btn'), 60),
        'mobile': '1' if data.get('mobile') in (True, 1, '1', 'true') else '0',
    }


def _cors(resp):
    origin = request.headers.get('Origin', '')
    if origin in _ALLOWED_ORIGINS:
        resp.headers['Access-Control-Allow-Origin'] = origin
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        resp.headers['Vary'] = 'Origin'
    return resp


@call_track_bp.route('/call-click', methods=['POST', 'OPTIONS'])
def call_click():
    if request.method == 'OPTIONS':
        return _cors(jsonify({}))
    origin = request.headers.get('Origin', '')
    if origin and origin not in _ALLOWED_ORIGINS:
        return _cors(jsonify({'ok': False})), 403
    rec = parse_click_payload(request.get_data(cache=False))
    if rec is None:
        return _cors(jsonify({'ok': False})), 400

    ip = (request.headers.get('X-Forwarded-For') or request.remote_addr or '').split(',')[0].strip()
    who = hashlib.sha256(f"{ip}|{request.headers.get('User-Agent', '')}".encode()).hexdigest()[:16]
    now = time.time()
    try:
        r = get_redis_client().redis
        if not r.set(f'call_click:dedup:{who}', '1', nx=True, ex=_DEDUP_TTL):
            return _cors(jsonify({'ok': True, 'dup': True}))
        cid = uuid.uuid4().hex[:12]
        rec['ts'] = str(int(now))
        key = f'call_click:{cid}'
        pipe = r.pipeline()
        pipe.hset(key, mapping=rec)
        pipe.expire(key, _RECORD_TTL)
        pipe.zadd(CALL_CLICK_INDEX, {cid: now})
        pipe.zremrangebyscore(CALL_CLICK_INDEX, 0, now - _RECORD_TTL)
        pipe.execute()
        logger.info(f"전화 클릭 기록: inflow={rec['inflow'] or '-'} page={rec['page']} mobile={rec['mobile']}")
    except Exception as exc:
        logger.warning(f"전화 클릭 기록 실패: {exc}")
        return _cors(jsonify({'ok': False})), 503
    return _cors(jsonify({'ok': True}))


def load_call_clicks(since_ts: float, until_ts: float) -> list[dict]:
    """기간 내 전화 클릭 레코드 (리포트용, 읽기 전용)."""
    r = get_redis_client().redis
    out = []
    for cid in r.zrangebyscore(CALL_CLICK_INDEX, since_ts, until_ts):
        cid = cid.decode() if isinstance(cid, bytes) else cid
        h = r.hgetall(f'call_click:{cid}')
        if h:
            out.append({(k.decode() if isinstance(k, bytes) else k): (v.decode() if isinstance(v, bytes) else v)
                        for k, v in h.items()})
    return out
