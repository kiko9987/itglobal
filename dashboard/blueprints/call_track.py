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

    ip = request.remote_addr or ''   # ProxyFix 가 실제 접속 IP 로 맞춰 줌 — X-Forwarded-For 첫 값은 위조 가능 (2026-10-29)
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


# ─────────────────────────────────────────────────────────────
# 홈페이지 카카오 상담 버튼 클릭 (2026-11-08 L-04501 계기)
# ─────────────────────────────────────────────────────────────
# 카카오톡 문의는 채널톡을 거쳐 오는데, 카카오는 고객이 어디서 왔는지(페이지·광고)를 넘겨주지 않는다.
# 그래서 홈페이지 카카오 버튼(pf.kakao.com 링크)을 누른 순간 유입경로를 따로 남기고, 몇 분 안에 들어온
# 새 카카오 문의와 시각으로 짝짓는다(channeltalk._kakao_inflow). 전화 클릭과 섞이지 않게 저장소를 나눔.
KAKAO_CLICK_INDEX = 'kakao_click:index'
KAKAO_MATCH_BEFORE = 15 * 60   # 버튼 누르고 첫 메시지까지 — 카카오 앱 열림·채널 추가·입력 시간
KAKAO_MATCH_AFTER = 2 * 60     # 서버·채널톡 시각 차이 여유


@call_track_bp.route('/kakao-click', methods=['POST', 'OPTIONS'])
def kakao_click():
    if request.method == 'OPTIONS':
        return _cors(jsonify({}))
    origin = request.headers.get('Origin', '')
    if origin and origin not in _ALLOWED_ORIGINS:
        return _cors(jsonify({'ok': False})), 403
    rec = parse_click_payload(request.get_data(cache=False))
    if rec is None:
        return _cors(jsonify({'ok': False})), 400

    ip = request.remote_addr or ''
    who = hashlib.sha256(f"{ip}|{request.headers.get('User-Agent', '')}".encode()).hexdigest()[:16]
    now = time.time()
    try:
        r = get_redis_client().redis
        if not r.set(f'kakao_click:dedup:{who}', '1', nx=True, ex=_DEDUP_TTL):
            return _cors(jsonify({'ok': True, 'dup': True}))
        cid = uuid.uuid4().hex[:12]
        rec['ts'] = str(int(now))
        key = f'kakao_click:{cid}'
        pipe = r.pipeline()
        pipe.hset(key, mapping=rec)
        pipe.expire(key, _RECORD_TTL)
        pipe.zadd(KAKAO_CLICK_INDEX, {cid: now})
        pipe.zremrangebyscore(KAKAO_CLICK_INDEX, 0, now - _RECORD_TTL)
        pipe.execute()
        logger.info(f"카카오 버튼 클릭 기록: inflow={rec['inflow'] or '-'} ref={rec['ref'] or '-'} page={rec['page']}")
    except Exception as exc:
        logger.warning(f"카카오 버튼 클릭 기록 실패: {exc}")
        return _cors(jsonify({'ok': False})), 503
    return _cors(jsonify({'ok': True}))


def kakao_click_marker(rec: dict) -> str:
    """클릭 기록 → [유입:…] 마커 본문. 홈페이지 폼의 숨은 '유입경로' 칸과 같은 규칙
    (광고 표시 → 그대로, 없으면 이전 사이트 ref/<도메인>, 그것도 없으면 direct)."""
    from dashboard.services.lead_helpers import normalize_inflow
    inflow = normalize_inflow(rec.get('inflow') or '')
    if inflow:
        return inflow
    ref = re.sub(r'[/|\s]', '', str(rec.get('ref') or ''))[:30]
    return normalize_inflow(f'ref/{ref}') if ref else 'direct'


def claim_kakao_click(at_ts: float, claimer: str) -> dict:
    """새 카카오 문의 시각(at_ts) 앞뒤의 홈페이지 카카오 버튼 기록을 찾아 하나를 차지한다.

    반환 {'status': 'match'|'none'|'ambiguous', 'marker': str, 'cid': str, 'n': int}
      - match: 아직 다른 문의가 안 가져간 기록이 있고, 여러 건이어도 유입이 모두 같을 때 → 가장 최근 것
      - ambiguous: 유입이 서로 다른 기록이 겹침 → 정하지 않음(틀린 유입을 붙이느니 비움)
      - none: 기록 없음 = 홈페이지 버튼을 거치지 않은 문의(카카오톡 검색·채널 추가 등)
    """
    out = {'status': 'none', 'marker': '', 'cid': '', 'n': 0}
    try:
        r = get_redis_client().redis
        cands = []
        for cid in r.zrangebyscore(KAKAO_CLICK_INDEX, at_ts - KAKAO_MATCH_BEFORE, at_ts + KAKAO_MATCH_AFTER):
            cid = cid.decode() if isinstance(cid, bytes) else cid
            h = {(k.decode() if isinstance(k, bytes) else k): (v.decode() if isinstance(v, bytes) else v)
                 for k, v in (r.hgetall(f'kakao_click:{cid}') or {}).items()}
            if h and not h.get('claim'):
                cands.append((float(h.get('ts') or 0), cid, h))
        out['n'] = len(cands)
        if not cands:
            return out
        markers = {kakao_click_marker(h) for _, _, h in cands}
        if len(markers) > 1:
            out['status'] = 'ambiguous'
            return out
        for _, cid, h in sorted(cands, key=lambda x: x[0], reverse=True):
            if r.hsetnx(f'kakao_click:{cid}', 'claim', claimer):
                out.update(status='match', marker=kakao_click_marker(h), cid=cid)
                return out
    except Exception as exc:
        logger.warning(f'카카오 버튼 기록 짝짓기 실패: {exc}')
    return out


def tag_kakao_click(cid: str, lead_no: str) -> None:
    """짝지은 클릭 기록에 리드 번호를 남김 (나중에 버튼 → 문의 → 계약 분석용)."""
    if not cid or not lead_no:
        return
    try:
        get_redis_client().redis.hset(f'kakao_click:{cid}', 'lead', lead_no)
    except Exception as exc:
        logger.debug(f'카카오 버튼 기록 리드 표시 실패: {exc}')


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
