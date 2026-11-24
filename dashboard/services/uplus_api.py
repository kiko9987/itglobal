"""U+커넥트 Open API (2026-11-24 미리 만듦 — 9건 신청, 승인 대기 중).

규격서: 「OPEN API 연동규격서 Ver1.2」(U+커넥트 자료실). 요점:
- REST POST·JSON, 호스트 https://3.34.90.76:4009 (검증·운영 같음).
- /v1/session (application_id + secure_key) → token. 하루마다 다시 받는다.
- API 마다 application_id 가 따로 나온다(승인 번호). secure_key 는 승인 뒤 문자로 옴.
- 응답: {"message": ..., "extensions": {"code": "20000000", "totalRecord": n, "data": [...]}}.

키를 아직 안 받았으면(환경변수 없음) enabled() 가 False — 부르는 쪽이 아무것도 안 한다.

환경변수(.env)
- UPLUS_API_HOST           기본 https://3.34.90.76:4009
- UPLUS_SECURE_KEY         승인 문자로 받은 키(API 마다 다르면 UPLUS_KEY_<이름>)
- UPLUS_APP_<이름>         API 마다 받은 application_id. 이름 = 아래 APIS 의 키를 대문자로(UPLUS_APP_STATUS 등)
- UPLUS_VERIFY_SSL         기본 1. 호스트가 IP 라 인증서가 안 맞으면 개발기 시험 때 0 으로(확인 필요)

⚠️ 승인 뒤 개발기 시험에서 확인할 것: created_at 시각대(+00:00 표기지만 예시는 한국 시각으로 보임),
   key_state 값의 뜻(문서에 없음 — 그래서 운행 종료는 '한 자리에 오래 서 있음'으로 판단), 호출 한도.
"""
from __future__ import annotations

import os
import threading
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

DEFAULT_HOST = 'https://3.34.90.76:4009'
OK_CODE = '20000000'
TIMEOUT = 15

# 이름 → 경로(규격서 5. API 목록)
APIS = {
    'location': '/v1/openapi/getRentCarLocation',          # 차량 현위치(마지막 수신 위치)
    'status': '/v1/openapi/getBizRentCarStatus',           # 차량 상태(위치·속도·주행거리·키 상태…)
    'route': '/v1/openapi/getRentCarDrivingLoute',         # 차량 경로(1분)
    'safe': '/v1/openapi/getSafeDrivingSearch',            # 안전 운전
    'eco': '/v1/openapi/getEchoDrivingSearch',             # 경제 운전(사용 시작·종료 시각)
    'vehicles': '/v1/openapi/getVehiclesListSearch',       # 차량 목록
    'devices': '/v1/openapi/getFMSDevicesSearch',          # 단말 목록
    'device_status': '/v1/openapi/getFMSDevicesStatusSearch',  # 단말 상태
    'route_day': '/v1/openapi/getDrivingLoutePerDay',      # 차량 경로(하루)
}

_lock = threading.Lock()
_tokens: Dict[str, Tuple[str, str]] = {}      # 이름 → (받은 날, token)


class UplusError(Exception):
    pass


def _host() -> str:
    return (os.getenv('UPLUS_API_HOST') or DEFAULT_HOST).rstrip('/')


def app(name: str) -> Optional[Tuple[int, str]]:
    """그 API 의 (application_id, secure_key). 아직 없으면 None."""
    raw = (os.getenv(f'UPLUS_APP_{name.upper()}') or '').strip()
    key = (os.getenv(f'UPLUS_KEY_{name.upper()}') or os.getenv('UPLUS_SECURE_KEY') or '').strip()
    if not raw or not key:
        return None
    try:
        return int(raw), key
    except ValueError:
        logger.warning(f'[U+] UPLUS_APP_{name.upper()} 가 숫자가 아니에요')
        return None


def enabled(name: str = 'status') -> bool:
    return name in APIS and app(name) is not None


def _verify() -> bool:
    return (os.getenv('UPLUS_VERIFY_SSL', '1').strip().lower() not in ('0', 'false', 'no'))


def _post(path: str, body: Dict) -> Dict:
    import requests
    verify = _verify()
    if not verify:
        try:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        except Exception:
            pass
    r = requests.post(_host() + path, json=body, timeout=TIMEOUT, verify=verify,
                      headers={'Content-Type': 'application/json;charset=UTF-8', 'Accept': 'application/json'})
    try:
        j = r.json()
    except ValueError:
        raise UplusError(f'{path} HTTP {r.status_code} (JSON 아님)')
    if not isinstance(j, dict):
        raise UplusError(f'{path} 응답 형식이 이상해요')
    j['_http'] = r.status_code
    return j


def _token(name: str, fresh: bool = False) -> str:
    a = app(name)
    if not a:
        raise UplusError(f'{name} API 키가 없어요')
    today = date.today().isoformat()
    with _lock:
        got = _tokens.get(name)
        if got and got[0] == today and not fresh:
            return got[1]
        j = _post('/v1/session', {'application_id': a[0], 'secure_key': a[1]})
        ext = j.get('extensions') or {}
        data = ext.get('data') or []
        tok = (data[0] or {}).get('token') if data and isinstance(data[0], dict) else ''
        if str(ext.get('code')) != OK_CODE or not tok:
            raise UplusError(f"{name} 인증 실패: {j.get('message') or ext.get('code')}")
        _tokens[name] = (today, tok)
        return tok


def call(name: str, body: Optional[Dict] = None) -> Dict:
    """API 한 번 — extensions 를 돌려준다. 토큰이 안 맞으면 한 번 새로 받아 다시."""
    if name not in APIS:
        raise UplusError(f'모르는 API: {name}')
    a = app(name)
    if not a:
        raise UplusError(f'{name} API 키가 없어요')
    for attempt in (0, 1):
        req = dict(body or {}, token=_token(name, fresh=bool(attempt)), application_id=a[0])
        j = _post(APIS[name], req)
        ext = j.get('extensions') or {}
        code = str(ext.get('code') or '')
        if code == OK_CODE:
            return ext
        msg = str(j.get('message') or '')
        if attempt == 0 and ('토큰' in msg or 'token' in msg.lower() or j.get('_http') == 401):
            continue
        if '없' in msg and code == '70000009':     # '조회 결과가 없습니다'
            return {'code': code, 'data': [], 'totalRecord': 0}
        raise UplusError(f'{name} 실패: {msg or code}')
    raise UplusError(f'{name} 실패')


def _num(x) -> Optional[float]:
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    return v


def parse_time(s) -> Optional[datetime]:
    """'2025-05-16 15:17:00.0' · '2025-03-17T20:04:27.000+00:00' → 한국 시각(naive).
    +00:00 표기지만 규격서 예시는 한국 시각으로 보여 시각대를 바꾸지 않는다(승인 뒤 확인)."""
    s = str(s or '').strip()
    if not s:
        return None
    s = s.replace('T', ' ')
    for cut in ('+', 'Z'):
        if cut in s[10:]:
            s = s[:10] + s[10:].split(cut)[0]
    s = s.split('.')[0]
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def car_status() -> List[Dict]:
    """모든 차 상태 — [{'plate', 'lat', 'lon', 'at'(datetime|None), 'speed', 'key_state', 'odometer'}].
    위치가 없거나 0 인 차는 lat/lon None."""
    out: List[Dict] = []
    page, limit = 1, 50
    while page <= 10:
        ext = call('status', {'page': page, 'limit': limit})
        data = ext.get('data') or []
        for d in data:
            if not isinstance(d, dict):
                continue
            lat, lon = _num(d.get('latitude')), _num(d.get('longitude') or d.get('logitude'))
            ok = lat and lon and abs(lat) > 1 and abs(lon) > 1
            out.append({
                'plate': str(d.get('vehicle_number') or '').replace(' ', ''),
                'lat': lat if ok else None, 'lon': lon if ok else None,
                'at': parse_time(d.get('created_at')),
                'speed': _num(d.get('velocity')),
                'key_state': str(d.get('key_state') or ''),
                'odometer': _num(d.get('cumulative_driving_distance') or d.get('odometer')),
            })
        total = int(_num(ext.get('totalRecord')) or 0)
        if len(data) < limit or page * limit >= total:
            break
        page += 1
    return out
