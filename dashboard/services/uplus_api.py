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

시각대(2026-11-25 개발기 시험으로 확인): 모두 '+00:00' 표기로 오지만 실제 기준이 API 마다 다르다.
   차량 상태(created_at)·경제/안전 운전(use_start_dt·use_end_dt) = 세계 표준시 → 9시간 더함(UTC_APIS).
   경로(하루, gpsDttm·created_at) = 한국 시각 그대로. 근거: 같은 운행이 경제 운전 09:05~09:57, 경로 18:05~18:57.
   단말 상태(disability_dt) = 한국 시각 그대로. 운영도 개발기와 같은 기준(11-25 U+ 문의 No.629 답변).
위치: 차량 상태 응답엔 위도·경도 칸이 없다(개발기) → 위치 API(getRentCarLocation)를 차 번호로 붙인다.
⚠️ 아직 확인할 것: key_state 값의 뜻(문서에 없음 — 그래서 운행 종료는 '한 자리에 오래 서 있음'으로 판단).
호출 주기: U+ 권고 1분 이상, 짧게 부르면 막힐 수 있음(No.629 답변). 지금은 화면 60초 캐시·스케줄 10~30분·경로 10분 캐시라 맞음.
"""
from __future__ import annotations

import os
import threading
from datetime import date, datetime, timedelta
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


# 잘못 적은 설정을 이미 ERROR 로 알린 값 — enabled() 가 10분마다·화면마다 불려 같은 값으로 알림이 쌓이지 않게
_bad_logged: set = set()


def config_error(name: str = 'status') -> str:
    """그 API 의 설정이 잘못 적혔으면 그 문구, 아니면 ''. 키가 아직 없는 것(승인 전)은 오류가 아니다.
    UPLUS_APP_<이름> 을 숫자가 아니게 적으면 예전엔 WARNING 만 남기고 '승인 전'과 똑같이 보여, 키를 넣고도
    아무것도 안 되는 걸 아무도 몰랐다(11-24 조용한 실패 점검)."""
    raw = (os.getenv(f'UPLUS_APP_{name.upper()}') or '').strip()
    key = (os.getenv(f'UPLUS_KEY_{name.upper()}') or os.getenv('UPLUS_SECURE_KEY') or '').strip()
    if not raw or not key:
        return ''
    try:
        int(raw)
        return ''
    except ValueError:
        return f'UPLUS_APP_{name.upper()} 가 숫자가 아니에요'


def config_errors() -> List[str]:
    """잘못 적힌 U+ 설정 전부(화면 머리 '설정 오류' 표시용)."""
    return [e for e in (config_error(n) for n in APIS) if e]


def app(name: str) -> Optional[Tuple[int, str]]:
    """그 API 의 (application_id, secure_key). 아직 없거나 잘못 적었으면 None(잘못 적은 건 ERROR — config_error)."""
    raw = (os.getenv(f'UPLUS_APP_{name.upper()}') or '').strip()
    key = (os.getenv(f'UPLUS_KEY_{name.upper()}') or os.getenv('UPLUS_SECURE_KEY') or '').strip()
    if not raw or not key:
        return None
    try:
        return int(raw), key
    except ValueError:
        if (name, raw) not in _bad_logged:
            _bad_logged.add((name, raw))
            logger.error(f'[U+] UPLUS_APP_{name.upper()} 가 숫자가 아니에요 — 이 API 는 키가 없는 것처럼 꺼져 있어요. '
                         f'.env 를 고친 뒤 재시작해 주세요', force_alert=True)
        return None


def enabled(name: str = 'status') -> bool:
    return name in APIS and app(name) is not None


def state(name: str = 'status') -> str:
    """'on'(키 있음) · 'config_error'(잘못 적음 — 관리자 확인 필요) · 'off'(승인 전, 키 없음)."""
    if name not in APIS:
        return 'off'
    if config_error(name):
        app(name)               # ERROR 한 번 남기기
        return 'config_error'
    return 'on' if app(name) is not None else 'off'


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


KST = timedelta(hours=9)
# 시각이 세계 표준시로 오는 API. 11-27 운영 확인: 경제·안전 운전만 세계 표준시(22:34 UTC 시작 = 경로 07:34),
#   차량 상태는 한국 시각(11:03 에 11:03:46 으로 옴 — 개발기 11-25 엔 세계 표준시였음). 바뀌면 .env UPLUS_UTC_APIS=eco,safe 처럼.
UTC_APIS_DEFAULT = ('eco', 'safe')


def _utc(api: str) -> bool:
    raw = os.getenv('UPLUS_UTC_APIS')
    names = UTC_APIS_DEFAULT if raw is None else tuple(x.strip() for x in raw.split(',') if x.strip())
    return api in names


def parse_time(s, api: str = '') -> Optional[datetime]:
    """'2025-05-16 15:17:00.0' · '2025-03-17T20:04:27.000+00:00' → 한국 시각(naive).
    표기(+00:00)는 믿지 않는다 — api 가 세계 표준시로 오는 API(_utc)면 9시간 더하고, 아니면 그대로."""
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
            t = datetime.strptime(s, fmt)
            return t + KST if api and _utc(api) else t
        except ValueError:
            pass
    return None


def _pages(name: str, body: Dict, limit: int = 100, max_pages: int = 50) -> List[Dict]:
    """여러 쪽에 걸친 목록을 모두."""
    out: List[Dict] = []
    for page in range(1, max_pages + 1):
        ext = call(name, dict(body, page=page, limit=limit))
        data = ext.get('data') or []
        if isinstance(data, dict):
            data = [data]
        out += [d for d in data if isinstance(d, dict)]
        total = int(_num(ext.get('totalRecord')) or 0)
        if len(data) < limit or (total and page * limit >= total):
            break
    return out


def _plate(d: Dict) -> str:
    return str(d.get('vehicle_number') or '').replace(' ', '')


def _int(x) -> int:
    v = _num(x)
    return int(v) if v is not None else 0


SAFE_EVENTS = (('rapid_acc', 'rapid_acceleration'), ('rapid_dec', 'rapid_deceleration'), ('sudden_stop', 'sudden_stop'),
               ('rapid_turn', 'rapid_turn'), ('rapid_u_turn', 'rapid_u_turn'), ('rapid_lane', 'rapid_lane_change'),
               ('rapid_pass', 'rapid_passing'))


def trips(plate: str, day: str) -> List[Dict]:
    """하루 운행 — 경제운전(사용 시작·끝, 주행 전·후 계기판, 거리, 연료) + 안전운전(급가속 등 횟수)을 시작 시각으로 합침.
    [{'start', 'end'(datetime), 'odo_start', 'odo_end', 'km', 'fuel_ml', 'eco', 'safe', 'avg_speed', 'max_speed',
      'rapid_acc', 'rapid_dec', 'sudden_stop', 'rapid_turn', 'rapid_u_turn', 'rapid_lane', 'rapid_pass'}]
    시작·끝이 같고 거리가 0 인 줄(시동만 켬)은 뺀다.
    시각이 세계 표준시라 한국 날짜 하루(전날 15시~그날 15시 UTC)가 조회 범위를 넘는다 — U+ 가 범위를 어느 시각대로
    거르는지 몰라 전날 0시부터 넓게 받고, 한국 시각으로 바꾼 시작 날짜로 거른다(11-25)."""
    prev = (date.fromisoformat(day) - timedelta(days=1)).isoformat()
    rng = {'vehicle_number': plate, 'start_date': f'{prev} 00:00:00', 'end_date': f'{day} 23:59:59'}
    eco = _pages('eco', rng) if enabled('eco') else []
    safe = {str(d.get('use_start_dt')): d for d in (_pages('safe', rng) if enabled('safe') else [])}
    out = []
    for d in eco:
        st, en = parse_time(d.get('use_start_dt'), 'eco'), parse_time(d.get('use_end_dt'), 'eco')
        if not st or st.date().isoformat() != day:
            continue
        o1, o2 = _num(d.get('before_odometer')), _num(d.get('after_odometer'))
        km = _num(d.get('mileage'))
        if not km and o1 is not None and o2 is not None:
            km = round(abs(o2 - o1), 1)
        km = km or 0.0
        if (not en or en <= st) and km < 0.1:
            continue
        s = safe.get(str(d.get('use_start_dt'))) or {}
        t = {
            'start': st, 'end': en, 'odo_start': o1, 'odo_end': o2, 'km': round(km, 1),
            'fuel_ml': _num(d.get('fuel_consumption')), 'eco': _num(d.get('echo_driving_index')),
            'safe': _num(s.get('safe_driving_index')),
            'avg_speed': _num(d.get('avg_speed') if d.get('avg_speed') is not None else s.get('avg_speed')),
            'max_speed': max(_num(d.get('max_speed')) or 0, _num(s.get('max_speed')) or 0),
        }
        # 안전운전 짝이 없는 운행(안전운전 API 실패·승인 전·시작 시각이 안 맞음)은 급가속 등을 모르는 것 — 0 회로 넣으면
        # 운전 습관 화면에 '급가속 0' 처럼 깨끗한 운전으로 보였다(11-24 조용한 실패 점검). 모르면 None(화면 '-').
        for k, src in SAFE_EVENTS:
            t[k] = _int(s.get(src)) if s else None
        out.append(t)
    out.sort(key=lambda t: t['start'])
    return out


def route_day(plate: str, day: str) -> List[Tuple[datetime, float, float]]:
    """하루 경로(5초마다 좌표) — [(시각, 위도, 경도)] 시각 순, 같은 점 겹침 뺌."""
    ext = call('route_day', {'vehicle_number': plate, 'select_day': day})
    data = ext.get('data') or {}
    if isinstance(data, list):
        data = data[0] if data else {}
    pts, last = [], None
    for p in (data.get('location') or []):
        t = parse_time(p.get('gpsDttm') or p.get('created_at'), 'route_day')
        lat, lon = _num(p.get('latitude')), _num(p.get('longitude') or p.get('logitude'))
        if not t or not lat or not lon or abs(lat) < 1:
            continue
        key = (t, lat, lon)
        if key == last:
            continue
        last = key
        pts.append(key)
    pts.sort(key=lambda x: x[0])
    return pts


def _location_map() -> Dict[str, Tuple[float, float]]:
    """위치 API — 차 번호 → (위도, 경도). 같은 번호가 여러 줄이면 첫 줄(개발기 시험 자료에 있음). 시각 칸은 없다."""
    out: Dict[str, Tuple[float, float]] = {}
    for d in _pages('location', {}):
        lat, lon = _num(d.get('latitude')), _num(d.get('longitude') or d.get('logitude'))
        if lat and lon and abs(lat) > 1 and abs(lon) > 1:
            out.setdefault(_plate(d), (lat, lon))
    return out


def locations() -> List[Dict]:
    """모든 차 마지막 위치 — [{'plate', 'lat', 'lon', 'at', 'speed'}] (차량 상태 + 위치 API 를 합친 것)."""
    return [{'plate': s['plate'], 'lat': s['lat'], 'lon': s['lon'], 'at': s['at'], 'speed': s['speed']}
            for s in car_status() if s['lat'] is not None]


def vehicles() -> List[Dict]:
    """U+ 에 등록된 차 — [{'plate', 'model_id', 'odometer', 'year', 'driver', 'imei'}]."""
    return [{'plate': _plate(d), 'model_id': d.get('vehicle_model_id'), 'odometer': _num(d.get('cumulative_driving_distance')),
             'year': f"{d.get('vehicle_year_year') or ''}", 'driver': d.get('driver_name') or '', 'imei': d.get('imei') or ''}
            for d in _pages('vehicles', {})]


def device_status(day: str) -> List[Dict]:
    """단말 상태(그날) — [{'plate', 'model', 'status', 'fault_at'(datetime|None), 'memo'}]."""
    rows = _pages('device_status', {'start_date': f'{day} 00:00:00', 'end_date': f'{day} 23:59:59'})
    return [{'plate': _plate(d), 'model': d.get('model_name') or d.get('model_no') or '', 'status': str(d.get('status') or ''),
             'fault_at': parse_time(d.get('disability_dt'), 'device_status'), 'memo': d.get('error_memo') or ''} for d in rows]


def car_status() -> List[Dict]:
    """모든 차 상태 — [{'plate', 'lat', 'lon', 'at'(datetime|None), 'speed', 'key_state', 'odometer',
    'fuel', 'voltage', 'tpms'(4칸), 'idle_sec'}]. 위치가 없거나 0 인 차는 lat/lon None."""
    out: List[Dict] = []
    for d in _pages('status', {}, limit=50, max_pages=10):
        lat, lon = _num(d.get('latitude')), _num(d.get('longitude') or d.get('logitude'))
        ok = lat and lon and abs(lat) > 1 and abs(lon) > 1
        out.append({
            'plate': _plate(d),
            'lat': lat if ok else None, 'lon': lon if ok else None,
            'at': parse_time(d.get('created_at'), 'status'),
            'speed': _num(d.get('velocity')),
            'key_state': str(d.get('key_state') or ''),
            'odometer': _num(d.get('cumulative_driving_distance') or d.get('odometer')),
            'fuel': _num(d.get('fuel_level')), 'voltage': _num(d.get('device_voltage')),
            'tpms': [_num(d.get(k)) for k in ('tpms_fl', 'tpms_fr', 'tpms_rl', 'tpms_rr')],
            'idle_sec': _num(d.get('idling_seconds')), 'soc': _num(d.get('soc')),
        })
    # 차량 상태엔 위도·경도가 안 온다(11-25 개발기: 500대 중 0대) → 위치 API 를 차 번호로 붙인다.
    #   위치 API 가 실패해도 상태(연료·전압·시동)는 그대로 돌려준다 — 위치만 비고, 남는 기록으로 알 수 있게.
    if out and enabled('location') and any(s['lat'] is None for s in out):
        try:
            locs = _location_map()
        except Exception as e:
            logger.warning(f'[U+] 위치 API 실패 — 이번엔 위치 없이 상태만: {e}')
            locs = {}
        for s in out:
            if s['lat'] is None and s['plate'] in locs:
                s['lat'], s['lon'] = locs[s['plate']]
    return out
