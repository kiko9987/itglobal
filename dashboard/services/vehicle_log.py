"""차량 운행 (2026-11-20 부장님) — U+커넥트 운행일지를 올려 차량별 운행 기록을 모으고,
같은 날 같은 지역 방문 일정과 맞춰 '그날 이 차를 쓴 사람'을 추정한다.

- U+커넥트 Open API(고객지원 › OpenAPI 신청)는 11-20 9건 신청해 승인 대기 — 그전까지(그리고 API 가 멈췄을 때) [운행일지 일괄 다운로드]
  zip(차량마다 국세청 양식 xlsx 1개) 또는 xlsx 하나를 그대로 올린다. 같은 운행(차량·날짜·출발 시각)은 한 번만 들어간다.
- U+커넥트 운행일지의 운전자(성명) 칸은 비어 있고 차량은 그날그날 바뀐다(부장님) → 리드 시트의
  방문 예정일·방문 주소·영업 담당자로 추정한다. 차량 배정(services/vehicle_assign)이 있으면 그 값이 먼저. 방문 주소는 도로명이라 동이 없어 '시·구' 단위로 맞춘다.
- 회사(김포 걸포동)로 들어오고 나가는 운행은 대조하지 않는다 — 김포 방문과 늘 겹친다.
- 위치는 U+커넥트가 준 동 단위 출발지·도착지뿐. 보는 사람은 인사 관리자(경영지원)만(blueprints/gw_vehicles.py).
"""
from __future__ import annotations

import io
import re
import zipfile
from collections import Counter
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, Iterable, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

HOME_AREA = ('김포', '', '걸포동')     # 회사 — 경기도 김포시 걸포로 134-84(걸포동)
MAX_FILES = 60
MAX_FILE_BYTES = 5 * 1024 * 1024
MAX_UPLOAD_BYTES = 30 * 1024 * 1024

_DDL = [
    '''CREATE TABLE IF NOT EXISTS vehicle_trips (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        plate TEXT NOT NULL,
        day TEXT NOT NULL,
        start_time TEXT NOT NULL,
        end_time TEXT,
        odo_start REAL,
        odo_end REAL,
        km REAL NOT NULL DEFAULT 0,
        from_addr TEXT,
        to_addr TEXT,
        commute_km REAL NOT NULL DEFAULT 0,
        business_km REAL NOT NULL DEFAULT 0,
        note TEXT,
        upload_id INTEGER,
        UNIQUE (plate, day, start_time)
    )''',
    'CREATE INDEX IF NOT EXISTS idx_vehicle_trips_day ON vehicle_trips(day)',
    '''CREATE TABLE IF NOT EXISTS vehicle_addr_dong (
        addr TEXT PRIMARY KEY,
        dong TEXT,
        checked_at TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicles (
        plate TEXT PRIMARY KEY,
        model TEXT,
        updated_at TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_uploads (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        file_name TEXT,
        period TEXT,
        vehicles INTEGER NOT NULL DEFAULT 0,
        added INTEGER NOT NULL DEFAULT 0,
        skipped INTEGER NOT NULL DEFAULT 0,
        uploaded_by TEXT,
        uploaded_at TEXT
    )''',
]
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


@contextmanager
def _conn():
    global _table_ready
    with _db()._get_connection() as conn:
        if not _table_ready:
            for ddl in _DDL:
                conn.execute(ddl)
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


# ── 엑셀 읽기 ──────────────────────────────────────────────────────────
_DAY_RE = re.compile(r'^(\d{4}-\d{2}-\d{2})')
_TIME_RE = re.compile(r'^\d{1,2}:\d{2}(:\d{2})?$')
NOT_LOG = 'U+커넥트 운행일지(국세청) 엑셀이 아니에요. U+커넥트 [운행일지 일괄 다운로드]로 받은 파일을 올려 주세요.'


def _num(v) -> float:
    s = re.sub(r'[^\d.\-]', '', str(v if v is not None else ''))
    try:
        return round(float(s), 1) if s not in ('', '-', '.') else 0.0
    except ValueError:
        return 0.0


def _s(v) -> str:
    return '' if v is None else str(v).strip()


def _below(rows: List[tuple], label: str) -> str:
    """머리 칸(label) 바로 아래 줄의 같은 칸 값 — 차종·자동차등록번호."""
    for i, r in enumerate(rows[:12]):
        for j, v in enumerate(r):
            if _s(v) == label and i + 1 < len(rows) and j < len(rows[i + 1]):
                return _s(rows[i + 1][j])
    return ''


def parse_sheet(rows: List[tuple]) -> Dict:
    """운행일지 한 장(한 차량) → {'plate','model','period','trips':[...]}. 양식이 아니면 ValueError."""
    rows = [tuple(r) for r in rows]
    plate = _below(rows, '자동차등록번호')
    model = _below(rows, '차종')
    period = _s(rows[0][1]) if rows and len(rows[0]) > 1 and _s(rows[0][0]) == '과세기간' else ''
    head = next((i for i, r in enumerate(rows) if r and _s(r[0]).startswith('사용일자')), None)
    if not plate or head is None:
        raise ValueError(NOT_LOG)
    trips = []
    for r in rows[head + 1:]:
        r = tuple(r) + (None,) * (13 - len(r))
        m = _DAY_RE.match(_s(r[0]))
        if not m:
            continue                        # 머리 둘째 줄(출퇴근용·일반 업무용)·합계 줄
        start = _s(r[8])
        if not _TIME_RE.match(start):
            continue
        trips.append({
            'day': m.group(1), 'start_time': start, 'end_time': _s(r[9]),
            'odo_start': _num(r[3]), 'odo_end': _num(r[4]), 'km': _num(r[5]),
            'from_addr': _s(r[6])[:80], 'to_addr': _s(r[7])[:80],
            'commute_km': _num(r[10]), 'business_km': _num(r[11]), 'note': _s(r[12])[:200],
        })
    return {'plate': plate[:20], 'model': model[:40], 'period': period[:40], 'trips': trips}


def _read_xlsx(data: bytes) -> Dict:
    import openpyxl
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception:
        raise ValueError(NOT_LOG)
    try:
        return parse_sheet(list(wb.worksheets[0].iter_rows(values_only=True)))
    finally:
        wb.close()


def _zip_name(info: zipfile.ZipInfo) -> str:
    n = info.filename
    if not info.flag_bits & 0x800:          # UTF-8 표시가 없으면 윈도 한글(cp949)일 수 있음
        try:
            n = n.encode('cp437').decode('cp949')
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return n


def parse_upload(file_name: str, data: bytes) -> List[Dict]:
    """올린 파일(zip 또는 xlsx) → 차량별 운행일지 목록."""
    if not data:
        raise ValueError('빈 파일이에요.')
    if len(data) > MAX_UPLOAD_BYTES:
        raise ValueError('파일이 너무 커요(30MB까지).')
    if zipfile.is_zipfile(io.BytesIO(data)):
        z = zipfile.ZipFile(io.BytesIO(data))
        if '[Content_Types].xml' in z.namelist():     # xlsx 도 속은 zip — 엑셀 하나를 올린 것
            return [_read_xlsx(data)]
        items = [i for i in z.infolist() if not i.is_dir() and _zip_name(i).lower().endswith('.xlsx')]
        if not items:
            raise ValueError('zip 안에 운행일지 엑셀(.xlsx)이 없어요.')
        if len(items) > MAX_FILES:
            raise ValueError(f'zip 안 파일이 너무 많아요({MAX_FILES}개까지).')
        out = []
        for i in items:
            if i.file_size > MAX_FILE_BYTES:
                raise ValueError(f'{_zip_name(i)} 파일이 너무 커요.')
            out.append(_read_xlsx(z.read(i)))
        return out
    if not (file_name or '').lower().endswith('.xlsx'):
        raise ValueError(NOT_LOG)
    return [_read_xlsx(data)]


def import_upload(file_name: str, data: bytes, by: str) -> Dict:
    """올린 운행일지를 장부에 — 이미 있는 운행(차량·날짜·출발 시각)은 건너뜀."""
    sheets = parse_upload(file_name, data)
    added = skipped = 0
    periods = sorted({s['period'] for s in sheets if s['period']})
    with _conn() as conn:
        uid = conn.execute('INSERT INTO vehicle_uploads (file_name, period, vehicles, uploaded_by, uploaded_at) '
                           'VALUES (?,?,?,?,?)', ((file_name or '')[:120], ', '.join(periods)[:120],
                                                  len(sheets), by, _now())).lastrowid
        for s in sheets:
            conn.execute('INSERT INTO vehicles (plate, model, updated_at) VALUES (?,?,?) '
                         'ON CONFLICT(plate) DO UPDATE SET model=excluded.model, updated_at=excluded.updated_at',
                         (s['plate'], s['model'], _now()))
            for t in s['trips']:
                cur = conn.execute(
                    'INSERT OR IGNORE INTO vehicle_trips (plate, day, start_time, end_time, odo_start, odo_end, km, '
                    'from_addr, to_addr, commute_km, business_km, note, upload_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
                    (s['plate'], t['day'], t['start_time'], t['end_time'], t['odo_start'], t['odo_end'], t['km'],
                     t['from_addr'], t['to_addr'], t['commute_km'], t['business_km'], t['note'], uid))
                if cur.rowcount:
                    added += 1
                else:
                    skipped += 1
        conn.execute('UPDATE vehicle_uploads SET added=?, skipped=? WHERE id=?', (added, skipped, uid))
    return {'vehicles': len(sheets), 'added': added, 'skipped': skipped, 'period': ', '.join(periods)}


# ── 지역 맞추기 ────────────────────────────────────────────────────────
_METRO = {'서울특별시': '서울', '서울시': '서울', '서울': '서울', '인천광역시': '인천', '인천': '인천',
          '부산광역시': '부산', '부산': '부산', '대구광역시': '대구', '대구': '대구', '광주광역시': '광주',
          '대전광역시': '대전', '대전': '대전', '울산광역시': '울산', '울산': '울산',
          '세종특별자치시': '세종', '세종': '세종'}
_PROV = {'경기도', '경기', '강원도', '강원특별자치도', '강원', '충청북도', '충북', '충청남도', '충남',
         '전라북도', '전북특별자치도', '전북', '전라남도', '전남', '경상북도', '경북', '경상남도', '경남',
         '제주특별자치도', '제주도', '제주'}


def area(addr: str) -> Tuple[str, str, str]:
    """주소 → (시, 구, 동·읍·면). 운행일지 '경기도 용인시 수지구 풍덕천동' → ('용인','수지구','풍덕천동'),
    방문 주소(축약) '용인 수지구 심곡로 11' → ('용인','수지구',''), '마포구 양화로 45' → ('서울','마포구','')."""
    toks = re.sub(r'[(),]', ' ', str(addr or '')).split()
    i, city = 0, ''
    if toks and toks[0] in _METRO:
        city, i = _METRO[toks[0]], 1
    elif toks and toks[0] in _PROV:
        i = 1
    if i >= len(toks):
        return (city, '', '')
    gu = dong = ''
    t = toks[i]
    if not city:
        if t.endswith('구') and len(t) >= 2 and not re.search(r'\d', t):
            city, gu, i = '서울', t, i + 1          # 축약 주소는 '서울'을 뺀다(normalize_display)
        else:
            c = t[:-1] if len(t) >= 3 and t.endswith(('시', '군')) else t
            if re.search(r'\d', c) or (len(t) >= 3 and t.endswith(('로', '길'))):   # 도로명부터 시작 — 시를 모름
                return ('', '', '')
            city, i = c, i + 1
    if not gu and i < len(toks) and toks[i].endswith(('구', '군')) and not re.search(r'\d', toks[i]):
        gu, i = toks[i], i + 1
    if i < len(toks) and toks[i].endswith(('동', '읍', '면', '가')) and not re.search(r'\d', toks[i][:-1] or 'x'):
        dong = toks[i]
    return (city, gu, dong)


def same_area(trip: Tuple[str, str, str], visit: Tuple[str, str, str]) -> bool:
    """운행 도착지와 방문지가 같은 곳인가. 같은 시(구가 둘 다 있으면 같은 구)이고,
    방문지 동을 알면 동까지 같아야 한다. 방문지 동을 모르면 회사가 있는 김포 안에서는 맞추지 않는다
    (출퇴근·회사 오가는 길이 김포 방문과 늘 겹친다)."""
    if not trip[0] or trip[0] != visit[0]:
        return False
    if trip[1] and visit[1] and trip[1] != visit[1]:
        return False
    if visit[2]:
        return trip[2] == visit[2]
    return trip[0] != HOME_AREA[0]


def is_home(addr_area: Tuple[str, str, str]) -> bool:
    return addr_area[0] == HOME_AREA[0] and addr_area[2] == HOME_AREA[2]


# ── 방문 일정 ──────────────────────────────────────────────────────────
_DATE_RE = re.compile(r'(\d{4})[-./](\d{1,2})[-./](\d{1,2})')
_NOT_VISITED = ('유선', '부재', '취소', '보류')     # 방문 예정일이 있어도 가지 않은 리드
_HAS_NAME = re.compile(r'[가-힣A-Za-z]')          # 글자가 하나라도 있어야 사람 이름('-'·'—'·'?' 제외)


def _visit_day(v) -> str:
    m = _DATE_RE.search(str(v or ''))
    if not m:
        return ''
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
    except ValueError:
        return ''


_MAX_SPAN_DAYS = 31     # 기간 방문을 날마다 펼치는 한도 — 끝 날짜 오기('2026-01-02~12-30')로 1년치가 생기지 않게


def _visit_days(v) -> List[str]:
    """방문 예정일 → 걸친 날들. '2026-11-24~26' 같은 기간은 날마다(11-24 부장님: 첫날에만 세어
    25일 코스·차 없음 경고에서 '하루종일 현장관리'가 빠지던 것). 기간 해석은 일정 목록(day_schedule)과 같은 함수.
    끝이 시작보다 앞이거나 한도를 넘으면 첫날만."""
    first = _visit_day(v)
    if not first:
        return []
    from dashboard.services.visit_assignment_sync import _parse_visit_date_end
    st = date.fromisoformat(first)
    en = _parse_visit_date_end(v)
    if not en or en <= st or (en - st).days > _MAX_SPAN_DAYS:
        return [first]
    return [(st + timedelta(days=i)).isoformat() for i in range((en - st).days + 1)]


def _visit_leads(leads: Iterable[dict], start: str, end: str):
    for lead in leads:
        days = [d for d in _visit_days(lead.get('방문 예정일')) if start <= d <= end]
        if not days:
            continue
        status = str(lead.get('상태') or '').strip()
        if any(w in status for w in _NOT_VISITED):
            continue
        # '-'·'—' 같은 빈칸 표시는 사람이 아니다(운전자 칸에 "-, 박민재"로 나오던 것, 11-21 점검)
        who = [w.strip() for w in re.split(r'[,/·\s]+', str(lead.get('영업 담당자') or ''))
               if _HAS_NAME.search(w)]
        addr = str(lead.get('방문 주소') or '').strip()
        if who and area(addr)[0]:
            for day in days:
                yield day, lead, who, addr


def place_label(ar: Tuple[str, str, str], addr: str = '') -> str:
    """갈 곳 짧은 이름 — (시, 구, 동) → '천안 불당동'. 동을 아직 못 찾았으면 '천안 서북구'."""
    city, gu, dong = ar
    return ' '.join(x for x in (city, dong or gu) if x) or str(addr or '')[:20]


def visits_by_day(leads: Iterable[dict], start: str, end: str) -> Dict[str, List[Dict]]:
    """리드 시트 → {날짜: [{'lead_no','who':[이름],'addr','area'}]} (기간 안, 담당자·주소 있는 방문만).
    방문 주소는 도로명이라 동이 없다 — resolve_dongs 로 미리 찾아 둔 동을 붙인다."""
    out: Dict[str, List[Dict]] = {}
    rows = list(_visit_leads(leads, start, end))
    dongs = _dong_cache([r[3] for r in rows])
    for day, lead, who, addr in rows:
        ar = area(addr)
        if not ar[2] and dongs.get(addr):
            ar = (ar[0], ar[1], dongs[addr])
        out.setdefault(day, []).append({'lead_no': str(lead.get('리드 No') or '').strip(), 'who': who,
                                        'addr': addr[:60], 'area': ar})
    return out


# ── 방문 주소의 동 찾기(행안부 도로명주소) ────────────────────────────────
_ROAD_NO = re.compile(r'^(.*?(?:로|길)\s*\d+(?:-\d+)?)')


def _dong_cache(addrs: Iterable[str]) -> Dict[str, str]:
    addrs = sorted({a for a in addrs if a})
    if not addrs:
        return {}
    out = {}
    with _conn() as conn:
        for i in range(0, len(addrs), 400):
            part = addrs[i:i + 400]
            q = 'SELECT addr, dong FROM vehicle_addr_dong WHERE addr IN (%s)' % ','.join('?' * len(part))
            out.update({r[0]: r[1] or '' for r in conn.execute(q, part).fetchall()})
    return out


def _lookup_dong(addr: str) -> Optional[str]:
    """도로명 주소 → 법정동(읍·면). 못 찾으면 '', 행안부가 잠시 안 되면 None(다음에 다시)."""
    from dashboard.services import address_resolver as ar
    m = _ROAD_NO.match(addr)
    q = (m.group(1) if m else addr)[:60]
    try:
        res = ar._juso_search_cached(q)
    except ar._JusoTransientError:
        return None
    except Exception:
        return None
    for road, jibun, _bd in res or ():
        d = area(jibun)[2]
        if d:
            return d
    return ''


def unresolved(leads: Iterable[dict], start: str, end: str) -> List[str]:
    """동을 아직 찾지 않은 방문 주소(기간 안)."""
    addrs = sorted({r[3] for r in _visit_leads(leads, start, end) if not area(r[3])[2]})
    known = _dong_cache(addrs)
    return [a for a in addrs if a not in known]


def resolve_dongs(addrs: Iterable[str], limit: int = 600, pause: float = 0.15) -> int:
    """방문 주소들의 동을 찾아 저장 — 운행일지 올린 뒤 뒤에서 돈다. 저장한 개수."""
    import time
    n = 0
    for a in list(addrs)[:limit]:
        d = _lookup_dong(a)
        if d is None:
            time.sleep(1.0)
            continue
        with _conn() as conn:
            conn.execute('INSERT OR REPLACE INTO vehicle_addr_dong (addr, dong, checked_at) VALUES (?,?,?)',
                         (a, d, _now()))
        n += 1
        if pause:
            time.sleep(pause)
    return n


# ── 화면 ───────────────────────────────────────────────────────────────
def month_range(ym: str) -> Tuple[str, str]:
    y, m = (int(x) for x in ym.split('-'))
    first = date(y, m, 1)
    last = (date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
    return first.isoformat(), last.isoformat()


def months() -> List[str]:
    with _conn() as conn:
        rows = conn.execute("SELECT DISTINCT substr(day,1,7) FROM vehicle_trips ORDER BY 1 DESC").fetchall()
    return [r[0] for r in rows]


def last_upload() -> Optional[Dict]:
    with _conn() as conn:
        r = conn.execute('SELECT file_name, period, vehicles, added, skipped, uploaded_by, uploaded_at '
                         'FROM vehicle_uploads ORDER BY id DESC LIMIT 1').fetchone()
    return dict(r) if r else None


def trip_range() -> Optional[Tuple[str, str]]:
    """장부에 있는 운행의 첫날·마지막 날(방문 주소 동 찾기 범위)."""
    with _conn() as conn:
        r = conn.execute('SELECT MIN(day), MAX(day) FROM vehicle_trips').fetchone()
    return (r[0], r[1]) if r and r[0] else None


def _trips(start: str, end: str, plate: str = '') -> List[Dict]:
    q = ('SELECT plate, day, start_time, end_time, km, from_addr, to_addr FROM vehicle_trips '
         'WHERE day BETWEEN ? AND ?' + (' AND plate = ?' if plate else '') + ' ORDER BY day, start_time')
    with _conn() as conn:
        return [dict(r) for r in conn.execute(q, (start, end, plate) if plate else (start, end)).fetchall()]


def _models() -> Dict[str, str]:
    with _conn() as conn:
        return {r[0]: r[1] or '' for r in conn.execute('SELECT plate, model FROM vehicles').fetchall()}


def _labels() -> Dict[str, str]:
    """차 이름 — 차량 배정 화면과 같은 이름('쏘렌토 8021'). 운행 화면이 번호판만 보여 줘 맞춰 보기 어려웠음(11-21)."""
    try:
        from dashboard.services import vehicle_assign as va
        return {c['plate']: c.get('label') or '' for c in va.fleet(include_inactive=True)}
    except Exception:   # 차량 목록을 못 읽어도 운행은 보여 준다(번호판만)
        return {}


def _match(trip: Dict, day_visits: List[Dict]) -> List[Dict]:
    """이 운행의 도착지와 같은 곳의 그날 방문(same_area). 회사(걸포동)로 오는 운행은 대조하지 않는다."""
    ar = area(trip.get('to_addr'))
    if not ar[0] or is_home(ar):
        return []
    return [v for v in day_visits if same_area(ar, v['area'])]


def _guess(matched: List[Dict]) -> List[str]:
    c = Counter(n for v in {v['lead_no']: v for v in matched}.values() for n in v['who'])
    return [n for n, _ in c.most_common(3)]


class _Assign:
    """차량 배정·실제 운전자를 읽은 결과 — 요청마다 따로 만들어 인자로 넘긴다.
    예전엔 오전·오후 운전자·운전자 답이 모듈 전역이라 두 사람이 다른 달·차를 동시에 열면
    서로 덮어썼다(11-24 조용한 실패 점검 23번).
    rows = (날짜, 번호판) → 탄 사람 이름 · fixed = 번호판 → 고정 사용자 이름
    shifts = (날짜, 번호판) → (오전 운전, 오후 운전, 나누는 시각) · confirmed = 저녁 DM 으로 실제 운전자를 받은 (날짜, 번호판)
    error = '' 정상 · 'assign' 배정을 못 읽음(운전자는 방문 지역 추정만) · 'driver' 운전자 답만 못 읽음(배정 기준)"""
    __slots__ = ('rows', 'fixed', 'shifts', 'confirmed', 'error')

    def __init__(self):
        self.rows: Dict[Tuple[str, str], List[str]] = {}
        self.fixed: Dict[str, str] = {}
        self.shifts: Dict[Tuple[str, str], Tuple[str, str, str]] = {}
        self.confirmed: set = set()
        self.error = ''


def _assignments(start: str, end: str) -> _Assign:
    """차량 배정(services/vehicle_assign) — (날짜, 번호판) → 탄 사람 이름. 고정 차는 배정 없으면 고정 사용자.
    배정이 있으면 방문 일정 추정보다 우선한다(11-20).
    못 읽으면 ERROR 를 남기고 error 에 표시 — 예전엔 말없이 빈 값이라 '방문 지역으로 추정한 운전자'가 정상처럼 보였다
    (11-24 조용한 실패 점검 23번)."""
    a = _Assign()
    try:
        from dashboard.services import vehicle_assign as va
        who = va._who_map()
        rows = {k: [va.display(c, who) for c in v] for k, v in va.assigned_people(start, end).items()}
        fixed = {p: va.display(u, who) for p, u in va.fixed_users().items()}
        # 오전·오후 나눈 차 — 13:00 이후 운행은 오후 운전자(11-22)
        shifts = {}
        for k, h in va.assigned_halves(start, end).items():
            am, pm = h['people'], va.pm_of(h)
            shifts[k] = (va.display(am[0], who) if am else '', va.display(pm[0], who) if pm else '', va.HALF_AT)
    except Exception as e:
        logger.error(f'[차량운행] {start}~{end} 차량 배정을 못 읽어 운전자를 방문 지역으로만 추정함: {e}', exc_info=True)
        a.error = 'assign'
        return a
    a.rows, a.fixed, a.shifts = rows, fixed, shifts
    # 그날 저녁 슬랙으로 받은 실제 운전자(services/vehicle_driver) — 있으면 배정보다 먼저(11-22)
    try:
        from dashboard.services import vehicle_driver as vd
        actual = vd.actual_drivers(start, end)
    except Exception as e:
        logger.error(f'[차량운행] {start}~{end} 운전자 답(저녁 슬랙)을 못 읽어 배정 기준으로만 보임: {e}', exc_info=True)
        a.error = 'driver'
        return a
    for k, d in actual.items():
        base = shifts.get(k, ('', '', va.HALF_AT))
        am_n = va.display(d['am'], who) if d['am'] else base[0]
        pm_n = va.display(d['pm'], who) if d['pm'] else base[1]
        shifts[k] = (am_n, pm_n, va.HALF_AT)
        a.confirmed.add(k)
        if k not in rows:
            rows[k] = [x for x in dict.fromkeys((am_n, pm_n)) if x]
    return a


def trip_driver(day: str, plate: str, start_time: str, a: _Assign) -> str:
    """운행 한 건의 운전자 — 오전 운전자, 13:00 이후 운행은 오후 운전자(오전·오후 같으면 같은 사람)."""
    sh = a.shifts.get((day, plate))
    if not sh:
        return ''
    am, pm, at = sh
    return (pm if (start_time or '')[:5] >= at else am) or ''


def _drivers_for(day: str, plate: str, a: _Assign, matched: List[Dict]) -> Tuple[List[str], str]:
    if (day, plate) in a.rows:
        ppl = a.rows[(day, plate)]
        sh = a.shifts.get((day, plate))
        if sh:                       # 오전 운전 → 오후 운전 → 나머지
            first = [x for x in dict.fromkeys(sh[:2]) if x]
            ppl = first + [p for p in ppl if p not in first]
        return ppl, ('confirmed' if (day, plate) in a.confirmed else 'assigned')
    if plate in a.fixed:
        return [a.fixed[plate]], 'fixed'
    return _guess(matched), 'guess'


def month_summary(ym: str, leads: Iterable[dict]) -> Dict:
    """월 차량별 합계 + 운전자(배정 > 고정 사용자 > 그날 같은 지역 방문으로 추정)."""
    start, end = month_range(ym)
    leads = list(leads)
    visits = visits_by_day(leads, start, end)
    asg = _assignments(start, end)
    models = _models()
    cars: Dict[str, Dict] = {}
    for t in _trips(start, end):
        c = cars.setdefault(t['plate'], {'plate': t['plate'], 'model': models.get(t['plate'], ''), 'km': 0.0,
                                         'trips': 0, 'days': set(), 'last': '', 'matched': [], 'known': Counter(),
                                         'known_days': set()})
        c['km'] += t['km'] or 0
        c['trips'] += 1
        c['days'].add(t['day'])
        c['last'] = max(c['last'], t['day'])
        m = _match(t, visits.get(t['day'], []))
        c['matched'] += m
        if t['day'] not in c['known_days']:
            names, src = _drivers_for(t['day'], t['plate'], asg, [])
            if src != 'guess':
                c['known_days'].add(t['day'])
                sh = asg.shifts.get((t['day'], t['plate']))
                c['known'].update([x for x in dict.fromkeys(sh[:2]) if x] if sh else names[:1])
    rows = []
    labels = _labels()
    for c in cars.values():
        if c['known']:
            drivers, src = [n for n, _ in c['known'].most_common(3)], 'assigned'
        else:
            drivers, src = _guess(c['matched']), 'guess'
        rows.append({'plate': c['plate'], 'label': labels.get(c['plate'], ''), 'model': c['model'],
                     'km': round(c['km'], 1), 'trips': c['trips'],
                     'days': len(c['days']), 'last': c['last'], 'drivers': drivers, 'source': src,
                     'known_days': len(c['known_days']),
                     'visits': len({v['lead_no'] for v in c['matched']})})
    rows.sort(key=lambda r: -r['km'])
    return {'month': ym, 'start': start, 'end': end, 'vehicles': rows,
            'total_km': round(sum(r['km'] for r in rows), 1),
            'pending_dongs': len(unresolved(leads, start, end)) if rows else 0,
            'assign_error': asg.error}      # 화면 경고 한 줄(11-24 조용한 실패 점검 23번)


def vehicle_days(ym: str, plate: str, leads: Iterable[dict]) -> Dict:
    """한 차량의 한 달 — 날짜별 운행과 같은 지역 방문(추정 운전자)."""
    start, end = month_range(ym)
    visits = visits_by_day(leads, start, end)
    days: Dict[str, Dict] = {}
    for t in _trips(start, end, plate):
        d = days.setdefault(t['day'], {'day': t['day'], 'km': 0.0, 'trips': [], 'matched': {}})
        m = _match(t, visits.get(t['day'], []))
        d['km'] += t['km'] or 0
        d['trips'].append({'start': t['start_time'][:5], 'end': (t['end_time'] or '')[:5], 'km': t['km'],
                           'from': t['from_addr'] or '', 'to': t['to_addr'] or '',
                           'home': is_home(area(t['to_addr'])),
                           'visits': [{'lead_no': v['lead_no'], 'who': v['who'], 'addr': v['addr']} for v in m]})
        for v in m:
            d['matched'][v['lead_no']] = v
    asg = _assignments(start, end)
    try:        # U+ API 로 받은 운행의 지표(급가속 등·최고 속도·출퇴근) — 경로 지도 단추도 이것이 있을 때만(11-24)
        from dashboard.services import vehicle_api_sync as vas
        mets = vas.metrics(start, end, plate)
    except Exception:
        mets = {}
    for d in days.values():          # 운행마다 운전자(배정한 날만 — 오후 교대 반영)
        for t in d['trips']:
            t['driver'] = trip_driver(d['day'], plate, t['start'], asg)
            m = mets.get((plate, d['day'], t['start']))
            if m:
                t['route'] = True
                t['max_speed'] = int(m['max_speed'] or 0)
                t['safe'] = m['safe']
                t['commute'] = bool(m['commute'])
                t['events'] = {k: m[k] or 0 for k in ('rapid_acc', 'rapid_dec', 'sudden_stop', 'rapid_turn',
                                                      'rapid_u_turn', 'rapid_lane', 'rapid_pass')}
    out = []
    for d in sorted(days.values(), key=lambda x: x['day'], reverse=True):
        drivers, src = _drivers_for(d['day'], plate, asg, list(d['matched'].values()))
        out.append({'day': d['day'], 'km': round(d['km'], 1), 'trips': d['trips'], 'drivers': drivers, 'source': src})
    return {'month': ym, 'plate': plate, 'model': _models().get(plate, ''), 'label': _labels().get(plate, ''), 'days': out,
            'assign_error': asg.error}


# ── 국세청 운행기록부 엑셀(11-24 부장님 — U+ '운행일지(국세청)' 내려받기를 그룹웨어에서) ─────────────
WD_KO = '월화수목금토일'


def logbook_rows(plate: str, start: str, end: str) -> List[Dict]:
    """운행기록부 줄 — 사용일자·부서·성명·주행 전/후 km·주행거리·출퇴근용·일반업무용·비고.
    성명 = 운전자 답 > 배정(13시 뒤 오후) > 고정 사용자. 모르면 빈칸(만들어 넣지 않는다)."""
    from dashboard.services import hr
    asg = _assignments(start, end)     # 못 읽으면 ERROR 알림이 가고 성명은 빈칸(지어 넣지 않음)
    depts = {r['name']: r.get('department') or '' for r in hr._rows("SELECT name, department FROM hr_people")}
    with _conn() as conn:
        rows = conn.execute('SELECT * FROM vehicle_trips WHERE plate = ? AND day BETWEEN ? AND ? ORDER BY day, start_time',
                            (plate, start, end)).fetchall()
    out = []
    for r in rows:
        name = trip_driver(r['day'], plate, r['start_time'], asg) or asg.fixed.get(plate, '')
        d = date.fromisoformat(r['day'])
        out.append({
            'day': f"{r['day']}({WD_KO[d.weekday()]})", 'dept': depts.get(name, ''), 'name': name,
            'odo_start': r['odo_start'], 'odo_end': r['odo_end'], 'km': r['km'] or 0,
            'commute': r['commute_km'] or 0, 'business': r['business_km'] or 0,
            'note': ' → '.join(x for x in (r['from_addr'] or '', r['to_addr'] or '') if x)[:120],
            'time': f"{(r['start_time'] or '')[:5]}~{(r['end_time'] or '')[:5]}",
        })
    return out


def logbook_xlsx(plates: List[str], start: str, end: str, corp: str = '') -> bytes:
    """차량마다 한 장 — 국세청 「업무용승용차 운행기록부」 칸 그대로. 마지막 줄에 합계·업무사용비율."""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    labels, models = _labels(), _models()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    thin = Side(style='thin', color='999999')
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill('solid', fgColor='F1F5F9')
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for plate in plates:
        rows = logbook_rows(plate, start, end)
        ws = wb.create_sheet(re.sub(r'[\/*?:\[\]]', '', (labels.get(plate) or plate))[:28] or plate)
        ws.append(['업무용승용차 운행기록부'])
        ws['A1'].font = Font(size=14, bold=True)
        ws.append(['과세기간', f'{start} ~ {end}', '', '법인명', corp])
        ws.append(['차종', models.get(plate) or labels.get(plate) or '', '', '자동차등록번호', plate])
        ws.append([])
        ws.append(['사용일자(요일)', '사용자', '', '운행 내역', '', '', '', '', '비고', '운행 시각'])
        ws.append(['', '부서', '성명', '주행 전 계기판의 거리(km)', '주행 후 계기판의 거리(km)', '주행거리(km)',
                   '업무용 사용거리(km)', '', '', ''])
        ws.append(['', '', '', '', '', '', '출·퇴근용(km)', '일반 업무용(km)', '', ''])
        for rng in ('A5:A7', 'B5:C5', 'D5:H5', 'I5:I7', 'J5:J7', 'B6:B7', 'C6:C7', 'D6:D7', 'E6:E7', 'F6:F7', 'G6:H6'):
            ws.merge_cells(rng)
        for row in ws.iter_rows(min_row=5, max_row=7, max_col=10):
            for c in row:
                c.fill, c.border, c.alignment = head_fill, box, center
                c.font = Font(bold=True)
        tot = com = bus = 0.0
        for r in rows:
            ws.append([r['day'], r['dept'], r['name'], r['odo_start'], r['odo_end'], r['km'], r['commute'] or None,
                       r['business'] or None, r['note'], r['time']])
            tot += r['km']
            com += r['commute']
            bus += r['business']
        last = ws.max_row
        ws.append(['합계', '', '', '', '', round(tot, 1), round(com, 1), round(bus, 1),
                   f'업무사용비율 {round((com + bus) / tot * 100, 1) if tot else 0}%', ''])
        for row in ws.iter_rows(min_row=8, max_row=ws.max_row, max_col=10):
            for c in row:
                c.border = box
                if c.column in (4, 5, 6, 7, 8):
                    c.number_format = '#,##0.0'
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
        for col, w in zip('ABCDEFGHIJ', (16, 12, 10, 14, 14, 11, 12, 12, 48, 13)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A8'
        if not rows:
            ws.cell(row=last + 1, column=9, value='이 기간 운행 기록이 없어요')
    if not wb.worksheets:
        wb.create_sheet('운행기록부')
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
