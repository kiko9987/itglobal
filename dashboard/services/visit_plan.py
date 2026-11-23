"""그룹웨어 방문 일정 (2026-11-22 부장님) — 다음 날 방문을 누가 가는지 그룹웨어에서 짜고 확정한다.

지금까지는 JW 가 슬랙 캔버스2에 'YM+LSH' 같은 팀 이름과 방문 줄을 적고 /일정확정을 실행했다.
→ 차량 배정 화면 ① 방문 일정 단계에서 JW 가 하던 순서 그대로(11-22 부장님):
   캔버스1 현장을 지역·일정별 '코스'로 묶고 순서를 정한다 → 코스마다 갈 사람을 넣는다 → 넣었다 뺐다 조정 → [일정 확정].
   (그 수작업 결과가 캔버스2였다. 코스 = 캔버스2의 'YM+LSH' 섹션 하나.)
   확정 처리(시트 영업 담당자·슬랙 List·담당자/당번 DM·캔버스1)는 /일정확정과 같은 본체
   (visit_assignment_sync.commit_plan)를 쓴다. 리드를 골라 넣으므로 전화번호 매칭 실패·이니셜 판독 불가가 없다.
- 현장은 캔버스1과 같은 기준(visit_canvas_sync.visit_open + 그날에 걸친 방문)으로 띄우고, 사람은 처음부터 배정한다
  (11-22 부장님: 시트 영업 담당자 = 캔버스2 결과라 그걸로 팀을 미리 묶으면 안 됨). 온라인 당번도 비워 두고 시작.
- 휴무·반차는 그룹웨어 휴가·토요 휴무표에서(캔버스2 '휴무' 칸 대신).
- 화면: 왼쪽 사람, 오른쪽 위 '코스에 안 넣은 현장'(지역 꼬리표, [지역별로 코스 나눠 보기]), 그 아래 코스 카드.
- 코스 메모는 그 코스 사람들의 담당자 DM 위쪽에 간다(캔버스2를 내리면 메모 볼 곳이 DM 뿐).
- 바꾸는 동안 슬랙 /일정확정도 그대로 둔다. 그룹웨어에서 확정한 날은 /일정확정이 막는다(덮어쓰기 방지).
- 확정은 30초쯤 걸려 뒤에서 돌리고, 화면은 상태를 다시 물어 본다(run_state).
"""
from __future__ import annotations

import json
import re
import threading
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Callable, Dict, Iterable, List, Optional

from dashboard.services import vehicle_assign as va
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

MAX_COURSES = 60
NAME_MAX = 20
MEMO_MAX = 200
SHIFTS = ('', '오전', '오후')
_RUN_STALE = timedelta(minutes=5)     # 서버 재시작으로 끊긴 확정은 이만큼 지나면 '끊김'으로 본다

_DDL = '''CREATE TABLE IF NOT EXISTS visit_plan_day (
    day TEXT PRIMARY KEY,
    plan TEXT,
    saved_by TEXT,
    saved_at TEXT,
    confirmed_by TEXT,
    confirmed_at TEXT,
    confirmed_plan TEXT,
    run_state TEXT,
    run_at TEXT,
    run_by TEXT,
    run_result TEXT
)'''
# 당일 배정(11-23 부장님) — 그날 생긴 현장을 온라인 당번·일정 담당이 근처 매니저에게 바로 맡긴 기록
_URGENT_DDL = '''CREATE TABLE IF NOT EXISTS visit_plan_urgent (
    day TEXT,
    lead_no TEXT,
    members TEXT,
    by TEXT,
    at TEXT,
    PRIMARY KEY (day, lead_no)
)'''
_table_ready = False
_run_lock = threading.Lock()


@contextmanager
def _conn():
    global _table_ready
    with va._db()._get_connection() as conn:
        if not _table_ready:
            conn.execute(_DDL)
            conn.execute(_URGENT_DDL)
            for col in ('state', 'result'):      # 11-23 뒤에서 처리(시트·List·DM) 상태 — 먼저 만든 표에 칸 추가
                try:
                    conn.execute(f'ALTER TABLE visit_plan_urgent ADD COLUMN {col} TEXT')
                except Exception:
                    pass
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


_STAFF_TTL = 5.0
_staff_memo: Dict[str, object] = {'fn': None, 'at': 0.0, 'val': None}


def _staff() -> List[Dict]:
    """현장 명단 — 한 번 불러올 때 여러 번 읽지 않게 몇 초만 기억(11-23 속도: 읽을 때마다 사용자 목록 로그가 147줄씩)."""
    import time
    m = _staff_memo
    if m['fn'] is va.staff and m['val'] is not None and time.monotonic() - m['at'] < _STAFF_TTL:
        return [dict(p) for p in m['val']]
    val = va.staff()
    m.update(fn=va.staff, at=time.monotonic(), val=[dict(p) for p in val])
    return [dict(p) for p in val]


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


# ── 권한·사람 ───────────────────────────────────────────────────────────
def can_plan(email: str) -> bool:
    from dashboard.services.visit_assignment_sync import PLANNER_EMAILS
    return (email or '').strip().lower() in PLANNER_EMAILS


def _name_of_email(email: str) -> str:
    from dashboard.services import hr
    p = hr.find_by_email(email or '')
    return (p or {}).get('name') or (email or '')


def _code_maps() -> Dict[str, str]:
    """이름 → 이니셜. 현장 명단(va.staff) 먼저, 없으면 로그인 계정 이니셜(사무직 등)."""
    from dashboard.services.visit_assignment_sync import _load_initial_maps
    _, name_to_ini = _load_initial_maps()
    out = dict(name_to_ini)
    out.update({p['name']: p['code'] for p in _staff()})
    return out


# ── 그날 방문 ───────────────────────────────────────────────────────────
def _open_leads(leads: Iterable[dict]) -> List[dict]:
    """캔버스1에 올라가는 리드만(방문 예약·공사 확정, 방문 완료 처리 전)."""
    from dashboard.services.visit_canvas_sync import visit_open
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception:
        rc = None
    return [l for l in leads if visit_open(l, rc)]


def _on_day(day: str, leads: Iterable[dict]) -> List[dict]:
    """방문 예정일이 그날에 걸친 리드만(여러 날짜 '2026-11-23~24' 포함) — day_schedule 과 같은 기준."""
    from dashboard.services import visit_assignment_sync as vas
    d = date.fromisoformat(day)
    out = []
    for l in leads:
        vd = l.get('방문 예정일')
        st = vas._parse_visit_date_start(vd)
        if not st:
            continue
        en = vas._parse_visit_date_end(vd)
        if st <= d <= (en if en and en >= st else st):
            out.append(l)
    return out


def visits(day: str, leads: Iterable[dict], done: bool = False) -> List[Dict]:
    """그날 현장(캔버스1과 같은 리드·묶음) 을 한 줄로 — 카드 하나 = 리드 하나.
    본인 방문 필수면 owner = 상담한 사람 이니셜(온라인 상담자).
    done=True(오늘·지난 날 보기)면 방문 완료한 것까지 — 그날 실제로 간 현장(11-23 부장님)."""
    from dashboard.services import vehicle_log as vl
    leads = _on_day(day, leads)         # 그날 방문만 먼저 — 방문 완료 확인(Redis)을 전체 791건이 아니라 몇 건만(11-23 속도)
    leads = leads if done else _open_leads(leads)
    by_name = _code_maps()
    by_no = {str(l.get('리드 No') or '').strip(): l for l in leads}
    out = []
    for g in va.day_schedule(day, leads):
        for it in g['items']:
            if not it.get('lead_no'):
                continue
            src = str((by_no.get(it['lead_no']) or {}).get('온라인 상담자') or '').strip()
            src = '' if src in ('-', '미정') else src
            must = '본인 방문 필수' in it['marks']
            owner = by_name.get(src, '') if must else ''
            # 누구 현장인지(11-23 부장님) — 거래처·기타는 늘, 온라인은 본인 방문 필수일 때만(그 밖엔 누가 가도 됨)
            contact = src if (must or g['cat'] in ('거래처', '기타')) else ''
            item = {k: v for k, v in it.items() if k != 'who'}      # 시트 담당자(=캔버스2 결과)는 보이지 않는다
            item['marks'] = [m for m in item['marks'] if m != '본인 방문 필수']     # 이름과 한 꼬리표로(아래 must)
            out.append({**item, 'cat': g['cat'], 'title': g['title'], 'owner': owner, 'contact': contact, 'must': must,
                        'reg_at': _reg_at((by_no.get(it['lead_no']) or {}).get('상담 시간')), **_where(it['addr'], vl)})
    return out


def _reg_at(raw) -> str:
    """시트 '상담 시간'('2026.11.23. 08:12') → '2026-11-23 08:12'. 못 읽으면 ''."""
    m = re.search(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})\D+(\d{1,2}):(\d{2})', str(raw or ''))
    if not m:
        return ''
    y, mo, d, h, mi = m.groups()
    return f'{y}-{int(mo):02d}-{int(d):02d} {int(h):02d}:{mi}'


_BIG_CITIES = ('서울', '인천', '부산', '대구', '광주', '대전', '울산')
# 경기 시·군(축약 주소엔 '경기'가 없음) — 현장 목록 지역 순서 서울 → 인천 → 경기 → 그 밖(11-23 부장님)
_GYEONGGI = ('수원', '성남', '고양', '용인', '부천', '안산', '안양', '남양주', '화성', '평택', '의정부', '시흥', '파주',
             '김포', '광명', '광주', '군포', '하남', '오산', '이천', '안성', '의왕', '양주', '구리', '포천', '동두천',
             '과천', '여주', '양평', '가평', '연천')


def _where(addr: str, vl) -> Dict[str, str]:
    """현장 꼬리표 — 도로명 주소에 적힌 단위만(시·구·읍·면). 지번 법정동(운행일지용 행안부 조회값)은 쓰지 않는다
    (11-22 부장님: 우리 주소는 도로명 기준). place = '서울 강남구'·'평택'·'화성 봉담읍',
    region = 같은 지역끼리 모아 보는 단위(경기는 시, 서울·인천 등 광역시는 구까지)."""
    city, gu, town = vl.area(addr)
    town = town if town.endswith(('읍', '면')) else ''          # '…동'은 도로명 주소의 단위가 아님
    region = (f'{city} {gu}'.strip() if city in _BIG_CITIES else city) or gu or ''
    place = ' '.join(x for x in (city, gu, town) if x) or ''
    # 광주는 구가 있으면 광주광역시(동구·서구 …), 없으면 경기 광주시
    rank = (0 if city == '서울' else 1 if city == '인천' else
            2 if city in _GYEONGGI and not (city == '광주' and gu) else 3 if city else 4)
    return {'place': place, 'region': region, 'region_rank': rank}


def people(vs: List[Dict], more: Iterable[str] = ()) -> List[Dict]:
    """팀에 넣을 수 있는 사람 — 현장 명단 + 본인 방문 필수 건을 상담한 다른 사람(사무직 등).
    more = 지난 날 시트 담당자처럼 명단 밖이어도 보여 줄 이니셜."""
    out = [{'code': p['code'], 'name': p['name']} for p in _staff()]
    have = {p['code'] for p in out}
    from dashboard.services.visit_assignment_sync import _load_initial_maps
    ini_to_name, _ = _load_initial_maps()
    for c in [v.get('owner') or '' for v in vs] + list(more):
        if c and c not in have and c in ini_to_name:
            have.add(c)
            out.append({'code': c, 'name': ini_to_name[c], 'extra': True})
    return out


def _sheet_plan(day: str, leads: Iterable[dict], vs: List[Dict]) -> Dict:
    """그룹웨어에 일정이 없는 오늘·지난 날(슬랙 /일정확정으로 정한 날) — 시트 영업 담당자대로 코스를 묶어 보여 준다.
    같은 담당자 조합 = 한 코스, 현장은 그날 순서대로. 담당자가 비어 있는 현장은 현장 목록에 남는다. 보기만(저장·DM 없음)."""
    by_name = _code_maps()
    nos = {v['lead_no'] for v in vs}
    groups: Dict[tuple, List[str]] = {}
    for g in va.day_schedule(day, leads):
        for it in g['items']:
            n = it.get('lead_no')
            codes = tuple(dict.fromkeys(by_name[w] for w in it.get('who') or [] if w in by_name))
            if n in nos and codes:
                groups.setdefault(codes, []).append(n)
    try:
        online = [{'code': o['code'], 'shift': o.get('shift') or ''} for o in va.online_duty(day)]
    except Exception:
        online = []
    return {'courses': [{'name': '', 'members': list(k), 'leads': v, 'memo': ''} for k, v in groups.items()],
            'online': online, 'online_shared': False}


def _empty_plan() -> Dict:
    """저장한 일정이 없는 날 — 모든 현장이 '코스에 안 넣은 현장'에서 시작(시트 담당자·캔버스2 당번은 쓰지 않음)."""
    return {'courses': [], 'online': [], 'online_shared': False}


# ── 저장된 일정 ─────────────────────────────────────────────────────────
def _row(day: str) -> Optional[Dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM visit_plan_day WHERE day = ?', (day,)).fetchone()
    return dict(r) if r else None


def _loads(raw: Optional[str]) -> Optional[Dict]:
    try:
        return json.loads(raw) if raw else None
    except ValueError:
        return None


def _old_courses(raw: Dict) -> Optional[list]:
    """11-22 앞 판 저장분 — 'teams'(팀 칸) 는 그대로 코스로, 'sites'(현장별 사람) 는 같은 사람끼리 묶어 코스로."""
    if isinstance(raw.get('teams'), list):
        return raw['teams']
    if isinstance(raw.get('sites'), dict):
        by: Dict[tuple, Dict] = {}
        for n, v in raw['sites'].items():
            v = v if isinstance(v, dict) else {}
            key = tuple(v.get('members') or [])
            c = by.setdefault(key, {'members': list(key), 'leads': [], 'memo': ''})
            c['leads'].append(n)
            if v.get('memo') and not c['memo']:
                c['memo'] = v['memo']
        return list(by.values())
    return None


def clean_plan(raw: Dict, visit_nos: Iterable[str], codes: Iterable[str], strict: bool = True) -> Dict:
    """화면에서 온 일정을 다듬는다. strict=False 면 모르는 현장·사람은 조용히 뺀다(불러올 때).
    plan = {'courses': [{'name', 'members': [이니셜], 'leads': [리드 No(가는 순서)], 'memo'}], 'online', 'online_shared'}
    사내 업무(11-23 부장님)는 {'kind': 'work', 'members', 'leads': [], 'memo'=할 일} — 회사 안에서 하는 일만
    (우체국·은행·차량 정비처럼 밖에 나가는 일은 기타 방문(ETC-)으로 등록돼 현장 목록에 나온다)."""
    if not isinstance(raw, dict):
        raise ValueError('일정 내용을 읽지 못했어요.')
    visit_nos, codes = set(visit_nos), set(codes)
    courses_in = raw.get('courses')
    if courses_in is None:
        courses_in = _old_courses(raw) or []
    if not isinstance(courses_in, list) or len(courses_in) > MAX_COURSES:
        raise ValueError('코스가 너무 많아요.')
    seen: set = set()
    courses = []
    for t in courses_in:
        if not isinstance(t, dict):
            continue
        members = []
        for c in t.get('members') or []:
            c = str(c or '').strip().upper()
            if c in codes and c not in members:
                members.append(c)
            elif c and c not in codes and strict:
                raise ValueError(f'명단에 없는 사람({c})이 코스에 있어요. 새로 고친 뒤 다시 해 주세요.')
        work = t.get('kind') == 'work'
        leads = []
        for n in ([] if work else t.get('leads') or []):
            n = str(n or '').strip()
            if n in visit_nos and n not in seen:
                seen.add(n)
                leads.append(n)
            elif n and n not in visit_nos and strict:
                raise ValueError('그사이 방문 일정이 바뀐 건이 있어요. 새로 고친 뒤 다시 해 주세요.')
        name = re.sub(r'\s+', ' ', str(t.get('name') or '')).strip()[:NAME_MAX]
        memo = re.sub(r'\s+', ' ', str(t.get('memo') or '')).strip()[:MEMO_MAX]
        if members or leads or memo or name:
            c = {'name': name, 'members': members, 'leads': leads, 'memo': memo}
            if work:
                c['kind'] = 'work'
            courses.append(c)
    online = []
    for o in raw.get('online') or []:
        if not isinstance(o, dict):
            continue
        c = str(o.get('code') or '').strip().upper()
        shift = str(o.get('shift') or '').strip()
        if c in codes and all(x['code'] != c for x in online):
            online.append({'code': c, 'shift': shift if shift in SHIFTS else ''})
    return {'courses': courses, 'online': online, 'online_shared': bool(raw.get('online_shared')) and not online}


def _staffed(plan: Dict) -> List[Dict]:
    """사람과 현장이 다 있는 코스 — 확정하면 DM 이 가는 것."""
    return [c for c in plan['courses'] if c['members'] and c['leads']]


def is_work(c: Dict) -> bool:
    return c.get('kind') == 'work'


def course_labels(plan: Dict) -> List[str]:
    """화면과 같은 이름 — 방문 코스는 '코스 N', 사내 업무는 '사내 업무 N'(종류마다 따로 셈)."""
    out, n = [], {'visit': 0, 'work': 0}
    for c in plan['courses']:
        k = 'work' if is_work(c) else 'visit'
        n[k] += 1
        out.append(c.get('name') or (f'사내 업무 {n[k]}' if k == 'work' else f'코스 {n[k]}'))
    return out


def editable(day: str, today: Optional[date] = None) -> bool:
    """확정은 내일부터 — 오늘·지난 날은 DM 대상이 아니다(방문 끝난 날은 담당 DM 이 안 나감)."""
    return va.valid_day(day) and date.fromisoformat(day) > (today or date.today())


def warnings(plan: Dict, vs: List[Dict], off: Dict[str, List[Dict]], names: Dict[str, str]) -> List[Dict]:
    """확정 전 점검 — 대부분 확인만 받는다(level 'warn').
    휴가자(반차·확정 전 토요 휴무 빼고)를 코스·당번에 넣었으면 'error' — 빼야 확정된다(11-23 부장님)."""
    out: List[Dict] = []
    by_no = {v['lead_no']: v for v in vs}
    done = {n for c in _staffed(plan) for n in c['leads']}
    left = [v for v in vs if v['lead_no'] not in done]
    if left:
        out.append({'level': 'warn', 'kind': 'left',
                    'text': f'코스·사람이 안 정해진 현장 {len(left)}건 — 담당 DM 이 가지 않아요.'})
    offs = {o['code']: o['why'] for o in off.get('off', [])}
    hard = {o['code'] for o in off.get('off', []) if not o.get('soft')}
    count: Dict[str, int] = {}
    labels = course_labels(plan)
    for i, c in enumerate(plan['courses']):
        label = labels[i]
        if is_work(c) and c['members'] and not c['memo']:
            out.append({'level': 'warn', 'kind': 'task', 'course': i, 'text': f'{label}에 할 일을 적어 주세요.'})
        if is_work(c) and c['memo'] and not c['members']:
            out.append({'level': 'warn', 'kind': 'nobody', 'course': i, 'text': f'{label}에 사람을 넣지 않았어요.'})
        if c['leads'] and not c['members']:
            out.append({'level': 'warn', 'kind': 'nobody', 'course': i,
                        'text': f'{label}에 사람을 넣지 않았어요(현장 {len(c["leads"])}건).'})
        for m in c['members']:
            count[m] = count.get(m, 0) + 1
            if m in hard:
                out.append({'level': 'error', 'kind': 'off', 'course': i,
                            'text': f'{label}의 {names.get(m, m)}님은 이날 {offs[m]}예요 — 빼야 확정돼요.'})
            elif m in offs:
                out.append({'level': 'warn', 'kind': 'off', 'course': i,
                            'text': f'{label}의 {names.get(m, m)}님은 이날 {offs[m]}예요.'})
        for n in c['leads'] if c['members'] else []:
            v = by_no.get(n) or {}
            if v.get('owner') and v['owner'] not in c['members']:
                out.append({'level': 'warn', 'kind': 'owner', 'course': i,
                            'text': f'{label}의 본인 방문 필수 현장({(v.get("addr") or n)[:30]})에 상담한 '
                                    f'{names.get(v["owner"], v["owner"])}님이 없어요.'})
    for m, n in count.items():
        if n > 1:
            out.append({'level': 'warn', 'kind': 'twice', 'text': f'{names.get(m, m)}님이 코스 {n}곳에 들어가 있어요.'})
    for o in plan['online']:
        if o['code'] in offs:
            bad = o['code'] in hard
            out.append({'level': 'error' if bad else 'warn', 'kind': 'off',
                        'text': f'온라인 당번 {names.get(o["code"], o["code"])}님은 이날 {offs[o["code"]]}예요'
                                + (' — 빼야 확정돼요.' if bad else '.')})
    # 코스·사내 업무·당번 어디에도 없는 현장 직원(11-23 부장님: 문의가 적어 남는 인원을 빠뜨리지 않게) — 막지는 않음
    if plan['courses'] and not plan['online_shared']:
        placed = {m for c in plan['courses'] for m in c['members']} | {o['code'] for o in plan['online']}
        idle = [p for p in _staff() if p['code'] not in placed and p['code'] not in offs]
        if idle:
            out.append({'level': 'warn', 'kind': 'idle',
                        'text': f'아무 데도 안 들어간 사람 {len(idle)}명({", ".join(p["name"] for p in idle)}) — '
                                f'남는 인원은 사내 업무에 넣어 주세요.'})
    return out


def _state(row: Optional[Dict]) -> Dict:
    """확정 진행 상태 — {'state': ''|'running'|'done'|'failed', 'at', 'by', 'result'}."""
    if not row or not row.get('run_state'):
        return {'state': ''}
    st = row['run_state']
    if st == 'running':
        try:
            if datetime.now() - datetime.strptime(row['run_at'], '%Y-%m-%d %H:%M:%S') > _RUN_STALE:
                st = 'stale'
        except (TypeError, ValueError):
            st = 'stale'
    return {'state': st, 'at': row.get('run_at') or '', 'by': _name_of_email(row.get('run_by') or ''),
            'result': _loads(row.get('run_result')) or {}}


def load(day: str, leads: Iterable[dict], me: str = '') -> Dict:
    leads = list(leads)
    past = not editable(day)          # 오늘·지난 날 — 방문 완료한 현장까지 보여 줌(고치지는 못함)
    vs = visits(day, leads, done=past)
    row = _row(day)
    saved = _loads((row or {}).get('plan'))
    sheet = _sheet_plan(day, leads, vs) if past and not saved else None
    ppl = people(vs, [m for c in (sheet or {}).get('courses', []) for m in c['members']])
    names = {p['code']: p['name'] for p in ppl}
    off = va.day_off(day)
    dropped = 0
    if sheet is not None:
        plan, source = clean_plan(sheet, [v['lead_no'] for v in vs], names, strict=False), 'sheet'
    elif saved:
        plan = clean_plan(saved, [v['lead_no'] for v in vs], names, strict=False)
        before = sum(len(c.get('leads') or []) for c in (saved.get('courses') or _old_courses(saved) or [])
                     if isinstance(c, dict))
        dropped = max(0, before - sum(len(c['leads']) for c in plan['courses']))   # 그사이 바뀌거나 취소된 현장
        source = 'saved'
    else:
        plan = _empty_plan()
        source = 'empty'
    confirmed_plan = _loads((row or {}).get('confirmed_plan'))
    urgent_map = _merge_urgent(day, plan, [v['lead_no'] for v in vs], names)
    return {
        'day': day, 'day_label': va.day_label(day), 'visits': vs, 'people': ppl,
        'off': off['off'], 'half': off['half'], 'plan': plan, 'source': source, 'dropped': dropped,
        'saved_at': (row or {}).get('saved_at') or '', 'saved_name': _name_of_email((row or {}).get('saved_by') or ''),
        'confirmed_at': (row or {}).get('confirmed_at') or '',
        'confirmed_name': _name_of_email((row or {}).get('confirmed_by') or ''),
        'changed_since_confirm': bool(confirmed_plan) and confirmed_plan != plan,
        'warnings': warnings(plan, vs, off, names),
        'run': _state(row),
        'editable': editable(day), 'can_plan': can_plan(me),
        'cars': cars_by_person(day),
        'urgent': urgent_map,
    }


def cars_by_person(day: str) -> Dict[str, List[Dict]]:
    """그날 차량 배정에 저장된 차 — {이니셜: [{plate, label}]}. 방문 일정 코스 제목 줄에 보기만(11-23 부장님).
    고치는 것은 차량 배정 화면에서만. 못 읽어도 방문 일정은 그대로 보이게 빈 값."""
    try:
        rows = va._day_rows(day)
        if not rows:
            return {}
        fl = {c['plate']: c for c in va.fleet(True)}
        out: Dict[str, List[Dict]] = {}
        for plate, r in rows.items():
            car = {'plate': plate, 'label': (fl.get(plate) or {}).get('label') or plate}
            for code in va.riders(r):
                if car not in out.setdefault(code, []):
                    out[code].append(car)
        return out
    except Exception as e:
        logger.warning(f'[방문일정] 차량 배정 읽기 실패: {e}')
        return {}


def save(day: str, raw: Dict, by: str, leads: Iterable[dict]) -> Dict:
    if not editable(day):
        raise ValueError('오늘과 지난 날 일정은 고칠 수 없어요.')
    vs = visits(day, leads)
    names = {p['code']: p['name'] for p in people(vs)}
    plan = clean_plan(raw, [v['lead_no'] for v in vs], names)
    with _conn() as conn:
        conn.execute('INSERT INTO visit_plan_day (day, plan, saved_by, saved_at) VALUES (?,?,?,?) '
                     'ON CONFLICT(day) DO UPDATE SET plan = excluded.plan, saved_by = excluded.saved_by, '
                     'saved_at = excluded.saved_at', (day, json.dumps(plan, ensure_ascii=False), by, _now()))
    return plan


# ── 확정 ────────────────────────────────────────────────────────────────
def start_confirm(day: str, raw: Dict, by: str, leads: Iterable[dict],
                  fresh_leads: Callable[[], List[dict]]) -> Dict:
    """저장 + 확정 시작(뒤에서 돈다). 막는 문제(error)가 있으면 ValueError."""
    leads = list(leads)
    plan = save(day, raw, by, leads)
    vs = visits(day, leads)
    names = {p['code']: p['name'] for p in people(vs)}
    errs = [w for w in warnings(plan, vs, va.day_off(day), names) if w['level'] == 'error']
    if errs:
        raise ValueError(errs[0]['text'])
    if not _staffed(plan):
        raise ValueError('사람과 현장을 넣은 코스가 없어요.')
    with _run_lock:
        st = _state(_row(day))
        if st['state'] == 'running':
            raise ValueError(f"{st['by']}님이 확정하는 중이에요. 잠시 뒤 새로 고쳐 주세요.")
        with _conn() as conn:
            conn.execute("UPDATE visit_plan_day SET run_state = 'running', run_at = ?, run_by = ?, run_result = NULL "
                         "WHERE day = ?", (_now(), by, day))
    threading.Thread(target=_run_confirm, args=(day, plan, by, fresh_leads),
                     name=f'visit-plan-{day}', daemon=True).start()
    return {'state': 'running'}


def build_parsed(day: str, plan: Dict, leads: Iterable[dict]) -> Dict:
    """일정 → visit_assignment_sync 가 받는 배정 모양(parse_assignment_canvas_full 과 같음).
    방문 순서 = 코스 순서 → 코스 안 가는 순서(DM 번호가 이 순서). 코스 메모는 배정마다 붙여 DM 위쪽에."""
    by_no = {str(l.get('리드 No') or '').strip(): l for l in leads}
    assignments = []
    for c in _staffed(plan):
        for n in c['leads']:
            lead = by_no.get(n)
            if lead is None:
                raise ValueError(f'리드 {n} 를 시트에서 찾지 못했어요. 새로 고친 뒤 다시 해 주세요.')
            assignments.append({'lead': lead, 'lead_no': n, 'phone': '', 'phone_digits': '',
                                'assign': list(c['members']), 'original': None, 'raw': '', 'memo': c['memo']})
    off = va.day_off(day)
    return {
        'assignments': assignments,
        'online_duty': [o['code'] for o in plan['online']],
        'online_duty_shifts': {o['code']: o['shift'] for o in plan['online'] if o['shift']},
        'online_shared': plan['online_shared'],
        'off_duty': [o['code'] for o in off['off']],
        'duplicates': [], 'unreadable': [],
    }


def _summary(res: Dict) -> Dict:
    dm = res.get('dm') or {}
    return {
        'sheet': res.get('updated_count', 0),
        'sheet_failed': [{'lead_no': a, 'error': str(b)[:120]} for a, b in (res.get('failed') or [])],
        'list': res.get('list_updated', 0), 'list_failed': res.get('list_failed', 0),
        'list_missing': res.get('list_missing') or [],
        'dm': dm.get('visit_mgr_sent', 0), 'dm_failed': dm.get('visit_mgr_failed', 0),
        'duty': dm.get('online_duty_sent', 0), 'deassign': dm.get('deassign_sent', 0),
        'target': dm.get('target_date') or '',
    }


def _finish(day: str, state: str, result: Dict, by: str = '', plan: Optional[Dict] = None) -> None:
    with _conn() as conn:
        if state == 'done':
            conn.execute('UPDATE visit_plan_day SET run_state = ?, run_result = ?, confirmed_by = ?, confirmed_at = ?, '
                         'confirmed_plan = ? WHERE day = ?',
                         (state, json.dumps(result, ensure_ascii=False), by, _now(),
                          json.dumps(plan, ensure_ascii=False), day))
        else:
            conn.execute('UPDATE visit_plan_day SET run_state = ?, run_result = ? WHERE day = ?',
                         (state, json.dumps(result, ensure_ascii=False), day))


def _work_map(plan: Optional[Dict]) -> Dict[str, Dict]:
    """사내 업무 — {이니셜: {'task', 'with': [같이 하는 사람]}} (할 일이 적힌 것만)."""
    out: Dict[str, Dict] = {}
    for c in (plan or {}).get('courses') or []:
        if is_work(c) and c.get('members') and c.get('memo'):
            for m in c['members']:
                out.setdefault(m, {'task': c['memo'], 'with': [x for x in c['members'] if x != m]})
    return out


def _send_work_dms(day: str, plan: Dict, prev: Optional[Dict]) -> Dict[str, int]:
    """사내 업무 DM — 처음 확정이면 모두, 다시 확정이면 할 일·같이 하는 사람이 바뀐 사람만(11-23 부장님)."""
    from dashboard.services.visit_assignment_sync import _dm_safe
    now, before = _work_map(plan), _work_map(prev)
    who = {p['code']: p for p in _staff()}
    sent = failed = 0
    for code, w in now.items():
        if before.get(code) == w:
            continue
        p = who.get(code)
        text = (f':office: *{va.day_label(day)} 사내 업무*\n할 일 : {_dm_safe(w["task"])}'
                + (f'\n함께 : {", ".join(_dm_safe(va.display(x, who)) for x in w["with"])}' if w['with'] else ''))
        if p and va._dm(p.get('slack_id') or '', text):
            sent += 1
        else:
            failed += 1
    return {'work_dm': sent, 'work_dm_failed': failed}


def _run_confirm(day: str, plan: Dict, by: str, fresh_leads: Callable[[], List[dict]]) -> None:
    from dashboard.services import visit_assignment_sync as vas
    prev = _loads((_row(day) or {}).get('confirmed_plan'))      # 다시 확정이면 사내 업무 DM 은 바뀐 사람만
    try:
        leads = fresh_leads()            # 확정은 시트를 새로 읽어서(지금 영업 담당자 값과 비교해야 함)
        if not leads:
            raise ValueError('리드 시트를 읽지 못했어요. 잠시 뒤 다시 해 주세요.')
        parsed = build_parsed(day, plan, leads)
        res = vas.commit_plan(parsed, target_date=date.fromisoformat(day))
        if not res.get('ok'):
            raise ValueError(str(res.get('reason') or '확정하지 못했어요.').lstrip('⏳ '))
    except ValueError as e:
        _finish(day, 'failed', {'message': str(e)})
        return
    except Exception as e:
        logger.error(f'[방문일정] {day} 확정 실패: {e}', exc_info=True)
        _finish(day, 'failed', {'message': '확정하지 못했어요. 잠시 뒤 다시 해 주세요.'})
        return
    summary = _summary(res)
    # /일정확정 뒤에 하던 것 — 온라인 당번 저장(차량 배정 명단), 차량 배정 요청 DM, 누락 점검
    try:
        va.record_online_duty(day, parsed['online_duty'], parsed['online_duty_shifts'])
    except Exception as e:
        logger.warning(f'[방문일정] 온라인 당번 저장 실패: {e}', exc_info=True)
    try:
        summary.update(_send_work_dms(day, plan, prev))
    except Exception as e:
        logger.warning(f'[방문일정] 사내 업무 DM 실패: {e}', exc_info=True)
    _finish(day, 'done', summary, by, plan)
    try:
        va.request_assignment(day)
    except Exception as e:
        logger.warning(f'[방문일정] 차량 배정 요청 DM 실패: {e}', exc_info=True)
    try:
        summary['cards'] = _audit(res, by)
        _finish(day, 'done', summary, by, plan)
    except Exception as e:
        logger.warning(f'[방문일정] 누락 점검 실패(확정은 정상): {e}', exc_info=True)
    logger.info(f'[방문일정] {day} 확정 — 시트 {summary["sheet"]} · DM {summary["dm"]} · 당번 {summary["duty"]} ({by})')


def _audit(res: Dict, by: str) -> List[Dict]:
    """진행 중 방문 카드인데 캔버스1에 없는 건 + List 선택지 없는 담당자 → 사장님 DM(있을 때만).
    캔버스2 대조는 하지 않는다(그룹웨어는 그날 방문을 다 보여 주고 '배정 안 된 방문'을 미리 띄움)."""
    from dashboard.services import visit_canvas_audit as vca
    au = {'cards': vca.check_cards_vs_canvas1(), 'c2_missing': [], 'c2_unmatched': [], 'target': None,
          'list_missing': list(res.get('list_missing') or [])}
    text = vca.format_commit_audit(au)
    if text:
        vca.send_admin_dm(f':clipboard: *방문 일정 확인 필요* (그룹웨어 확정: {_name_of_email(by)})\n{text}')
    return au['cards']


def run_state(day: str) -> Dict:
    return _state(_row(day))


# ── 당일 배정(11-23 부장님) ──────────────────────────────────────────────
#    그날 생긴 현장은 온라인 당번이 고객과 통화한 뒤 근처 매니저에게 바로 보낸다 — 그 결과를 시트·List 에 남기고
#    매니저에게 긴급 DM(주소·연락처·문의). 이미 정한 코스는 건드리지 않는다. 오늘만.
def _local(email: str) -> str:
    return (email or '').strip().lower().split('@')[0].upper()


def can_urgent(email: str, day: str) -> bool:
    """당일 배정 — 오늘만, 일정 담당 또는 그날 온라인 당번."""
    if day != date.today().isoformat() or not email:
        return False
    if can_plan(email):
        return True
    try:
        return _local(email) in {o['code'] for o in va.online_duty(day)}
    except Exception:
        return False


def is_duty_today(email: str) -> bool:
    """오늘 온라인 당번 — 방문 일정 메뉴·화면을 열어 준다(당일 배정하려고)."""
    return can_urgent(email, date.today().isoformat()) and not can_plan(email)


def _urgent_rows(day: str) -> List[Dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute('SELECT * FROM visit_plan_urgent WHERE day = ? ORDER BY at', (day,)).fetchall()]


def _merge_urgent(day: str, plan: Dict, nos: List[str], names: Dict[str, str]) -> Dict[str, Dict]:
    """당일 배정 기록을 코스에 붙인다(같은 사람 코스가 있으면 그 끝, 없으면 새 코스) → {리드 No: {at, by_name, members}}."""
    out: Dict[str, Dict] = {}
    try:
        rows = _urgent_rows(day)
    except Exception as e:
        logger.warning(f'[방문일정] 당일 배정 읽기 실패: {e}')
        return out
    placed = {n for c in plan['courses'] for n in c['leads']}
    for r in rows:
        members = [m for m in (r.get('members') or '').split(',') if m in names]
        n = r['lead_no']
        out[n] = {'at': (r.get('at') or '')[11:16], 'by_name': _name_of_email(r.get('by') or ''), 'members': members,
                  'state': r.get('state') or 'done', 'result': _loads(r.get('result')) or {}}
        if out[n]['state'] == 'failed' or n in placed or n not in nos or not members:
            continue
        c = next((c for c in plan['courses'] if not is_work(c) and c['members'] == members), None)
        if c is None:
            c = {'name': '', 'members': members, 'leads': [], 'memo': ''}
            plan['courses'].append(c)
        c['leads'].append(n)
        placed.add(n)
    return out


def _slack_of(code: str, name: str) -> str:
    p = next((p for p in _staff() if p['code'] == code), None)
    if p and p.get('slack_id'):
        return p['slack_id']
    try:
        from dashboard.services import hr
        r = hr._rows("SELECT slack_id FROM hr_people WHERE name = ? AND status = '재직'", (name,))
        return (r[0].get('slack_id') or '') if r else ''
    except Exception:
        return ''


def urgent_assign(day: str, lead_no: str, members: List[str], by: str, leads: Iterable[dict],
                  background: bool = True) -> Dict:
    """당일 배정 — 확인하고 기록만 남긴 뒤 바로 돌려준다. 시트 영업 담당자·슬랙 List·캔버스1·긴급 DM 은 뒤에서
    (11-23 부장님: 한 번에 11초쯤 걸려 화면이 멈춘 줄 알았음). 막는 문제는 ValueError / PermissionError."""
    if not can_urgent(by, day):
        raise PermissionError('당일 배정은 오늘 일정만, 일정 담당이나 오늘 온라인 당번이 할 수 있어요.')
    leads = list(leads)
    vs = visits(day, leads, done=True)
    v = next((x for x in vs if x['lead_no'] == lead_no), None)
    lead = next((l for l in leads if str(l.get('리드 No') or '').strip() == lead_no), None)
    if not v or not lead:
        raise ValueError('그 현장을 오늘 방문 일정에서 찾지 못했어요. 새로 고친 뒤 다시 해 주세요.')
    if str(lead.get('영업 담당자') or '').strip() not in ('', '-'):
        raise ValueError('이미 담당자가 정해진 현장이에요. 새로 고친 뒤 확인해 주세요.')
    saved = _loads((_row(day) or {}).get('plan')) or {}
    if any(lead_no in (c.get('leads') or []) for c in saved.get('courses') or []):
        raise ValueError('이미 코스에 들어 있는 현장이에요.')
    ppl = people(vs)
    names = {p['code']: p['name'] for p in ppl}
    members = [m for m in dict.fromkeys(str(x or '').strip().upper() for x in members or []) if m]
    if not members or any(m not in names for m in members):
        raise ValueError('맡길 사람을 골라 주세요.')
    hard = {o['code']: o['why'] for o in va.day_off(day)['off'] if not o.get('soft')}
    bad = [m for m in members if m in hard]
    if bad:
        raise ValueError(f'{names[bad[0]]}님은 오늘 {hard[bad[0]]}예요.')
    with _conn() as conn:          # 두 사람이 동시에 눌러도 한 번만(먼저 들어온 것). 실패했던 건은 다시 할 수 있게
        conn.execute("DELETE FROM visit_plan_urgent WHERE day = ? AND lead_no = ? AND state = 'failed'", (day, lead_no))
        try:
            conn.execute("INSERT INTO visit_plan_urgent (day, lead_no, members, by, at, state) VALUES (?,?,?,?,?,'running')",
                         (day, lead_no, ','.join(members), by, _now()))
        except Exception:
            raise ValueError('방금 다른 사람이 이 현장을 당일 배정했어요. 새로 고쳐 주세요.')
    names_txt = ','.join(names[m] for m in members)
    job = dict(day=day, lead_no=lead_no, lead=lead, v=v, members=members, names=names, names_txt=names_txt, by=by)
    if background:
        threading.Thread(target=_urgent_work, kwargs=job, daemon=True, name=f'urgent-{lead_no}').start()
        return {'names': names_txt, 'state': 'running'}
    return {'names': names_txt, **_urgent_work(**job)}


def _urgent_done(day: str, lead_no: str, state: str, result: Dict) -> None:
    with _conn() as conn:
        conn.execute('UPDATE visit_plan_urgent SET state = ?, result = ? WHERE day = ? AND lead_no = ?',
                     (state, json.dumps(result, ensure_ascii=False), day, lead_no))


def _urgent_work(day: str, lead_no: str, lead: Dict, v: Dict, members: List[str], names: Dict[str, str],
                 names_txt: str, by: str) -> Dict:
    """뒤에서 — 시트 영업 담당자 → 슬랙 List → 캔버스1 → 긴급 DM. 결과는 visit_plan_urgent.state/result."""
    from dashboard.services import visit_assignment_sync as vas
    from dashboard.services.lead_service import update_lead
    try:
        res = update_lead(lead_no, {'영업 담당자': names_txt})
        if isinstance(res, dict) and res.get('success') is False:
            raise RuntimeError(res.get('message') or '시트 저장 실패')
    except Exception as e:
        logger.error(f'[방문일정] 당일 배정 시트 기록 실패 {lead_no}: {e}', exc_info=True)
        out = {'message': '시트에 담당자를 적지 못했어요. 다시 해 주세요.'}
        _urgent_done(day, lead_no, 'failed', out)
        return {'dm': 0, 'dm_failed': 0, **out}
    try:       # 슬랙 List 담당자 칸도 같이(바뀌면 슬랙 쪽이 시트를 List 값으로 덮어써서 맞춰 둔다)
        vas._update_slack_list_managers([{'lead': lead, 'lead_no': lead_no, 'phone': '', 'phone_digits': '',
                                          'assign': members, 'original': None, 'raw': ''}], {}, sheet_leads=None)
    except Exception as e:
        logger.warning(f'[방문일정] 당일 배정 List 실패 {lead_no}: {e}')
    try:
        from dashboard.services.visit_canvas_sync import rebuild_canvas_async
        rebuild_canvas_async()
    except Exception as e:
        logger.warning(f'[방문일정] 당일 배정 캔버스 실패: {e}')
    link = ''
    try:
        client = vas._get_visit_client()
        link = vas._get_visit_card_permalink(client, lead_no) if client else ''
    except Exception:
        link = ''
    safe = vas._dm_safe
    by_name = _name_of_email(by)
    sent = failed = 0
    for m in members:
        others = [names[x] for x in members if x != m]
        text = (f':rotating_light: *{va.day_label(day)} 당일 배정 — 오늘 추가 방문*\n'
                f'{safe(v.get("addr") or "-")}\n'
                + (f'연락처 : {safe(v["phone"])}\n' if v.get('phone') else '')
                + (f'{safe(v["text"])}\n' if v.get('text') else '')
                + (f'함께 : {safe(", ".join(others))}\n' if others else '')
                + f'배정 : {safe(by_name)} · {lead_no}'
                + (f'\n<{link}|방문 카드 보기>' if link else ''))
        if va._dm(_slack_of(m, names[m]), text):
            sent += 1
        else:
            failed += 1
    out = {'dm': sent, 'dm_failed': failed}
    _urgent_done(day, lead_no, 'done', out)
    logger.info(f'[방문일정] {day} 당일 배정 {lead_no} → {names_txt} (DM {sent}/{len(members)}, {by})')
    return out


# ── 다른 곳에서 쓰는 것 ──────────────────────────────────────────────────
def confirmed_upcoming_day() -> Optional[Dict]:
    """슬랙 /일정확정이 덮어쓰지 않게 — 다음 근무일 일정을 그룹웨어에서 확정했으면 {'day','label','by'}."""
    day = va.next_workday()
    row = _row(day)
    if not row or not row.get('confirmed_at'):
        return None
    return {'day': day, 'label': va.day_label(day), 'by': _name_of_email(row.get('confirmed_by') or '')}


def confirmed_info(day: str) -> Optional[Dict]:
    """그 날 방문 일정을 그룹웨어에서 확정했으면 {'confirmed_at', 'confirmed_name', 'courses'} — 차량 배정 위 한 줄."""
    row = _row(day)
    plan = _loads((row or {}).get('confirmed_plan'))
    if not row or not row.get('confirmed_at') or not plan:
        return None
    courses = [c for c in plan.get('courses') or _old_courses(plan) or [] if c.get('members') and c.get('leads')]
    return {'confirmed_at': row['confirmed_at'], 'confirmed_name': _name_of_email(row.get('confirmed_by') or ''),
            'courses': len(courses)}


def confirmed_work(day: str) -> List[Dict]:
    """그룹웨어에서 확정한 날의 사내 업무 — [{'code', 'task'}]. 차량 배정은 이 사람들을 '배정 없음'에서 뺀다(회사 안 일)."""
    plan = _loads((_row(day) or {}).get('confirmed_plan'))
    return [{'code': c, 'task': w['task']} for c, w in _work_map(plan).items()]


def confirmed_teams(day: str) -> Optional[List[Dict]]:
    """그룹웨어에서 확정한 날의 팀(같이 가는 사람 묶음) — 차량 배정이 시트(캐시)보다 먼저 쓴다. 없으면 None.
    [{'members', 'visits', 'big'}] (va.teams 와 같은 모양·같은 묶는 규칙)"""
    from collections import Counter
    plan = _loads((_row(day) or {}).get('confirmed_plan'))
    if not plan:
        return None
    combos: Counter = Counter()
    for c in plan.get('courses') or _old_courses(plan) or []:
        if c.get('members') and c.get('leads'):
            combos[tuple(c['members'])] += len(c['leads'])
    return va.teams_from_combos(combos)
