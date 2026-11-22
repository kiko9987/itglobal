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
_table_ready = False
_run_lock = threading.Lock()


@contextmanager
def _conn():
    global _table_ready
    with va._db()._get_connection() as conn:
        if not _table_ready:
            conn.execute(_DDL)
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


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
    out.update({p['name']: p['code'] for p in va.staff()})
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


def visits(day: str, leads: Iterable[dict]) -> List[Dict]:
    """그날 현장(캔버스1과 같은 리드·묶음) 을 한 줄로 — 카드 하나 = 리드 하나.
    본인 방문 필수면 owner = 상담한 사람 이니셜(온라인 상담자)."""
    from dashboard.services import vehicle_log as vl
    leads = _open_leads(leads)
    by_name = _code_maps()
    by_no = {str(l.get('리드 No') or '').strip(): l for l in leads}
    out = []
    try:
        dongs = vl._dong_cache([str(l.get('방문 주소') or '').strip() for l in leads])
    except Exception:
        dongs = {}
    for g in va.day_schedule(day, leads):
        for it in g['items']:
            if not it.get('lead_no'):
                continue
            owner = ''
            if '본인 방문 필수' in it['marks']:
                owner = by_name.get(str((by_no.get(it['lead_no']) or {}).get('온라인 상담자') or '').strip(), '')
            item = {k: v for k, v in it.items() if k != 'who'}      # 시트 담당자(=캔버스2 결과)는 보이지 않는다
            out.append({**item, 'cat': g['cat'], 'title': g['title'], 'owner': owner, **_where(it['addr'], dongs, vl)})
    return out


_BIG_CITIES = ('서울', '인천', '부산', '대구', '광주', '대전', '울산')


def _where(addr: str, dongs: Dict[str, str], vl) -> Dict[str, str]:
    """현장 꼬리표 — place = '평택 고덕동'(동 모르면 구), region = 코스 나눠 보기 단위(경기는 시, 서울·인천은 구까지)."""
    city, gu, dong = vl.area(addr)
    if not dong and dongs.get(addr):
        dong = dongs[addr]
    region = (f'{city} {gu}'.strip() if city in _BIG_CITIES else city) or gu or ''
    return {'place': vl.place_label((city, gu, dong), addr), 'region': region}


def people(vs: List[Dict]) -> List[Dict]:
    """팀에 넣을 수 있는 사람 — 현장 명단 + 본인 방문 필수 건을 상담한 다른 사람(사무직 등)."""
    out = [{'code': p['code'], 'name': p['name']} for p in va.staff()]
    have = {p['code'] for p in out}
    from dashboard.services.visit_assignment_sync import _load_initial_maps
    ini_to_name, _ = _load_initial_maps()
    for v in vs:
        c = v.get('owner') or ''
        if c and c not in have and c in ini_to_name:
            have.add(c)
            out.append({'code': c, 'name': ini_to_name[c], 'extra': True})
    return out


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
    plan = {'courses': [{'name', 'members': [이니셜], 'leads': [리드 No(가는 순서)], 'memo'}], 'online', 'online_shared'}"""
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
        leads = []
        for n in t.get('leads') or []:
            n = str(n or '').strip()
            if n in visit_nos and n not in seen:
                seen.add(n)
                leads.append(n)
            elif n and n not in visit_nos and strict:
                raise ValueError('그사이 방문 일정이 바뀐 건이 있어요. 새로 고친 뒤 다시 해 주세요.')
        name = re.sub(r'\s+', ' ', str(t.get('name') or '')).strip()[:NAME_MAX]
        memo = re.sub(r'\s+', ' ', str(t.get('memo') or '')).strip()[:MEMO_MAX]
        if members or leads or memo or name:
            courses.append({'name': name, 'members': members, 'leads': leads, 'memo': memo})
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


def course_label(c: Dict, i: int) -> str:
    return c.get('name') or f'코스 {i + 1}'


def editable(day: str, today: Optional[date] = None) -> bool:
    """확정은 내일부터 — 오늘·지난 날은 DM 대상이 아니다(방문 끝난 날은 담당 DM 이 안 나감)."""
    return va.valid_day(day) and date.fromisoformat(day) > (today or date.today())


def warnings(plan: Dict, vs: List[Dict], off: Dict[str, List[Dict]], names: Dict[str, str]) -> List[Dict]:
    """확정 전 점검 — 모두 확인만 받는다(level 'warn')."""
    out: List[Dict] = []
    by_no = {v['lead_no']: v for v in vs}
    done = {n for c in _staffed(plan) for n in c['leads']}
    left = [v for v in vs if v['lead_no'] not in done]
    if left:
        out.append({'level': 'warn', 'kind': 'left',
                    'text': f'코스·사람이 안 정해진 현장 {len(left)}건 — 담당 DM 이 가지 않아요.'})
    offs = {o['code']: o['why'] for o in off.get('off', [])}
    count: Dict[str, int] = {}
    for i, c in enumerate(plan['courses']):
        label = course_label(c, i)
        if c['leads'] and not c['members']:
            out.append({'level': 'warn', 'kind': 'nobody', 'course': i,
                        'text': f'{label}에 사람을 넣지 않았어요(현장 {len(c["leads"])}건).'})
        for m in c['members']:
            count[m] = count.get(m, 0) + 1
            if m in offs:
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
            out.append({'level': 'warn', 'kind': 'off',
                        'text': f'온라인 당번 {names.get(o["code"], o["code"])}님은 이날 {offs[o["code"]]}예요.'})
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
    vs = visits(day, leads)
    ppl = people(vs)
    names = {p['code']: p['name'] for p in ppl}
    off = va.day_off(day)
    row = _row(day)
    saved = _loads((row or {}).get('plan'))
    dropped = 0
    if saved:
        plan = clean_plan(saved, [v['lead_no'] for v in vs], names, strict=False)
        before = sum(len(c.get('leads') or []) for c in (saved.get('courses') or _old_courses(saved) or [])
                     if isinstance(c, dict))
        dropped = max(0, before - sum(len(c['leads']) for c in plan['courses']))   # 그사이 바뀌거나 취소된 현장
        source = 'saved'
    else:
        plan = _empty_plan()
        source = 'empty'
    confirmed_plan = _loads((row or {}).get('confirmed_plan'))
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
    }


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


def _run_confirm(day: str, plan: Dict, by: str, fresh_leads: Callable[[], List[dict]]) -> None:
    from dashboard.services import visit_assignment_sync as vas
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


# ── 다른 곳에서 쓰는 것 ──────────────────────────────────────────────────
def confirmed_upcoming_day() -> Optional[Dict]:
    """슬랙 /일정확정이 덮어쓰지 않게 — 다음 근무일 일정을 그룹웨어에서 확정했으면 {'day','label','by'}."""
    day = va.next_workday()
    row = _row(day)
    if not row or not row.get('confirmed_at'):
        return None
    return {'day': day, 'label': va.day_label(day), 'by': _name_of_email(row.get('confirmed_by') or '')}


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
