"""그룹웨어 방문 일정 (2026-11-22 부장님) — 다음 날 방문을 누가 가는지 그룹웨어에서 짜고 확정한다.

지금까지는 JW 가 슬랙 캔버스2에 'YM+LSH' 같은 팀 이름과 방문 줄을 적고 /일정확정을 실행했다.
→ 차량 배정 화면 ① 방문 일정 단계에서 방문 카드를 팀에 넣고 [일정 확정]을 누른다.
   확정 처리(시트 영업 담당자·슬랙 List·담당자/당번 DM·캔버스1)는 /일정확정과 같은 본체
   (visit_assignment_sync.commit_plan)를 쓴다. 리드를 골라 넣으므로 전화번호 매칭 실패·이니셜 판독 불가가 없다.
- 휴무·반차는 그룹웨어 휴가·토요 휴무표에서(캔버스2 '휴무' 칸 대신).
- 팀 메모는 담당자 DM 에 함께 간다(캔버스2를 내리면 메모 볼 곳이 DM 뿐).
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

MAX_TEAMS = 40
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


def _who_codes(names: Iterable[str], by_name: Dict[str, str]) -> List[str]:
    return list(dict.fromkeys(by_name[n] for n in names if n in by_name))


# ── 그날 방문 ───────────────────────────────────────────────────────────
def visits(day: str, leads: Iterable[dict]) -> List[Dict]:
    """그날 방문(캔버스1과 같은 묶음, va.day_schedule) 을 한 줄로 — 카드 하나 = 리드 하나.
    본인 방문 필수면 owner = 상담한 사람 이니셜(온라인 상담자)."""
    leads = list(leads)
    by_name = _code_maps()
    by_no = {str(l.get('리드 No') or '').strip(): l for l in leads}
    out = []
    for g in va.day_schedule(day, leads):
        for it in g['items']:
            if not it.get('lead_no'):
                continue
            owner = ''
            if '본인 방문 필수' in it['marks']:
                owner = by_name.get(str((by_no.get(it['lead_no']) or {}).get('온라인 상담자') or '').strip(), '')
            out.append({**it, 'cat': g['cat'], 'who_codes': _who_codes(it['who'], by_name), 'owner': owner})
    return out


def people(vs: List[Dict]) -> List[Dict]:
    """팀에 넣을 수 있는 사람 — 현장 명단 + 그날 방문에 이미 적힌 다른 사람(사무직 등)."""
    out = [{'code': p['code'], 'name': p['name']} for p in va.staff()]
    have = {p['code'] for p in out}
    from dashboard.services.visit_assignment_sync import _load_initial_maps
    ini_to_name, _ = _load_initial_maps()
    for v in vs:
        for c in v['who_codes']:
            if c not in have and c in ini_to_name:
                have.add(c)
                out.append({'code': c, 'name': ini_to_name[c], 'extra': True})
    return out


def _prefill(vs: List[Dict]) -> Dict:
    """저장한 일정이 없는 날 — 시트에 이미 적힌 영업 담당자대로 팀을 묶어 둔다(없으면 배정 안 됨 칸)."""
    teams: Dict[tuple, Dict] = {}
    for v in vs:
        key = tuple(v['who_codes'])
        if not key:
            continue
        teams.setdefault(key, {'members': list(key), 'leads': [], 'memo': ''})['leads'].append(v['lead_no'])
    return {'teams': list(teams.values()), 'online': [], 'online_shared': False}


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


def clean_plan(raw: Dict, visit_nos: Iterable[str], codes: Iterable[str], strict: bool = True) -> Dict:
    """화면에서 온 일정을 다듬는다. strict=False 면 모르는 방문·사람은 조용히 뺀다(불러올 때)."""
    if not isinstance(raw, dict):
        raise ValueError('일정 내용을 읽지 못했어요.')
    visit_nos, codes = set(visit_nos), set(codes)
    teams_in = raw.get('teams') or []
    if not isinstance(teams_in, list) or len(teams_in) > MAX_TEAMS:
        raise ValueError('팀이 너무 많아요.')
    seen_leads: set = set()
    teams = []
    for t in teams_in:
        if not isinstance(t, dict):
            continue
        members = []
        for c in t.get('members') or []:
            c = str(c or '').strip().upper()
            if c in codes and c not in members:
                members.append(c)
            elif c and c not in codes and strict:
                raise ValueError(f'명단에 없는 사람({c})이 팀에 있어요. 새로 고친 뒤 다시 해 주세요.')
        leads = []
        for n in t.get('leads') or []:
            n = str(n or '').strip()
            if n in visit_nos and n not in seen_leads:
                seen_leads.add(n)
                leads.append(n)
            elif n and n not in visit_nos and strict:
                raise ValueError('그사이 방문 일정이 바뀐 건이 있어요. 새로 고친 뒤 다시 해 주세요.')
        memo = re.sub(r'\s+', ' ', str(t.get('memo') or '')).strip()[:MEMO_MAX]
        if members or leads or memo:
            teams.append({'members': members, 'leads': leads, 'memo': memo})
    online = []
    for o in raw.get('online') or []:
        if not isinstance(o, dict):
            continue
        c = str(o.get('code') or '').strip().upper()
        shift = str(o.get('shift') or '').strip()
        if c in codes and all(x['code'] != c for x in online):
            online.append({'code': c, 'shift': shift if shift in SHIFTS else ''})
    return {'teams': teams, 'online': online, 'online_shared': bool(raw.get('online_shared')) and not online}


def editable(day: str, today: Optional[date] = None) -> bool:
    """확정은 내일부터 — 오늘·지난 날은 DM 대상이 아니다(방문 끝난 날은 담당 DM 이 안 나감)."""
    return va.valid_day(day) and date.fromisoformat(day) > (today or date.today())


def warnings(plan: Dict, vs: List[Dict], off: Dict[str, List[Dict]], names: Dict[str, str]) -> List[Dict]:
    """확정 전 점검 — level 'error' 는 확정을 막고, 'warn' 은 확인만 받는다."""
    out: List[Dict] = []
    by_no = {v['lead_no']: v for v in vs}
    placed = {n for t in plan['teams'] for n in t['leads']}
    left = [v for v in vs if v['lead_no'] not in placed]
    if left:
        out.append({'level': 'warn', 'kind': 'left',
                    'text': f'배정 안 된 방문 {len(left)}건 — 담당 DM 이 가지 않아요.'})
    for i, t in enumerate(plan['teams']):
        if t['leads'] and not t['members']:
            out.append({'level': 'error', 'kind': 'nobody', 'team': i,
                        'text': f'사람을 넣지 않은 팀에 방문 {len(t["leads"])}건이 있어요.'})
    offs = {o['code']: o['why'] for o in off.get('off', [])}
    count: Dict[str, int] = {}
    for i, t in enumerate(plan['teams']):
        for c in t['members']:
            count[c] = count.get(c, 0) + 1
            if c in offs:
                out.append({'level': 'warn', 'kind': 'off', 'team': i,
                            'text': f'{names.get(c, c)}님은 이날 {offs[c]}예요.'})
        for n in t['leads']:
            v = by_no.get(n) or {}
            if v.get('owner') and v['owner'] not in t['members']:
                out.append({'level': 'warn', 'kind': 'owner', 'team': i,
                            'text': f'본인 방문 필수 건({(v.get("addr") or n)[:30]})인데 상담한 '
                                    f'{names.get(v["owner"], v["owner"])}님이 팀에 없어요.'})
    for c, n in count.items():
        if n > 1:
            out.append({'level': 'warn', 'kind': 'twice',
                        'text': f'{names.get(c, c)}님이 팀 {n}곳에 들어가 있어요.'})
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
        before = sum(len(t.get('leads') or []) for t in saved.get('teams') or [])
        plan = clean_plan(saved, [v['lead_no'] for v in vs], names, strict=False)
        dropped = before - sum(len(t['leads']) for t in plan['teams'])   # 그사이 날짜가 바뀌거나 취소된 방문
        source = 'saved'
    else:
        plan = _prefill(vs)
        plan['online'] = [{'code': o['code'], 'shift': o['shift']} for o in va.online_duty(day)]
        source = 'sheet' if plan['teams'] else 'empty'
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
    if not any(t['leads'] and t['members'] for t in plan['teams']):
        raise ValueError('팀에 넣은 방문이 없어요.')
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
    방문 순서 = 팀 순서 → 팀 안 순서(DM 번호가 이 순서)."""
    by_no = {str(l.get('리드 No') or '').strip(): l for l in leads}
    assignments = []
    for t in plan['teams']:
        if not t['members']:
            continue
        for n in t['leads']:
            lead = by_no.get(n)
            if lead is None:
                raise ValueError(f'리드 {n} 를 시트에서 찾지 못했어요. 새로 고친 뒤 다시 해 주세요.')
            assignments.append({'lead': lead, 'lead_no': n, 'phone': '', 'phone_digits': '',
                                'assign': list(t['members']), 'original': None, 'raw': '', 'memo': t['memo']})
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
    """그룹웨어에서 확정한 날의 팀 — 차량 배정이 시트(캐시)보다 먼저 쓴다. 없으면 None.
    [{'members', 'visits', 'big'}] (va.teams 와 같은 모양)"""
    row = _row(day)
    plan = _loads((row or {}).get('confirmed_plan'))
    if not plan:
        return None
    out = []
    for t in plan.get('teams') or []:
        if t.get('members') and t.get('leads'):
            out.append({'members': list(t['members']), 'visits': len(t['leads']),
                        'big': len(t['members']) >= va.BIG_TEAM})
    return out
