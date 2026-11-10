"""개인 원가 탭(박용구·박정우) ↔ 메인 '공사 현황' 자동 동기화 (2026-11-10).

경영지원(SB)이 영업 담당자 두 사람의 원가를 대량으로 넣을 때는 PM 화면보다 시트 탭이 편하다.
그래서 탭은 남기되, SB 는 원가 4칸(제품대·도급비·자재비·기타비)만 넣고 나머지는 서버가 맞춘다.

  · 탭 줄 = 그 담당자의 공사 하나. 코드(A열)는 수식이 아닌 값 — 메인에서 행이 지워져도 밀리지 않는다.
    (예전 탭은 메인 행 위치를 그대로 비춰서, 메인 행 삭제 뒤 원가가 다른 공사 옆으로 밀렸다.)
  · 보기 칸(사업자·주소·공사 내용·금액·수금 …)은 서버가 메인 값을 그대로 넣는다.
  · 순익·마진율은 탭 안에서 같은 줄 수식으로 바로 계산 (메인 AJ·AK 와 같은 식).
  · 원가는 양방향: 탭에서 바꾸면 메인(=PM)으로, PM 에서 바꾸면 탭으로. 마지막으로 맞춘 값을
    Redis 에 두고 어느 쪽이 바뀌었는지 판단한다. 양쪽이 같이 바뀌면 탭 값을 따른다(원가 입력 주인이 SB).
  · 새 공사가 생기면 맨 아래에 줄을 붙이고, 담당자가 바뀌었거나 없어진 코드 줄은 원가가 이미 맞춰진
    경우에만 지운다(안 맞춰진 원가가 있으면 그대로 두고 기록만).
  · 메인 쓰기는 시트 쓰기 대기열(personal_cost_push)로 — 바뀐 원가 4칸 + _version+1 만 쓴다.
  · 한 번에 너무 많은 원가가 메인으로 가려 하면(실수로 칸을 지운 경우 등) 쓰지 않고 관리자에게 알린다.

탭 머리글이 아래 HEADERS 와 다르면(아직 재구성 전) 그 탭은 건드리지 않는다.
처음 만들기·다시 만들기: scripts/rebuild_personal_cost_tabs.py (rebuild_tab).
"""

import json
import os
import time
from typing import Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

COST = '__cost__'
CALC = '__calc__'

# (탭 머리글, 메인 머리글 | COST | CALC) — 38열 A~AL
LAYOUT = [
    ('프로젝트 No.', '프로젝트 코드'),
    ('사업자', '사업자'),
    ('담당자', '담당자'),
    ('거래처', '거래처명'),
    ('사업자명', '사업자명'),
    ('현장 주소', '현장 주소'),
    ('공사 구분', '공사 구분'),
    ('기계 분류', '기계 분류'),
    ('브랜드', '브랜드'),
    ('공사 시작', '공사 시작'),
    ('공사 종료', '공사 종료'),
    ('공사 내용', '공사 내용'),
    ('도급 구분', '도급 구분'),
    ('시공자', '시공자'),
    ('현장 담당자', '발주처 담당자'),
    ('담당자 연락처', '발주처 연락처'),
    ('담당자 이메일', '발주처 이메일'),
    ('총액1', '총액 1'),
    ('부가세', '부가세'),
    ('총액2', '총액 2'),
    ('계약금', '계약금'),
    ('중도금', '중도금'),
    ('잔금', '잔금'),
    ('미수금', '미수금'),
    ('계산서', '계산서'),
    ('수금 날짜', '수금 날짜'),
    ('수금 확인', '수금 확인'),
    ('제품대', COST),
    ('도급비', COST),
    ('자재비', COST),
    ('기타비', COST),
    ('순익', CALC),
    ('마진율', CALC),
    ('비고', '수금 관련 특이사항'),
    ('계약금 부담자', '계약금 부담자'),
    ('중도금 부담자', '중도금 부담자'),
    ('잔금 부담자', '잔금 부담자'),
    ('견적서 및 계약서 폴더 경로', '견적서 및 계약서 폴더 경로'),
]
HEADERS = [h for h, _ in LAYOUT]
WIDTH = len(LAYOUT)
COST_FIELDS = ('제품대', '도급비', '자재비', '기타비')            # 메인 머리글 = 탭 머리글
COST_IDX = [HEADERS.index(c) for c in COST_FIELDS]           # 27..30 (AB..AE)
TOTAL_IDX = HEADERS.index('총액1')                            # 17 (R)
PROFIT_IDX, MARGIN_IDX = HEADERS.index('순익'), HEADERS.index('마진율')   # 31·32 (AF·AG)

LAST_KEY = 'personal_cost_tab:last'        # hash code → json [제품대,도급비,자재비,기타비] 마지막으로 맞춘 값
PENDING_KEY = 'personal_cost_tab:pending'  # hash code → json 원가 — 메인 쓰기 대기열에 들어간 것(쓰이면 지움)
PENDING_TTL = 900
QUEUE_OP = 'personal_cost_push'


def col_letter(i: int) -> str:
    s, i = '', i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


TAB_LAST_COL = col_letter(WIDTH - 1)       # AL


def tab_names() -> list:
    raw = os.getenv('PERSONAL_COST_TABS', '박용구,박정우')
    return [t.strip() for t in raw.split(',') if t.strip()]


def max_push() -> int:
    try:
        return int(os.getenv('PERSONAL_COST_TAB_MAX_PUSH', '40'))
    except ValueError:
        return 40


# ── 값 다루기 ─────────────────────────────────────────────

def cost_num(v) -> int:
    """원가 칸 값 → 정수 원. 빈칸·글자·True/False 는 0. '1,200,000원' 같은 글자 숫자도 읽는다."""
    if v is None or isinstance(v, bool):
        return 0
    if isinstance(v, (int, float)):
        return int(round(v))
    s = str(v).strip().replace(',', '').replace('원', '').replace(' ', '')
    if not s:
        return 0
    try:
        return int(round(float(s)))
    except ValueError:
        return 0


def costs_of(row: list, idx: list) -> list:
    return [cost_num(row[i] if i < len(row) else '') for i in idx]


def same_value(a, b) -> bool:
    """보기 칸 비교 — True 와 1 을 같게 보지 않고, 3.0 과 3 은 같게 본다."""
    def k(v):
        if v is None:
            return ('s', '')
        if isinstance(v, bool):
            return ('b', v)
        if isinstance(v, (int, float)):
            return ('n', round(float(v), 6))
        return ('s', str(v))
    return k(a) == k(b)


def profit_formula(r: int) -> str:
    # 메인 AJ = S-(AF+AG+AH+AI) 와 같은 식 (탭: R=총액1, AB~AE=원가)
    return f'=IF($A{r}="","",N($R{r})-SUM($AB{r}:$AE{r}))'


def margin_formula(r: int) -> str:
    # 메인 AK = IF(OR(S=0,AJ=0),0,AJ/S)
    return f'=IF($A{r}="","",IF(OR(N($R{r})=0,N($AF{r})=0),0,$AF{r}/$R{r}))'


def decide(t: list, m: list, last: Optional[list]) -> str:
    """원가 한 줄 판단 — 'same' | 'to_main' | 'to_tab' | 'conflict_tab'.

    t=탭, m=메인, last=마지막으로 맞춘 값(없으면 None).
    """
    if t == m:
        return 'same'
    if last is None:
        # 맞춘 기록이 없음(재구성 직후 탭에만 있던 원가, Redis 유실 등): 탭이 비었으면 메인을 살리고, 아니면 탭.
        return 'to_tab' if not any(t) else 'to_main'
    if m == last:
        return 'to_main'
    if t == last:
        return 'to_tab'
    return 'conflict_tab'


def main_row_values(main_row: list, main_idx: dict) -> list:
    """메인 행 → 탭 줄 값 (원가 칸은 메인 원가, 계산 칸은 빈값 — 수식은 따로 씀)."""
    out = []
    for h, src in LAYOUT:
        if src == CALC:
            out.append('')
        else:
            key = h if src == COST else src
            i = main_idx.get(key)
            v = main_row[i] if i is not None and i < len(main_row) else ''
            out.append(cost_num(v) if src == COST else ('' if v is None else v))
    return out


# ── 계획 (순수 함수 — 시트·Redis 안 건드림) ─────────────────────

def plan_tab(tab: str, tab_vals: list, tab_formulas: list, main_by_code: dict, own_codes: list,
             main_idx: dict, last: dict, pending: dict) -> dict:
    """탭 하나를 메인에 맞추는 계획.

    tab_vals: 탭 A1:AL (UNFORMATTED, 0번=머리글). tab_formulas: AF2:AG (FORMULA) 줄 목록.
    main_by_code: 코드 → 메인 행 값. own_codes: 이 담당자 공사 코드(메인 순서). last/pending: code → [4].
    Returns dict(cells=[(row, col, value)], formulas=[(row, col, formula)], appends=[[values]],
                 pushes=[(code, costs)], last=[(code, costs)], deletes=[row], notes=[str])
    """
    plan = {'cells': [], 'formulas': [], 'appends': [], 'pushes': [], 'last': [], 'deletes': [], 'notes': []}
    own = set(own_codes)
    seen = {}
    for r, row in enumerate(tab_vals[1:], start=2):
        code = str(row[0]).strip() if row and row[0] is not None else ''
        if not code:
            continue
        if code in seen:
            plan['notes'].append(f'{tab} {r}행 {code} — 같은 코드 줄이 {seen[code]}행에 이미 있어 건너뜀')
            continue
        seen[code] = r
        t = costs_of(row, COST_IDX)
        mrow = main_by_code.get(code)
        if code not in own or mrow is None:
            # 담당자가 바뀌었거나 메인에서 없어진 코드 — 원가가 이미 맞춰졌으면 지움
            mc = costs_of(mrow, [main_idx[c] for c in COST_FIELDS]) if mrow is not None else None
            if not any(t) or t == last.get(code) or (mc is not None and t == mc):
                plan['deletes'].append(r)
                if mrow is None:              # 다른 사람 탭으로 옮겨 간 코드면 그쪽 기록은 그대로
                    plan['last'].append((code, None))
            else:
                plan['notes'].append(f'{tab} {r}행 {code} — 이 담당자 공사가 아닌데 메인에 안 넘어간 원가 {t}가 있어 남겨 둠')
            continue
        if code in pending:
            continue                        # 메인 쓰기 대기 중 — 쓰인 뒤 다음 회차에 판단
        m = costs_of(mrow, [main_idx[c] for c in COST_FIELDS])
        lst = last.get(code)
        act = decide(t, m, lst)
        if act == 'same':
            if lst != t:
                plan['last'].append((code, t))
        elif act in ('to_main', 'conflict_tab'):
            plan['pushes'].append((code, t))
            plan['last'].append((code, t))
            if act == 'conflict_tab':
                plan['notes'].append(f'{tab} {code} — 탭({t})과 PM({m})이 둘 다 바뀜 → 탭 값으로')
        else:   # to_tab
            for i, ci in enumerate(COST_IDX):
                plan['cells'].append((r, ci, m[i]))
            plan['last'].append((code, m))
        # 보기 칸
        want = main_row_values(mrow, main_idx)
        for ci, (h, src) in enumerate(LAYOUT):
            if src in (COST, CALC) or ci == 0:
                continue
            cur = row[ci] if ci < len(row) else ''
            if not same_value(cur, want[ci]):
                plan['cells'].append((r, ci, want[ci]))
        # 순익·마진율 수식
        fr = tab_formulas[r - 2] if r - 2 < len(tab_formulas) else []
        for ci, fn in ((PROFIT_IDX, profit_formula), (MARGIN_IDX, margin_formula)):
            j = ci - PROFIT_IDX
            if (fr[j] if j < len(fr) else '') != fn(r):
                plan['formulas'].append((r, ci, fn(r)))
    for code in own_codes:
        if code not in seen and code in main_by_code:
            vals = main_row_values(main_by_code[code], main_idx)
            plan['appends'].append(vals)
            plan['last'].append((code, [vals[i] for i in COST_IDX]))
    return plan


# ── 시트 읽기 ─────────────────────────────────────────────

def _svc():
    from dashboard.services.project_service import get_sheets_manager
    return get_sheets_manager().service


def read_main(svc, sid: str, main_name: str) -> tuple:
    """메인 → (머리글 index, 코드 → 행 값, 코드 → 행 번호, 담당자 → [코드](메인 순서))."""
    vals = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{main_name}'!A1:AZ",
        valueRenderOption='UNFORMATTED_VALUE').execute().get('values', [])
    hdr = [str(h).strip() for h in (vals[0] if vals else [])]
    idx = {h: i for i, h in enumerate(hdr) if h}
    need = ['프로젝트 코드', '담당자'] + list(COST_FIELDS) + [s for _, s in LAYOUT if s not in (COST, CALC)]
    missing = [h for h in need if h not in idx]
    if missing:
        raise RuntimeError(f'메인 머리글 없음: {missing}')
    ci, oi = idx['프로젝트 코드'], idx['담당자']
    by_code, rownum, owners, dup = {}, {}, {}, set()
    for r, row in enumerate(vals[1:], start=2):
        code = str(row[ci]).strip() if ci < len(row) and row[ci] is not None else ''
        if not code:
            continue
        if code in by_code:
            dup.add(code)
            continue
        by_code[code], rownum[code] = row, r
        owner = str(row[oi]).strip() if oi < len(row) and row[oi] is not None else ''
        owners.setdefault(owner, []).append(code)
    for code in dup:      # 코드가 두 줄이면 어느 쪽인지 몰라 건드리지 않음 (정합성 점검이 따로 알림)
        by_code.pop(code, None)
        rownum.pop(code, None)
    for owner in owners:
        owners[owner] = [c for c in owners[owner] if c not in dup]
    return idx, by_code, rownum, owners


def _redis():
    from dashboard.utils.redis_client import get_redis_client
    return get_redis_client().redis


def _load_hash(key: str) -> dict:
    out = {}
    for k, v in (_redis().hgetall(key) or {}).items():
        try:
            out[k] = json.loads(v)
        except Exception:
            pass
    return out


def _notify_admin(key: str, text: str) -> None:
    try:
        from dashboard.services.sync_scheduler import _notify_admin as _n
        _n(key, text)
    except Exception as exc:
        logger.warning(f'[COST_TAB] 관리자 알림 실패: {exc}')


_noted: set = set()


# ── 동기화 한 회차 ─────────────────────────────────────────

def run_sync(dry_run: bool = False) -> dict:
    """예약 작업 본체 (5분마다). dry_run 이면 계획만 돌려줌."""
    from dashboard.utils.google_sheets import main_sheet_name
    svc = _svc()
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    main_name = main_sheet_name()
    main_idx, by_code, _rownum, owners = read_main(svc, sid, main_name)
    tabs = tab_names()
    got = svc.spreadsheets().values().batchGet(
        spreadsheetId=sid, ranges=[f"'{t}'!A1:{TAB_LAST_COL}" for t in tabs],
        valueRenderOption='UNFORMATTED_VALUE').execute().get('valueRanges', [])
    gotf = svc.spreadsheets().values().batchGet(
        spreadsheetId=sid, ranges=[f"'{t}'!AF2:AG" for t in tabs],
        valueRenderOption='FORMULA').execute().get('valueRanges', [])
    last, pending = _load_hash(LAST_KEY), _load_hash(PENDING_KEY)

    plans = {}
    for t, vr, fr in zip(tabs, got, gotf):
        vals = vr.get('values', [])
        hdr = [str(h).strip() for h in (vals[0] if vals else [])][:WIDTH]
        if hdr != HEADERS:
            if t not in _noted:
                _noted.add(t)
                logger.info(f'[COST_TAB] {t} 탭 머리글이 새 형식이 아님 — 건너뜀 (재구성 전)')
            continue
        plans[t] = (vals, plan_tab(t, vals, fr.get('values', []), by_code, owners.get(t, []),
                                   main_idx, last, pending))

    pushes = [p for _, (_, pl) in plans.items() for p in pl['pushes']]
    summary = {t: {k: len(v) for k, v in pl.items()} for t, (_, pl) in plans.items()}
    if dry_run:
        return {'summary': summary, 'plans': {t: pl for t, (_, pl) in plans.items()}}
    if len(pushes) > max_push():
        _notify_admin('cost_tab_mass_push',
                      f':warning: 개인 원가 탭 → PM 으로 원가 {len(pushes)}건이 한 번에 바뀌려 해서 멈췄습니다 '
                      f'(한도 {max_push()}건). 탭에서 원가 칸을 실수로 지웠거나 붙여 넣었는지 확인해 주세요. '
                      f'예: {", ".join(c for c, _ in pushes[:5])}')
        logger.warning(f'[COST_TAB] 원가 {len(pushes)}건 한꺼번에 → 중단 (한도 {max_push()})')
        return {'summary': summary, 'blocked': len(pushes)}

    rc = _redis()
    for t, (vals, pl) in plans.items():
        try:
            _apply_tab(svc, sid, t, vals, pl)
        except _TabMoved as exc:
            logger.info(f'[COST_TAB] {t} 탭 줄이 쓰는 사이 바뀜 — 다음 회차에 다시: {exc}')
            continue
        for code, c in pl['last']:
            if code in [p[0] for p in pl['pushes']]:
                continue          # 메인에 쓰인 뒤 핸들러가 기록
            if c is None:
                rc.hdel(LAST_KEY, code)
            else:
                rc.hset(LAST_KEY, code, json.dumps(c))
        for n in pl['notes']:
            if n not in _noted:
                _noted.add(n)
                logger.warning(f'[COST_TAB] {n}')
        if pl['pushes']:
            from dashboard.services.sheet_write_queue import enqueue
            items = [{'code': c, 'costs': v, 'tab': t} for c, v in pl['pushes']]
            for it in items:
                rc.hset(PENDING_KEY, it['code'], json.dumps(it['costs']))
            rc.expire(PENDING_KEY, PENDING_TTL)
            enqueue(QUEUE_OP, {'sheet_id': sid, 'sheet_name': main_name, 'items': items},
                    meta={'user_email': f'개인 원가 탭({t})'})
            logger.info(f'[COST_TAB] {t} → PM 원가 {len(items)}건 대기열: '
                        + ', '.join(f"{i['code']} {i['costs']}" for i in items[:10]))
    return {'summary': summary}


class _TabMoved(Exception):
    pass


def _apply_tab(svc, sid: str, tab: str, vals: list, pl: dict) -> None:
    if not (pl['cells'] or pl['formulas'] or pl['appends'] or pl['deletes']):
        return
    # 읽은 뒤 SB 가 줄을 넣거나 지웠으면 행 번호가 틀어짐 → 코드 열을 다시 읽어 같을 때만 씀
    now = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{tab}'!A1:A",
                                          valueRenderOption='UNFORMATTED_VALUE').execute().get('values', [])
    a = lambda rows: [str(r[0]).strip() if r and r[0] is not None else '' for r in rows]
    before = a(vals)
    after = a(now)
    while before and not before[-1]:
        before.pop()
    while after and not after[-1]:
        after.pop()
    if before != after:
        raise _TabMoved(f'{len(before)}→{len(after)}줄')

    nrows = len(vals)            # 머리글 포함 마지막으로 값이 있는 줄
    data, fdata, freqs = [], [], []
    for r, ci, v in pl['cells']:
        data.append({'range': f"'{tab}'!{col_letter(ci)}{r}", 'values': [[v]]})
    for r, ci, f in pl['formulas']:
        fdata.append({'range': f"'{tab}'!{col_letter(ci)}{r}", 'values': [[f]]})
    if pl['appends']:
        start = nrows + 1
        end = start + len(pl['appends']) - 1
        meta = svc.spreadsheets().get(spreadsheetId=sid, ranges=[f"'{tab}'"],
                                      fields='sheets(properties(sheetId,gridProperties(rowCount)))').execute()
        p = meta['sheets'][0]['properties']
        if p['gridProperties']['rowCount'] < end + 20:
            svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': [{'appendDimension': {
                'sheetId': p['sheetId'], 'dimension': 'ROWS',
                'length': end + 100 - p['gridProperties']['rowCount']}}]}).execute()
        # 새 줄도 정렬·숫자 서식·체크박스·드롭다운
        freqs = data_format_requests(p['sheetId'], start - 1, end) + \
            dropdown_requests(svc, sid, p['sheetId'], start - 1, end)
        data.append({'range': f"'{tab}'!A{start}:{TAB_LAST_COL}{end}", 'values': pl['appends']})
        for k in range(len(pl['appends'])):
            r = start + k
            fdata.append({'range': f"'{tab}'!AF{r}:AG{r}", 'values': [[profit_formula(r), margin_formula(r)]]})
    if data:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=sid, body={'valueInputOption': 'RAW', 'data': data}).execute()
    if fdata:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=sid, body={'valueInputOption': 'USER_ENTERED', 'data': fdata}).execute()
    if freqs:
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': freqs}).execute()
    if pl['deletes']:
        meta = svc.spreadsheets().get(spreadsheetId=sid, ranges=[f"'{tab}'"],
                                      fields='sheets(properties(sheetId))').execute()
        gid = meta['sheets'][0]['properties']['sheetId']
        reqs = [{'deleteDimension': {'range': {'sheetId': gid, 'dimension': 'ROWS',
                                               'startIndex': r - 1, 'endIndex': r}}}
                for r in sorted(pl['deletes'], reverse=True)]
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': reqs}).execute()
        # 지운 줄 아래 순익·마진율 수식은 Sheets 가 행 번호를 맞춰 줌 — 다음 회차 점검에서도 확인
    logger.info(f"[COST_TAB] {tab} 탭 갱신: 칸 {len(pl['cells'])} · 수식 {len(pl['formulas'])} · "
                f"새 줄 {len(pl['appends'])} · 지운 줄 {len(pl['deletes'])}")


# ── 메인 쓰기 (시트 쓰기 대기열 핸들러) ──────────────────────────

def _register_queue_handler():
    from dashboard.services.sheet_write_queue import register

    @register(QUEUE_OP)
    def _handle_cost_push(payload: dict) -> None:
        """탭 원가 → 메인 원가 4칸 + _version+1. 행은 쓰는 순간 코드로 다시 찾는다."""
        svc = _svc()
        sid, name = payload['sheet_id'], payload['sheet_name']
        hdr = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{name}'!1:1").execute()['values'][0]
        hdr = [str(h).strip() for h in hdr]
        L = {c: col_letter(hdr.index(c)) for c in COST_FIELDS}
        vcol = col_letter(hdr.index('_version'))
        code_col = col_letter(hdr.index('프로젝트 코드'))
        got = svc.spreadsheets().values().batchGet(
            spreadsheetId=sid, ranges=[f"'{name}'!{code_col}2:{code_col}", f"'{name}'!{vcol}2:{vcol}"],
            valueRenderOption='UNFORMATTED_VALUE').execute()['valueRanges']
        codes = [str(r[0]).strip() if r else '' for r in got[0].get('values', [])]
        vers = got[1].get('values', [])
        rows = {}
        for i, c in enumerate(codes):
            if c:
                rows.setdefault(c, []).append(i)
        data, done = [], []
        for it in payload['items']:
            hit = rows.get(it['code'], [])
            if len(hit) != 1:
                logger.warning(f"[COST_TAB] 메인에서 {it['code']} 줄을 {len(hit)}개 찾음 — 원가 안 씀")
                continue
            i = hit[0]
            r = i + 2
            try:
                ver = int(float(vers[i][0])) if i < len(vers) and vers[i] and str(vers[i][0]).strip() else 0
            except (TypeError, ValueError):
                ver = 0
            for c, v in zip(COST_FIELDS, it['costs']):
                data.append({'range': f"'{name}'!{L[c]}{r}", 'values': [[v]]})
            data.append({'range': f"'{name}'!{vcol}{r}", 'values': [[str(ver + 1)]]})
            done.append(it)
        if data:
            svc.spreadsheets().values().batchUpdate(
                spreadsheetId=sid, body={'valueInputOption': 'RAW', 'data': data}).execute()
        rc = _redis()
        for it in payload['items']:
            cur = rc.hget(PENDING_KEY, it['code'])
            if it in done:
                rc.hset(LAST_KEY, it['code'], json.dumps(it['costs']))
            if cur is not None and json.loads(cur) == it['costs']:
                rc.hdel(PENDING_KEY, it['code'])
        if done:
            from dashboard.services.project_service import invalidate_project_cache
            invalidate_project_cache(None, trigger_refresh=True)
            logger.info('[COST_TAB] 메인 원가 반영: ' + ', '.join(f"{i['code']} {i['costs']}" for i in done))


_register_queue_handler()


# ── 칸 서식 ─────────────────────────────────────────────

# 메인 시트와 같게: 주소·폴더 경로는 왼쪽, 금액(원가·순익 포함)·마진율은 오른쪽, 나머지는 가운데
LEFT_COLS = ('현장 주소', '견적서 및 계약서 폴더 경로')
MONEY_COLS = ('총액1', '총액2', '계약금', '중도금', '잔금', '미수금') + COST_FIELDS + ('순익',)
DATE_COLS = ('공사 시작', '공사 종료', '수금 날짜')
CHECK_COLS = ('부가세', '수금 확인')


def column_alignment(i: int) -> str:
    h = HEADERS[i]
    if h in LEFT_COLS:
        return 'LEFT'
    if h in MONEY_COLS or h == '마진율':
        return 'RIGHT'
    return 'CENTER'


def data_format_requests(gid: int, r0: int, r1: int, check_end: int | None = None) -> list:
    """데이터 줄(0부터 센 r0 이상 r1 미만) 서식 — 재구성 때와 새 줄을 붙일 때 같이 쓴다.

    체크박스는 check_end(기본 r1) 줄까지만 — 빈 줄에 빈 체크박스가 보이지 않게.
    """
    def rng(c0, c1, end=r1):
        return {'sheetId': gid, 'startColumnIndex': c0, 'endColumnIndex': c1,
                'startRowIndex': r0, 'endRowIndex': end}
    reqs = []
    for i in range(WIDTH):
        h = HEADERS[i]
        f = {'horizontalAlignment': column_alignment(i), 'verticalAlignment': 'MIDDLE'}
        fields = ['horizontalAlignment', 'verticalAlignment']
        if h in DATE_COLS:
            f['numberFormat'] = {'type': 'DATE', 'pattern': 'yyyy-mm-dd'}
        elif h in MONEY_COLS:
            f['numberFormat'] = {'type': 'NUMBER', 'pattern': '#,##0'}
        elif h == '마진율':
            f['numberFormat'] = {'type': 'PERCENT', 'pattern': '0.0%'}
        if 'numberFormat' in f:
            fields.append('numberFormat')
        if i in COST_IDX:
            f['backgroundColor'] = COST_BG
            fields.append('backgroundColor')
        reqs.append({'repeatCell': {'range': rng(i, i + 1), 'cell': {'userEnteredFormat': f},
                                    'fields': 'userEnteredFormat(' + ','.join(fields) + ')'}})
        if h in CHECK_COLS and (check_end is None or check_end > r0):
            reqs.append({'setDataValidation': {'range': rng(i, i + 1, r1 if check_end is None else check_end),
                                               'rule': {'condition': {'type': 'BOOLEAN'}}}})
    return reqs


# 옛 탭·메인과 같은 모양: 머리글 검은 바탕·흰 굵은 글씨·흰 줄, 원가 값 칸만 연노랑(#FFF2CC)
COST_BG = {'red': 1, 'green': 0.949, 'blue': 0.8}
_WHITE = {'red': 1, 'green': 1, 'blue': 1}


def header_format_requests(gid: int) -> list:
    line = {'style': 'SOLID', 'color': _WHITE}
    return [{'repeatCell': {
        'range': {'sheetId': gid, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': WIDTH},
        'cell': {'userEnteredFormat': {
            'backgroundColor': {'red': 0, 'green': 0, 'blue': 0},
            'textFormat': {'bold': True, 'foregroundColor': _WHITE},
            'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE',
            'borders': {'top': line, 'bottom': line, 'left': line, 'right': line}}},
        'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)'}}]


# 메인과 같은 드롭다운(색깔 칩)을 쓰는 칸 — 규칙·칩 색은 API 로 못 만들어서 메인 칸에서 그대로 복사
DROPDOWN_COLS = ('사업자', '담당자', '거래처', '사업자명', '공사 구분', '기계 분류', '브랜드', '도급 구분', '계산서')


def dropdown_requests(svc, sid: str, tab_gid: int, r0: int, r1: int) -> list:
    """메인 마지막 공사 줄의 드롭다운을 탭 r0~r1(0부터, r1 미포함) 줄에 붙여 넣는 요청."""
    if r1 <= r0:
        return []
    from dashboard.utils.google_sheets import main_sheet_name
    name = main_sheet_name()
    meta = svc.spreadsheets().get(spreadsheetId=sid, ranges=[f"'{name}'"],
                                  fields='sheets(properties(sheetId))').execute()
    mgid = meta['sheets'][0]['properties']['sheetId']
    got = svc.spreadsheets().values().batchGet(
        spreadsheetId=sid, ranges=[f"'{name}'!1:1", f"'{name}'!A1:A"]).execute()['valueRanges']
    hdr = [str(h).strip() for h in got[0].get('values', [[]])[0]]
    src = len(got[1].get('values', [])) - 1          # 0부터 센 마지막 공사 줄
    if src < 1:
        return []
    src_of = dict(LAYOUT)
    reqs = []
    for h in DROPDOWN_COLS:
        ci, mh = HEADERS.index(h), src_of[h]
        if mh not in hdr:
            continue
        mi = hdr.index(mh)
        reqs.append({'copyPaste': {
            'source': {'sheetId': mgid, 'startRowIndex': src, 'endRowIndex': src + 1,
                       'startColumnIndex': mi, 'endColumnIndex': mi + 1},
            'destination': {'sheetId': tab_gid, 'startRowIndex': r0, 'endRowIndex': r1,
                            'startColumnIndex': ci, 'endColumnIndex': ci + 1},
            'pasteType': 'PASTE_DATA_VALIDATION'}})
    return reqs


def apply_formats(tab: str) -> int:
    """이미 만든 탭의 데이터 줄 서식만 다시 입힌다(값은 안 건드림). 줄 수를 돌려준다."""
    svc = _svc()
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    meta = svc.spreadsheets().get(spreadsheetId=sid, ranges=[f"'{tab}'"],
                                  fields='sheets(properties(sheetId,gridProperties(rowCount)))').execute()
    p = meta['sheets'][0]['properties']
    rows = p['gridProperties']['rowCount']
    codes = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{tab}'!A1:A").execute().get('values', [])
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={
        'requests': data_format_requests(p['sheetId'], 1, rows, check_end=len(codes))
        + dropdown_requests(svc, sid, p['sheetId'], 1, len(codes))
        + header_format_requests(p['sheetId'])}).execute()
    return len(codes) - 1


# ── 처음 만들기·다시 만들기 ─────────────────────────────────

def rebuild_tab(tab: str, backup_dir: str, apply: bool = False) -> dict:
    """탭을 새 형식으로 다시 만든다. 원가는 메인 값, 메인이 0 이고 옛 탭에만 있으면 옛 탭 값(다음 동기화가 메인으로).

    옛 탭은 값·수식·서식 정보를 backup_dir 에 JSON 으로 남긴다. apply=False 면 비교 결과만.
    """
    from dashboard.utils.google_sheets import main_sheet_name
    svc = _svc()
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    main_idx, by_code, _rownum, owners = read_main(svc, sid, main_sheet_name())
    codes = owners.get(tab, [])
    old_v = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{tab}'!A1:{TAB_LAST_COL}",
                                            valueRenderOption='UNFORMATTED_VALUE').execute().get('values', [])
    old_hdr = [str(h).strip() for h in (old_v[0] if old_v else [])]
    if not old_hdr or old_hdr[0] != '프로젝트 No.' or [old_hdr[i] if i < len(old_hdr) else '' for i in COST_IDX] != list(COST_FIELDS):
        raise RuntimeError(f'{tab} 탭 원가 칸 위치가 예상과 다름: {old_hdr[:5]} … {[old_hdr[i] for i in COST_IDX if i < len(old_hdr)]}')
    old_costs = {}
    for row in old_v[1:]:
        code = str(row[0]).strip() if row and row[0] is not None else ''
        if code and code not in old_costs:
            old_costs[code] = costs_of(row, COST_IDX)
    rows, report = [], {'rows': len(codes), 'tab_only': [], 'differ': [], 'main_only': 0}
    for code in codes:
        vals = main_row_values(by_code[code], main_idx)
        m = [vals[i] for i in COST_IDX]
        o = old_costs.get(code, [0, 0, 0, 0])
        if any(m) and any(o) and o != m:
            report['differ'].append((code, o, m))
        elif any(m) and not any(o):
            report['main_only'] += 1
        elif any(o) and not any(m):
            report['tab_only'].append((code, o))
            for i, ci in enumerate(COST_IDX):
                vals[ci] = o[i]
        rows.append(vals)
    if not apply:
        return report

    meta = svc.spreadsheets().get(spreadsheetId=sid, ranges=[f"'{tab}'"],
                                  fields='sheets(properties,protectedRanges,conditionalFormats,basicFilter,filterViews)'
                                  ).execute()['sheets'][0]
    old_f = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{tab}'!A1:{TAB_LAST_COL}",
                                            valueRenderOption='FORMULA').execute().get('values', [])
    os.makedirs(backup_dir, exist_ok=True)
    bk = os.path.join(backup_dir, f"개인탭_재구성전_{tab}_{time.strftime('%Y%m%d_%H%M%S')}.json")
    with open(bk, 'w', encoding='utf-8') as f:
        json.dump({'meta': meta, 'values': old_v, 'formulas': old_f}, f, ensure_ascii=False)
    gid = meta['properties']['sheetId']
    n = len(rows) + 1
    grid_rows = n + 100

    reqs = [{'deleteProtectedRange': {'protectedRangeId': p['protectedRangeId']}}
            for p in meta.get('protectedRanges', [])]
    for _ in meta.get('conditionalFormats', []):
        reqs.append({'deleteConditionalFormatRule': {'sheetId': gid, 'index': 0}})
    if meta.get('basicFilter'):
        reqs.append({'clearBasicFilter': {'sheetId': gid}})
    reqs.append({'updateCells': {'range': {'sheetId': gid},
                                 'fields': 'userEnteredValue,userEnteredFormat,note,dataValidation'}})
    reqs.append({'updateSheetProperties': {'properties': {
        'sheetId': gid, 'gridProperties': {'rowCount': grid_rows, 'columnCount': WIDTH,
                                           'frozenRowCount': 1, 'frozenColumnCount': 1}},
        'fields': 'gridProperties(rowCount,columnCount,frozenRowCount,frozenColumnCount)'}})
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': reqs}).execute()

    data = [{'range': f"'{tab}'!A1:{TAB_LAST_COL}1", 'values': [HEADERS]}]
    for k in range(0, len(rows), 500):
        data.append({'range': f"'{tab}'!A{k + 2}:{TAB_LAST_COL}{k + 1 + len(rows[k:k + 500])}",
                     'values': rows[k:k + 500]})
    svc.spreadsheets().values().batchUpdate(spreadsheetId=sid, body={'valueInputOption': 'RAW', 'data': data}).execute()
    svc.spreadsheets().values().batchUpdate(spreadsheetId=sid, body={'valueInputOption': 'USER_ENTERED', 'data': [
        {'range': f"'{tab}'!AF2:AG{n}", 'values': [[profit_formula(r), margin_formula(r)] for r in range(2, n + 1)]}]}
    ).execute()

    def rng(c0, c1, r0=1, r1=None):
        d = {'sheetId': gid, 'startColumnIndex': c0, 'endColumnIndex': c1, 'startRowIndex': r0}
        if r1 is not None:
            d['endRowIndex'] = r1
        return d

    reqs = (data_format_requests(gid, 1, grid_rows, check_end=n) + dropdown_requests(svc, sid, gid, 1, n)
            + header_format_requests(gid))
    for ci in (PROFIT_IDX, MARGIN_IDX):     # 적자면 빨간 굵은 글씨 (옛 탭 규칙)
        reqs.append({'addConditionalFormatRule': {'index': 0, 'rule': {
            'ranges': [rng(ci, ci + 1)],
            'booleanRule': {'condition': {'type': 'NUMBER_LESS', 'values': [{'userEnteredValue': '0'}]},
                            'format': {'textFormat': {'bold': True, 'foregroundColor': {'red': 0.85}}}}}}})
    reqs.append({'setBasicFilter': {'filter': {'range': rng(0, WIDTH, 0)}}})
    note = '서버가 메인(공사 현황)에서 자동으로 채우는 칸입니다. 원가(제품대·도급비·자재비·기타비)만 입력하세요.'
    for c0, c1 in ((0, COST_IDX[0]), (PROFIT_IDX, WIDTH)):
        reqs.append({'addProtectedRange': {'protectedRange': {
            'range': rng(c0, c1, 0), 'description': note, 'warningOnly': True}}})
    reqs.append({'addProtectedRange': {'protectedRange': {
        'range': rng(0, WIDTH, 0, 1), 'description': '머리글 — 바꾸면 자동 동기화가 멈춥니다.', 'warningOnly': True}}})
    reqs.append({'updateCells': {'range': rng(COST_IDX[0], COST_IDX[0] + 1, 0, 1), 'rows': [{'values': [{
        'note': '노란 칸(제품대~기타비)만 입력하면 5분 안에 PM·메인 시트에 반영됩니다. '
                'PM 에서 고친 원가도 이 탭에 다시 채워집니다. 나머지 칸은 서버가 메인 값으로 채웁니다.'}]}],
        'fields': 'note'}})
    svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={'requests': reqs}).execute()

    rc = _redis()
    tab_only = {c for c, _ in report['tab_only']}
    for code, vals in zip(codes, rows):
        if code in tab_only:
            rc.hdel(LAST_KEY, code)        # 맞춘 기록 없음 → 다음 동기화가 탭 값을 메인으로
        else:
            rc.hset(LAST_KEY, code, json.dumps([vals[i] for i in COST_IDX]))
    report['backup'] = bk
    return report
