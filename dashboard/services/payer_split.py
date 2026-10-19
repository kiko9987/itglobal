"""나눠 내는 공사(부담자 여러 명) — '계약금/중도금/잔금 부담자' 칸 해석 (2026-10-19).

계기: 오남프라자(G4144~G4160)는 한 호실을 임대인·임차인이 나눠 내고, 사람마다 사업자·계산서·
결제 방식이 다르다. 공사 한 줄의 칸 3개(계약금·중도금·잔금)를 '부담자 칸'으로 쓰고,
옛 '○○ 입금자명' 열(AL·AM·AN, 2025-03 이후 미사용)을 '○○ 부담자'로 바꿔 누가 그 칸을 내는지 적는다.

칸 형식 (구분 기호 '·', 순서 자유 — 첫 토큰만 구분):
    임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서
    임차인 · 기종원(위너스골프) · bell8517@nate.com · 18,471,200 · 계산서
  - 첫 토큰 = 구분(임대인/임차인/본사 …), 금액 = 부담액(부가세 포함 실제로 받을 돈)
  - 사업자번호·이메일·결제 방식(계산서/현금/카드)은 있으면 인식, 결제 방식 없으면 '계산서'
  - 구분 기호가 없거나 금액·이름이 없으면 부담자가 아님(None) → 옛 입금자명 값('25/02/11 고희원 R')은 무시
  하나라도 해석되면 그 공사는 '나눠 내는 공사'. 받은 돈 = 그 칸의 금액(입금 메모 합), 남은 돈 = 부담액 − 받은 돈.

입력: PM [부담자 설정](매니저 포함, invoice_mgmt /api/invoice/payers) — 사람은 고르기·숫자만, plan_payers 가
칸을 정하고 save_payers 가 이 칸과 **칸 메모 맨 윗줄('◆ 임대인 이희승')**을 함께 기록한다(시트 메모 = PM 메모).
입금 파서(payment_sync)는 '◆' 줄을 건너뛴다.

JS 미러: dashboard/src/js/utils/payerSplit.js (규칙 바꾸면 같이 바꿀 것).
"""
import re
from typing import Dict, List, Optional, Tuple

STAGES = ('계약금', '중도금', '잔금')
PAYER_FIELD = {s: f'{s} 부담자' for s in STAGES}
# 칸 메모 맨 윗줄 부담자 표시 — 시트 메모와 PM 메모가 같게(2026-10-19 사용자 요청).
#   payment_sync 입금 파서는 이 기호로 시작하는 줄을 건너뛴다(입금 아님).
PAYER_NOTE_PREFIX = '◆'
# 저장 시 칸 자동 배정 순서: 첫째 = 계약금 칸(임대인), 둘째 = 잔금 칸(임차인), 셋째 = 중도금 칸
ASSIGN_ORDER = ('계약금', '잔금', '중도금')
METHODS = ('계산서', '현금', '카드')

_SEP_RE = re.compile(r'\s*[·|｜]\s*')
_BIZNO_RE = re.compile(r'^(\d{3})-?(\d{2})-?(\d{5})$')
_AMT_RE = re.compile(r'^(?:₩\s*)?(\d[\d,]*)\s*원?$')
_EMAIL_RE = re.compile(r'^[^\s@]+@[^\s@]+\.[^\s@]+$')
_CASH_WORDS = {'현금', '계산서x', '계산서없음', '계산서미발행', 'n입금', 'x'}
_METHOD_WORDS = {'계산서': '계산서', '세금계산서': '계산서', '카드': '카드', '미정': '미정'}


def _num(v) -> int:
    try:
        s = re.sub(r'[^\d.\-]', '', str(v if v is not None else ''))
        return int(float(s)) if s not in ('', '-', '.') else 0
    except (ValueError, TypeError):
        return 0


def _method_of(tok: str) -> str:
    t = re.sub(r'\s+', '', tok).lower()
    if t in _CASH_WORDS:
        return '현금'
    return _METHOD_WORDS.get(t, '')


def parse_payer_cell(text) -> Optional[dict]:
    """부담자 칸 → {'role','name','biz_no','email','amount','method','note','raw'} 또는 None."""
    s = str(text if text is not None else '').strip()
    if not s or s == '-' or not _SEP_RE.search(s):
        return None
    toks = [t.strip() for t in _SEP_RE.split(s) if t.strip()]
    if len(toks) < 3:
        return None
    role = toks[0]
    if (len(role) > 12 or _AMT_RE.match(role) or _BIZNO_RE.match(role) or '@' in role
            or _method_of(role)):
        return None
    name = biz_no = email = method = ''
    amount = 0
    notes = []
    for t in toks[1:]:
        m_biz = _BIZNO_RE.match(t)
        m_amt = _AMT_RE.match(t)
        if not biz_no and m_biz:
            biz_no = f'{m_biz.group(1)}-{m_biz.group(2)}-{m_biz.group(3)}'
        elif not email and _EMAIL_RE.match(t):
            email = t
        elif not amount and m_amt and _num(m_amt.group(1)) >= 1000:
            amount = _num(m_amt.group(1))
        elif not method and _method_of(t):
            method = _method_of(t)
        elif not name:
            name = t
        else:
            notes.append(t)
    if not name or amount <= 0:
        return None
    return {'role': role, 'name': name, 'biz_no': biz_no, 'email': email, 'amount': amount,
            'method': method or '계산서', 'note': ' · '.join(notes), 'raw': s}


def stage_payers(rec: dict) -> Dict[str, dict]:
    """프로젝트 레코드 → {단계: 부담자}. 받은 돈(received)·남은 돈(remaining) 포함."""
    out = {}
    for s in STAGES:
        p = parse_payer_cell((rec or {}).get(PAYER_FIELD[s]))
        if p:
            got = _num((rec or {}).get(s))
            p = dict(p, stage=s, received=got, remaining=p['amount'] - got)
            out[s] = p
    return out


def is_split(rec: dict) -> bool:
    return bool(stage_payers(rec))


def stage_label(stage: str, payer: Optional[dict]) -> str:
    """'계약금' → '계약금(임대인)' (부담자 없으면 그대로)."""
    return f"{stage}({payer['role']})" if payer else stage


def payer_line(payer: dict) -> str:
    """'임대인 이희승 · 부담 6,887,100 · 남음 1,377,420' (카드·툴팁용 한 줄)."""
    rem = int(payer.get('remaining', payer['amount']))
    tail = '완납' if rem <= 0 else f'남음 {rem:,}'
    return f"{payer['role']} {payer['name']} · 부담 {payer['amount']:,} · {tail}"


def _norm(s) -> str:
    s = re.sub(r'\(주\)|\(유\)|㈜|주식회사|유한회사', '', str(s or ''))
    return re.sub(r'[\s()\[\]·.,/\-]', '', s)


def _name_keys(name: str) -> List[str]:
    """'기종원(위너스골프)' → ['기종원위너스골프', '기종원', '위너스골프'] (3자 이상만)."""
    keys = {_norm(name)}
    for part in re.split(r'[()\[\]/]', str(name or '')):
        keys.add(_norm(part))
    return [k for k in keys if len(k) >= 3]


def name_matches(partner: str, payer: dict) -> bool:
    """입금자명(은행 문자, 잘림·메모 붙음)에 부담자 이름이 들어 있나."""
    p = _norm(partner)
    if len(p) < 2:
        return False
    for k in _name_keys(payer.get('name', '')):
        if k in p or (len(p) >= 4 and p in k):
            return True
    return False


def payer_for_biz(payers: Dict[str, dict], biz: str = '', biz_no: str = '') -> Optional[dict]:
    """계산서 요청 사업자(이름·번호) → 부담자. 번호 일치 우선, 이름은 정규화 일치·포함(3자↑)."""
    digits = re.sub(r'\D', '', biz_no or '')
    if digits:
        for p in payers.values():
            if p.get('biz_no') and re.sub(r'\D', '', p['biz_no']) == digits:
                return p
    b = _norm(biz)
    if len(b) < 2:
        return None
    hits = [p for p in payers.values()
            if any(k == b or (len(b) >= 3 and (k in b or b in k)) for k in _name_keys(p.get('name', '')))]
    return hits[0] if len(hits) == 1 else None


def match_deposit_stage(payers: Dict[str, dict], partner: str, amount) -> str:
    """입금(입금자명·금액) → 이 공사의 어느 부담자 칸인지. 확실할 때만 단계, 아니면 ''.

    ① 이름이 맞고 남은 돈 안에 드는 부담자가 하나면 그 칸
    ② 여럿이면 남은 돈과 금액이 같은(마지막 입금) 부담자가 하나일 때만
    ③ 이름이 안 맞으면 남은 돈과 금액이 정확히 같은 부담자가 하나일 때만
    """
    amt = _num(amount)
    if amt <= 0 or not payers:
        return ''
    fits = [p for p in payers.values() if amt <= p.get('remaining', p['amount']) + 1000]
    named = [p for p in fits if name_matches(partner, p)]
    if len(named) == 1:
        return named[0]['stage']
    pool = named if named else list(payers.values())
    exact = [p for p in pool if abs(p.get('remaining', p['amount']) - amt) <= 1]
    return exact[0]['stage'] if len(exact) == 1 else ''


def guess_split_payer(partner: str, amount, recs: list) -> str:
    """미지정 입금 → 나눠 내는 공사의 '코드 구분' 추정(예 'G4157-JW 임대인'). 확실할 때만, 아니면 ''.

    이름이 맞고 남은 돈 안에 드는 부담자 후보가 공사 전체에서 하나면 그것, 여럿이면 남은 돈과
    금액이 같은(마지막 입금) 후보가 하나일 때만. (같은 임대인이 여러 호실을 가진 경우 대비)
    """
    amt = _num(amount)
    if amt <= 0 or not partner:
        return ''
    cands = []
    for r in recs or []:
        payers = stage_payers(r)
        if not payers:
            continue
        code = str(r.get('프로젝트 코드') or '').strip()
        for p in payers.values():
            if name_matches(partner, p) and amt <= p['remaining'] + 1000:
                cands.append((code, p))
    if len(cands) > 1:
        cands = [c for c in cands if abs(c[1]['remaining'] - amt) <= 1]
    if len(cands) == 1:
        code, p = cands[0]
        return f"{code} {p['role']}"
    return ''


def payer_note_line(payer: dict) -> str:
    """칸 메모 맨 윗줄 — '◆ 임대인 이희승'."""
    return f"{PAYER_NOTE_PREFIX} {payer['role']} {payer['name']}"


def apply_note_line(note: str, line: str) -> str:
    """메모의 '◆' 줄을 line 하나로 맞춘다(맨 위, 아래 빈 줄 1개). line='' 이면 '◆' 줄만 지운다.
    입금 기록 줄은 그대로 둔다."""
    body = [l for l in str(note or '').split('\n') if not l.strip().startswith(PAYER_NOTE_PREFIX)]
    while body and not body[0].strip():
        body.pop(0)
    text = '\n'.join(body).rstrip()
    if not line:
        return text
    return f"{line}\n\n{text}" if text else line


def format_payer_cell(p: dict) -> str:
    """부담자 → 칸 문자열 '임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서' (parse_payer_cell 역)."""
    parts = [str(p['role']).strip(), str(p['name']).strip()]
    m = _BIZNO_RE.match(re.sub(r'\s', '', str(p.get('biz_no') or '')))
    if m:
        parts.append(f'{m.group(1)}-{m.group(2)}-{m.group(3)}')
    email = str(p.get('email') or '').strip()
    if email and _EMAIL_RE.match(email):
        parts.append(email)
    parts.append(f"{_num(p.get('amount')):,}")
    parts.append(p.get('method') if p.get('method') in METHODS else '계산서')
    return ' · '.join(parts)


def _same_person(a: dict, b: dict) -> bool:
    da, db = re.sub(r'\D', '', a.get('biz_no') or ''), re.sub(r'\D', '', b.get('biz_no') or '')
    if da and db:
        return da == db
    return _norm(a.get('name')) == _norm(b.get('name')) and bool(_norm(a.get('name')))


def plan_payers(rows: list, rec: dict) -> Tuple[Dict[str, dict], List[str], List[str]]:
    """PM [부담자 설정] 입력 → (칸 배정 {단계: 부담자}, 오류 목록, 경고 목록). 순수 함수.

    rows: [{role, name, biz_no, email, amount, method, stage?}] — 빈 줄(이름·금액 모두 없음)은 무시, 최대 3명.
    칸: 요청한 stage(중복 불가) → 기존에 같은 사람(번호·이름)이 있던 칸 → 남은 칸(계약금→잔금→중도금).
    오류: 구분·이름·금액 누락, 4명 이상, 같은 칸 중복, 입금이 있는 칸에 부담자가 빠짐(그 돈 주인 미상).
    경고: 부담액 합계 ≠ 총액 2.  rows 가 비면 '나눠 내기 해제'(입금 있어도 허용 — 평범한 공사로 복귀).
    """
    errors, warns = [], []
    clean = []
    for i, r in enumerate(rows or [], start=1):
        r = r or {}
        name = str(r.get('name') or '').strip()
        amount = _num(r.get('amount'))
        role = str(r.get('role') or '').strip()
        if not name and not amount:
            continue
        if not role:
            errors.append(f"{i}번째 줄: 구분(임대인·임차인 등)을 정해 주세요.")
        elif len(role) > 12 or _SEP_RE.search(role):
            errors.append(f"{i}번째 줄: 구분은 12자 이내, '·' 없이 적어 주세요.")
        if not name:
            errors.append(f"{i}번째 줄: 사업자명을 입력해 주세요.")
        elif _SEP_RE.search(name):
            errors.append(f"{i}번째 줄: 사업자명에 '·' 기호는 쓸 수 없습니다.")
        if amount <= 0:
            errors.append(f"{i}번째 줄: 부담액을 입력해 주세요.")
        method = r.get('method') if r.get('method') in METHODS else '계산서'
        stage = r.get('stage') if r.get('stage') in STAGES else ''
        clean.append({'role': role, 'name': name, 'biz_no': str(r.get('biz_no') or '').strip(),
                      'email': str(r.get('email') or '').strip(), 'amount': amount,
                      'method': method, 'stage': stage})
    if len(clean) > 3:
        errors.append('부담자는 3명까지 정할 수 있습니다(칸이 계약금·중도금·잔금 3개).')
    if errors:
        return {}, errors, warns
    if not clean:
        return {}, [], []
    taken = {}
    for p in clean:
        if p['stage']:
            if p['stage'] in taken:
                errors.append(f"{p['stage']} 칸에 두 사람을 넣을 수 없습니다.")
            taken[p['stage']] = p
    existing = stage_payers(rec or {})
    for p in clean:
        if p['stage']:
            continue
        for s, old in existing.items():
            if s not in taken and _same_person(p, old):
                p['stage'] = s
                taken[s] = p
                break
    for p in clean:
        if p['stage']:
            continue
        s = next((x for x in ASSIGN_ORDER if x not in taken), '')
        p['stage'] = s
        taken[s] = p
    for s in STAGES:
        got = _num((rec or {}).get(s))
        if got > 0 and s not in taken:
            errors.append(f"{s} 칸에 이미 입금 {got:,}원이 있습니다 — 그 돈을 낸 부담자를 {s} 칸에 지정해 주세요.")
    total2 = _num((rec or {}).get('총액 2'))
    share = sum(p['amount'] for p in clean)
    if total2 > 0 and abs(total2 - share) > 1:
        warns.append(f"부담액 합계 {share:,}원이 공사 총액(총액 2) {total2:,}원과 {total2 - share:+,}원 다릅니다.")
    return ({p['stage']: p for p in clean} if not errors else {}), errors, warns


def save_payers(code: str, plan: Dict[str, dict]) -> dict:
    """칸 배정 결과를 시트에 기록 — '○○ 부담자' 칸(AL·AM·AN) + 칸 메모 맨 윗줄('◆ 구분 이름').

    빈 칸은 '-' 와 '◆' 줄 제거. 메모 쓰기는 입금 기록과 같은 락(slack_bot._INTAKE_SHEET_LOCK)으로
    묶어 동시 입금 기록과 메모가 엉키지 않게 한다. 입금 파서는 '◆' 줄을 무시하므로 카드 영향 없음.
    Returns: 바뀐 필드 {'계약금 부담자': …, '계약금_메모': …} (PM 화면·캐시 갱신용).
    """
    import os
    from dashboard.services.lead_service import get_sheets_manager
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    mgr = get_sheets_manager()
    row = mgr.find_row_by_project_code(sid, code, f"{sn}!A:A")
    if not row:
        raise ValueError(f'{code} 행을 시트에서 못 찾음')
    f2l = mgr.get_field_to_letter()
    try:
        from dashboard.blueprints.slack_bot import _INTAKE_SHEET_LOCK as _lock
    except Exception:          # 테스트·단독 실행
        import threading
        _lock = threading.Lock()
    changed = {}
    with _lock:
        for s in STAGES:
            p = plan.get(s)
            cell_val = format_payer_cell(p) if p else '-'
            pcol = f2l.get(PAYER_FIELD[s])
            if pcol:
                cur = str(mgr.get_cell_value(sid, sn, f'{pcol}{row}') or '').strip()
                # 부담자를 빼는 경우 '-' 로 지우되, 옛 '입금자명' 값(2025-03 이전, 부담자 형식 아님)은 보존
                keep_old = not p and (cur in ('', '-') or parse_payer_cell(cur) is None)
                if cur != cell_val and not keep_old:
                    if not mgr.update_cell_value(sid, sn, f'{pcol}{row}', cell_val):
                        raise RuntimeError(f'{PAYER_FIELD[s]} 기록 실패')
                    changed[PAYER_FIELD[s]] = cell_val
            ncol = f2l.get(s)
            if ncol:
                note = mgr.get_cell_note(sid, sn, f'{ncol}{row}') or ''
                new_note = apply_note_line(note, payer_note_line(p) if p else '')
                if new_note.strip() != note.strip():
                    if not mgr.update_cell_note(sid, sn, f'{ncol}{row}', new_note):
                        raise RuntimeError(f'{s} 메모 기록 실패')
                    changed[f'{s}_메모'] = new_note
    return changed


def split_problems(rec: dict) -> List[str]:
    """정합성 점검용 — 나눠 내는 공사의 문제 문구 목록(없으면 [])."""
    payers = stage_payers(rec)
    out = []
    for s in STAGES:
        raw = str((rec or {}).get(PAYER_FIELD[s]) or '').strip()
        if raw and raw != '-' and s not in payers and _SEP_RE.search(raw):
            out.append(f"{PAYER_FIELD[s]} 칸 형식 확인 필요: {raw[:40]!r}")
    if not payers:
        return out
    total2 = _num(rec.get('총액 2'))
    share = sum(p['amount'] for p in payers.values())
    if total2 > 0 and abs(total2 - share) > 1:
        out.append(f"부담액 합계 {share:,}원 ≠ 총액 2 {total2:,}원 (차이 {total2 - share:+,}원)")
    for s in STAGES:
        if s not in payers and _num(rec.get(s)) > 0:
            out.append(f"{s} 칸에 입금 {_num(rec.get(s)):,}원이 있는데 부담자가 없음")
    return out
