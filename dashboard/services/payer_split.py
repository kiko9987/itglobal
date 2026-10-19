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

JS 미러: dashboard/src/js/utils/payerSplit.js (규칙 바꾸면 같이 바꿀 것).
"""
import re
from typing import Dict, List, Optional

STAGES = ('계약금', '중도금', '잔금')
PAYER_FIELD = {s: f'{s} 부담자' for s in STAGES}

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
