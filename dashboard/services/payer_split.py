"""나눠 내는 공사(부담자 여러 명) — '계약금/중도금/잔금 부담자' 칸 해석 (2026-10-19).

계기: 오남프라자(G4144~G4160)는 한 호실을 임대인·임차인이 나눠 내고, 사람마다 사업자·계산서·
결제 방식이 다르다. 공사 한 줄의 칸 3개(계약금·중도금·잔금)를 '부담자 칸'으로 쓰고,
옛 '○○ 입금자명' 열(AL·AM·AN, 2025-03 이후 미사용)을 '○○ 부담자'로 바꿔 누가 그 칸을 내는지 적는다.

칸 형식 (구분 기호 '·', 순서 자유 — 첫 토큰만 구분):
    임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서
    임차인 · 기종원(위너스골프) · bell8517@nate.com · 18,471,200 · 계산서
  - 첫 토큰 = 구분(임대인/임차인/본사 …), 금액 = 부담액(부가세 포함 실제로 받을 돈)
  - 사업자번호·이메일·결제 방식(계산서/현금/카드)은 있으면 인식, 결제 방식 없으면 '계산서'
  - 현금 뒤 '부가세 없음' = 부가세 없이 받는 사람(2026-10-21). PM 창은 부가세 별도로 보여 주고 입력받되
    (다른 화면과 같게) 칸에는 받을 돈으로 환산해 적는다 — 이 표시가 있으면 1.1 환산을 하지 않음.
  - 구분 기호가 없거나 금액·이름이 없으면 부담자가 아님(None) → 옛 입금자명 값('25/02/11 고희원 R')은 무시
  하나라도 해석되면 그 공사는 '나눠 내는 공사'. 받은 돈 = 그 칸의 금액(입금 메모 합), 남은 돈 = 부담액 − 받은 돈.

입력: PM [부담자 설정](매니저 포함, invoice_mgmt /api/invoice/payers) — 사람은 고르기·숫자만, plan_payers 가
칸을 정하고 save_payers 가 이 칸과 **칸 메모 맨 윗줄('◆ 임대인 이희승')**을 함께 기록한다(시트 메모 = PM 메모).
입금 파서(payment_sync)는 '◆' 줄을 건너뛴다.

2026-10-20 차수·범용화 (오남 입금 칸 어긋남 계기):
  - 칸 = 사람(고정). 사람마다 여러 번 나눠 내는 '차수'(계약금·중도금·잔금·일시불)는 칸이 아니라 그 입금 메모
    블록 안의 '◇ 차수 중도금' 줄로 남긴다(입금 파서는 '◇' 줄도 건너뜀).
  - 부담자는 임대인·임차인만이 아니다(한 법인의 두 사업자, 병원 + 원장 개인사업자 …) — 구분은 자유 문구.
  - 계산서 상태는 사람별: '-'(통합 발행 포함) 규칙·'수금완료 + 마지막 칸 발행 = 전체 완료' 규칙을 쓰지 않고,
    그 칸에 발행한 금액(계산서_메모 줄) ↔ 부담액·받은 돈으로 판단(payer_bill_state).

JS 미러: dashboard/src/js/utils/payerSplit.js (규칙 바꾸면 같이 바꿀 것).
"""
import re
from typing import Dict, List, Optional, Tuple

STAGES = ('계약금', '중도금', '잔금')
PAYER_FIELD = {s: f'{s} 부담자' for s in STAGES}
# 칸 메모 맨 윗줄 부담자 표시 — 시트 메모와 PM 메모가 같게(2026-10-19 사용자 요청).
#   payment_sync 입금 파서는 이 기호로 시작하는 줄을 건너뛴다(입금 아님).
PAYER_NOTE_PREFIX = '◆'
# 저장 시 칸 자동 배정 순서: 첫째 = 계약금 칸, 둘째 = 잔금 칸, 셋째 = 중도금 칸 (구분 이름과 무관, 줄 순서)
ASSIGN_ORDER = ('계약금', '잔금', '중도금')
METHODS = ('계산서', '현금', '카드')
# 차수 — 한 사람이 여러 번 나눠 낼 때 몇 번째 돈인지. 입금 메모 블록 안 '◇ 차수 중도금' 줄.
INSTALLMENTS = ('계약금', '중도금', '잔금', '일시불')
INSTALLMENT_PREFIX = '◇'
_INST_RE = re.compile(r'^\s*◇\s*차수\s*(계약금|중도금|잔금|일시불)')

_SEP_RE = re.compile(r'\s*[·|｜]\s*')
_BIZNO_RE = re.compile(r'^(\d{3})-?(\d{2})-?(\d{5})$')
_AMT_RE = re.compile(r'^(?:₩\s*)?(\d[\d,]*)\s*원?$')
_EMAIL_RE = re.compile(r'^[^\s@]+@[^\s@]+\.[^\s@]+$')
_CASH_WORDS = {'현금', '계산서x', '계산서없음', '계산서미발행', 'n입금', 'x'}
_METHOD_WORDS = {'계산서': '계산서', '세금계산서': '계산서', '카드': '카드', '미정': '미정'}
# 현금을 부가세 없이 받는 사람 표시(2026-10-21) — 칸 금액은 그대로 '실제로 받을 돈'이고, 이 표시가 있으면
#   부가세 별도 화면에서 금액을 1.1로 나누지 않는다(G4148 임대인 600만 현금 = 부가세 없음, G4151·G4152 현금은 부가세 포함).
VAT_FREE_TOKEN = '부가세 없음'
_VAT_FREE_WORDS = {'부가세없음', 'vat없음'}


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
    vat_free = False
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
        elif not vat_free and re.sub(r'\s+', '', t).lower() in _VAT_FREE_WORDS:
            vat_free = True
        elif not name:
            name = t
        else:
            notes.append(t)
    if not name or amount <= 0:
        return None
    method = method or '계산서'
    return {'role': role, 'name': name, 'biz_no': biz_no, 'email': email, 'amount': amount,
            'method': method, 'vat_free': vat_free and method == '현금',
            'note': ' · '.join(notes), 'raw': s}


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


def payer_title(payer: Optional[dict], stage: str = '') -> str:
    """카드·안내용 사람 이름 — '임차인 김지민(다정약국)' (부담자 없으면 단계 이름)."""
    if not payer:
        return stage
    return f"{payer['role']} {payer['name']}"


def installment_line(inst: str) -> str:
    """입금 메모 블록에 넣는 차수 줄 — '◇ 차수 중도금' (모르는 값이면 '')."""
    return f"{INSTALLMENT_PREFIX} 차수 {inst}" if inst in INSTALLMENTS else ''


def parse_installment(text) -> str:
    """메모 블록(여러 줄)에서 차수 — '중도금' (없으면 '')."""
    for ln in str(text or '').split('\n'):
        m = _INST_RE.match(ln)
        if m:
            return m.group(1)
    return ''


def resolve_deposit_column(payers: Dict[str, dict], stage: str, partner: str, amount,
                           installment: str = '') -> dict:
    """나눠 내는 공사 입금 → 기록할 칸·차수 결정 (지정 확인·기록 공용). 순수 함수.

    installment 가 있으면(새 지정 창: 사람 + 차수) stage = 사람 칸 그대로, 사람 칸이 아니면 오류.
    없으면(옛 지정: '수금 단계'를 들어온 순서로 고른 것) 입금자명·금액으로 사람을 찾아 그 칸에 기록하고,
    지정된 단계는 차수로 남긴다. 사람을 못 찾으면 지정 칸이 사람 칸일 때만 그대로(대리 입금 등), 아니면 오류.

    Returns {'stage': 기록 칸 or '', 'installment': 차수 or '', 'moved': 지정 칸과 다르게 기록하나,
             'error': 문구 or ''}
    """
    out = {'stage': '', 'installment': '', 'moved': False, 'error': ''}
    names = ', '.join(f"{payer_title(p)}({s} 칸)" for s, p in payers.items())
    if installment in INSTALLMENTS:
        if stage not in payers:
            out['error'] = f"{stage} 칸에는 부담자가 없습니다 — 누가 낸 돈인지 골라 주세요: {names}"
            return out
        out.update(stage=stage, installment=installment)
        return out
    hit = match_deposit_stage(payers, partner, amount)
    if hit:
        out.update(stage=hit, installment=stage if stage in INSTALLMENTS else '', moved=hit != stage)
        return out
    if stage in payers:
        out.update(stage=stage)
        return out
    out['error'] = (f"누가 낸 돈인지 알 수 없습니다(입금자 {partner or '미상'}) — [재지정]에서 부담자와 차수를 "
                    f"골라 주세요: {names}")
    return out


def _bill_tok(t) -> str:
    """계산서 칸 값 정규화 — slack_bot._bill_norm_token 과 같은 규칙('-'·빈칸 = '')."""
    s = str(t if t is not None else '').strip()
    if s in ('', '-'):
        return ''
    return {'카드결제': '카드', '일반': '발행', '혼합': '기타', '확인필요': '기타'}.get(s, s)


def payer_bill_state(payer: dict, token, issued_amt: int = 0, has_lines: bool = False) -> dict:
    """한 사람(칸)의 계산서 상태 — 그 칸에 발행한 금액 ↔ 부담액·받은 돈 (나눠 내는 공사 전용).

    issued_amt = 계산서_메모에서 그 칸(단계)으로 적힌 장 금액 합(VAT 포함), has_lines = 그런 줄이 있나.
    줄이 없는데 칸이 '발행'이면 옛·수기 기록으로 보고 부담액 전액 발행으로 친다.
    state:
      issued   — 부담액만큼 발행 (preissued=True 면 아직 다 안 받음 → 선발행·입금 대기 표시)
      partial  — 일부만 발행(to_issue 남음). need=True 면 받은 돈이 발행액보다 많고 그 사람 완납 → ⚠️
      needed   — 발행 없음 + 그 사람 완납 → ⚠️ 미발행
      pending  — 발행 없음 + 아직 다 안 받음 → 발행예정(조용)
      cash / card / other — 계산서 대상 아님(현금·카드·기타)
      none     — 받은 돈도 발행도 없음
    """
    amount = int(payer.get('amount') or 0)
    received = int(payer.get('received') or 0)
    remaining = amount - received
    tok = _bill_tok(token)
    base = {'state': 'none', 'issued': 0, 'to_issue': 0, 'preissued': False, 'need': False,
            'pending': 0}
    if tok == '카드':
        return dict(base, state='card')
    if tok == 'N입금':
        return dict(base, state='cash')
    if tok == '기타':
        return dict(base, state='other')
    iss = int(issued_amt or 0) if has_lines else (amount if tok == '발행' else 0)
    if iss <= 0 and payer.get('method') == '현금':
        return dict(base, state='cash')
    if iss <= 0 and payer.get('method') == '카드':
        return dict(base, state='card')
    if iss >= amount - 1 and amount > 0:
        return dict(base, state='issued', issued=iss, preissued=remaining > 1,
                    pending=max(0, iss - received))
    if iss > 0:
        need = received > iss + 1 and remaining <= 1
        return dict(base, state='partial', issued=iss, to_issue=amount - iss, need=need,
                    pending=max(0, iss - received))
    if received > 0:
        return dict(base, state='needed' if remaining <= 1 else 'pending')
    return base


def payer_bill_states(rec: dict, active_lines: list) -> Dict[str, dict]:
    """나눠 내는 공사 → {단계: payer_bill_state(...) + payer}. active_lines = 취소 안 된 계산서 장 줄
    [{'stage','amt'}] (slack_bot._active_invoice_lines 결과 — 호출부가 넘김, 이 모듈은 순수 유지)."""
    payers = stage_payers(rec)
    out = {}
    for s, p in payers.items():
        mine = [it for it in (active_lines or []) if it.get('stage') == s and int(it.get('amt') or 0) > 0]
        st = payer_bill_state(p, (rec or {}).get(f'{s} 계산서'),
                              sum(int(it['amt']) for it in mine), bool(mine))
        out[s] = dict(st, payer=p)
    return out


def split_y_summary(states: Dict[str, dict]) -> str:
    """나눠 내는 공사의 계산서 요약(Y, 드롭다운 '{칸} - {상태}' 형식 유지) — 사람별 상태로.

    손볼 사람(⚠️ 미발행·덜 발행) → '{칸} - 미발행' 우선, 다음 아직 발행 안 한 입금(발행예정) → '{칸} - 미발행',
    그 외는 마지막으로 움직인 칸(계약금 → 잔금 → 중도금 순)의 상태. 움직임 없으면 '미발행'.
    """
    order = [s for s in ASSIGN_ORDER if s in states]
    urgent = [s for s in order if states[s]['state'] == 'needed'
              or (states[s]['state'] == 'partial' and states[s]['need'])]
    if urgent:
        return f"{urgent[0]} - 미발행"
    def _uninvoiced_paid(st):   # 받은 돈 중 아직 계산서 안 끊은 몫이 있나
        got = int(st['payer'].get('received') or 0)
        return (st['state'] == 'pending' and got > 0) or (st['state'] == 'partial' and got > st['issued'] + 1)
    quiet = [s for s in order if _uninvoiced_paid(states[s])]
    if quiet:
        return f"{quiet[0]} - 미발행"
    # 현금·카드로 내기로 한 사람은 아직 한 푼도 안 냈으면 움직임 없음(10-21 305호: 입금 전인데 'N입금'으로 잡힘)
    active = [s for s in order if states[s]['state'] != 'none'
              and not (states[s]['state'] in ('cash', 'card') and int(states[s]['payer'].get('received') or 0) <= 0)]
    if not active:
        return '미발행'
    s = active[-1]
    label = {'issued': '발행완료', 'partial': '발행완료', 'cash': 'N입금', 'card': '카드결제'}.get(
        states[s]['state'], '기타')
    return f"{s} - {label}"


def payer_line(payer: dict) -> str:
    """'임대인 이희승 · 부담 6,887,100 · 남음 1,377,420' (카드·툴팁용 한 줄).
    현금을 부가세 없이 받는 사람은 '· 현금·부가세 없음'을 붙인다(2026-10-21, 입금 지정 모달·PM 툴팁)."""
    rem = int(payer.get('remaining', payer['amount']))
    tail = '완납' if rem <= 0 else f'남음 {rem:,}'
    vf = ' · 현금·부가세 없음' if payer.get('vat_free') else ''
    return f"{payer['role']} {payer['name']} · 부담 {payer['amount']:,}{vf} · {tail}"


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


_ROOM_RE = re.compile(r'(?<!\d)(\d{3,4})\s*호')


def guess_intake_target(partner: str, amount, recs: list) -> Optional[Tuple[str, str]]:
    """미지정 입금 → (공사 코드, 부담자 칸) — 입금 지정 창 미리 채움용(2026-10-21). 확실할 때만, 아니면 None.

    ① 입금자명에 호실('301호임대인')이 있으면: 현장 주소에 그 호실이 있는 나눠 내는 공사의 부담자 중
       입금자명에 든 구분(임대인 등) → 없으면 이름이 맞는 사람 → 없으면 전원에서
    ② 호실이 없으면: 이름이 맞는 부담자(guess_split_payer 와 같은 규칙)
    금액이 남은 돈 안에 드는 후보만. 여럿이면 남은 돈과 금액이 같은(마지막 입금) 후보가 하나일 때만.
    """
    amt = _num(amount)
    if amt <= 0 or not partner:
        return None
    pn = _norm(partner)
    rooms = set(_ROOM_RE.findall(str(partner)))
    cands = []
    for r in recs or []:
        payers = stage_payers(r)
        if not payers:
            continue
        code = str(r.get('프로젝트 코드') or '').strip()
        named = [p for p in payers.values() if name_matches(partner, p)]
        if rooms:
            addr = str(r.get('현장 주소') or '')
            if not any(re.search(rf'(?<!\d){rm}\s*호', addr) for rm in rooms):
                continue
            roled = [p for p in payers.values() if _norm(p['role']) and _norm(p['role']) in pn]
            pool = roled or named or list(payers.values())
        else:
            pool = named
        cands += [(code, p) for p in pool if amt <= p['remaining'] + 1000]
    if len(cands) > 1:
        cands = [c for c in cands if abs(c[1]['remaining'] - amt) <= 1]
    if len(cands) == 1:
        return cands[0][0], cands[0][1]['stage']
    return None


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
    method = p.get('method') if p.get('method') in METHODS else '계산서'
    parts.append(method)
    if p.get('vat_free') and method == '현금':
        parts.append(VAT_FREE_TOKEN)
    return ' · '.join(parts)


def vat_flag(rec) -> bool:
    """부가세 별도 공사인가 — '부가세' 칸(시트 체크박스). 칸이 없는 레코드는 총액 2 > 총액 1 로 판단.
    project_service._parse_vat_flag 와 같은 값을 참으로 본다."""
    rec = rec if rec is not None else {}
    v = rec.get('부가세')
    if v is None or (isinstance(v, float) and v != v) or (isinstance(v, str) and not v.strip()):
        t1, t2 = _num(rec.get('총액 1')), _num(rec.get('총액 2'))
        return t1 > 0 and t2 > t1
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v == 1
    return str(v).strip().lower() in ('포함', 'true', 'yes', '1', 'vat 포함', 'y')


def vat_ratio(rec: dict) -> float:
    """부담액(받을 돈) ↔ 부가세 별도 금액 환산 비율 — 부가세 별도 공사면 1.1, 아니면 1.
    (2026-10-21: 총액2/총액1 은 '부가세 없음' 사람이 있으면 총액 2 가 줄어 비율이 틀어지므로 10% 고정)"""
    return 1.1 if vat_flag(rec) else 1.0


def payer_supply(payer: dict, ratio: float) -> int:
    """부담액(받을 돈) → 부가세 별도 금액. 현금·부가세 없음은 받을 돈 그대로."""
    amt = int(payer.get('amount') or 0)
    return amt if payer.get('vat_free') or ratio <= 1 else int(round(amt / ratio))


def vat_free_amount(rec) -> int:
    """부가세 별도 공사에서 부가세 없이 받는 몫의 합('현금 · 부가세 없음' 부담자 금액, 2026-10-21).
    총액 2 는 이 몫엔 부가세를 붙이지 않는다 — 시트 계산식(total2_formula)·서버 계산(project_service) 공용."""
    if not vat_flag(rec):
        return 0
    return sum(int(p['amount']) for p in stage_payers(rec).values() if p.get('vat_free'))


def total2_formula(row: int, r_col: str, s_col: str, vat_free: int = 0) -> str:
    """시트 총액 2 칸 계산식. vat_free=0 이면 기존 표준식과 글자 하나 다르지 않게(projects.py 신규 행과 같음).
    vat_free>0: 총액 1 중 그 몫(MIN(총액1, 몫))은 그대로, 나머지에만 부가세(절사 + 끝자리 1/9원 보정)."""
    R, S = f'{r_col}{row}', f'{s_col}{row}'
    if not vat_free or vat_free <= 0:
        return (f'=IF({S}=TRUE, {R}+FLOOR({R}*0.1,1) + IF(MOD({R}+FLOOR({R}*0.1,1), 10)=1, -1, '
                f'IF(MOD({R}+FLOOR({R}*0.1,1), 10)=9, 1, 0)), {R})')
    vf = int(vat_free)
    x = f'MAX({R}-{vf},0)'
    b = f'{x}+FLOOR({x}*0.1,1)'
    return f'=IF({S}=TRUE, MIN({R},{vf})+{b} + IF(MOD({b}, 10)=1, -1, IF(MOD({b}, 10)=9, 1, 0)), {R})'


def share_check(payers: list, rec: dict) -> str:
    """부담액 합계 점검 문구(맞으면 '') — PM 저장 경고·정합성 점검 공용 (2026-10-21 부가세 별도 기준).

    부가세 별도 공사: ① 부가세 별도 금액 합 ↔ 총액 1  ② 받을 돈 합 ↔ 총액 2 (총액 2 만 따로 고쳐진 경우).
    ②는 '현금 · 부가세 없음' 사람이 있으면 건너뜀 — 총액 2 는 저장 때 시스템이 그 몫을 빼고 맞추므로(저장 전
    시트의 옛 총액 2 와 비교하면 엉뚱한 경고). 부가세 없는 공사나 총액 1 이 없으면: 받을 돈 합 ↔ 총액 2.
    """
    rec = rec or {}
    t1, t2 = _num(rec.get('총액 1')), _num(rec.get('총액 2'))
    ratio = vat_ratio(rec)
    gross = sum(int(p.get('amount') or 0) for p in payers)
    tol = max(1, len(payers))   # 사람마다 원 단위 반올림
    if ratio > 1 and t1 > 0:
        supply = sum(payer_supply(p, ratio) for p in payers)
        if abs(t1 - supply) > tol:
            return f"부담액 합계(VAT 별도) {supply:,}원이 공사 총액(총액1) {t1:,}원과 {t1 - supply:+,}원 다릅니다."
        if t2 > 0 and abs(t2 - gross) > tol and not any(p.get('vat_free') for p in payers):
            # 표기는 PM 화면과 같게 '총액2', 금액과는 괄호로 구분('총액 2 16,446,200원'이 2억으로 읽힘, 10-21)
            return f"부담액 합계(받을 돈) {gross:,}원이 총액2({t2:,}원)와 {t2 - gross:+,}원 다릅니다."
        return ''
    if t2 > 0 and abs(t2 - gross) > 1:
        return f"부담액 합계 {gross:,}원이 공사 총액 {t2:,}원과 {t2 - gross:+,}원 다릅니다."
    return ''


def _same_person(a: dict, b: dict) -> bool:
    da, db = re.sub(r'\D', '', a.get('biz_no') or ''), re.sub(r'\D', '', b.get('biz_no') or '')
    if da and db:
        return da == db
    return _norm(a.get('name')) == _norm(b.get('name')) and bool(_norm(a.get('name')))


def plan_payers(rows: list, rec: dict) -> Tuple[Dict[str, dict], List[str], List[str]]:
    """PM [부담자 설정] 입력 → (칸 배정 {단계: 부담자}, 오류 목록, 경고 목록). 순수 함수.

    rows: [{role, name, biz_no, email, amount, method, vat_free?, stage?}] — amount = 받을 돈(부가세 포함, 화면이
    부가세 별도 입력을 환산해 보냄). 빈 줄(이름·금액 모두 없음)은 무시, 최대 3명.
    칸: 요청한 stage(중복 불가) → 기존에 같은 사람(번호·이름)이 있던 칸 → 남은 칸(계약금→잔금→중도금).
    오류: 구분·이름·금액 누락, 4명 이상, 같은 칸 중복, 입금이 있는 칸에 부담자가 빠짐(그 돈 주인 미상).
    경고: share_check(부가세 별도 합 ≠ 총액 1 등).  rows 가 비면 '나눠 내기 해제'(입금 있어도 허용).
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
            errors.append(f"{i}번째 줄: 구분을 적어 주세요(예: 임대인·임차인·본점).")
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
                      'method': method, 'vat_free': bool(r.get('vat_free')) and method == '현금',
                      'stage': stage})
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
    msg = share_check(clean, rec)
    if msg:
        warns.append(msg)
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
        changed.update(sync_total2_formula(mgr, sid, sn, row, f2l, plan))
    return changed


def _norm_fx(s) -> str:
    return re.sub(r'\s+', '', str(s or '')).upper()


def sync_total2_formula(mgr, sid: str, sn: str, row: int, f2l: dict, plan: Dict[str, dict]) -> dict:
    """총액 2 칸 계산식을 부담자 설정에 맞춤(2026-10-21) — '현금 · 부가세 없음' 몫엔 부가세를 안 붙인다.

    부가세 없는 몫이 있으면 그 몫을 뺀 계산식, 없어졌으면 표준식으로 되돌림. 처음부터 표준식(또는 값)이고
    부가세 없는 몫도 없으면 칸을 건드리지 않음. 바뀌면 새 총액 2·미수금(서버 계산 규칙과 같음)을 돌려줘
    PM 캐시·화면에 바로 반영. Returns {'총액 2': float, '미수금': float} 또는 {}.
    """
    tcol, rcol, scol = f2l.get('총액 2'), f2l.get('총액 1'), f2l.get('부가세')
    if not (tcol and rcol and scol):
        return {}
    vf = sum(int(p.get('amount') or 0) for p in (plan or {}).values() if p and p.get('vat_free'))
    cur = str(mgr.get_cell_formula(sid, sn, f'{tcol}{row}') or '')
    if vf <= 0 and 'MAX(' not in cur.upper():
        return {}
    want = total2_formula(row, rcol, scol, vf)
    if _norm_fx(cur) == _norm_fx(want):
        return {}
    if not mgr.update_cell_value(sid, sn, f'{tcol}{row}', want):
        raise RuntimeError('총액 2 계산식 기록 실패')
    rm = mgr.get_row_map(sid, sn, row) or {}
    from dashboard.services.project_service import _calculate_total2
    t1 = _num(rm.get(rcol))
    flag = vat_flag({'부가세': rm.get(scol), '총액 1': t1})
    total2 = _calculate_total2(float(t1), flag, vf if flag else 0)
    paid = sum(_num(rm.get(f2l.get(s))) for s in STAGES if f2l.get(s))
    return {'총액 2': float(total2), '미수금': float(total2 - paid)}


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
    msg = share_check(list(payers.values()), rec)
    if msg:
        out.append(msg)
    for s in STAGES:
        if s not in payers and _num(rec.get(s)) > 0:
            out.append(f"{s} 칸에 입금 {_num(rec.get(s)):,}원이 있는데 부담자가 없음")
    for s, p in payers.items():
        if p['remaining'] < -1:
            out.append(f"{payer_title(p)}({s} 칸) 부담액 {p['amount']:,}원보다 {-p['remaining']:,}원 더 받음")
        memo_key = f'{s}_메모'
        if memo_key in (rec or {}):
            first = next((l.strip() for l in str(rec.get(memo_key) or '').split('\n') if l.strip()), '')
            if first != payer_note_line(p):
                out.append(f"{s} 칸 메모 맨 윗줄에 부담자 표시('{payer_note_line(p)}')가 없음")
    return out
