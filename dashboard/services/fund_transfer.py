# -*- coding: utf-8 -*-
"""법인 간 이체(자금 이동) — 계좌·사업자 불일치 입금의 자금 이동 요청 추적·짝짓기·처리.

원본(대장) = 시트 U/V/W 셀 메모 (2026-10-04 사용자 결정: 메모가 원장, 서버는 찾기용 색인만).
  대기 = '⚠️ 계좌·사업자 불일치 … — 자금 이동 요청' 줄이 있는 입금 블록 중, 같은 셀 메모에서
         그 뒤로 같은 방향(출발→도착) 매출이동 줄('YYYY-MM-DD R>G …')이 아직 없는 것.
  완료 = 매출이동 줄이 붙음 — #입금_관리 [법인 간 이체로 처리] 버튼이 자동으로 붙인다.
  이동 없음 = 요청 줄 바로 아래 '⚠️ 자금 이동 없음: …' 줄 (2026-10-26 G3982: 그룹으로 발행·수금해
         이체가 필요 없었는데 SB 댓글은 기록에 안 남아 정산 알림에 계속 나왔음). 붙는 길 두 가지:
         ① 자동 — 그 단계 세금계산서의 공급자 법인 = 돈이 들어온 법인(close_by_invoice)
         ② 경영지원 — #입금_관리 카드 스레드에 댓글(문구 무관)을 달면 뜨는 [이동 없이 처리] 버튼

짝짓기 = 이체로 들어온 입금 문자 ↔ 대기 요청: 도착 사업자(계좌)·금액이 같고, 최초 입금자 이름의
  앞부분이 같음. 경영지원이 이체할 때 받는 분 표시에 최초 입금자명을 그대로 넣고, 은행은 이름을
  자르기도 한다(실데이터 G4125-YM: 하나 '대연이엔지주식회' → 기업 '대연이엔지').
  보조 증거 = 출발 계좌의 같은 금액·이름 출금 문자(sms_inbound.record_outflow 서버 보관).
  고객 환불 → 고객 재입금도 이름·금액이 같아 문자만으론 구분 불가 → 처리는 경영지원이 버튼으로 확정.
"""
import json
import os
import re
from datetime import datetime
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_REQ_LINE_RE = re.compile(r'^\s*⚠.*계좌·사업자 불일치.*자금 이동 요청')
_NO_MOVE_LINE_RE = re.compile(r'^\s*⚠.*자금 이동 없음')
_STAGES = ('계약금', '중도금', '잔금')
_BANK_CODE = {'기업': 'G', '하나': 'R', '농협': 'N'}
_CODE_BANK = {'G': '기업', 'R': '하나', 'N': '농협'}
ENTITY = {'G': '글로벌', 'R': '글로벌그룹', 'P': '플렌트'}
_CACHE_KEY = 'fund_move_pending'
_CACHE_TTL = 300        # 색인 5분 — 원본은 메모, 필요 시 force 재계산


def norm_name(s: str) -> str:
    """이름 비교용 정규화 — 법인 표기·공백·기호 제거."""
    s = re.sub(r'\(주\)|㈜|주식회사|주식회|\(유\)|유한회사|\(사\)', '', s or '')
    s = re.sub(r'[\s.,\-()·]', '', s)
    return s.lower()


def names_match(a: str, b: str) -> bool:
    """앞부분 일치(은행이 이름을 자르는 경우 대응). 2글자 미만은 판정 안 함."""
    na, nb = norm_name(a), norm_name(b)
    if len(na) < 2 or len(nb) < 2:
        return False
    return na.startswith(nb) or nb.startswith(na)


def _blocks(note: str) -> List[str]:
    return [b for b in re.split(r'\n\s*\n', note or '') if b.strip()]


def find_pending_in_note(note: str, code: str, row: int, stage: str, col: str) -> List[Dict]:
    """한 셀 메모에서 대기 중인 자금 이동 요청 추출.

    블록 순서대로 요청을 방향별 대기열에 넣고, 뒤에 나오는 같은 방향 매출이동 줄이 가장 이른
    요청부터 하나씩 완료 처리(한 셀에 잘못 들어온 입금이 여러 건이어도 짝이 맞게).
    """
    from dashboard.services.payment_sync import _TRANSFER_RE, _parse_memo_block
    to = (code or '')[:1].upper()
    open_reqs: List[Dict] = []
    moves: List[tuple] = []     # 요청 뒤에 따로 기록된 출금(-)·입금(+) 블록 (법인 코드, 금액)
    for blk in _blocks(note):
        req_line = next((l.strip() for l in blk.splitlines() if _REQ_LINE_RE.match(l)), None)
        if req_line:
            if any(_NO_MOVE_LINE_RE.match(l) for l in blk.splitlines()):
                continue        # 경영지원이 '자금 이동 없이 처리'로 닫은 요청
            p = _parse_memo_block(blk) or {}
            frm = (p.get('acct_code') or _BANK_CODE.get(p.get('bank') or '', '')).upper()
            if frm and frm != to:
                yr = str(p.get('date_year') or '')
                open_reqs.append({
                    'code': code, 'row': row, 'stage': stage, 'col': col,
                    'amount': abs(int(p.get('amount') or 0)),
                    'partner': (p.get('partner') or '').strip(),
                    'bank': p.get('bank') or _CODE_BANK.get(frm, ''),
                    'from': frm, 'to': to,
                    'date': f"{yr}/{p.get('date_md')}" if len(yr) == 4 else (p.get('date_md') or ''),
                    'req_line': req_line,
                    '_after': len(moves),
                })
            continue
        m = _TRANSFER_RE.search(blk)
        if m:
            d = (m.group(1).upper(), m.group(2).upper())
            for i, r in enumerate(open_reqs):
                if (r['from'], r['to']) == d:
                    open_reqs.pop(i)
                    break
            continue
        p = _parse_memo_block(blk) or {}
        ent = (p.get('acct_code') or _BANK_CODE.get(p.get('bank') or '', '')).upper()
        if ent and int(p.get('amount') or 0):
            moves.append((ent, int(p.get('amount') or 0)))
    # '법인 간 이체' 버튼 대신 이체 문자 두 장(보낸 법인 출금 + 받은 법인 입금, 같은 금액)을 각각 이 칸에
    #   지정한 경우도 이체 완료 (2026-10-28 G4166-YM: 하나 출금·기업 입금 15:26 기록됐는데 정산 알림에 계속 남음)
    used = set()
    still = []
    for r in open_reqs:
        later = [(i, e, a) for i, (e, a) in enumerate(moves) if i >= r['_after'] and i not in used]
        out = next((i for i, e, a in later if e == r['from'] and a == -r['amount']), None)
        inn = next((i for i, e, a in later if e == r['to'] and a == r['amount']), None)
        if out is not None and inn is not None:
            used.update((out, inn))
            continue
        still.append(r)
    for r in still:
        r.pop('_after', None)
    return still


def _scan_sheet() -> List[Dict]:
    from dashboard.services.payment_sync import _get_payment_service, _VALID_PROJECT_RE
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    svc = _get_payment_service()
    if not (svc and sid and sn):
        return []
    from dashboard.utils.google_sheets import main_field_to_letter
    _f2l = main_field_to_letter()                 # 열 글자는 헤더 지도에서 (2026-10-22, 예전 U/V/W 고정)
    _cols = tuple(_f2l[s] for s in _STAGES)
    codes = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{sn}'!A2:A10000").execute().get('values', [])
    nd = svc.spreadsheets().get(
        spreadsheetId=sid, ranges=[f"'{sn}'!{_cols[0]}2:{_cols[-1]}10000"],
        fields='sheets.data.rowData.values.note', includeGridData=True).execute()
    rows = (((nd.get('sheets') or [{}])[0].get('data') or [{}])[0].get('rowData') or [])
    out: List[Dict] = []
    for i, rd in enumerate(rows):
        code = str(codes[i][0]).strip() if i < len(codes) and codes[i] else ''
        if not _VALID_PROJECT_RE.match(code):
            continue
        for j, v in enumerate(rd.get('values', [])[:3]):
            note = v.get('note') or ''
            if '자금 이동 요청' in note:
                out += find_pending_in_note(note, code, i + 2, _STAGES[j], _cols[j])
    return out


def pending_requests(force: bool = False) -> List[Dict]:
    """대기 중인 자금 이동 요청 전체 (메모 스캔, Redis 색인 5분)."""
    rc = None
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        if not force:
            raw = rc.get(_CACHE_KEY)
            if raw:
                return json.loads(raw)
    except Exception:
        rc = None
    try:
        reqs = _scan_sheet()
    except Exception as exc:
        logger.warning(f'[FUND_MOVE] 대기 요청 스캔 실패: {exc}')
        return []
    if rc is not None:
        try:
            rc.set(_CACHE_KEY, json.dumps(reqs, ensure_ascii=False), ex=_CACHE_TTL)
        except Exception:
            pass
    return reqs


def invalidate():
    try:
        from dashboard.utils.redis_client import get_redis_client
        get_redis_client().redis.delete(_CACHE_KEY)
    except Exception:
        pass


def deposit_entity(text: str, preview: Optional[dict] = None) -> str:
    """입금 문자가 들어온 계좌의 사업자 코드 (G/R/N/'')."""
    try:
        from dashboard.services.itg_accounts import match_account
        a = match_account(text or '')
        if a:
            return a.code
    except Exception:
        pass
    return ((preview or {}).get('acct_code') or '').upper()


def match_deposit(text: str, preview: dict, reqs: Optional[List[Dict]] = None) -> List[Dict]:
    """이체로 들어온 입금 ↔ 대기 요청 후보 (도착 사업자·금액·이름 앞부분 일치)."""
    to = deposit_entity(text, preview)
    if to not in ('G', 'R'):
        return []
    amt = abs(int((preview or {}).get('amount') or 0))
    name = (preview or {}).get('partner') or ''
    reqs = pending_requests() if reqs is None else reqs
    return [r for r in reqs
            if r['to'] == to and r['amount'] == amt and names_match(r['partner'], name)]


def outflow_evidence(req: Dict, since_ts: int) -> Optional[Dict]:
    """출발 계좌의 같은 금액·이름 출금 문자(서버 보관) — 법인 간 이체 보조 증거."""
    try:
        from dashboard.blueprints.sms_inbound import recent_outflows
        for o in reversed(recent_outflows(since_ts)):
            if (o.get('acct_code') == req['from'] and int(o.get('amount') or 0) == req['amount']
                    and names_match(o.get('partner', ''), req['partner'])):
                return o
    except Exception:
        pass
    return None


def transfer_line(req: Dict, deposit_text: str, preview: dict, checker: str) -> str:
    """메모에 붙일 매출이동 줄 — 날짜는 이체 입금 문자 날짜(연도 없으면 올해)."""
    pv = preview or {}
    md, yr = (pv.get('date_md') or '').strip(), str(pv.get('date_year') or '').strip()
    try:
        mm, dd = (int(x) for x in md.split('/'))
        from dashboard.services.sms_intake import year_for_md   # 연도 없으면 올해, 미래면 작년 (2026-10-29)
        date = f"{int(yr) if len(yr) == 4 else year_for_md(mm, dd):04d}-{mm:02d}-{dd:02d}"
    except Exception:
        date = datetime.now().strftime('%Y-%m-%d')
    try:
        from dashboard.services.itg_accounts import match_account
        a = match_account(deposit_text or '')
        to_bank = a.bank if a else _CODE_BANK.get(req['to'], '')
    except Exception:
        to_bank = _CODE_BANK.get(req['to'], '')
    return (f"{date} {req['from']}>{req['to']} 매출이동 "
            f"({req.get('bank') or _CODE_BANK.get(req['from'], '')} → {to_bank} 법인 간 이체 · 처리 {checker})")


def insert_transfer_line(note: str, req_line: str, line: str) -> str:
    """요청 줄이 있는 입금 블록 끝에 빈 줄을 두고 매출이동 줄 삽입(파서가 별도 블록으로 읽어
    원 입금에 'R → G' 표시를 붙이고 연도도 유지). 요청 줄을 못 찾으면 메모 끝에 붙인다."""
    note = note or ''
    idx = note.find(req_line) if req_line else -1
    if idx < 0:
        return f"{note.rstrip()}\n\n{line}" if note.strip() else line
    m = re.compile(r'\n\s*\n').search(note, idx)
    end = m.start() if m else len(note.rstrip())
    return f"{note[:end].rstrip()}\n\n{line}{note[end:]}"


def no_move_line(reason: str, checker: str, when: Optional[datetime] = None) -> str:
    """'이동 없이 처리' 메모 줄 — '⚠'로 시작해 수금 파서가 통째로 건너뛰고, 'YYYY-MM-DD R>G' 꼴이
    아니라 매출이동으로 읽히지 않는다(댓글의 '>'는 '→'로)."""
    reason = re.sub(r'\s+', ' ', (reason or '').replace('>', '→')).strip()[:60]
    return f"⚠️ 자금 이동 없음: {reason} (확인 {checker} {(when or datetime.now()):%m/%d})"


def insert_no_move_line(note: str, req_line: str, line: str) -> Optional[str]:
    """요청 줄 바로 아래(같은 입금 블록)에 '이동 없음' 줄 삽입. 요청 줄을 못 찾으면 None —
    다른 블록에 붙으면 어느 요청을 닫는지 모호해지므로 쓰지 않는다."""
    note = note or ''
    idx = note.find(req_line) if req_line else -1
    if idx < 0:
        return None
    end = note.find('\n', idx)
    end = len(note) if end < 0 else end
    return f"{note[:end]}\n{line}{note[end:]}"


def apply_no_move(req: Dict, reason: str, checker: str, when: Optional[datetime] = None) -> str:
    """대기 요청을 '이동 없이 처리'로 닫음 — 원 입금 셀 메모의 요청 줄 아래에 기록(시트 금액은 그대로).
    Returns 기록한 줄. 호출부가 같은 셀 RMW 직렬화(_INTAKE_SHEET_LOCK)를 책임진다."""
    from dashboard.services.lead_service import get_sheets_manager
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    mgr = get_sheets_manager()
    row = mgr.find_row_by_project_code(sid, req['code'], f"{sn}!A:A")
    if not row:
        raise RuntimeError(f"{req['code']} 행을 시트에서 못 찾음")
    from dashboard.utils.google_sheets import main_field_to_letter
    cell = f"{main_field_to_letter().get(req.get('stage')) or req['col']}{row}"
    line = no_move_line(reason, checker, when)
    note = mgr.get_cell_note(sid, sn, cell) or ''
    new_note = insert_no_move_line(note, req.get('req_line', ''), line)
    if new_note is None:
        raise RuntimeError(f"{cell} 메모에서 자금 이동 요청 줄을 못 찾음")
    if not mgr.update_cell_note(sid, sn, cell, new_note):
        raise RuntimeError(f"{cell} 메모 기록 실패")
    invalidate()
    logger.info(f"[FUND_MOVE] 자금 이동 없이 처리: {req['code']} {req['stage']} {cell} — {line}")
    return line


_ISSUER_KEY = 'invoice_issuer'   # Redis hash '코드|단계' → 그 단계 세금계산서를 낸 법인(G/R)


def record_invoice_issuer(code: str, stages: str, entity: str) -> None:
    """발행한 세금계산서의 공급자 법인 기억 — 계산서가 입금보다 먼저 나간 경우(G3982: 10/22 발행, 10/23 입금)
    나중에 생긴 자금 이동 요청도 정산 알림 전에 대조해 닫기 위해. 수정발행이면 최신 값으로 덮음."""
    st = [s.strip() for s in str(stages or '').split(',') if s.strip()]
    if not (code and st and entity in ('G', 'R')):
        return
    try:
        from dashboard.utils.redis_client import get_redis_client
        get_redis_client().redis.hset(_ISSUER_KEY, mapping={f'{code}|{s}': entity for s in st})
    except Exception as exc:
        logger.warning(f'[FUND_MOVE] 계산서 발행 법인 기록 실패({code}): {exc}')


def invoice_issuer(code: str, stage: str) -> str:
    try:
        from dashboard.utils.redis_client import get_redis_client
        v = get_redis_client().redis.hget(_ISSUER_KEY, f'{code}|{stage}')
        return (v.decode() if isinstance(v, bytes) else v) or ''
    except Exception:
        return ''


def close_by_recorded_issuers(when: Optional[datetime] = None) -> List[Dict]:
    """대기 요청 중 그 단계 계산서를 이미 돈이 들어온 법인 이름으로 낸 것 → '이동 없음' 기록.
    정산 알림 직전에 부른다. 호출부가 같은 셀 RMW 직렬화(_INTAKE_SHEET_LOCK)를 책임진다."""
    done = []
    for r in pending_requests(force=True):
        sup = invoice_issuer(r['code'], r['stage'])
        if sup and sup == r['from']:
            apply_no_move(r, f"{ENTITY[sup]} 이름으로 계산서 발행 — 돈과 계산서 사업자가 같음", '자동', when)
            done.append(r)
    return done


def close_by_invoice(code: str, stages: str, supplier: str, when: Optional[datetime] = None) -> List[Dict]:
    """발행한 세금계산서의 공급자 법인(supplier G/R)이 돈이 들어온 법인과 같으면, 그 공사·그 단계의 자금 이동
    요청은 옮길 이유가 없다(돈과 계산서가 같은 사업자) → '이동 없음' 기록. 다른 단계 요청은 그대로.
    Returns 닫은 요청. 호출부가 같은 셀 RMW 직렬화(_INTAKE_SHEET_LOCK)를 책임진다."""
    st = {s.strip() for s in str(stages or '').split(',') if s.strip()}
    if not (code and st and supplier in ('G', 'R')):
        return []
    done = []
    for r in pending_requests(force=True):
        if r['code'] == code and r['stage'] in st and r['from'] == supplier:
            apply_no_move(r, f"{ENTITY[supplier]} 이름으로 계산서 발행 — 돈과 계산서 사업자가 같음", '자동', when)
            done.append(r)
    return done


def apply_transfer(req: Dict, deposit_text: str, preview: dict, checker: str) -> str:
    """원 입금 셀 메모에 매출이동 줄 기록(시트 금액은 그대로) → 색인 무효화 → 수금 카드 재렌더.
    Returns 기록한 줄. 호출부가 같은 셀 RMW 직렬화(_INTAKE_SHEET_LOCK)를 책임진다."""
    from dashboard.services.lead_service import get_sheets_manager
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    mgr = get_sheets_manager()
    row = mgr.find_row_by_project_code(sid, req['code'], f"{sn}!A:A")
    if not row:
        raise RuntimeError(f"{req['code']} 행을 시트에서 못 찾음")
    # 열은 단계로 다시 찾음 — 색인(5분 캐시)에 남은 옛 열 글자로 쓰지 않게 (2026-10-22)
    from dashboard.utils.google_sheets import main_field_to_letter
    cell = f"{main_field_to_letter().get(req.get('stage')) or req['col']}{row}"
    line = transfer_line(req, deposit_text, preview, checker)
    note = mgr.get_cell_note(sid, sn, cell) or ''
    if line in note:
        return line    # 재처리 멱등
    if not mgr.update_cell_note(sid, sn, cell, insert_transfer_line(note, req.get('req_line', ''), line)):
        raise RuntimeError(f"{cell} 메모 기록 실패")
    invalidate()
    try:
        from dashboard.services.payment_sync import rerender_stage_card
        rerender_stage_card(req['code'], req['stage'])
    except Exception as exc:
        logger.warning(f"[FUND_MOVE] 수금 카드 재렌더 실패(무시, 메모는 기록됨): {exc}")
    logger.info(f"[FUND_MOVE] 법인 간 이체 기록: {req['code']} {req['stage']} {cell} — {line}")
    return line
