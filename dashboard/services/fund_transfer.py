# -*- coding: utf-8 -*-
"""법인 간 이체(자금 이동) — 계좌·사업자 불일치 입금의 자금 이동 요청 추적·짝짓기·처리.

원본(대장) = 시트 U/V/W 셀 메모 (2026-10-04 사용자 결정: 메모가 원장, 서버는 찾기용 색인만).
  대기 = '⚠️ 계좌·사업자 불일치 … — 자금 이동 요청' 줄이 있는 입금 블록 중, 같은 셀 메모에서
         그 뒤로 같은 방향(출발→도착) 매출이동 줄('YYYY-MM-DD R>G …')이 아직 없는 것.
  완료 = 매출이동 줄이 붙음 — #입금_관리 [법인 간 이체로 처리] 버튼이 자동으로 붙인다.

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
_STAGES = ('계약금', '중도금', '잔금')
_COLS = ('U', 'V', 'W')
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
    for blk in _blocks(note):
        req_line = next((l.strip() for l in blk.splitlines() if _REQ_LINE_RE.match(l)), None)
        if req_line:
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
                })
            continue
        m = _TRANSFER_RE.search(blk)
        if m:
            d = (m.group(1).upper(), m.group(2).upper())
            for i, r in enumerate(open_reqs):
                if (r['from'], r['to']) == d:
                    open_reqs.pop(i)
                    break
    return open_reqs


def _scan_sheet() -> List[Dict]:
    from dashboard.services.payment_sync import _get_payment_service, _VALID_PROJECT_RE
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    svc = _get_payment_service()
    if not (svc and sid and sn):
        return []
    codes = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{sn}'!A2:A10000").execute().get('values', [])
    nd = svc.spreadsheets().get(
        spreadsheetId=sid, ranges=[f"'{sn}'!U2:W10000"],
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
                out += find_pending_in_note(note, code, i + 2, _STAGES[j], _COLS[j])
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
        date = f"{int(yr) if len(yr) == 4 else datetime.now().year:04d}-{mm:02d}-{dd:02d}"
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
    cell = f"{req['col']}{row}"
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
