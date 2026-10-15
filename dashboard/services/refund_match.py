# -*- coding: utf-8 -*-
"""입금 반환(환불) 자동 짝짓기 — 서버 보관 출금 문자 ↔ 시트에 기록된 입금 (2026-10-08).

계기: R4163-TH 계약금 200,000원(10/03 하나·김보라)을 현금거래 전환으로 10/08 출금 반환했는데
출금 문자는 서버 보관만 되고(sms_inbound.record_outflow) 시트엔 기록 경로가 없어 미수금이 틀어짐.
같은 사례 R4011-TH(8월)는 경영지원이 수기로 계약금 메모에 출금 문자를 붙이고 값을 0으로 기록.

후보 조건 (전부 충족):
  ① 입금받았던 같은 계좌(사업자 코드, 없으면 은행)에서 출금
  ② 같은 금액   ③ 같은 이름(앞부분 일치 — 은행 문자는 이름이 잘림)
  ④ 입금 후 90일 이내
  ⑤ 금액 검증 — 그 프로젝트가 출금액만큼 더 받은 상태:
     입금 합계 − 총액2 = 출금액 (과입금 반환), 또는 총액2 = 0(공사 취소)이고 입금 합계 = 출금액
후보가 정확히 1건일 때만 #입금_관리 카드 → 경영지원 [↩️ 반환으로 기록] 이 최종 판단.
환불이 먼저 나가고 현금 수령 기록이 나중에 들어오는 순서도 있어(R4163) 주기 스캔으로 재평가.

받은 돈 돌려줌(kind='reroute', 2026-10-15 R4101-JSH): ①~④는 맞는데 ⑤가 아님 — 더 받은 돈이 없는
공사에서 받은 입금을 그대로 돌려준 출금(사업자 통장으로 받은 계약금을 돌려주고 현금으로 다시 받는
경우 등). 기록하면 미수금이 다시 생김. 과입금·취소 후보가 없을 때만, 1건일 때만 제안.
보관 출금 37건(10/6~) 백테스트: 이 규칙으로 새로 걸리는 다른 출금 0건.
"""
import json
import logging
import os
import re
import time
from datetime import date, datetime
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

MAX_DAYS = 90
OUTFLOW_LOOKBACK_SEC = 60 * 60 * 24 * 14     # 출금 문자 서버 보관 기간과 동일
_STAGES = ('계약금', '중도금', '잔금')
_COLS = ('U', 'V', 'W')
_HANDLED = 'refund_match:handled:'           # 제안·무시·기록된 출금 id (재제안 방지)
_CAND = 'refund_match:cand:'                 # 제안 카드 페이로드
_DONE = 'refund_match:done:'                 # 기록 완료 (멱등)
_KEY_TTL = 60 * 60 * 24 * 30
_SCAN_LOCK = 'refund_match:scan_lock'


def _num(v) -> int:
    try:
        x = float(str(v).replace(',', '').replace('₩', '').strip() or 0)
        return 0 if x != x else int(round(x))
    except (ValueError, TypeError):
        return 0


def _same_account(dep: Dict, out: Dict) -> bool:
    """사업자 코드(G/R/N) 우선, 둘 중 하나가 없으면 은행명으로. 둘 다 없으면 불일치(정밀도 우선)."""
    da, oa = (dep.get('acct_code') or '').strip(), (out.get('acct_code') or '').strip()
    if da and oa:
        return da == oa
    db, ob = (dep.get('bank') or '').strip(), (out.get('bank') or '').strip()
    return bool(db and ob and db == ob)


def _deposit_date(p: Dict, ref: date) -> Optional[date]:
    m = re.match(r'^\s*(\d{1,2})/(\d{1,2})\s*$', str(p.get('date_md') or ''))
    if not m:
        return None
    yr = _num(p.get('date_year')) or ref.year
    try:
        d = date(yr, int(m.group(1)), int(m.group(2)))
    except ValueError:
        return None
    if not p.get('date_year') and d > ref:      # 연도 없는 메모 — 출금일 이후면 작년
        try:
            d = date(ref.year - 1, d.month, d.day)
        except ValueError:
            return None
    return d


def amount_check(paid: int, total2: int, amount: int) -> str:
    """⑤ 금액 검증 → 사유 문자열('' = 불충족)."""
    if total2 > 0 and abs((paid - total2) - amount) <= 1:
        return f'받은 돈 {paid:,}원 − 총액 {total2:,}원 = {amount:,}원'
    if total2 == 0 and paid > 0 and abs(paid - amount) <= 1:
        return f'공사 취소(총액 0) · 받은 돈 {paid:,}원 전액'
    return ''


def find_candidates(outflow: Dict, rows: List[Dict], include_reroute: bool = False) -> List[Dict]:
    """출금 1건 → 조건 ①~⑤ 를 모두 만족하는 (프로젝트·단계·입금) 후보 목록.

    rows: [{'code', 'row', 'vals'(A..AD 값 리스트), 'notes'([U,V,W 메모])}]
    include_reroute: ⑤(과입금·취소)가 아니어도 ①~④가 맞으면 kind='reroute'(받은 돈 돌려줌) 후보로 포함.
    후보 kind: 'overpay'(⑤ 충족) | 'reroute'. unpaid_before/after = 기록 전후 미수금.
    """
    from dashboard.services.fund_transfer import names_match
    from dashboard.services.payment_sync import _parse_notes
    amt = abs(_num(outflow.get('amount')))
    name = (outflow.get('partner') or '').strip()
    if amt <= 0 or not name:
        return []
    o_date = datetime.fromtimestamp(int(outflow.get('ts') or time.time())).date()
    out: List[Dict] = []
    for r in rows:
        vals = list(r.get('vals') or []) + [''] * 30
        u, v, w, total2 = _num(vals[20]), _num(vals[21]), _num(vals[22]), _num(vals[19])
        paid = u + v + w
        reason = amount_check(paid, total2, amt)           # ⑤ 먼저 — 대부분 행을 싸게 걸러냄
        kind = 'overpay'
        if not reason:
            # 받은 돈 돌려줌 — 총액이 있고 더 받은 돈이 없는 공사만 (과입금인데 금액이 다른 건 제외)
            if not (include_reroute and total2 > 0 and paid <= total2 + 1 and paid >= amt):
                continue
            kind = 'reroute'
            reason = f'받은 돈 {paid:,}원 / 총액 {total2:,}원 — 더 받은 돈 없음'
        pays = _parse_notes((list(r.get('notes') or []) + ['', '', ''])[:3],
                            stage_vals={'계약금': u, '중도금': v, '잔금': w})
        for stage, col in zip(_STAGES, _COLS):
            stage_pays = [p for p in pays if p.get('stage') == stage]
            # 이미 이 반환이 기록된 단계(같은 금액·이름의 반환 블록)는 제외
            if any(p.get('is_refund') and abs(abs(_num(p.get('amount'))) - amt) <= 1
                   and names_match(p.get('partner', ''), name) for p in stage_pays):
                continue
            for p in stage_pays:
                if p.get('is_refund') or abs(_num(p.get('amount')) - amt) > 1:
                    continue                                            # ②
                if not _same_account(p, outflow):                       # ①
                    continue
                if not names_match(p.get('partner', ''), name):         # ③
                    continue
                d = _deposit_date(p, o_date)
                if not d or d > o_date or (o_date - d).days > MAX_DAYS:  # ④
                    continue
                if kind == 'reroute' and {'계약금': u, '중도금': v, '잔금': w}[stage] < amt:
                    continue            # 그 단계에 돌려줄 만큼 남아 있지 않음
                out.append({
                    'kind': kind,
                    'unpaid_before': max(total2 - paid, 0), 'unpaid_after': max(total2 - (paid - amt), 0),
                    'code': r['code'], 'row': r['row'], 'stage': stage, 'col': col,
                    'amount': amt, 'stage_value': {'계약금': u, '중도금': v, '잔금': w}[stage],
                    'paid': paid, 'total2': total2, 'reason': reason,
                    'dep_date': d.isoformat(), 'dep_partner': p.get('partner', ''),
                    'dep_bank': p.get('bank', ''), 'dep_acct': p.get('acct_code', ''),
                    # 단계 계산서 칸(Z/AA/AB) — 발행돼 있으면 반환 시 취소·수정발행 안내
                    'stage_invoice': str(vals[25 + _STAGES.index(stage)] or '').strip(),
                })
    return out


def load_rows() -> List[Dict]:
    """공사현황 전 행(A..AD 값 + U/V/W 메모) — 입금 폴러 전용 service(스레드별)."""
    from dashboard.services.payment_sync import _get_payment_service, _VALID_PROJECT_RE
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    svc = _get_payment_service()
    if not (svc and sid and sn):
        return []
    vals = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{sn}'!A2:AD10000",
        valueRenderOption='UNFORMATTED_VALUE').execute().get('values', [])
    nd = svc.spreadsheets().get(
        spreadsheetId=sid, ranges=[f"'{sn}'!U2:W10000"],
        fields='sheets.data.rowData.values.note', includeGridData=True).execute()
    note_rows = (((nd.get('sheets') or [{}])[0].get('data') or [{}])[0].get('rowData') or [])
    rows = []
    for i, v in enumerate(vals):
        code = str(v[0]).strip() if v else ''
        if not _VALID_PROJECT_RE.match(code):
            continue
        rd = note_rows[i] if i < len(note_rows) else {}
        notes = [(x.get('note') or '') for x in (rd.get('values') or [])[:3]]
        rows.append({'code': code, 'row': i + 2, 'vals': v, 'notes': notes})
    return rows


def load_row(code: str, row: int) -> List[Dict]:
    """후보 한 행만 다시 읽기(반환 기록 재검증용) — load_rows 형식 리스트, 코드가 안 맞으면 [].

    전체 시트(값+메모)를 읽는 load_rows 대신 그 행만 (약 2.7초 → 1초 미만, 2026-10-15)."""
    from dashboard.services.payment_sync import _get_payment_service, _fetch_row_notes
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    svc = _get_payment_service()
    if not (svc and sid and sn and row):
        return []
    vals = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{sn}'!A{row}:AD{row}",
        valueRenderOption='UNFORMATTED_VALUE').execute().get('values', [[]])[0]
    if not vals or str(vals[0]).strip() != code:
        return []                      # 행이 밀렸으면 호출부가 전체 조회로
    return [{'code': code, 'row': row, 'vals': vals,
             'notes': list(_fetch_row_notes(sid, sn, row) or ['', '', ''])[:3]}]


def build_processing_blocks(outflow: Dict, c: Dict) -> list:
    """[반환으로 기록] 누른 직후 — 버튼 없이 '⏳ 반환 기록 중' (재클릭 차단, 처리 중임을 바로 보여줌)."""
    from dashboard.services.sms_intake import INTAKE_SEP, quoted_body
    lines = ["⠀", f">⏳ *반환 기록 중…*  `{c['code']}`  ·  *{c['stage']}*  ·  *{c['amount']:,}원*",
             f">{INTAKE_SEP}", *quoted_body(outflow.get('text') or ''), f">{INTAKE_SEP}"]
    return [{"type": "section", "text": {"type": "mrkdwn", "text": "\n".join(lines)}}]


def _fmt_md(iso_or_ts) -> str:
    try:
        if isinstance(iso_or_ts, (int, float)):
            return datetime.fromtimestamp(int(iso_or_ts)).strftime('%m/%d %H:%M')
        return datetime.fromisoformat(str(iso_or_ts)).strftime('%m/%d')
    except Exception:
        return '-'


def build_candidate_blocks(oid: str, outflow: Dict, c: Dict) -> list:
    """#입금_관리 반환 후보 카드 (인입 카드와 같은 인용 레이아웃)."""
    from dashboard.services.sms_intake import INTAKE_SEP, quoted_body
    amt = c['amount']
    after = c['stage_value'] - amt
    reroute = c.get('kind') == 'reroute'
    if reroute:
        # 받은 돈 돌려줌 — 더 받은 돈이 아니라 받은 입금을 그대로 돌려준 출금 (2026-10-15 R4101-JSH)
        head = f">↩️ *반환 후보 — 받은 돈을 돌려준 출금*  `{c['code']}`  ·  *{c['stage']}*  ·  *{amt:,}원*"
        check = [f">받은 돈 : {c['paid']:,}원 / 총액 {c['total2']:,}원 (더 받은 돈 없음)",
                 f">기록하면 : {c['stage']} {c['stage_value']:,}원 → {after:,}원  ·  "
                 f"미수금 {c['unpaid_before']:,}원 → {c['unpaid_after']:,}원 (다시 받을 돈)"]
    else:
        head = f">↩️ *반환 후보 — 기록 확인 필요*  `{c['code']}`  ·  *{c['stage']}*  ·  *{amt:,}원*"
        check = [f">금액 검증 : {c['reason']}"]
    lines = [
        "⠀",
        head,
        f">{INTAKE_SEP}",
        f">입금 : {_fmt_md(c['dep_date'])}  {c['dep_bank']}  {amt:,}원  {c['dep_partner']}",
        f">출금 : {_fmt_md(outflow.get('ts'))}  {outflow.get('bank', '')}  {amt:,}원  "
        f"{outflow.get('partner', '')}",
        *check,
        f">{INTAKE_SEP}",
        *quoted_body(outflow.get('text') or ''),
        f">{INTAKE_SEP}",
    ]
    confirm_text = (f"`{c['code']}` {c['stage']} 메모에 출금 내역을 추가하고 "
                    f"{c['stage']} {c['stage_value']:,}원 → {after:,}원으로 기록할까요?")
    context = (f"경영지원 확인 후 기록 — {c['stage']} 메모에 출금 내역 추가 · "
               f"{c['stage']} {c['stage_value']:,}원 → {after:,}원"
               + (" · 계산서 칸 '-'" if after == 0 else ''))
    if reroute:
        confirm_text += f" 미수금이 {c['unpaid_after']:,}원이 됩니다."
        context = ("다시 받을 돈이면 [반환으로 기록] (예: 사업자 통장으로 받은 돈을 돌려주고 "
                   "현금으로 다시 받기로 한 경우), 이 공사와 상관없는 출금이면 [반환 아님]")
    return [
        {"type": "section", "text": {"type": "mrkdwn", "text": "\n".join(lines)}},
        {"type": "actions", "elements": [
            {"type": "button", "action_id": "payment_refund_record", "style": "primary",
             "text": {"type": "plain_text", "text": "↩️ 반환으로 기록"}, "value": oid,
             "confirm": {
                 "title": {"type": "plain_text", "text": "반환 기록"},
                 "text": {"type": "mrkdwn", "text": confirm_text},
                 "confirm": {"type": "plain_text", "text": "기록"},
                 "deny": {"type": "plain_text", "text": "취소"}}},
            {"type": "button", "action_id": "payment_refund_dismiss",
             "text": {"type": "plain_text", "text": "반환 아님"}, "value": oid},
        ]},
        {"type": "context", "elements": [{"type": "mrkdwn", "text": context}]},
    ]


def build_done_blocks(outflow: Dict, c: Dict, old: int, new: int, checker: str) -> list:
    """기록 완료 카드 — 버튼 제거, 처리자·값 변화 표시."""
    from dashboard.services.sms_intake import INTAKE_SEP, quoted_body
    lines = [
        "⠀",
        f">✅ *반환 기록 완료*  `{c['code']}`  ·  *{c['stage']}*  ·  *{c['amount']:,}원*",
        f">처리 : 확인 {checker}  ·  {c['stage']} {old:,}원 → {new:,}원"
        + (f"  ·  미수금 {c['unpaid_after']:,}원 (다시 받을 돈)" if c.get('kind') == 'reroute' else ''),
        f">{INTAKE_SEP}",
        f">입금 : {_fmt_md(c['dep_date'])}  {c['dep_bank']}  {c['amount']:,}원  {c['dep_partner']}",
        f">출금 : {_fmt_md(outflow.get('ts'))}  {outflow.get('bank', '')}  {c['amount']:,}원  "
        f"{outflow.get('partner', '')}",
        f">{INTAKE_SEP}",
        *quoted_body(outflow.get('text') or ''),
        f">{INTAKE_SEP}",
    ]
    return [{"type": "section", "text": {"type": "mrkdwn", "text": "\n".join(lines)}}]


def build_dismissed_blocks(outflow: Dict, c: Dict, checker: str) -> list:
    from dashboard.services.sms_intake import INTAKE_SEP, quoted_body
    lines = ["⠀", f">~반환 후보~ — *반환 아님 처리*  `{c['code']}`  ·  {c['stage']}  ·  {c['amount']:,}원",
             f">처리 : {checker}  (시트 기록 없음)", f">{INTAKE_SEP}",
             *quoted_body(outflow.get('text') or ''), f">{INTAKE_SEP}"]
    return [{"type": "section", "text": {"type": "mrkdwn", "text": "\n".join(lines)}}]


def _payment_client():
    tok = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    if not tok:
        return None
    from slack_sdk import WebClient
    return WebClient(token=tok)


def post_candidate(oid: str, outflow: Dict, c: Dict, slack=None) -> bool:
    from dashboard.utils.redis_client import get_redis_client
    rc = get_redis_client().redis
    ch = os.getenv('SLACK_PAYMENT_INTAKE_CHANNEL', '').strip()
    slack = slack or _payment_client()
    if not (slack and ch):
        return False
    if not rc.set(_HANDLED + oid, 'proposing', nx=True, ex=_KEY_TTL):
        return False                                     # 이미 제안·처리됨
    try:
        resp = slack.chat_postMessage(
            channel=ch, text=f"↩️ 반환 후보: {c['code']} {c['stage']} {c['amount']:,}원",
            blocks=build_candidate_blocks(oid, outflow, c), unfurl_links=False)
        ts = (resp or {}).get('ts', '')
        rc.set(_CAND + oid, json.dumps({'outflow': dict(outflow, id=oid), 'cand': c,
                                        'channel': ch, 'ts': ts}, ensure_ascii=False), ex=_KEY_TTL)
        rc.set(_HANDLED + oid, f'proposed:{ts}', ex=_KEY_TTL)
        if ts:
            try:
                slack.pins_add(channel=ch, timestamp=ts)   # 미처리 박제 (기록 시 해제)
            except Exception:
                pass
        logger.info(f"[REFUND_MATCH] 반환 후보 카드: {c['code']} {c['stage']} {c['amount']:,}원 "
                    f"(출금 {oid}, {c['reason']})")
        return True
    except Exception as exc:
        rc.delete(_HANDLED + oid)                         # 실패 → 다음 스캔 재시도
        logger.warning(f"[REFUND_MATCH] 후보 카드 발송 실패 ({oid}): {exc}")
        return False


def load_candidate(oid: str) -> Optional[Dict]:
    from dashboard.utils.redis_client import get_redis_client
    raw = get_redis_client().redis.get(_CAND + oid)
    return json.loads(raw) if raw else None


def scan_and_propose(slack=None) -> Dict:
    """서버 보관 출금(14일) 중 미처리분을 시트 입금과 짝지어 후보 카드 발송. 10분 주기 + 출금 도착 시."""
    from dashboard.blueprints.sms_inbound import recent_outflows
    from dashboard.utils.redis_client import get_redis_client
    if os.getenv('REFUND_MATCH_ENABLED', '1').strip() in ('0', 'false', 'off'):
        return {'ok': True, 'skipped': 'disabled'}
    rc = get_redis_client().redis
    if not rc.set(_SCAN_LOCK, '1', nx=True, ex=120):
        return {'ok': True, 'skipped': 'busy'}
    try:
        outs = [o for o in recent_outflows(int(time.time()) - OUTFLOW_LOOKBACK_SEC)
                if abs(_num(o.get('amount'))) > 0 and not rc.exists(_HANDLED + o['id'])]
        if not outs:
            return {'ok': True, 'outflows': 0, 'proposed': 0}
        rows = load_rows()
        proposed = 0
        for o in outs:
            allc = find_candidates(o, rows, include_reroute=True)
            # 과입금·취소 후보 우선, 없을 때만 '받은 돈 돌려줌' 후보
            cands = [x for x in allc if x['kind'] != 'reroute'] or allc
            if len(cands) == 1:
                proposed += int(post_candidate(o['id'], o, cands[0], slack))
            elif len(cands) > 1:
                logger.info(f"[REFUND_MATCH] 후보 {len(cands)}건 — 모호해서 제안 안 함 (출금 {o['id']}: "
                            f"{[(x['code'], x['stage'], x['kind']) for x in cands]})")
        return {'ok': True, 'outflows': len(outs), 'proposed': proposed}
    finally:
        rc.delete(_SCAN_LOCK)


def scan_async() -> None:
    """출금 도착 직후 백그라운드 스캔. 테스트 환경(FLASK_ENV=testing·pytest)에선 실행 안 함 —
    record_outflow 를 부르는 테스트가 실제 시트 조회·슬랙 카드 발송으로 이어지지 않게."""
    import threading
    if (os.getenv('FLASK_ENV') == 'testing' or 'PYTEST_CURRENT_TEST' in os.environ
            or os.getenv('REFUND_MATCH_ENABLED', '1').strip() in ('0', 'false', 'off')):
        return

    def _run():
        try:
            scan_and_propose()
        except Exception as exc:
            logger.warning(f"[REFUND_MATCH] 스캔 실패: {exc}")
    threading.Thread(target=_run, daemon=True, name='refund-match-scan').start()
