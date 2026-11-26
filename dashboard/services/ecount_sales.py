"""이카운트 매출전표 목록 — 오늘 이카운트에 넣을 매출전표를 경영지원(SB) DM 으로 보낸다 (보고서 모드, 2026-10-30).

경영지원(SB)은 공사가 확정되면 확정일에 매출전표 I 을 한 장, 금액·부가세가 바뀌면 바뀐 날 차액을
한 장 넣는다. 사람이 기억해서 넣다 보니 부가세 여부만 바뀐 건의 차액이 빠졌다(G4152·G4147·R4163·
G4208, 10-29 대조). 그래서 시트와 '이카운트에 들어 있다고 보는 금액'을 매일 비교해 목록으로 띄운다.

- 이카운트 API 는 조회가 안 되므로 '들어 있다고 보는 금액'은 우리 장부(users.db ecount_sales_entries)에 둔다.
  처음 값(기준선)은 scripts/ecount_sales_seed.py 로 한 번 넣는다.
- 넣을 전표 = 시트 목표 − 장부 합계. 목표 공급가 = 총액 1, 부가세 = 총액 2 − 총액 1(부가세 있음일 때).
  공사 취소는 매출 0원 반영(총액 1 = 0) 뒤에 '전액 마이너스'로 잡힌다.
- 경영지원이 이카운트에 손으로 넣고 [이카운트에 넣음] 을 누르면 장부에 적혀 다음 알림에서 빠진다.
- 이카운트에는 아무것도 보내지 않는다. 자동 입력은 설계 문서 3·4단계(API 키·승인 카드)에서.
"""
from __future__ import annotations

import json
import re
import sqlite3
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

ACTION_DONE = 'ecount_sales_done'
KIND_BASELINE = '기준선'
MAX_ITEMS = 40          # 슬랙 메시지 블록 50개 한도 — 헤더·꼬리 몫을 남긴다
MIN_BASELINE_ROWS = 1000  # 기준선이 없으면(장부가 비면) 공사 4천 건이 전부 '확정'으로 뜬다 → 발송 안 함

_DDL = '''
    CREATE TABLE IF NOT EXISTS ecount_sales_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_code TEXT NOT NULL,
        supply INTEGER NOT NULL,
        vat INTEGER NOT NULL,
        kind TEXT NOT NULL,
        entry_key TEXT UNIQUE,
        recorded_by TEXT,
        recorded_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        note TEXT
    )
'''
# 공사 금액 수정 요청의 사유 — 요청 카드(Redis pending)는 반영(✅) 때 지워져서, 이카운트 '수정 사유' 칸에 쓰려고 따로 남긴다(11-26)
_DDL_REASONS = '''
    CREATE TABLE IF NOT EXISTS ecount_change_reasons (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_code TEXT NOT NULL,
        reason TEXT NOT NULL,
        source TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )
'''
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


def _ensure_table(conn) -> None:
    global _table_ready
    if _table_ready:
        return
    conn.execute(_DDL)
    conn.execute('CREATE INDEX IF NOT EXISTS idx_ecount_sales_code ON ecount_sales_entries(project_code)')
    conn.execute(_DDL_REASONS)
    conn.execute('CREATE INDEX IF NOT EXISTS idx_ecount_reason_code ON ecount_change_reasons(project_code)')
    conn.commit()
    _table_ready = True


# ── 금액 ─────────────────────────────────────────────────────────────

def _num(v) -> int:
    from dashboard.services.payer_split import _num as _n
    return _n(v)


def target_of(rec: dict) -> Tuple[int, int]:
    """시트가 말하는 이카운트 매출 (공급가, 부가세).

    부가세는 시트 총액 2 − 총액 1 — 3,800,000원처럼 부가세 포함 금액으로 정한 공사는 공급가의
    10% 반올림과 1원 다르고(48건), 부가세 없이 받는 부담자 몫(G4148)도 총액 2 가 이미 뺀다.
    총액 2 가 아직 안 채워졌으면 10% 반올림."""
    from dashboard.services.payer_split import vat_flag
    supply = _num(rec.get('총액 1'))
    if not supply or not vat_flag(rec):
        return supply, 0
    t2 = _num(rec.get('총액 2'))
    return supply, (t2 - supply if t2 > supply else round(supply * 0.1))


def _classify(in_ledger: bool, sent: Tuple[int, int], target: Tuple[int, int]) -> str:
    if not in_ledger:
        return '공사 확정'
    if target == (0, 0):
        return '공사 취소'
    if sent == (0, 0):
        return '공사 재개'
    if sent[0] == target[0]:
        return '부가세 변경'
    return '금액 변경'


def entry_key(code: str, sent: Tuple[int, int], target: Tuple[int, int]) -> str:
    """같은 전표를 두 번 적지 않게 — 장부가 바뀌면(누가 먼저 눌렀으면) 키가 달라진다."""
    return f'{code}|{sent[0]},{sent[1]}>{target[0]},{target[1]}'


def compute_pending(recs: List[dict], ledger: Dict[str, Tuple[int, int]],
                    change_dates: Optional[Dict[str, str]] = None) -> List[dict]:
    """시트 레코드 + 장부 → 넣을 전표 목록 (확정일·변경일 순).

    장부에 없는 공사는 '공사 확정'(전액). 공사 확정 날짜가 없으면 매출 0 — 아직 확정 전이거나,
    취소돼 날짜가 지워진 것(perform_cancel 이 즉시 지우고 총액 1=0 은 경영지원 ✅ 뒤). 그래서
    장부에 금액이 있는데 날짜가 비면 '공사 취소'(전액 마이너스), 장부에 없으면 건너뛴다.
    """
    from dashboard.services.payer_split import is_split
    change_dates = change_dates or {}
    out = []
    for r in recs:
        code = str(r.get('프로젝트 코드') or '').strip()
        if not code:
            continue
        confirmed = str(r.get('공사 확정') or '').strip()[:10]
        in_ledger = code in ledger
        sent = ledger.get(code, (0, 0))
        if confirmed:
            target = target_of(r)
        elif in_ledger and sent != (0, 0):
            target = (0, 0)
        else:
            continue
        d_supply, d_vat = target[0] - sent[0], target[1] - sent[1]
        if not d_supply and not d_vat:
            continue
        # 1원 반올림 차이(공급가 −1·부가세 +1 같은)는 넣을 전표가 아님(11-26 G4160-JW, 부장님 '제외 필요 없는 건만')
        if abs(d_supply) <= 1 and abs(d_vat) <= 1:
            continue
        kind = _classify(in_ledger, sent, target)
        out.append({
            'code': code, 'kind': kind,
            'supply': d_supply, 'vat': d_vat,
            'sent': sent, 'target': target,
            'date': confirmed if kind == '공사 확정' else (change_dates.get(code) or ''),
            'partner': str(r.get('유입 구분') or '').strip(),
            'addr': str(r.get('현장 주소') or '').strip(),
            'note': str(r.get('수금 관련 특이사항') or '').strip(),
            'split': is_split(r),
            'key': entry_key(code, sent, target),
        })
    out.sort(key=lambda e: (e['date'] or '9999', e['code']))
    return out


# ── 장부 ─────────────────────────────────────────────────────────────

def ledger_totals() -> Dict[str, Tuple[int, int]]:
    """공사별 이카운트에 들어 있다고 보는 금액 합계."""
    with _db()._get_connection() as conn:
        _ensure_table(conn)
        rows = conn.execute(
            'SELECT project_code, SUM(supply), SUM(vat) FROM ecount_sales_entries GROUP BY project_code'
        ).fetchall()
    return {r[0]: (int(r[1] or 0), int(r[2] or 0)) for r in rows}


def record_entry(code: str, supply: int, vat: int, kind: str, key: str,
                 by: str, note: str = '') -> bool:
    """장부에 전표 한 장 적기. 같은 키가 이미 있으면 False (버튼 두 번 눌림)."""
    with _db()._get_connection() as conn:
        _ensure_table(conn)
        try:
            conn.execute(
                'INSERT INTO ecount_sales_entries (project_code, supply, vat, kind, entry_key, recorded_by, note) '
                'VALUES (?, ?, ?, ?, ?, ?, ?)',
                (code, int(supply), int(vat), kind, key, by, note))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False


def remember_reason(code: str, reason: str, source: str = '금액 수정 요청') -> None:
    """금액·부가세 수정 요청의 사유를 남긴다 — 이카운트 올리기 '수정 사유 : …' 에 씀. 실패해도 요청은 그대로."""
    code, reason = (code or '').strip(), re.sub(r'\s+', ' ', reason or '').strip()
    if not code or not reason:
        return
    try:
        with _db()._get_connection() as conn:
            _ensure_table(conn)
            conn.execute('INSERT INTO ecount_change_reasons (project_code, reason, source) VALUES (?, ?, ?)',
                         (code, reason[:200], source))
            conn.commit()
    except Exception as exc:
        logger.warning(f'[ECOUNT] 수정 사유 저장 실패({code}): {exc}')


def pending_reasons(codes: List[str]) -> Dict[str, str]:
    """공사별, 장부에 마지막으로 적힌 뒤 들어온 수정 요청 사유(여러 번이면 ' / '로 이어 붙임).
    장부보다 먼저 들어온 사유는 이미 이카운트에 넣은 변경의 것이라 쓰지 않는다."""
    if not codes:
        return {}
    out: Dict[str, List[str]] = {}
    with _db()._get_connection() as conn:
        _ensure_table(conn)
        q = ','.join('?' * len(codes))
        rows = conn.execute(
            f"SELECT r.project_code, r.reason FROM ecount_change_reasons r "
            f"WHERE r.project_code IN ({q}) AND r.created_at > COALESCE("
            f"(SELECT MAX(e.recorded_at) FROM ecount_sales_entries e WHERE e.project_code = r.project_code), '') "
            f"ORDER BY r.id", list(codes)).fetchall()
    for code, reason in rows:
        if reason not in out.setdefault(code, []):
            out[code].append(reason)
    return {k: ' / '.join(v) for k, v in out.items()}


def seed_baseline(rows: List[Tuple[str, int, int, str]], by: str) -> int:
    """기준선 넣기 — rows=[(코드, 공급가, 부가세, 메모)]. 이미 장부에 있는 공사는 건너뛴다."""
    have = set(ledger_totals())
    n = 0
    with _db()._get_connection() as conn:
        _ensure_table(conn)
        for code, supply, vat, note in rows:
            if code in have:
                continue
            conn.execute(
                'INSERT INTO ecount_sales_entries (project_code, supply, vat, kind, entry_key, recorded_by, note) '
                'VALUES (?, ?, ?, ?, ?, ?, ?)',
                (code, int(supply), int(vat), KIND_BASELINE, f'{code}|baseline', by, note))
            n += 1
        conn.commit()
    return n


def change_dates(codes: List[str]) -> Dict[str, str]:
    """공사별 마지막 금액·부가세 변경일(한국 날짜) — 감사 기록에서. 변경 전표의 날짜 칸 참고용."""
    if not codes:
        return {}
    out: Dict[str, str] = {}
    with _db()._get_connection() as conn:
        q = ','.join('?' * len(codes))
        rows = conn.execute(
            f"SELECT project_code, MAX(timestamp) FROM audit_logs "
            f"WHERE (field_name IN ('총액 1', '부가세') OR action IN ('CANCEL_PROJECT', 'RESUME_PROJECT')) "
            f"AND project_code IN ({q}) GROUP BY project_code",
            list(codes)).fetchall()
    for code, ts in rows:
        try:
            out[code] = (datetime.fromisoformat(str(ts)[:19]) + timedelta(hours=9)).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return out


# ── 슬랙 ─────────────────────────────────────────────────────────────

def _won(n: int) -> str:
    return f'{n:,}원'


def _md(d: str) -> str:
    return f'{int(d[5:7])}/{int(d[8:10])}' if len(d) >= 10 else ''


def item_text(e: dict) -> str:
    when = f' · {_md(e["date"])}' if e.get('date') else ''
    memo = e['addr'] if e['kind'] == '공사 확정' else e['kind']
    if e['kind'] == '공사 취소' and e.get('note'):
        # 취소 시 특이사항 = '공사 취소' + 사유 — SB 적요 '취소 사유 : 고객변심'
        reason = e['note'].replace('공사 취소', '', 1).strip(' -:·\n')
        memo = f'취소 사유 : {reason[:40]}' if reason else memo
    lines = [
        f"*{e['code']}* · {e['kind']}{when}",
        f"공급가: {_won(e['supply'])} · 부가세: {_won(e['vat'])} · 합계: {_won(e['supply'] + e['vat'])}",
        f"거래처: {e['partner'] or '-'} · 적요: {memo or '-'}",
    ]
    if e['kind'] != '공사 확정':
        lines.append(f"(이카운트 {_won(sum(e['sent']))} → 시트 {_won(sum(e['target']))})")
    if e.get('split'):
        lines.append(':warning: 나눠 내는 공사 — 부담자별로 나눠 넣을지 확인')
    return '\n'.join(lines)


def amount_req_text(amount_reqs: List[dict]) -> str:
    """금액 수정·매출 0원 요청 중 아직 PM 에 반영 안 된 것 — DM 맨 위 (pin_remind.collect_amount_requests)."""
    lines = [f':pencil2: *금액 반영 대기 — PM 에 반영한 뒤 요청 카드에 ✅ ({len(amount_reqs)}건)*']
    for e in amount_reqs[:MAX_ITEMS]:
        link = f'  |  <{e["permalink"]}|바로가기>' if e.get('permalink') else ''
        lines.append(f'• {e["summary"]}{link}')
    return '\n'.join(lines)[:2900]   # 섹션 글자 한도 3000


def build_blocks(items: List[dict], amount_reqs: Optional[List[dict]] = None) -> List[dict]:
    blocks: List[dict] = []
    if amount_reqs:
        blocks += [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': amount_req_text(amount_reqs)}}]
        if items:
            blocks.append({'type': 'divider'})
    if not items:
        return blocks
    shown = items[:MAX_ITEMS - len(blocks)]
    blocks += [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': (
            f':ledger: *이카운트 매출전표 — 넣을 것 {len(items)}건*\n'
            '매출전표 I 에 넣은 뒤 [넣음]을 눌러 주세요. 누른 건은 다음 알림에서 빠집니다.')}},
        {'type': 'divider'},
    ]
    for e in shown:
        value = json.dumps({'c': e['code'], 's': e['supply'], 'v': e['vat'],
                            'k': e['key'], 'kd': e['kind']}, ensure_ascii=False)
        blocks.append({
            'type': 'section', 'block_id': f"ecs:{e['key']}"[:255],
            'text': {'type': 'mrkdwn', 'text': item_text(e)},
            'accessory': {'type': 'button', 'action_id': ACTION_DONE, 'style': 'primary',
                          'text': {'type': 'plain_text', 'text': '넣음'}, 'value': value},
        })
    if len(items) > len(shown):
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
                       f'나머지 {len(items) - len(shown)}건은 위 건을 처리하면 다음 알림에 나옵니다.'}]})
    return blocks


def mark_done_blocks(blocks: List[dict], key: str, who: str, now: datetime) -> List[dict]:
    """[넣음] 누른 줄 → 버튼 대신 처리 표시."""
    out = []
    for b in blocks or []:
        if b.get('block_id') == f'ecs:{key}'[:255]:
            b = {k: v for k, v in b.items() if k != 'accessory'}
            txt = b.get('text', {}).get('text', '')
            b['text'] = {'type': 'mrkdwn',
                         'text': f"~{txt.splitlines()[0]}~\n:white_check_mark: {who} 넣음 {now:%m/%d %H:%M}"}
        out.append(b)
    return out


def collect_items(force_refresh: bool = True) -> Optional[List[dict]]:
    """넣을 전표 목록. 기준선이 없으면 None."""
    from dashboard.services.project_service import get_project_records
    ledger = ledger_totals()
    if len(ledger) < MIN_BASELINE_ROWS:
        logger.warning(f'[ECOUNT] 장부 기준선 없음({len(ledger)}건) — scripts/ecount_sales_seed.py 먼저')
        return None
    recs = get_project_records(force_refresh=force_refresh)
    if not recs:
        return None
    items = compute_pending(recs, ledger)
    changed = [e['code'] for e in items if e['kind'] != '공사 확정']
    dates = change_dates(changed)
    try:
        reasons = pending_reasons(changed)
    except Exception as exc:          # 사유가 없어도 목록은 띄운다(적요는 금액으로 채움)
        logger.warning(f'[ECOUNT] 수정 사유 읽기 실패: {exc}')
        reasons = {}
    for e in items:
        if e['kind'] != '공사 확정':
            e['date'] = dates.get(e['code'], '')
            e['reason'] = reasons.get(e['code'], '')
    items.sort(key=lambda e: (e['date'] or '9999', e['code']))
    return items


def send_todo_dm() -> dict:
    """경영지원(SB) 개인 DM — 금액 반영 대기 + 이카운트 매출전표 목록 (2026-10-30).

    스케줄러 진입점, 미처리 정산 요약(#영업_관리)과 같은 10:00·15:30. 한 사람의 할 일이라 채널이 아니라
    본인 DM (채널에 올리면 감시처럼 느껴진다, 사용자). 영업일만, 둘 다 0건이면 안 보냄."""
    try:
        from dashboard.services.absent_remind import _is_business_day
        if not _is_business_day(date.today()):
            return {'ok': True, 'total': 0, 'reason': 'non_business_day'}
    except Exception:
        pass
    from dashboard.services.pin_remind import collect_amount_requests, open_todo_dm
    try:
        amount_reqs = collect_amount_requests()
    except Exception as exc:
        logger.warning(f'[ECOUNT] 금액 반영 대기 조회 실패(무시): {exc}')
        amount_reqs = []
    items = collect_items() or []   # 기준선이 없거나 시트 조회 실패면 매출전표 목록만 빠짐
    total = len(amount_reqs) + len(items)
    if not total:
        return {'ok': True, 'total': 0, 'reason': None}
    c, ch = open_todo_dm()
    if not c or not ch:
        return {'ok': False, 'total': total, 'reason': 'dm_open_failed'}
    try:
        r = c.chat_postMessage(
            channel=ch, text=f'경영지원 할 일 — 금액 반영 대기 {len(amount_reqs)}건, 이카운트 매출전표 {len(items)}건',
            blocks=build_blocks(items, amount_reqs), unfurl_links=False, unfurl_media=False)
        logger.info(f'[ECOUNT] 경영지원 DM 발송 (금액 {len(amount_reqs)}, 매출전표 {len(items)}) ts={r.get("ts")}')
        return {'ok': bool(r.get('ok')), 'total': total, 'ts': r.get('ts', '')}
    except Exception as exc:
        logger.error(f'[ECOUNT] 경영지원 DM 발송 실패: {exc}', exc_info=True)
        return {'ok': False, 'total': total, 'reason': str(exc)}


# ── 그룹웨어 '이카운트 올리기'(2026-11-26 부장님·SB) ─────────────────────────
# SB 답: '전일 공사현황의 프로젝트명·주소·금액·유입구분·부가세 여부를 이카운트 양식에 맞게 만들어 주면 바로 등록'.
# 이카운트 웹자료올리기는 엑셀에서 복사해 표에 붙여넣는 방식 — 칸 순서를 그대로 맞춘다(11-26 화면에서 확인).
#   프로젝트등록:   프로젝트코드 · 프로젝트명 · ㅇㅇ · 프로젝트그룹1 · 프로젝트그룹2 · 적요
#   매출전표 I:     일자 · 부가세유형 · 거래처코드 · 거래처명 · 공급가액 · 외화공급가액 · 환율 · 부가세 · 적요 ·
#                   매출계정코드 · 입금계좌코드 · 수수료 · 프로젝트코드
# SB 가 넣던 모양(회계거래조회 10/27~11/25 134건): 거래처 = 유입 구분, 적요 = 현장 주소(바뀐 금액은 '수정 사유 : …'),
#   프로젝트 = 공사 코드, 부가세 없는 공사는 부가세 칸을 비움. 매출계정 4049 제품매출 · 입금계좌 1089 외상매출금(기본값).
PROJECT_COLS = ('프로젝트코드', '프로젝트명', 'ㅇㅇ', '프로젝트그룹1', '프로젝트그룹2', '적요')
SALES_COLS = ('일자', '부가세유형', '거래처코드', '거래처명', '공급가액', '외화공급가액', '환율', '부가세', '적요',
              '매출계정코드', '입금계좌코드', '수수료', '프로젝트코드')
VAT_TYPE = '11'          # 세금계산서(이카운트 부가세유형 코드)
VAT_TYPE_NONE = '1Z'     # 기타매출 — 부가세 없는 공사(G4310-YG 전표에서 확인, 부가세 0)
SALES_ACCOUNT = '4049'   # 제품매출
AR_ACCOUNT = '1089'      # 외상매출금
DATE_FMT = '%Y%m%d'      # 웹자료올리기 일자 칸
# 유입 구분 → 이카운트 거래처코드(11-26 거래처등록에서 확인). 없는 유입 구분은 거래처명만 넣고 '확인 필요'
PARTNER_CODES = {
    '거래처': '99237', '홈페이지': '99655', '전화': '99656', '당근': '99644', '기타': '99657', '카카오톡': '99658',
    '채널톡': '99661', '큐플레이스': '99662', '메일': '99663', '소개': '99051', '온라인': '99015', '숨고': '99295',
}


def _auto_reason(e: dict) -> str:
    """수정 요청 사유가 없을 때 — 무엇이 바뀌었는지(SB 가 이카운트에서 다시 고치지 않게, 11-26 부장님)."""
    (s0, v0), (s1, v1) = e['sent'], e['target']
    parts = []
    if s0 != s1:
        parts.append(f"공급가액 {s0:,} → {s1:,} {'상향' if s1 > s0 else '하향'}")
    if v0 != v1 and not (s0 != s1 and v0 and v1):     # 공급가가 바뀌어 부가세가 따라 바뀐 건 공급가만 적음
        parts.append('부가세 별도 → 없음' if not v1 else '부가세 없음 → 별도' if not v0 else f'부가세 {v0:,} → {v1:,}')
    return ', '.join(parts)


def memo_of(e: dict) -> str:
    """적요 — 확정·재개는 현장 주소, 취소는 '취소 사유 : …', 금액·부가세 변경은 '수정 사유 : …'.
    수정 사유는 금액 수정 요청에 적힌 사유, 없으면 바뀐 금액(11-26 부장님 '수정 사유 : 만 있으면 이상')."""
    if e['kind'] in ('공사 확정', '공사 재개'):
        return e.get('addr') or ''
    if e['kind'] == '공사 취소':
        reason = (e.get('note') or '').replace('공사 취소', '', 1).strip(' -:·\n')
        m = re.match(r'^\(?\s*사유\s*[:：]\s*(.*?)\)?$', reason, re.S)   # 특이사항 '(사유: 고객변심)' → '고객변심'
        reason = (m.group(1) if m else reason).strip()
        return f'취소 사유 : {reason[:60]}' if reason else '취소 사유 : '
    reason = (e.get('reason') or '').strip() or _auto_reason(e)
    return f'수정 사유 : {reason[:80]}'.strip()


def _ymd(d: str) -> str:
    try:
        return datetime.strptime(d[:10], '%Y-%m-%d').strftime(DATE_FMT)
    except (TypeError, ValueError):
        return ''


def upload_rows(items: List[dict], today: Optional[date] = None) -> Dict[str, List[dict]]:
    """넣을 전표 목록 → 이카운트 웹자료올리기 줄. {'sales': [...], 'projects': [...]}.
    줄마다 'cells'(칸 순서 그대로, 복사용)와 화면용 값·'warn'(확인 필요 사유)을 담는다."""
    today = today or date.today()
    sales, projects = [], []
    for e in items:
        partner = e.get('partner') or ''
        pcode = PARTNER_CODES.get(partner, '')
        warn = []
        if not pcode:
            warn.append(f"거래처 없음('{partner or '빈칸'}')")
        if e.get('split'):
            warn.append('나눠 내는 공사')
        day = _ymd(e.get('date') or '') or today.strftime(DATE_FMT)
        if not e.get('date'):
            warn.append('바뀐 날 모름(오늘로 넣음)')
        vat = '' if not e['vat'] else str(e['vat'])
        vtype = VAT_TYPE if (e['target'][1] or e['vat']) else VAT_TYPE_NONE
        if e['code'].endswith('-SH'):
            warn.append('SH 정산 공사')
        # 어디서 어디로 — 이카운트에 있다고 보는 금액 → 공사 현황(새 공사는 없음). 확인할 점 칸에 회색 글로
        trace = '' if e['kind'] == '공사 확정' else (
            f"이카운트 {sum(e['sent']):,}원 → 시트 {sum(e['target']):,}원")
        memo = memo_of(e)
        cells = [day, vtype, pcode, partner, str(e['supply']), '', '', vat, memo, SALES_ACCOUNT, AR_ACCOUNT, '',
                 e['code']]
        sales.append({'key': e['key'], 'code': e['code'], 'kind': e['kind'], 'date': e.get('date') or '',
                      'partner': partner, 'supply': e['supply'], 'vat': e['vat'], 'memo': memo, 'vat_type': vtype,
                      'sent': list(e['sent']), 'target': list(e['target']), 'warn': warn, 'trace': trace,
                      'cells': cells})
        if e['kind'] == '공사 확정':
            projects.append({'key': e['key'], 'code': e['code'], 'name': e.get('addr') or '',
                             'warn': [] if e.get('addr') else ['현장 주소가 비어 있어요'],
                             'cells': [e['code'], e.get('addr') or '', '', '', '', '']})
    return {'sales': sales, 'projects': projects}


KIND_EXCLUDED = '제외'


def mark_done(keys: List[str], by: str, items: Optional[List[dict]] = None, exclude: bool = False) -> Dict[str, int]:
    """화면에서 [넣음 표시] — 지금 목록에 있는 전표만 장부에 적는다(목록이 바뀐 뒤 옛 키는 건너뜀).
    exclude=True 는 [제외](11-26 부장님) — 이카운트에 넣지 않기로 한 줄(SH 정산 공사 등). 장부엔 kind '제외'로 남겨
    목록에서 빠지게 한다(이카운트엔 없으니 나중에 이카운트와 맞출 땐 '제외' 줄을 빼고 본다)."""
    items = items if items is not None else (collect_items(force_refresh=False) or [])
    by_key = {e['key']: e for e in items}
    done = skipped = 0
    for k in keys:
        e = by_key.get(k)
        if not e:
            skipped += 1
            continue
        if record_entry(e['code'], e['supply'], e['vat'], KIND_EXCLUDED if exclude else e['kind'], k, by=by,
                        note=f"제외 — 원래 {e['kind']}" if exclude else ''):
            done += 1
        else:
            skipped += 1
    if done:
        logger.info(f"[ECOUNT] 그룹웨어 [{'제외' if exclude else '넣음 표시'}] {done}건 by {by}")
    return {'done': done, 'skipped': skipped}
