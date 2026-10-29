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


def build_blocks(items: List[dict]) -> List[dict]:
    shown = items[:MAX_ITEMS]
    blocks: List[dict] = [
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
    dates = change_dates([e['code'] for e in items if e['kind'] != '공사 확정'])
    for e in items:
        if e['kind'] != '공사 확정':
            e['date'] = dates.get(e['code'], '')
    items.sort(key=lambda e: (e['date'] or '9999', e['code']))
    return items


def send_ecount_sales_list() -> dict:
    """스케줄러 진입점 (정산 핀 리마인드와 같은 10:00·15:30). 영업일만, 0건이면 안 보냄."""
    try:
        from dashboard.services.absent_remind import _is_business_day
        if not _is_business_day(date.today()):
            return {'ok': True, 'total': 0, 'reason': 'non_business_day'}
    except Exception:
        pass
    items = collect_items()
    if items is None:
        return {'ok': False, 'total': 0, 'reason': 'no_baseline_or_sheet'}
    if not items:
        return {'ok': True, 'total': 0, 'reason': None}
    from dashboard.services.pin_remind import open_todo_dm
    c, ch = open_todo_dm()   # 경영지원(SB) 개인 DM — 채널에 올리지 않는다 (2026-10-30)
    if not c or not ch:
        return {'ok': False, 'total': len(items), 'reason': 'dm_open_failed'}
    try:
        r = c.chat_postMessage(channel=ch, text=f'이카운트 매출전표 — 넣을 것 {len(items)}건',
                               blocks=build_blocks(items), unfurl_links=False, unfurl_media=False)
        logger.info(f'[ECOUNT] 매출전표 목록 발송 {len(items)}건 ts={r.get("ts")}')
        return {'ok': bool(r.get('ok')), 'total': len(items), 'ts': r.get('ts', '')}
    except Exception as exc:
        logger.error(f'[ECOUNT] 매출전표 목록 발송 실패: {exc}', exc_info=True)
        return {'ok': False, 'total': len(items), 'reason': str(exc)}
