"""이카운트 대조 — 이카운트 매출과 공사 현황이 어긋난 공사를 그룹웨어에서 하나씩 처리(2026-11-27 부장님).

이카운트는 조회 API 가 없어서, 이카운트 회계거래조회를 읽어 공사별로 합산한 결과를 '한 번의 대조(batch)'로
저장해 둔다(11-26 첫 대조: 2025-01~ 매출전표 3,135장 → 공사 2,351개, 어긋남 + 코드 다름). 새로 읽으면 새 batch.
경영지원(SB)은 줄마다 [이카운트 고침] / [시트가 틀림] / [그대로 둠] 을 누르고, 처리한 줄은 아래로 내려간다.
[시트가 틀림]은 시트를 고쳐야 한다는 표시일 뿐 — 시트는 사람이 고친다(금액 수정 게이트·버전 규칙 그대로).
앞으로 생기는 확정·금액 변경은 이카운트 업로드 ①·② 탭(ecount_sales)이 잡는다. 이 표는 지난 차이만.
"""
from __future__ import annotations

from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_DDL = '''
    CREATE TABLE IF NOT EXISTS ecount_audit_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch TEXT NOT NULL,
        item_key TEXT NOT NULL,
        category TEXT NOT NULL,
        code TEXT NOT NULL,
        ecount_code TEXT,
        site TEXT,
        eco_supply INTEGER,
        eco_vat INTEGER,
        sheet_supply INTEGER,
        sheet_vat INTEGER,
        evidence TEXT,
        note TEXT,
        status TEXT NOT NULL DEFAULT 'open',
        memo TEXT,
        resolved_by TEXT,
        resolved_at TIMESTAMP,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(batch, item_key)
    )
'''

# 묶음 — 화면 순서. (키, 제목, 설명)
CATEGORIES = (
    ('dup', '중복 입력 의심', '같은 날 같은 금액 전표가 여러 장이에요. 한 장만 남기고 지워야 할 수 있어요.'),
    ('supply', '공급가가 다름', '이카운트 매출 합계와 공사 현황 금액이 달라요.'),
    ('code', '프로젝트 코드가 다름', '이카운트 프로젝트 코드와 공사 현황 코드가 달라요. 다음 전표 때 프로젝트를 못 찾을 수 있어요.'),
    ('vat_sheet', '시트는 부가세 있음, 이카운트는 없음', '계산서를 발행했는데 이카운트 매출에 부가세가 없어요.'),
    ('vat_eco', '이카운트는 부가세 있음, 시트는 없음', '시트 계산서 칸은 대부분 N입금(계산서 없이 N 통장 입금)이에요.'),
    ('split', '합계는 같고 나눔만 다름', '받은 돈 합계는 같아요. 나눠 내는 공사처럼 일부 몫만 부가세가 있는 경우예요.'),
    ('round', '1만 원 이하 반올림 차이', '부가세 포함 금액으로 정한 공사의 반올림 차이예요.'),
)
CATEGORY_KEYS = tuple(c[0] for c in CATEGORIES)

STATUSES = {
    'open': '처리 전',
    'ecount_fixed': '이카운트 고침',
    'sheet_wrong': '시트가 틀림',
    'keep': '그대로 둠',
}


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


def _conn():
    conn = _db()._get_connection()
    return conn


def _ensure(conn) -> None:
    conn.execute(_DDL)
    conn.execute('CREATE INDEX IF NOT EXISTS idx_ecount_audit_batch ON ecount_audit_items(batch)')
    # 처리 기록(11-27 부장님 '근거는 기존 거 밑에 히스토리') — 되돌릴 때 앞 처리·메모를 여기로 옮기고, 다시 대조한 결과도 한 줄씩 쌓음.
    cols = {r[1] for r in conn.execute('PRAGMA table_info(ecount_audit_items)').fetchall()}
    if 'history' not in cols:
        conn.execute('ALTER TABLE ecount_audit_items ADD COLUMN history TEXT')


def _md(ts) -> str:
    """resolved_at(UTC 'YYYY-MM-DD HH:MM:SS') → 한국 시각 'M/D'. 못 읽으면 오늘."""
    from datetime import datetime, timedelta
    try:
        d = datetime.strptime(str(ts)[:19], '%Y-%m-%d %H:%M:%S') + timedelta(hours=9)
    except (TypeError, ValueError):
        d = datetime.now()
    return f'{d.month}/{d.day}'


def _append_history(conn, item_id: int, line: str) -> None:
    row = conn.execute('SELECT history FROM ecount_audit_items WHERE id = ?', (item_id,)).fetchone()
    old = (row[0] if row else '') or ''
    conn.execute('UPDATE ecount_audit_items SET history = ? WHERE id = ?',
                 ((old + '\n' if old else '') + line.strip(), item_id))


def add_history(item_id: int, line: str) -> None:
    """근거 칸 아래 기록에 한 줄 더함(다시 대조한 결과 등). 상태·메모는 그대로."""
    with _conn() as conn:
        _ensure(conn)
        if not conn.execute('SELECT 1 FROM ecount_audit_items WHERE id = ?', (item_id,)).fetchone():
            raise LookupError('대조 항목을 찾지 못했어요.')
        _append_history(conn, item_id, line)
        conn.commit()


def import_items(batch: str, items: List[Dict]) -> int:
    """대조 결과 넣기 — 같은 batch·item_key 는 건너뜀(두 번 넣어도 처리 상태가 지워지지 않게). 넣은 줄 수."""
    added = 0
    with _conn() as conn:
        _ensure(conn)
        for it in items:
            cat = it.get('category')
            if cat not in CATEGORY_KEYS or not it.get('code'):
                raise ValueError(f'대조 항목 형식이 맞지 않아요: {it}')
            key = it.get('item_key') or f"{cat}|{it['code']}"
            cur = conn.execute(
                'INSERT OR IGNORE INTO ecount_audit_items (batch, item_key, category, code, ecount_code, site, '
                'eco_supply, eco_vat, sheet_supply, sheet_vat, evidence, note) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
                (batch, key, cat, it['code'], it.get('ecount_code') or it['code'], it.get('site') or '',
                 it.get('eco_supply'), it.get('eco_vat'), it.get('sheet_supply'), it.get('sheet_vat'),
                 it.get('evidence') or '', it.get('note') or ''))
            added += cur.rowcount or 0
        conn.commit()
    return added


def latest_batch() -> Optional[str]:
    with _conn() as conn:
        _ensure(conn)
        row = conn.execute('SELECT MAX(batch) FROM ecount_audit_items').fetchone()
    return row[0] if row and row[0] else None


def list_items(batch: Optional[str] = None) -> Dict:
    """{batch, groups:[{key,label,desc,items}], counts:{open,done,total}} — 묶음 안에서 처리 전이 위, 처리한 줄이 아래."""
    batch = batch or latest_batch()
    if not batch:
        return {'batch': None, 'groups': [], 'counts': {'open': 0, 'done': 0, 'total': 0}}
    with _conn() as conn:
        _ensure(conn)
        cur = conn.execute('SELECT * FROM ecount_audit_items WHERE batch = ? ORDER BY code', (batch,))
        names = [d[0] for d in cur.description]
        rows = [dict(zip(names, tuple(r))) for r in cur.fetchall()]
    groups = []
    for key, label, desc in CATEGORIES:
        its = [r for r in rows if r['category'] == key]
        if not its:
            continue
        its.sort(key=lambda r: (r['status'] != 'open', r['code']))
        for r in its:
            r['status_label'] = STATUSES.get(r['status'], r['status'])
        groups.append({'key': key, 'label': label, 'desc': desc, 'items': its,
                       'open': sum(1 for r in its if r['status'] == 'open')})
    n_open = sum(1 for r in rows if r['status'] == 'open')
    return {'batch': batch, 'groups': groups, 'counts': {'open': n_open, 'done': len(rows) - n_open, 'total': len(rows)}}


def set_status(item_id: int, status: str, memo: str, by: str) -> Dict:
    """줄 처리 — open 으로 되돌리면 처리한 사람·시각을 지움. 메모는 그대로 덮어씀(빈칸이면 지움)."""
    if status not in STATUSES:
        raise ValueError('처리 상태가 맞지 않아요.')
    memo = (memo or '').strip()[:200]
    with _conn() as conn:
        _ensure(conn)
        row = conn.execute('SELECT id, code, status, memo, resolved_by, resolved_at FROM ecount_audit_items WHERE id = ?',
                           (item_id,)).fetchone()
        if not row:
            raise LookupError('대조 항목을 찾지 못했어요. 새로 고쳐 주세요.')
        if row[2] != 'open':                    # 앞 처리는 기록으로 남김(되돌려도 누가 무엇을 했는지 보이게)
            _append_history(conn, item_id, f'{_md(row[5])} {row[4] or ""} {STATUSES.get(row[2], row[2])}'
                                           + (f' — {row[3]}' if row[3] else ''))
        if status == 'open':
            conn.execute("UPDATE ecount_audit_items SET status='open', memo=?, resolved_by=NULL, resolved_at=NULL "
                         'WHERE id = ?', (memo or None, item_id))
        else:
            conn.execute('UPDATE ecount_audit_items SET status=?, memo=?, resolved_by=?, resolved_at=CURRENT_TIMESTAMP '
                         'WHERE id = ?', (status, memo or None, by, item_id))
        conn.commit()
    logger.info(f'[ECOUNT_AUDIT] {row[1]} → {STATUSES[status]} by {by}{" · " + memo if memo else ""}')
    return {'id': item_id, 'status': status, 'status_label': STATUSES[status]}
