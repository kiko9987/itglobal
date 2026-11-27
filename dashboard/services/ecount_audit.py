"""이카운트 대조 — 이카운트 매출과 공사 현황이 어긋난 공사를 그룹웨어에서 하나씩 처리(2026-11-27 부장님).

이카운트는 조회 API 가 없어서, 이카운트 회계거래조회를 읽어 공사별로 합산한 결과를 '한 번의 대조(batch)'로
저장해 둔다(11-26 첫 대조: 2025-01~ 매출전표 3,135장 → 공사 2,351개, 어긋남 + 코드 다름). 새로 읽으면 새 batch.
경영지원(SB)은 줄마다 [이카운트 고침] / [시트가 틀림] / [그대로 둠] 을 누르고, 처리한 줄은 아래로 내려간다.
[시트가 틀림]은 시트를 고쳐야 한다는 표시일 뿐 — 시트는 사람이 고친다(금액 수정 게이트·버전 규칙 그대로).
앞으로 생기는 확정·금액 변경은 이카운트 업로드 ①·② 탭(ecount_sales)이 잡는다. 이 표는 지난 차이만.

밤 자동 대조(11-28 부장님, services/ecount_recon.py): 매일 자정 이카운트를 읽어 batch 'auto' 한 묶음을 계속 고친다.
  - 매출(side 'sales')·매입(side 'buy')을 나눠 둔다. 공사 하나에 한 줄(item_key = side|코드).
  - 버튼은 'SB 가 고쳤다고 함' 표시일 뿐 — 다음 대조에서 실제로 맞으면 '맞음'(완료), 아직 다르면 '처리 전'으로
    되돌리고 근거를 처리 기록(history)에 남긴다(apply_nightly).
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
    # 매입(11-28) — 이카운트 매입전표(공급가) 공사별 합 ↔ 공사 현황 원가(제품대·도급비·자재비·기타비 합)
    ('buy_code', '공사 현황에 없는 코드', '이카운트 매입전표의 프로젝트 코드가 공사 현황에 없어요. 같은 번호의 공사로 옮겨야 할 수 있어요.'),
    ('buy_cancel', '취소된 공사에 남은 매입', '공사 현황에서 취소된 공사인데 이카운트에 매입이 남아 있어요.'),
    ('buy_noeco', '이카운트 매입이 없음', '공사 현황에는 원가가 있는데 이카운트에 매입전표가 없어요.'),
    # 시트 원가(11-28 부장님 'SB 할 일은 30건만') — 이카운트는 맞고 공사 현황 원가 칸을 채우거나 고칠 일. 보기만, DM 에 안 넣음
    ('buy_nocost', '공사 현황 원가가 비어 있음', '이카운트에는 매입이 있는데 공사 현황 원가 칸(제품대·도급비·자재비·기타비)이 모두 비어 있어요. 시트에 원가를 넣을 일이에요.'),
    ('buy_diff', '원가가 다름', '이카운트 매입 합계(공급가)와 공사 현황 원가 합계가 달라요. 어느 쪽이 맞는지 보고 시트 원가를 고칠 일이에요.'),
)
CATEGORY_KEYS = tuple(c[0] for c in CATEGORIES)
BUY_CATEGORIES = ('buy_code', 'buy_cancel', 'buy_noeco')          # 매입 — 이카운트를 고칠 일(SB)
COST_CATEGORIES = ('buy_nocost', 'buy_diff')                     # 시트 원가 — 공사 현황 원가 칸을 채우거나 고칠 일
SIDES = (('sales', '매출'), ('buy', '매입'), ('cost', '시트 원가'))
WORK_SIDES = ('sales', 'buy')                                   # SB 할 일(탭 숫자·아침 DM)
AUTO_BATCH = 'auto'              # 밤 자동 대조 묶음(계속 고쳐 씀). 손으로 넣은 대조는 날짜 이름 batch('2026-11-26')
AUTO_BY = '자동 대조'
RESOLVED_SHOW_DAYS = 7           # '맞음'으로 끝난 줄은 이 기간만 화면 아래에 남김

STATUSES = {
    'open': '처리 전',
    'ecount_fixed': '이카운트 고침',
    'sheet_wrong': '시트가 틀림',
    'keep': '그대로 둠',
    'resolved': '맞음',           # 밤 자동 대조에서 이카운트·공사 현황이 같아진 것(사람이 누르는 버튼 아님)
}
MANUAL_STATUSES = ('open', 'ecount_fixed', 'sheet_wrong', 'keep')


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
    # 밤 자동 대조(11-28) — 매출/매입 구분, 처음·마지막으로 본 시각
    for col, ddl in (('side', "TEXT NOT NULL DEFAULT 'sales'"), ('first_seen', 'TIMESTAMP'), ('last_seen', 'TIMESTAMP')):
        if col not in cols:
            conn.execute(f'ALTER TABLE ecount_audit_items ADD COLUMN {col} {ddl}')


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
    """화면에 보일 대조 묶음 — 밤 자동 대조('auto')가 있으면 그것, 없으면 가장 최근 날짜 묶음."""
    with _conn() as conn:
        _ensure(conn)
        if conn.execute('SELECT 1 FROM ecount_audit_items WHERE batch = ? LIMIT 1', (AUTO_BATCH,)).fetchone():
            return AUTO_BATCH
        row = conn.execute('SELECT MAX(batch) FROM ecount_audit_items WHERE batch != ?', (AUTO_BATCH,)).fetchone()
    return row[0] if row and row[0] else None


def _rows(conn, where: str, args) -> List[Dict]:
    cur = conn.execute(f'SELECT * FROM ecount_audit_items WHERE {where}', args)
    names = [d[0] for d in cur.description]
    return [dict(zip(names, tuple(r))) for r in cur.fetchall()]


def list_items(batch: Optional[str] = None) -> Dict:
    """{batch, sides:[{key,label,open,total}], groups:[{key,side,label,desc,items,open}], counts:{open,done,total}}
    — 묶음 안에서 처리 전이 위, 처리한 줄이 아래. 자동 대조의 '맞음' 줄은 RESOLVED_SHOW_DAYS 동안만."""
    batch = batch or latest_batch()
    if not batch:
        return {'batch': None, 'sides': [], 'groups': [], 'counts': {'open': 0, 'done': 0, 'total': 0}}
    with _conn() as conn:
        _ensure(conn)
        if batch == AUTO_BATCH:
            rows = _rows(conn, "batch = ? AND (status != 'resolved' OR resolved_at >= datetime('now', ?)) ORDER BY code",
                         (batch, f'-{RESOLVED_SHOW_DAYS} days'))
        else:
            rows = _rows(conn, 'batch = ? ORDER BY code', (batch,))
    groups = []
    for key, label, desc in CATEGORIES:
        its = [r for r in rows if r['category'] == key]
        if not its:
            continue
        its.sort(key=lambda r: (r['status'] != 'open', r['code']))
        for r in its:
            r['status_label'] = STATUSES.get(r['status'], r['status'])
        side = 'buy' if key in BUY_CATEGORIES else 'cost' if key in COST_CATEGORIES else 'sales'
        groups.append({'key': key, 'side': side, 'label': label, 'desc': desc,
                       'items': its, 'open': sum(1 for r in its if r['status'] == 'open')})
    sides = []
    for sk, sl in SIDES:
        g = [x for x in groups if x['side'] == sk]
        sides.append({'key': sk, 'label': sl, 'open': sum(x['open'] for x in g), 'total': sum(len(x['items']) for x in g)})
    n_open = sum(1 for r in rows if r['status'] == 'open')
    return {'batch': batch, 'sides': sides, 'groups': groups,
            'counts': {'open': n_open, 'done': len(rows) - n_open, 'total': len(rows)}}


# ── 밤 자동 대조(11-28) ─────────────────────────────────────────────────

_AMT_COLS = ('eco_supply', 'eco_vat', 'sheet_supply', 'sheet_vat')
_INFO_COLS = ('category', 'ecount_code', 'site', 'evidence', 'note') + _AMT_COLS


def _amts(r: Dict) -> tuple:
    return tuple(int(r.get(k) or 0) for k in _AMT_COLS)


def _amt_text(side: str, a: tuple) -> str:
    es, ev, ss, sv = a
    if side == 'buy':
        return f'이카운트 {es:,}원 · 공사 현황 {ss:,}원'
    return f'이카운트 {es:,}/{ev:,} · 공사 현황 {ss:,}/{sv:,}'


def _batch_md(batch: str) -> str:
    return f'{int(batch[5:7])}/{int(batch[8:10])}' if len(batch) >= 10 and batch[:4].isdigit() else batch


def apply_nightly(side: str, found: Dict[str, Dict], md: str) -> Dict:
    """밤 대조 결과 반영. found = {공사 코드: 대조 항목(category·금액·근거)} — 이번에 어긋난 공사 전부.

    - 처음 보는 공사: '처리 전'으로 넣음. 손 대조(11-26 등)에서 [그대로 둠]·[시트가 틀림]으로 둔 같은 금액이면 그 상태를 이어받음.
    - 이미 있는 줄: 금액·근거를 새로 고침. [이카운트 고침] 표시인데 아직 다르면 '처리 전'으로 되돌리고 기록.
      [그대로 둠]·[시트가 틀림]은 금액이 바뀌었을 때만 다시 '처리 전'. '맞음'이던 줄이 다시 다르면 '처리 전'.
    - found 에 없는(= 이번엔 같은) 줄: '맞음'(완료)으로 바꾸고 기록.
    돌려줌: {'new': [...], 'reopened': [...], 'resolved': [...], 'baseline': 이 쪽 첫 대조인지}"""
    if side not in dict(SIDES):
        raise ValueError('side')
    out = {'new': [], 'reopened': [], 'resolved': [], 'baseline': False}
    with _conn() as conn:
        _ensure(conn)
        have = {r['code']: r for r in _rows(conn, 'batch = ? AND side = ?', (AUTO_BATCH, side))}
        out['baseline'] = not have
        old = {}
        if side == 'sales':   # 손 대조 묶음의 마지막 처리 — 같은 코드의 가장 최근 batch
            for r in _rows(conn, 'batch != ? ORDER BY batch', (AUTO_BATCH,)):
                old[r['code']] = r
        for code, it in found.items():
            cat = it.get('category')
            if cat not in CATEGORY_KEYS:
                raise ValueError(f'대조 항목 형식이 맞지 않아요: {it}')
            new_a = _amts(it)
            row = have.get(code)
            hist: List[str] = []
            if row is None:
                status, memo, by, at = 'open', None, None, None
                o = old.get(code)
                if o:
                    if o.get('history'):
                        hist.extend(x for x in o['history'].split('\n') if x.strip())
                    olabel = STATUSES.get(o['status'], o['status'])
                    if o['status'] in ('keep', 'sheet_wrong') and _amts(o) == new_a:
                        status, memo, by, at = o['status'], o['memo'], o['resolved_by'], o['resolved_at']
                        hist.append(f"{md} 자동 대조 — {_batch_md(o['batch'])} 대조의 '{olabel}'을 이어받음(금액 그대로)")
                    elif o['status'] != 'open':
                        hist.append(f"{md} 자동 대조 — {_batch_md(o['batch'])} '{olabel}'"
                                    + (f"({o['memo']})" if o.get('memo') else '') + ' 뒤에도 아직 달라 처리 전으로')
                conn.execute(
                    'INSERT INTO ecount_audit_items (batch, item_key, side, category, code, ecount_code, site, eco_supply, '
                    'eco_vat, sheet_supply, sheet_vat, evidence, note, status, memo, resolved_by, resolved_at, history, '
                    'first_seen, last_seen) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP,CURRENT_TIMESTAMP)',
                    (AUTO_BATCH, f'{side}|{code}', side, cat, code, it.get('ecount_code') or code, it.get('site') or '',
                     *new_a, it.get('evidence') or '', it.get('note') or '', status, memo, by, at, '\n'.join(hist) or None))
                if status == 'open':
                    out['new'].append(code)
                continue
            st, changed = row['status'], _amts(row) != new_a
            reopen = False
            if st == 'ecount_fixed':
                hist.append(f"{md} 자동 대조 — '이카운트 고침' 뒤에도 아직 달라 처리 전으로")
                reopen = True
            elif st == 'resolved':
                hist.append(f'{md} 자동 대조 — 다시 달라져 처리 전으로({_amt_text(side, new_a)})')
                reopen = True
            elif st in ('keep', 'sheet_wrong') and changed:
                hist.append(f"{md} 자동 대조 — '{STATUSES[st]}' 뒤 금액이 바뀌어 처리 전으로({_amt_text(side, _amts(row))} → "
                            f'{_amt_text(side, new_a)})')
                reopen = True
            elif st == 'open' and changed:
                hist.append(f'{md} 자동 대조 — 금액 바뀜({_amt_text(side, _amts(row))} → {_amt_text(side, new_a)})')
            if reopen and st in ('ecount_fixed', 'keep', 'sheet_wrong'):
                # 사람이 남긴 처리·메모도 기록으로(되돌리기와 같은 모양)
                hist.insert(0, f"{_md(row['resolved_at'])} {row['resolved_by'] or ''} {STATUSES[st]}"
                               + (f" — {row['memo']}" if row['memo'] else ''))
            for line in hist:
                _append_history(conn, row['id'], line)
            vals = [int(it.get(c) or 0) if c in _AMT_COLS else (it.get(c) or '') for c in _INFO_COLS]
            vals[_INFO_COLS.index('ecount_code')] = it.get('ecount_code') or code
            sets = ', '.join(f'{c} = ?' for c in _INFO_COLS)
            conn.execute(f'UPDATE ecount_audit_items SET {sets}, last_seen = CURRENT_TIMESTAMP WHERE id = ?',
                         (*vals, row['id']))
            if reopen:
                conn.execute("UPDATE ecount_audit_items SET status = 'open', resolved_by = NULL, resolved_at = NULL "
                             'WHERE id = ?', (row['id'],))
                out['reopened'].append(code)
        for code, row in have.items():
            if code in found or row['status'] == 'resolved':
                continue
            was = STATUSES.get(row['status'], row['status'])
            _append_history(conn, row['id'], f'{md} 자동 대조 — 이카운트와 공사 현황이 같아져 맞음'
                                             + (f"('{was}' 확인)" if row['status'] != 'open' else ''))
            conn.execute("UPDATE ecount_audit_items SET status = 'resolved', resolved_by = ?, "
                         'resolved_at = CURRENT_TIMESTAMP, last_seen = CURRENT_TIMESTAMP WHERE id = ?', (AUTO_BY, row['id']))
            out['resolved'].append(code)
        conn.commit()
    logger.info(f"[ECOUNT_AUDIT] 자동 대조 {side}: 새로 {len(out['new'])} · 다시 열림 {len(out['reopened'])} · "
                f"맞음 {len(out['resolved'])}{' (첫 대조)' if out['baseline'] else ''}")
    return out


def open_items(side: str) -> List[Dict]:
    """자동 대조에서 지금 '처리 전'인 줄(DM 용)."""
    with _conn() as conn:
        _ensure(conn)
        return _rows(conn, "batch = ? AND side = ? AND status = 'open' ORDER BY code", (AUTO_BATCH, side))


def set_status(item_id: int, status: str, memo: str, by: str) -> Dict:
    """줄 처리 — open 으로 되돌리면 처리한 사람·시각을 지움. 메모는 그대로 덮어씀(빈칸이면 지움)."""
    if status not in MANUAL_STATUSES:            # '맞음'은 밤 자동 대조만 붙임
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
