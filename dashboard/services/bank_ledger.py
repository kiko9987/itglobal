"""입출금 장부(2026-11-26 부장님 — 재무 > 입출금, 다음 단계 자금일보).

- 은행 문자(입금·출금)가 들어올 때마다 날짜·시각·계좌·금액·상대방·**잔액**을 영구 보관한다.
  예전엔 입금 문자는 카드 처리 뒤 지우고, 출금은 14일만, 잔액은 받자마자 버렸다(직원 노출 금지 원칙).
  11-26 부장님 '잔액 보관' — 잔액은 이 장부(관리자 전용 재무 화면)에만 남기고 슬랙·시트에는 지금처럼 안 남긴다.
- 문자가 안 오는 계좌(하나 G·우리·수협·기업 P·국민 N)는 SB(관리자)가 화면에서 직접 넣는다(부장님 '해결 방법 있기 전까지').
- 입금 지정 → 경영지원 [확인 후 기록]이 끝나면 그 입금 줄에 공사 코드·단계(적요 'G4160-JW 매출수금 중도금')를 붙인다.
- 시작은 2026-11-26 지금부터(과거 출금·잔액은 남아 있지 않음). 계좌마다 첫 잔액은 [잔액 맞추기]로 한 번 넣는다.
- 같은 문자가 폰 3대·붙여넣기로 여러 번 와도 한 줄(uid = sms_intake.dedup_hash, 본문 정규화 기준).
- 장부 기록이 실패해도 입금 카드·환불 짝짓기 같은 기존 흐름은 그대로 간다(부르는 쪽이 예외를 삼킴).
"""
from __future__ import annotations

import re
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)
_LOG = '[입출금]'

IN, OUT, BAL = 'in', 'out', 'bal'          # 입금 · 출금 · 잔액 맞추기(금액 0, 잔액만)
DIR_LABEL = {IN: '입금', OUT: '출금', BAL: '잔액 맞추기'}

# 자금일보 계좌(11-26 SB 자금일보 양식 그대로 — 위쪽 계좌 표 8줄, code = 상세 내역 '구분' 칸 값) + 상세에만 나오는 계좌.
# 농협(N통장)은 위 표에 없고 상세 내역에 '농협'으로 적는다. 하도급지킴이는 자금일보에 안 나옴(문자만 보관)
SEED_ACCOUNTS = (
    # code, 사업자, 은행, 계좌번호, 위쪽 계좌 표에 있나, 문자 자동
    ('1011', 'G', '기업은행', '452-039388-01-011', 1, 1),
    ('1012', 'G', '하나은행', '255-910014-78704', 1, 0),
    ('1013', 'G', '우리은행', '1005-804-501966', 1, 0),
    ('1014', 'G', '수협', '1010-2378-7755', 1, 0),
    ('9011', 'R', '하나은행', '255-910014-30704', 1, 0),
    ('9012', 'R', '하나은행', '255-910014-31304', 1, 1),
    ('4011', 'P', '기업은행', '452-027746-04-011', 1, 0),
    ('5011', 'N', '국민은행', '165801-04-393444', 1, 0),
    ('농협', 'N', '농협(N통장)', '352-****-1682-33', 0, 1),
    ('하도급', 'R', '기업은행(하도급지킴이)', '452-047865-04-039', 0, 1),
)

DDL = [
    '''CREATE TABLE IF NOT EXISTS bank_accounts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT NOT NULL UNIQUE,
        entity TEXT NOT NULL,
        bank TEXT NOT NULL,
        number TEXT NOT NULL,
        in_report INTEGER NOT NULL DEFAULT 1,
        sms INTEGER NOT NULL DEFAULT 0,
        sort INTEGER NOT NULL DEFAULT 0,
        note TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS bank_tx (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        uid TEXT NOT NULL UNIQUE,
        account_id INTEGER,
        at TEXT NOT NULL,
        direction TEXT NOT NULL,
        amount INTEGER NOT NULL DEFAULT 0,
        balance INTEGER,
        counterparty TEXT,
        title TEXT,
        memo TEXT,
        project_code TEXT,
        source TEXT NOT NULL,
        raw TEXT,
        created_by TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT,
        updated_by TEXT,
        deleted_at TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_bank_tx_at ON bank_tx(at)',
    'CREATE INDEX IF NOT EXISTS idx_bank_tx_acct ON bank_tx(account_id, at)',
]
_table_ready = False


class LedgerError(ValueError):
    """화면에 그대로 보여 줄 입력 오류."""


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


@contextmanager
def _conn():
    """⚠️ 연결 안에서 다시 _conn() 을 열지 않는다."""
    global _table_ready
    with _db()._get_connection() as conn:
        try:
            if not _table_ready:
                for q in DDL:
                    conn.execute(q)
                for i, (code, ent, bank, num, rep, sms) in enumerate(SEED_ACCOUNTS):
                    conn.execute('INSERT OR IGNORE INTO bank_accounts (code, entity, bank, number, in_report, sms, sort) '
                                 'VALUES (?, ?, ?, ?, ?, ?, ?)', (code, ent, bank, num, rep, sms, i))
                conn.commit()
                _table_ready = True
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def accounts() -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute('SELECT * FROM bank_accounts ORDER BY sort, id')]


def _digits(s: str) -> str:
    return re.sub(r'\D', '', s or '')


def account_for_text(text: str) -> Optional[dict]:
    """문자 속 계좌 → 장부 계좌. 문자는 가운데를 가리므로(452***38801011) 앞 3자리 + 보이는 끝자리로 맞춘다."""
    for m in re.finditer(r'(?<!\d)(\d{3})[\d*\-]{6,}', text or ''):
        tok = m.group(0).replace('-', '').rstrip('*')
        visible = tok.rsplit('*', 1)[-1] if '*' in tok else tok
        if len(visible) < 5:
            continue
        for a in accounts():
            num = a['number'].replace('-', '')
            if '*' in num:                                   # 등록 번호도 가려져 있으면(농협) 끝자리끼리
                if tok[:3] == num[:3] and num.rsplit('*', 1)[-1].endswith(visible[-6:]):
                    return a
                continue
            if num.startswith(tok[:3]) and num.endswith(visible):
                return a
    return None


# ── 문자 읽기 ─────────────────────────────────────────────────────────
_BAL_RE = re.compile(r'(?:잔액|잔고|현재\s*잔액|이체\s*후\s*잔액)\s*[:：]?\s*([\d,]+)\s*원?')
_WHEN_RE = re.compile(r'(?:(\d{4})[/.-])?(\d{1,2})/(\d{1,2}),?\s+(\d{1,2}):(\d{2})')


def parse_when(text: str, now: Optional[datetime] = None) -> str:
    """'2026/08/10 14:49'·'하나,08/07, 15:33'·'10/26 09:15' → 'YYYY-MM-DD HH:MM'. 없으면 받은 시각.
    연도가 없으면 올해 — 다만 받은 날보다 한 달 넘게 뒤면 작년(1월에 받은 12월 문자)."""
    now = now or datetime.now()
    m = _WHEN_RE.search(text or '')
    if not m:
        return now.strftime('%Y-%m-%d %H:%M')
    y, mo, d, hh, mm = m.groups()
    mo, d, hh, mm = int(mo), int(d), int(hh), int(mm)
    year = int(y) if y else now.year
    try:
        dt = datetime(year, mo, d, hh, mm)
    except ValueError:
        return now.strftime('%Y-%m-%d %H:%M')
    if not y and dt - now > timedelta(days=31):
        dt = dt.replace(year=year - 1)
    return dt.strftime('%Y-%m-%d %H:%M')


def parse_balance(text: str) -> Optional[int]:
    m = _BAL_RE.search(text or '')
    return int(m.group(1).replace(',', '')) if m else None


def parse_sms(raw: str, direction: str, now: Optional[datetime] = None) -> dict:
    """은행 문자 원문(잔액 포함) → 장부 한 줄 값. raw 에는 잔액을 뺀 본문만 남긴다."""
    from dashboard.services.sms_intake import normalize_deposit_layout, parse_preview, strip_balance
    clean = strip_balance(raw)
    try:
        clean = normalize_deposit_layout(clean)
    except Exception:
        pass
    pv = parse_preview(clean) or {}
    acct = account_for_text(raw)
    return {
        'account_id': acct['id'] if acct else None,
        'at': parse_when(raw, now),
        'direction': direction,
        'amount': abs(int(pv.get('amount') or 0)),
        'balance': parse_balance(raw),
        'counterparty': (pv.get('partner') or '').strip()[:60],
        'raw': clean[:500],
    }


def record_sms(raw: str, direction: str, uid: str, source: str, now: Optional[datetime] = None) -> Optional[int]:
    """은행 문자 한 건 → 장부(같은 uid 는 한 번만). 실패해도 예외를 올리지 않는다(기존 흐름 보호)."""
    try:
        v = parse_sms(raw, direction, now)
        if not v['amount']:
            return None
        with _conn() as conn:
            cur = conn.execute(
                'INSERT OR IGNORE INTO bank_tx (uid, account_id, at, direction, amount, balance, counterparty, source, raw, '
                'created_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                (uid, v['account_id'], v['at'], direction, v['amount'], v['balance'], v['counterparty'], source,
                 v['raw'], source, _now()))
            return cur.lastrowid if cur.rowcount else None
    except Exception as e:
        logger.error(f'{_LOG} 문자 장부 기록 실패({direction}, uid={uid}): {e}', exc_info=True)
        return None


def designate(uid: str, project_code: str, stage: str, amount: Optional[int] = None) -> None:
    """입금 지정이 시트에 기록된 뒤 — 그 입금 줄에 공사 코드·적요. 분할이면 줄마다 이어 붙인다."""
    try:
        piece = f'{project_code} 매출수금 {stage}' + (f' {int(amount):,}' if amount else '')
        with _conn() as conn:
            r = conn.execute('SELECT id, memo, project_code FROM bank_tx WHERE uid=? AND deleted_at IS NULL', (uid,)).fetchone()
            if not r:
                return
            memo = r['memo'] or ''
            if piece in memo:
                return
            codes = [c for c in (r['project_code'] or '').split(',') if c]
            if project_code not in codes:
                codes.append(project_code)
            conn.execute('UPDATE bank_tx SET memo=?, project_code=?, title=COALESCE(NULLIF(title, \'\'), ?), updated_at=?, '
                         'updated_by=? WHERE id=?',
                         ((memo + ' / ' if memo else '') + piece, ','.join(codes), '매출수금', _now(), '입금 지정', r['id']))
    except Exception as e:
        logger.error(f'{_LOG} 입금 지정 장부 반영 실패(uid={uid} {project_code}): {e}', exc_info=True)


# ── 화면 ─────────────────────────────────────────────────────────────
def _row(r) -> dict:
    d = dict(r)
    d['direction_label'] = DIR_LABEL.get(d['direction'], d['direction'])
    return d


def list_tx(start: str, end: str, account_id: Optional[int] = None, direction: str = '', unassigned: bool = False,
            q: str = '') -> List[dict]:
    """기간(YYYY-MM-DD ~ YYYY-MM-DD, 끝 날짜 포함) 거래 — 최신 먼저."""
    sql = ('SELECT t.*, a.code AS acct_code, a.bank AS acct_bank, a.entity AS acct_entity FROM bank_tx t '
           'LEFT JOIN bank_accounts a ON a.id=t.account_id WHERE t.deleted_at IS NULL AND t.at >= ? AND t.at < ?')
    args: list = [start, (date.fromisoformat(end) + timedelta(days=1)).isoformat()]
    if account_id:
        sql += ' AND t.account_id=?'
        args.append(account_id)
    if direction in (IN, OUT, BAL):
        sql += ' AND t.direction=?'
        args.append(direction)
    if unassigned:
        sql += " AND t.direction IN ('in','out') AND COALESCE(t.memo,'')=''"
    if q:
        sql += ' AND (t.counterparty LIKE ? OR t.memo LIKE ? OR t.title LIKE ? OR t.project_code LIKE ?)'
        args += [f'%{q}%'] * 4
    sql += ' ORDER BY t.at DESC, t.id DESC'
    with _conn() as conn:
        return [_row(r) for r in conn.execute(sql, args)]


def summary(day: Optional[str] = None) -> List[dict]:
    """계좌마다 마지막으로 아는 잔액(그 날 끝까지)·그 시각·오늘 입금·출금 합."""
    day = day or date.today().isoformat()
    upto = (date.fromisoformat(day) + timedelta(days=1)).isoformat()
    out = []
    with _conn() as conn:
        for a in conn.execute('SELECT * FROM bank_accounts ORDER BY sort, id').fetchall():
            last = conn.execute('SELECT balance, at FROM bank_tx WHERE account_id=? AND deleted_at IS NULL AND balance IS NOT NULL '
                                'AND at < ? ORDER BY at DESC, id DESC LIMIT 1', (a['id'], upto)).fetchone()
            sums = conn.execute("SELECT direction, SUM(amount) FROM bank_tx WHERE account_id=? AND deleted_at IS NULL "
                                "AND at >= ? AND at < ? GROUP BY direction", (a['id'], day, upto)).fetchall()
            s = {d: int(v or 0) for d, v in sums}
            out.append(dict(a, balance=last['balance'] if last else None, balance_at=last['at'] if last else '',
                            day_in=s.get(IN, 0), day_out=s.get(OUT, 0)))
    return out


def _clean_input(data: dict, partial: bool = False) -> dict:
    v = {}
    if 'account_id' in data or not partial:
        try:
            v['account_id'] = int(data.get('account_id') or 0)
        except (TypeError, ValueError):
            v['account_id'] = 0
        if not v['account_id'] or v['account_id'] not in {a['id'] for a in accounts()}:
            raise LedgerError('계좌를 골라 주세요.')
    if 'at' in data or not partial:
        at = str(data.get('at') or '').strip().replace('T', ' ')[:16]
        try:
            datetime.strptime(at, '%Y-%m-%d %H:%M')
        except ValueError:
            raise LedgerError('날짜와 시각을 확인해 주세요.')
        v['at'] = at
    if 'direction' in data or not partial:
        if data.get('direction') not in (IN, OUT, BAL):
            raise LedgerError('입금·출금·잔액 맞추기 중에서 골라 주세요.')
        v['direction'] = data['direction']

    def num(k):
        raw = str(data.get(k) if data.get(k) is not None else '').replace(',', '').strip()
        if raw == '':
            return None
        if not re.fullmatch(r'-?\d{1,13}', raw):
            raise LedgerError('금액은 숫자로 적어 주세요.')
        return int(raw)
    if 'amount' in data or not partial:
        v['amount'] = abs(num('amount') or 0)
    if 'balance' in data or not partial:
        v['balance'] = num('balance')
    for k, n in (('counterparty', 60), ('title', 30), ('memo', 200), ('project_code', 60)):
        if k in data or not partial:
            v[k] = str(data.get(k) or '').strip()[:n]
    d = v.get('direction')
    if d in (IN, OUT) and not v.get('amount') and not partial:
        raise LedgerError('금액을 적어 주세요.')
    if d == BAL and v.get('balance') is None:
        raise LedgerError('맞출 잔액을 적어 주세요.')
    if d == BAL:
        v['amount'] = 0
    return v


def add_manual(data: dict, by: str, admin: bool) -> int:
    if not admin:
        raise PermissionError('관리자만 넣을 수 있어요.')
    v = _clean_input(data)
    import uuid
    with _conn() as conn:
        return conn.execute(
            'INSERT INTO bank_tx (uid, account_id, at, direction, amount, balance, counterparty, title, memo, project_code, '
            'source, created_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
            ('manual:' + uuid.uuid4().hex, v['account_id'], v['at'], v['direction'], v['amount'], v['balance'],
             v['counterparty'], v['title'], v['memo'], v['project_code'], 'manual', by, _now())).lastrowid


def update(tid: int, data: dict, by: str, admin: bool) -> None:
    """고치기 — 문자로 들어온 줄은 계정·적요·공사 코드만(금액·시각·잔액은 은행 문자 그대로)."""
    if not admin:
        raise PermissionError('관리자만 고칠 수 있어요.')
    with _conn() as conn:
        r = conn.execute('SELECT * FROM bank_tx WHERE id=? AND deleted_at IS NULL', (tid,)).fetchone()
    if not r:
        raise LookupError('없는 거래예요. 새로 고쳐 주세요.')
    allowed = ('title', 'memo', 'project_code') if r['source'] != 'manual' else (
        'account_id', 'at', 'direction', 'amount', 'balance', 'counterparty', 'title', 'memo', 'project_code')
    v = _clean_input({k: data[k] for k in allowed if k in data}, partial=True)
    if not v:
        return
    sets = ', '.join(f'{k}=?' for k in v)
    with _conn() as conn:
        conn.execute(f'UPDATE bank_tx SET {sets}, updated_at=?, updated_by=? WHERE id=?', (*v.values(), _now(), by, tid))


def delete(tid: int, by: str, admin: bool) -> None:
    """지우기 — 직접 넣은 줄만(문자로 들어온 줄은 은행 기록이라 못 지움)."""
    if not admin:
        raise PermissionError('관리자만 지울 수 있어요.')
    with _conn() as conn:
        r = conn.execute('SELECT source FROM bank_tx WHERE id=? AND deleted_at IS NULL', (tid,)).fetchone()
        if not r:
            raise LookupError('없는 거래예요. 새로 고쳐 주세요.')
        if r['source'] != 'manual':
            raise LedgerError('은행 문자로 들어온 거래는 지울 수 없어요. 계정·적요만 고쳐 주세요.')
        conn.execute('UPDATE bank_tx SET deleted_at=?, updated_by=? WHERE id=?', (_now(), by, tid))


def record_from_text(text: str, source: str) -> None:
    """인입 문자 공통 입구 — 장부 계좌로 확인되는 입금·출금 문자면 기록(사업자 통장 목록에 없는 계좌도).
    입금 카드·출금 보관은 부르는 쪽이 그대로 한다."""
    try:
        from dashboard.services.sms_intake import dedup_hash, looks_like_payment, looks_like_withdrawal
        if looks_like_payment(text):
            direction = IN
        elif looks_like_withdrawal(text):
            direction = OUT
        else:
            return
        if not account_for_text(text):
            return
        record_sms(text, direction, dedup_hash(source, text), source.split(':', 1)[0] or 'sms')
    except Exception as e:
        logger.error(f'{_LOG} 문자 장부 입구 실패: {e}', exc_info=True)

