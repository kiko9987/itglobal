"""이카운트 밤 자동 대조(2026-11-28 부장님) — 구글 시트(공사 현황) = 이카운트 = 그룹웨어 장부.

매일 자정 이카운트 회계거래조회를 읽어(tools/ecount_reader/read_ecount.mjs, 헤드리스 Edge) 공사별로 합산하고
공사 현황·장부(ecount_sales_entries)와 맞춘다. 이카운트에는 아무것도 쓰지 않는다.

1) 장부 맞춤 — 이카운트에서 **실제로 읽은 금액이 시트와 같을 때만** 장부를 이카운트 금액으로 맞춘다(조정 줄 덧붙임,
   kind '이카운트 대조'). [넣음 표시] 없이 이카운트에 들어간 전표도 이렇게 장부에 들어가 ①·② 목록에서 빠진다 —
   이중 입력을 막는 것이 목적. 버튼 처리([이카운트 고침] 등)는 근거로 쓰지 않는다(11-27 부장님).
2) 대조 — 이카운트 ≠ 시트인 공사를 ③ 이카운트 대조(ecount_audit, batch 'auto')에 매출·매입으로 나눠 남긴다.
   지난밤 알려 준 줄이 이번에 같아졌으면 '맞음'(완료), [이카운트 고침] 표시인데 아직 다르면 '처리 전'으로 되돌림.
   단, 이카운트 = 장부 ≠ 시트(= ② 매출전표 목록에 넣을 전표로 이미 뜬 것)는 정상 대기라 ③에 따로 올리지 않는다.
3) 아침 DM — 경영지원(SB)에게 고칠 것 목록, 없으면 '수정할 것 없음'(send_morning_dm).

매입: 이카운트 매입·지출 전표(공급가) 공사별 합 ↔ 공사 현황 원가(제품대·도급비·자재비·기타비 합). 매입 장부는 없다.
SB 는 전날 것을 다음 날 고치고 당일 건은 손대지 않는다(11-28) — 자정 대조가 '어제까지' 기준이 된다.
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import time
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

ROOT = Path(__file__).resolve().parents[2]
READER = ROOT / 'tools' / 'ecount_reader' / 'read_ecount.mjs'
OUT_DIR = ROOT / 'instance' / 'ecount_recon'
READ_FROM = '2025-01-01'           # 이카운트를 읽는 시작일
AUDIT_FROM = '2025-02-01'          # 이 날 이후 확정한 공사만 맞춤(그 전 확정은 2025-01 이전 전표가 빠져 있음, 11-26)
SALES_TOL = 1                      # 공급가·부가세 1원 차이는 같음(ecount_sales 와 같은 규칙)
BUY_TOL = 10
KIND_RECON = '이카운트 대조'
BY_RECON = 'auto:recon'
COST_FIELDS = ('제품대', '도급비', '자재비', '기타비')
SALES_TYPE = '매출전표 I'
NOT_BUY_TYPES = (SALES_TYPE, '입금보고서')   # 계정 없이 읽으면 섞여 나옴(11-28 실제 읽기: 입금보고서 451줄, 코드 붙은 것 8줄)
CODE_RE = re.compile(r'^([GRN])(\d{4})-([A-Z]{2,3})$')
KEEP_FILES_DAYS = 14
MIN_SALES_ROWS = 1000              # 이보다 적게 읽혔으면 읽기가 잘못된 것(11-27 3,138장)
MIN_SHEET_ROWS = 3000              # 공사 현황 행 수 하한(11-28 약 4,300행)
DROP_RATIO = 0.97                  # 지난번보다 3% 넘게 줄었으면 실패로 봄(전표가 한꺼번에 사라질 일은 없음)


def _env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, '') or default)
    except ValueError:
        return default


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


_DDL = '''
    CREATE TABLE IF NOT EXISTS ecount_recon_runs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        started_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        finished_at TIMESTAMP,
        status TEXT NOT NULL DEFAULT 'running',
        read_file TEXT,
        sales_rows INTEGER,
        buy_rows INTEGER,
        summary TEXT,
        eco_mismatch TEXT,
        error TEXT
    )
'''


def _ensure(conn) -> None:
    conn.execute(_DDL)


# ── 읽기 ─────────────────────────────────────────────────────────────

def _node() -> Optional[str]:
    from dashboard.services.doc_viewer import _node_path
    return _node_path()


def run_reader(timeout: int = 45 * 60) -> Tuple[Optional[Path], str]:
    """이카운트 읽기 도구 실행 → (결과 JSON 경로, 실패 사유). 비밀번호는 도구가 .env 에서 직접 읽는다."""
    node = _node()
    if not node or not READER.exists():
        return None, f'읽기 도구 없음(node={node}, reader={READER.exists()})'
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    try:
        proc = subprocess.run([node, str(READER), '--out', str(OUT_DIR), '--from', READ_FROM],
                              cwd=str(ROOT), capture_output=True, timeout=timeout,
                              creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    except subprocess.TimeoutExpired:
        return None, f'읽기 도구가 {timeout // 60}분 안에 끝나지 않음'
    out = proc.stdout.decode('utf-8', 'replace')
    m = re.search(r'^RESULT (.+)$', out, re.M)
    path = Path(m.group(1).strip()) if m else None
    if proc.returncode != 0 or not path or not path.exists():
        tail = ' / '.join(l for l in out.strip().splitlines()[-4:] if l and not l.startswith('RESULT'))
        return path if path and path.exists() else None, f'읽기 실패(코드 {proc.returncode}): {tail[:300]}'
    return path, ''


def load_read(path: Path) -> dict:
    data = json.loads(Path(path).read_text(encoding='utf-8'))
    if not data.get('ok'):
        raise ValueError(f"이카운트 읽기 실패: {data.get('error') or '사유 없음'}")
    if not data.get('sales') or not data.get('buy'):
        raise ValueError('매출·매입 중 한쪽을 읽지 않은 결과예요.')
    return data


def _ymd(d: str) -> str:
    """'2026/11/27 -3' → '2026-11-27'"""
    m = re.match(r'(\d{4})/(\d{2})/(\d{2})', d or '')
    return f'{m.group(1)}-{m.group(2)}-{m.group(3)}' if m else ''


def aggregate(rows: List[dict], buy: bool = False) -> Dict[str, dict]:
    """전표 줄 → 공사 코드별 {s, v, n, first, vouchers:[(날짜, 공급가, 부가세, 거래처)], site}."""
    out: Dict[str, dict] = {}
    for r in rows:
        code = str(r.get('code') or '').strip()
        if not code:
            continue
        if buy and r.get('typ') in NOT_BUY_TYPES:     # 매출전표·입금보고서(받은 돈)는 매입이 아님
            continue
        a = out.setdefault(code, {'s': 0, 'v': 0, 'n': 0, 'first': '', 'vouchers': [], 'site': ''})
        s, v, d = int(r.get('s') or 0), int(r.get('v') or 0), _ymd(r.get('d') or '')
        a['s'] += s
        a['v'] += v
        a['n'] += 1
        a['first'] = min(a['first'] or d, d) if d else a['first']
        a['vouchers'].append((d, s, v, str(r.get('partner') or '')))
        a['site'] = a['site'] or str(r.get('site') or '')
    return out


# ── 시트 ─────────────────────────────────────────────────────────────

def _num(v) -> int:
    from dashboard.services.payer_split import _num as n
    return n(v)


def _confirm(r: dict) -> str:
    return str(r.get('공사 확정') or '').strip()[:10]


def _target(r: dict) -> Tuple[int, int]:
    from dashboard.services.ecount_sales import target_of
    return target_of(r) if _confirm(r) else (0, 0)


def _canceled(r: dict) -> bool:
    return not _confirm(r) and '취소' in str(r.get('수금 관련 특이사항') or '')


def _cost(r: dict) -> int:
    return sum(_num(r.get(k)) for k in COST_FIELDS)


def _same(a: Tuple[int, int], b: Tuple[int, int], tol: int = SALES_TOL) -> bool:
    return abs(a[0] - b[0]) <= tol and abs(a[1] - b[1]) <= tol


def _yymmdd(d: str) -> str:
    return f'{d[2:4]}.{d[5:7]}.{d[8:10]}' if len(d) >= 10 else d


def _vline(label: str, vouchers: List[tuple], with_partner: bool = False, limit: int = 12) -> str:
    """화면 근거 칸이 읽는 모양 — '이카운트 공급가 전표: 26.05.11 6,000,000 · …'(gw-finance-ecount.js faWhy)."""
    vs = sorted(vouchers)
    parts = [f'{_yymmdd(d)} {s:,}' + (f' {p}' if with_partner and p else '') for d, s, _v, p in vs[:limit]]
    more = f' · 외 {len(vs) - limit}장' if len(vs) > limit else ''
    return f'{label}: ' + ' · '.join(parts) + more if parts else ''


def _inv(r: dict) -> str:
    v = [str(r.get(k) or '-') for k in ('계약금 계산서', '중도금 계산서', '잔금 계산서')]
    return f'계산서: 계약금 {v[0]} · 중도금 {v[1]} · 잔금 {v[2]}'


def _twins(recs: Dict[str, dict]) -> Dict[str, List[str]]:
    by_num: Dict[str, List[str]] = defaultdict(list)
    for c in recs:
        m = CODE_RE.match(c)
        if m:
            by_num[m.group(2)].append(c)
    return by_num


def _twin_map(eco: Dict[str, dict], recs: Dict[str, dict], since: str) -> Tuple[Dict[str, List[str]], List[str]]:
    """공사 현황에 없는 이카운트 코드 → 같은 번호의 공사 현황 코드(하나일 때만). (twin_of, 짝 없는 코드)."""
    by_num = _twins(recs)
    twin_of: Dict[str, List[str]] = defaultdict(list)
    orphans = []
    for ec, a in eco.items():
        if ec in recs:
            continue
        m = CODE_RE.match(ec)
        if not m or (a['first'] or '') < since:
            continue
        if not a['s'] and not a['v']:      # 넣었다가 취소해 합이 0인 옛 코드(R4286-JSH 중복 등)는 짝이 아님
            continue
        cands = by_num.get(m.group(2), [])
        if len(cands) == 1:
            twin_of[cands[0]].append(ec)
        elif a['s'] or a['v']:
            orphans.append(ec)
    return twin_of, orphans


# ── 장부 맞춤 ─────────────────────────────────────────────────────────

def _in_scope(code: str, r: Optional[dict], eco: Dict[str, dict]) -> bool:
    """맞출 공사 — 2025-02 이후 확정, 또는 확정 날짜가 없는데(취소 등) 이카운트 전표가 모두 2025-02 이후."""
    if not r:
        return False
    c = _confirm(r)
    if c:
        return c >= AUDIT_FROM
    # 확정 날짜 없음(취소 등) — 원래 매출전표(+)부터 읽은 기간 안에 있어야 함. 2025년 전 공사의 취소 전표(−)만
    # 보이는 경우(G0621-JW 등, 11-28 시험)는 원래 전표가 안 읽혀 맞춰 볼 수 없다
    a = eco.get(code)
    return bool(a) and (a['first'] or '') >= AUDIT_FROM and any(v[1] > 0 for v in a['vouchers'])


def ledger_fixes(eco_sales: Dict[str, dict], recs: Dict[str, dict], ledger: Dict[str, Tuple[int, int]]) -> List[dict]:
    """이카운트 = 시트인데 장부만 다른 공사 → 장부를 이카운트 금액으로(차액 조정 줄)."""
    out = []
    for code, r in recs.items():
        if not _in_scope(code, r, eco_sales):
            continue
        a = eco_sales.get(code)
        eco = (a['s'], a['v']) if a else (0, 0)
        if not _same(eco, _target(r)):
            continue
        cur = ledger.get(code, (0, 0))
        if _same(cur, eco):
            continue
        out.append({'code': code, 'cur': cur, 'eco': eco, 'supply': eco[0] - cur[0], 'vat': eco[1] - cur[1]})
    return out


def _backup_ledger() -> Path:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUT_DIR / f'ledger_backup_{datetime.now():%Y%m%d_%H%M%S}.json'
    with _db()._get_connection() as conn:
        cur = conn.execute('SELECT * FROM ecount_sales_entries')
        names = [d[0] for d in cur.description]
        rows = [dict(zip(names, tuple(r))) for r in cur.fetchall()]
    path.write_text(json.dumps(rows, ensure_ascii=False, default=str), encoding='utf-8')
    return path


def apply_ledger_fixes(fixes: List[dict], stamp: str) -> int:
    from dashboard.services import ecount_sales as es
    n = 0
    for f in fixes:
        es_, ev = f['eco']
        if es.record_entry(f['code'], f['supply'], f['vat'], KIND_RECON, f"{f['code']}|auto-{stamp}|{es_},{ev}", BY_RECON,
                           note=f"{stamp} 자동 대조: 이카운트 = 공사 현황 {es_:,}/{ev:,} → 장부 {f['cur'][0]:,}/{f['cur'][1]:,} 에서 맞춤"):
            n += 1
    return n


# ── 매출 대조 ─────────────────────────────────────────────────────────

def _classify_sales(eco: Tuple[int, int], tgt: Tuple[int, int], vouchers: List[tuple]) -> str:
    groups: Dict[tuple, int] = defaultdict(int)
    for d, s, _v, _p in vouchers:
        if s > 0:
            groups[(d, s)] += 1
    if any(n > 1 for n in groups.values()) and eco[0] > tgt[0]:
        return 'dup'
    tot = (eco[0] + eco[1]) - (tgt[0] + tgt[1])
    if abs(tot) <= 2:
        return 'split'
    if abs(eco[0] - tgt[0]) <= SALES_TOL:
        if eco[1] and not tgt[1]:
            return 'vat_eco'
        if tgt[1] and not eco[1]:
            return 'vat_sheet'
        return 'split'
    if abs(tot) <= 10000:
        return 'round'
    return 'supply'


def sales_diffs(eco_sales: Dict[str, dict], recs: Dict[str, dict], ledger: Dict[str, Tuple[int, int]],
                watch: Optional[set] = None) -> Tuple[Dict[str, dict], Dict[str, List[int]]]:
    """(③ 대조 항목 {코드: 항목}, 이카운트 ≠ 장부인 공사 {코드: [공급가, 부가세]}).

    watch = 이미 ③에 올라 있거나 손 대조에 있던 공사 — '② 정상 대기'라도 빼지 않고 계속 본다."""
    watch = watch or set()
    twin_of, orphans = _twin_map(eco_sales, recs, AUDIT_FROM)
    found: Dict[str, dict] = {}
    mismatch: Dict[str, List[int]] = {}
    codes = {c for c, r in recs.items() if _in_scope(c, r, eco_sales)} | set(twin_of)
    for code in sorted(codes):
        r = recs[code]
        tgt = _target(r)
        ecs = [code] + twin_of.get(code, [])
        parts = [eco_sales[c] for c in ecs if c in eco_sales]
        eco = (sum(a['s'] for a in parts), sum(a['v'] for a in parts))
        vouchers = [v for a in parts for v in a['vouchers']]
        own = eco_sales.get(code)
        led = ledger.get(code, (0, 0))
        if own and own['n'] and not _same(led, (own['s'], own['v'])):   # 이카운트에 전표가 있는데 장부와 다름
            mismatch[code] = [own['s'], own['v']]
        twin = twin_of.get(code)
        if not twin and _same(eco, tgt):
            continue
        if not parts and tgt == (0, 0):
            continue
        pending_only = not twin and _same(led, eco)      # 이카운트 = 장부 ≠ 시트 → ② 목록의 넣을 전표
        if pending_only and code not in watch:
            continue
        notes = []
        if twin:
            cat = 'code'
            notes.append(f"이카운트에는 {', '.join(twin)}{'(으)로도' if own else '(으)로'} 들어가 있어요."
                         + ('' if _same(eco, tgt) else ' 금액도 달라요.'))
        else:
            cat = _classify_sales(eco, tgt, vouchers)
            if cat == 'dup':
                notes.append('같은 날 같은 금액 전표가 여러 장이에요. 한 장이 중복일 수 있어요.')
            if not parts:
                notes.append('이카운트에 이 코드 매출전표가 없어요.')
            elif _same(led, tgt) and not _same(led, eco):
                notes.append('장부에는 넣은 것으로 돼 있는데 이카운트 금액이 달라요.')
        vl = _vline('이카운트 공급가 전표', vouchers)
        if vl:
            notes.append(vl)
        if code.endswith('-SH'):
            notes.append('SH 정산 공사예요.')
        if pending_only:
            notes.append('이카운트 업로드 ② 매출전표 목록에도 있어요.')
        found[code] = {'category': cat, 'code': code, 'ecount_code': twin[0] if twin else code,
                       'site': str(r.get('현장 주소') or '')[:40], 'eco_supply': eco[0], 'eco_vat': eco[1],
                       'sheet_supply': tgt[0], 'sheet_vat': tgt[1], 'evidence': _inv(r), 'note': ' '.join(notes)}
    for ec in orphans:
        a = eco_sales[ec]
        found[ec] = {'category': 'code', 'code': ec, 'ecount_code': ec, 'site': a['site'][:40],
                     'eco_supply': a['s'], 'eco_vat': a['v'], 'sheet_supply': None, 'sheet_vat': None, 'evidence': '',
                     'note': '공사 현황에 이 코드가 없어요(같은 번호 공사도 하나로 정해지지 않아요). '
                             + _vline('이카운트 공급가 전표', a['vouchers'])}
    return found, mismatch


# ── 매입 대조 ─────────────────────────────────────────────────────────

def buy_diffs(eco_buy: Dict[str, dict], recs: Dict[str, dict], today: date,
              since: Optional[str] = None, grace_days: Optional[int] = None) -> Dict[str, dict]:
    """매입 대조 항목 {코드: 항목}. 매입계산서는 월말에 몰려 오므로 확정 뒤 grace_days(기본 45일)가 지난 공사만
    원가를 맞춰 본다. 코드 다름·취소 공사 매입은 바로."""
    since = since or os.getenv('ECOUNT_RECON_BUY_FROM', '').strip() or AUDIT_FROM
    grace = grace_days if grace_days is not None else _env_int('ECOUNT_RECON_BUY_GRACE_DAYS', 45)
    ripe = (today - timedelta(days=grace)).isoformat()
    twin_of, orphans = _twin_map(eco_buy, recs, since)
    found: Dict[str, dict] = {}

    def item(cat, code, r, eco_amt, vouchers, cost, note, ecount_code=None):
        ev = ('원가: ' + ' · '.join(f'{k} {_num(r.get(k)):,}' for k in COST_FIELDS)) if r else ''
        vl = _vline('이카운트 매입 전표', vouchers, with_partner=True)
        found[code] = {'category': cat, 'code': code, 'ecount_code': ecount_code or code,
                       'site': str((r or {}).get('현장 주소') or '')[:40] or (eco_buy.get(code) or {}).get('site', '')[:40],
                       'eco_supply': eco_amt, 'eco_vat': 0, 'sheet_supply': cost, 'sheet_vat': 0, 'evidence': ev,
                       'note': ' '.join(x for x in (note, vl, 'SH 정산 공사예요.' if code.endswith('-SH') else '') if x)}

    for code, r in recs.items():
        twin = twin_of.get(code)
        ecs = [code] + (twin or [])
        parts = [eco_buy[c] for c in ecs if c in eco_buy]
        eco = sum(a['s'] for a in parts)
        vouchers = [v for a in parts for v in a['vouchers']]
        conf = _confirm(r)
        if twin:
            item('buy_code', code, r, eco, vouchers, _cost(r) if conf else None,
                 f"이카운트 매입이 {', '.join(twin)}(으)로 들어가 있어요. 공사 현황 코드는 {code}예요.", ecount_code=twin[0])
            continue
        if _canceled(r):
            if eco and parts and min(a['first'] for a in parts) >= since:
                item('buy_cancel', code, r, eco, vouchers, None, '공사 현황에서 취소된 공사예요.')
            continue
        if not conf or conf < since or conf > ripe:
            continue
        cost = _cost(r)
        if abs(eco - cost) <= BUY_TOL:
            continue
        if not parts:
            item('buy_noeco', code, r, 0, [], cost, '')
        elif not cost:
            item('buy_nocost', code, r, eco, vouchers, 0, '')
        else:
            item('buy_diff', code, r, eco, vouchers, cost, f'차이 {eco - cost:+,}원.')
    for ec in orphans:
        a = eco_buy[ec]
        item('buy_code', ec, None, a['s'], a['vouchers'], None, '공사 현황에 이 코드가 없어요(같은 번호 공사도 하나로 정해지지 않아요).')
    return found


# ── 실행 기록 ─────────────────────────────────────────────────────────

def _start_run() -> int:
    with _db()._get_connection() as conn:
        _ensure(conn)
        cur = conn.execute("INSERT INTO ecount_recon_runs (status) VALUES ('running')")
        conn.commit()
        return cur.lastrowid


def _finish_run(run_id: int, status: str, **kw) -> None:
    cols = {k: (json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v) for k, v in kw.items()}
    sets = ', '.join(f'{k} = ?' for k in cols)
    with _db()._get_connection() as conn:
        _ensure(conn)
        conn.execute(f'UPDATE ecount_recon_runs SET status = ?, finished_at = CURRENT_TIMESTAMP'
                     f'{", " + sets if sets else ""} WHERE id = ?', (status, *cols.values(), run_id))
        conn.commit()


def last_run(ok_only: bool = False) -> Optional[dict]:
    import sqlite3
    with _db()._get_connection() as conn:
        try:      # 읽기만 — 표가 아직 없으면(첫 실행 전) None. 시험 실행이 운영 DB 에 표를 만들지 않게
            cur = conn.execute('SELECT * FROM ecount_recon_runs' + (" WHERE status = 'ok'" if ok_only else '')
                               + ' ORDER BY id DESC LIMIT 1')
        except sqlite3.OperationalError:
            return None
        row = cur.fetchone()
        if not row:
            return None
        d = dict(zip([c[0] for c in cur.description], tuple(row)))
    for k in ('summary', 'eco_mismatch'):
        try:
            d[k] = json.loads(d[k]) if d.get(k) else {}
        except ValueError:
            d[k] = {}
    return d


def run_info() -> Optional[dict]:
    """화면 머리말용 — 마지막 실행과 마지막 성공(한국 시각 'M/D HH:MM')."""
    last, ok = last_run(), last_run(ok_only=True)
    if not last:
        return None

    def kst(ts):
        try:
            d = datetime.strptime(str(ts)[:19], '%Y-%m-%d %H:%M:%S') + timedelta(hours=9)
        except ValueError:
            return ''
        return f'{d.month}/{d.day} {d:%H:%M}'
    return {'status': last['status'], 'at': kst(last.get('finished_at') or last['started_at']),
            'error': last.get('error') or '', 'ok_at': kst(ok['finished_at']) if ok else '',
            'sales_rows': (ok or {}).get('sales_rows'), 'buy_rows': (ok or {}).get('buy_rows')}


def eco_mismatch() -> Dict[str, List[int]]:
    """마지막 성공 대조에서 이카운트 ≠ 장부였던 공사 — ② 목록에 '이카운트에 이미 있음' 경고(ecount_sales.upload_rows)."""
    r = last_run(ok_only=True)
    if not r or not r.get('finished_at'):
        return {}
    try:   # 이틀 넘은 결과는 쓰지 않음(대조가 멈췄으면 경고도 낡음)
        if datetime.utcnow() - datetime.strptime(str(r['finished_at'])[:19], '%Y-%m-%d %H:%M:%S') > timedelta(hours=36):
            return {}
    except ValueError:
        return {}
    return r.get('eco_mismatch') or {}


def _cleanup_files() -> None:
    if not OUT_DIR.exists():
        return
    cut = time.time() - KEEP_FILES_DAYS * 86400
    for p in OUT_DIR.iterdir():
        try:
            if p.is_file() and p.stat().st_mtime < cut:
                p.unlink()
        except OSError:
            pass


class _Lock:
    """두 번 겹쳐 돌지 않게(스케줄러 + 손 실행). 2시간 넘은 잠금은 버림."""
    def __init__(self):
        self.path = OUT_DIR / 'recon.lock'

    def __enter__(self):
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        try:
            if self.path.exists() and time.time() - self.path.stat().st_mtime > 7200:
                self.path.unlink()
            fd = os.open(str(self.path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
        except FileExistsError:
            raise RuntimeError('이카운트 대조가 이미 돌고 있어요.')
        return self

    def __exit__(self, *a):
        try:
            self.path.unlink()
        except OSError:
            pass


# ── 밤 실행 ───────────────────────────────────────────────────────────

def _watch_codes(ensure: bool) -> set:
    """이미 ③에 올라 있거나(자동 대조, '맞음' 빼고) 손 대조에서 '이카운트 고침' 말고 다른 처리를 한 매출 공사."""
    import sqlite3
    from dashboard.services import ecount_audit as ea
    with _db()._get_connection() as conn:
        if ensure:
            ea._ensure(conn)
        try:
            rows = conn.execute(
                "SELECT code FROM ecount_audit_items WHERE side = 'sales' AND "
                "((batch = ? AND status != 'resolved') OR (batch != ? AND status != 'ecount_fixed'))",
                (ea.AUTO_BATCH, ea.AUTO_BATCH)).fetchall()
        except sqlite3.OperationalError:          # 시험 실행 — side 칸이 아직 없는 운영 DB
            try:
                rows = conn.execute("SELECT code FROM ecount_audit_items WHERE status != 'ecount_fixed'").fetchall()
            except sqlite3.OperationalError:
                rows = []
    return {r[0] for r in rows}


def run_nightly(read_path: Optional[str] = None, apply: bool = True, today: Optional[date] = None) -> dict:
    """스케줄러 진입점(매일 00:10). read_path 가 있으면 읽기를 건너뛰고 그 결과로 대조(손 실행·시험).
    apply=False 면 DB 를 바꾸지 않고 결과만 돌려준다."""
    from dashboard.services import ecount_audit as ea
    from dashboard.services import ecount_sales as es
    from dashboard.services.project_service import get_project_records
    today = today or date.today()
    md = f'{today.month}/{today.day}'
    import contextlib
    with (_Lock() if apply else contextlib.nullcontext()):
        run_id = _start_run() if apply else 0
        try:
            if read_path:
                path, err = Path(read_path), ''
            else:
                path, err = run_reader()
            if err or not path:
                raise ValueError(err or '읽기 결과 없음')
            data = load_read(path)
            sales_rows, buy_rows = data['sales']['rows'], data['buy']['rows']
            prev = last_run(ok_only=True)
            if len(sales_rows) < MIN_SALES_ROWS:
                raise ValueError(f'매출 전표가 {len(sales_rows)}줄뿐이에요(읽기 이상).')
            if prev and prev.get('sales_rows') and len(sales_rows) < prev['sales_rows'] * DROP_RATIO:
                raise ValueError(f"매출 전표가 지난번 {prev['sales_rows']}줄 → {len(sales_rows)}줄로 줄었어요(읽기 이상).")
            if prev and prev.get('buy_rows') and len(buy_rows) < prev['buy_rows'] * DROP_RATIO:
                raise ValueError(f"매입 전표가 지난번 {prev['buy_rows']}줄 → {len(buy_rows)}줄로 줄었어요(읽기 이상).")
            eco_sales, eco_buy = aggregate(sales_rows), aggregate(buy_rows, buy=True)
            recs_list = get_project_records(force_refresh=True)
            # 시트를 못 읽으면 로컬 백업(몇 줄)으로 조용히 바뀐다(11-28 시험에서 1행) — 그대로 맞추면 전부 '다름'
            if not recs_list or len(recs_list) < MIN_SHEET_ROWS:
                raise ValueError(f'공사 현황이 {len(recs_list or [])}줄뿐이에요(시트 읽기 이상).')
            if prev and (prev.get('summary') or {}).get('sheet_rows') and                     len(recs_list) < prev['summary']['sheet_rows'] * DROP_RATIO:
                raise ValueError(f"공사 현황이 지난번 {prev['summary']['sheet_rows']}줄 → {len(recs_list)}줄로 줄었어요.")
            recs = {str(r.get('프로젝트 코드') or '').strip(): r for r in recs_list if str(r.get('프로젝트 코드') or '').strip()}

            # 1) 장부 맞춤 — 이카운트 = 시트일 때만
            ledger = es.ledger_totals()
            fixes = ledger_fixes(eco_sales, recs, ledger)
            cap = _env_int('ECOUNT_RECON_MAX_LEDGER_FIXES', 150)
            fixed, backup, held = 0, '', False
            if fixes and apply:
                if len(fixes) > cap:
                    held = True
                    logger.error(f'[ECOUNT_RECON] 장부 맞출 공사가 {len(fixes)}건으로 한도 {cap}건을 넘어 장부는 그대로 둠 — 확인 필요',
                                 extra={'force_alert': True})
                else:
                    backup = str(_backup_ledger())
                    fixed = apply_ledger_fixes(fixes, today.strftime('%Y%m%d'))
                    ledger = es.ledger_totals()
            elif fixes:
                for f in fixes:      # 시험 실행 — 맞춘 것처럼 계산만
                    ledger[f['code']] = f['eco']

            # 2) 대조 — 이미 보고 있는 공사는 ② 대기라도 계속
            watch = _watch_codes(ensure=apply)
            s_found, mismatch = sales_diffs(eco_sales, recs, ledger, watch)
            b_found = buy_diffs(eco_buy, recs, today)
            # 매입은 SB 가 이카운트를 고칠 것(코드·취소·없음)과 시트 원가 확인(빈칸·다름)으로 나눔(11-28 부장님)
            c_found = {c: i for c, i in b_found.items() if i['category'] in ea.COST_CATEGORIES}
            b_found = {c: i for c, i in b_found.items() if i['category'] not in ea.COST_CATEGORIES}
            if apply:
                s_res = ea.apply_nightly('sales', s_found, md)
                b_res = ea.apply_nightly('buy', b_found, md)
                c_res = ea.apply_nightly('cost', c_found, md)
            else:
                s_res = b_res = c_res = {}
            summary = {
                'md': md, 'upto': upto_label(today), 'read_file': str(path), 'ledger_fixes': len(fixes), 'ledger_fixed': fixed,
                'ledger_held': held, 'ledger_backup': backup, 'ledger_fix_codes': [f['code'] for f in fixes][:200],
                'sales': {'found': len(s_found), **{k: v for k, v in s_res.items()}},
                'buy': {'found': len(b_found), **{k: v for k, v in b_res.items()}},
                'cost': {'found': len(c_found), **{k: v for k, v in c_res.items()}},
                'eco_codes': len(eco_sales), 'buy_codes': len(eco_buy), 'sheet_rows': len(recs_list),
            }
            if apply:
                _finish_run(run_id, 'ok', read_file=str(path), sales_rows=len(sales_rows), buy_rows=len(buy_rows),
                            summary=summary, eco_mismatch=mismatch)
                _cleanup_files()
            logger.info(f"[ECOUNT_RECON] 대조 끝 — 장부 맞춤 {fixed}/{len(fixes)}, 매출 차이 {len(s_found)}, 매입 차이 {len(b_found)}")
            return {'ok': True, 'summary': summary, 'sales_found': s_found, 'buy_found': b_found, 'cost_found': c_found,
                    'mismatch': mismatch,
                    'fixes': fixes}
        except Exception as exc:
            if apply:
                _finish_run(run_id, 'failed', error=str(exc)[:500])
            logger.error(f'[ECOUNT_RECON] 이카운트 밤 대조 실패: {exc}', exc_info=True, extra={'force_alert': True})
            return {'ok': False, 'error': str(exc)}


# ── 아침 DM(경영지원) ───────────────────────────────────────────────────

DM_LIST_MAX = 12


def _won(n) -> str:
    return f'{int(n or 0):,}원'


def _dm_line(r: dict) -> str:
    """한 줄 요약 — 코드 · 무엇이 다른지."""
    side = r.get('side') or 'sales'
    es_, ev, ss, sv = (int(r.get(k) or 0) for k in ('eco_supply', 'eco_vat', 'sheet_supply', 'sheet_vat'))
    if r.get('ecount_code') and r['ecount_code'] != r['code']:
        what = f"이카운트 코드 {r['ecount_code']}"
    elif r.get('category') == 'buy_cancel':
        what = f'취소된 공사에 이카운트 매입 {_won(es_)} 남음'
    elif r.get('category') == 'buy_noeco':
        what = f'이카운트 매입 없음 · 공사 현황 원가 {_won(ss)}'
    elif side != 'sales':
        what = f'이카운트 매입 {_won(es_)} · 공사 현황 원가 {_won(ss)}'
    else:
        what = f'이카운트 {_won(es_ + ev)} · 공사 현황 {_won(ss + sv)}'
    return f"• `{r['code']}` {what}"


def build_dm(run: Optional[dict], upload: Dict[str, int], sales_open: List[dict], buy_open: List[dict]) -> Tuple[str, List[dict]]:
    """(알림 글, 블록). upload = {'projects': ①, 'sales': ②}.

    고칠 것 = ①·② 건수 + ③ 매출·매입(이카운트를 고칠 것) 처리 전. 시트 원가 확인(원가 빈칸·원가 다름)은 SB 할 일이
    아니라 넣지 않는다(11-28 부장님 'SB 할 일은 30건만')."""
    base = ((os.getenv('GW_PUBLIC_URL') or '').strip() or 'https://pm.itg-aircon.com').rstrip('/')
    link = f'{base}/finance/ecount'
    if not run or run.get('status') != 'ok':
        text = ('오늘 아침 이카운트 대조는 이카운트를 읽지 못해 쉬어요. '
                '이카운트 업로드 화면의 ① 프로젝트 등록·② 매출전표 목록은 그대로 쓸 수 있어요.')
        return text, [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': f':warning: {text}\n<{link}|이카운트 업로드 열기>'}}]
    summ = run.get('summary') or {}
    resolved = len((summ.get('sales') or {}).get('resolved') or []) + len((summ.get('buy') or {}).get('resolved') or [])
    total = upload.get('projects', 0) + upload.get('sales', 0) + len(sales_open) + len(buy_open)
    head = (f":ledger: *이카운트 대조 · {summ['upto']}까지 기준*" if summ.get('upto')
            else f":ledger: *이카운트 대조 · {summ.get('md', '')} 자정 기준*")
    tail = []
    if resolved:
        tail.append(f'어제 알려 드린 것 중 {resolved}건은 이카운트와 공사 현황이 같아져 완료로 바꿨어요.')
    tail.append(f'<{link}|이카운트 업로드 열기>')
    if not total:
        text = '수정할 것 없음 — 이카운트·공사 현황·장부가 모두 맞아요.'
        return text, [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': f'{head}\n{text}'}},
                      {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': ' '.join(tail)}]}]
    lines = [head, f'고칠 것이 {total}건 있어요.']
    if upload.get('projects') or upload.get('sales'):
        lines.append(f"• 이카운트 업로드 — ① 프로젝트 등록 {upload.get('projects', 0)}건 · ② 매출전표 {upload.get('sales', 0)}건")
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': '\n'.join(lines)}}]
    for label, rows in (('매출', sales_open), ('매입', buy_open)):
        if not rows:
            continue
        fresh = [r for r in rows if r.get('_fresh')]
        show = (fresh + [r for r in rows if not r.get('_fresh')])[:DM_LIST_MAX]
        t = [f'*③ 이카운트 대조 · {label} {len(rows)}건*' + (f' (새로 생긴 것 {len(fresh)}건)' if fresh else '')]
        t += [_dm_line(r) for r in show]
        if len(rows) > len(show):
            t.append(f'그 밖에 {len(rows) - len(show)}건은 그룹웨어에서 볼 수 있어요.')
        blocks.append({'type': 'section', 'text': {'type': 'mrkdwn', 'text': '\n'.join(t)[:2900]}})
    blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': ' '.join(tail)}]})
    return f'이카운트 대조 — 고칠 것 {total}건', blocks


def upload_counts() -> Dict[str, int]:
    """① 프로젝트 등록·② 매출전표 건수 — 화면과 같은 기준(오늘 확정·변경분은 내일로)."""
    from dashboard.services import ecount_sales as es
    items = es.collect_items(force_refresh=True) or []
    shown, _later = es.split_today(items)
    rows = es.upload_rows(shown)
    return {'projects': len(rows['projects']), 'sales': len(rows['sales'])}


def _is_bd(d: date) -> bool:
    try:
        from dashboard.services.absent_remind import _is_business_day
        return _is_business_day(d)
    except Exception:
        return d.weekday() < 5


def prev_business_day(today: date) -> date:
    d = today - timedelta(days=1)
    for _ in range(14):
        if _is_bd(d):
            return d
        d -= timedelta(days=1)
    return today - timedelta(days=1)


def fresh_for(run: Optional[dict], today: date) -> bool:
    """오늘 아침 DM 에 쓸 대조인지 — 직전 영업일 다음 날 이후(한국 날짜)에 돈 것.
    주말엔 대조를 안 돌려(화~토 00:10, 11-28 부장님) 월요일엔 토요일 00:10(금요일 밤) 대조를 쓴다."""
    if not run or not run.get('started_at'):
        return False
    try:
        ran = (datetime.strptime(str(run['started_at'])[:19], '%Y-%m-%d %H:%M:%S') + timedelta(hours=9)).date()
    except ValueError:
        return False
    return prev_business_day(today) < ran <= today


_WD = '월화수목금토일'


def upto_label(run_day: date) -> str:
    """자정 대조가 담은 마지막 날 — 00:10 에 돌면 전날까지. 'M/D(요일)'"""
    d = run_day - timedelta(days=1)
    return f'{d.month}/{d.day}({_WD[d.weekday()]})'


def send_morning_dm(today: Optional[date] = None) -> dict:
    """경영지원(SB) 아침 DM(스케줄러 07:50, 영업일). 고칠 것 목록 또는 '수정할 것 없음'."""
    from dashboard.services import ecount_audit as ea
    today = today or date.today()
    try:
        from dashboard.services.absent_remind import _is_business_day
        if not _is_business_day(today):
            return {'ok': True, 'reason': 'non_business_day'}
    except Exception:
        pass
    run = last_run()
    if not fresh_for(run, today):
        run = None
    if run and run.get('status') == 'ok':
        summ = run.get('summary') or {}
        sales_open, buy_open = ea.open_items('sales'), ea.open_items('buy')
        for side, rows in (('sales', sales_open), ('buy', buy_open)):
            s = summ.get(side) or {}
            newset = set() if s.get('baseline') else set((s.get('new') or []) + (s.get('reopened') or []))
            for r in rows:
                r['_fresh'] = r['code'] in newset
        try:
            upload = upload_counts()
        except Exception as exc:
            logger.warning(f'[ECOUNT_RECON] ①·② 건수 읽기 실패(DM 은 보냄): {exc}')
            upload = {}
    else:
        sales_open, buy_open, upload = [], [], {}
        logger.error('[ECOUNT_RECON] 오늘 자정 이카운트 대조 결과가 없어 SB 에게 쉬는 안내만 보냄', extra={'force_alert': True})
    text, blocks = build_dm(run, upload, sales_open, buy_open)
    from dashboard.services.pin_remind import open_todo_dm
    c, ch = open_todo_dm()
    if not c or not ch:
        return {'ok': False, 'reason': 'dm_open_failed'}
    try:
        r = c.chat_postMessage(channel=ch, text=text, blocks=blocks, unfurl_links=False, unfurl_media=False)
        logger.info(f'[ECOUNT_RECON] 경영지원 아침 DM 발송: {text}')
        return {'ok': bool(r.get('ok')), 'text': text}
    except Exception as exc:
        logger.error(f'[ECOUNT_RECON] 경영지원 아침 DM 실패: {exc}', exc_info=True)
        return {'ok': False, 'reason': str(exc)}
