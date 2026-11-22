"""차량 관리 (2026-11-22 부장님) — 차량별 보험·검사 만기, 정비·소모품·유류·주차·통행 비용, 계기판 km.

U+커넥트에도 정비비·유류비·하이패스·과태료·보험료 화면이 있지만 직접 입력하는 칸이라(접속마다 문자 인증)
아무도 안 써서 비어 있다(11-22 확인). 그래서 사람이 따로 입력하지 않아도 되게:
- 비용은 이미 올리는 결재에서 자동으로 모은다 — 플렉스 보관함(법인카드 지출결의·개인경비·자금집행·비품)과
  슬랙 전자결재(approvals). 내역에 차량 번호(끝 4자리 포함)나 한 대뿐인 차종 이름(카니발·스파크 등)이 있는 줄만.
- 같은 결재 줄은 한 번만 들어간다(source_key). 한 결재에 차가 여러 대면 금액을 나눠 적고 '나눔'으로 표시.
- 보험 만기는 공유 드라이브 '★법인 차량관리폴더' 증권 파일 이름의 날짜(가장 늦은 것)로 처음 채웠다(11-22).
  레이 3대는 롯데렌탈 장기대여 — 자동차공제 증명서 기간으로 채움.
- 파나메라·벤츠·랜드로버는 U+ 단말이 없어 계기판 km 를 사람이 넣는다(odometer).
"""
from __future__ import annotations

import json
import re
import time
from contextlib import contextmanager
from datetime import date, datetime
from typing import Dict, Iterable, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

KINDS = ('정비', '유류', '주차', '통행료', '세차', '검사', '보험', '과태료', '소모품', '기타')

# (번호판) → (소유, 보험사, 보험 만기, 메모) — 11-22 증권·증명서 기준
_INFO_SEED = {
    '237허5358': ('렌트(롯데렌탈)', '자동차공제', '2027-04-01', '장기대여 — 공제 증명서 기준'),
    '231허6854': ('렌트(롯데렌탈)', '자동차공제', '2026-10-22', '장기대여 — 갱신 증명서를 받아야 함(폴더엔 10-22 만기분만)'),
    '231허6873': ('렌트(롯데렌탈)', '자동차공제', '2026-10-22', '장기대여 — 갱신 증명서를 받아야 함(폴더엔 10-22 만기분만)'),
    '03모8021': ('법인 소유', '현대해상', '2027-03-07', ''),
    '16소4041': ('법인 소유', '현대해상', '2027-02-02', ''),
    '817루4904': ('법인 소유', 'KB손해보험', '2027-04-14', ''),
    '821저4919': ('법인 소유', '현대해상', '2027-01-03', ''),
    '90머9131': ('법인 소유', '현대해상', '2027-07-25', ''),
    '53우1783': ('법인 소유', '', '', '증권이 폴더에 없음'),
    '60거5484': ('법인 소유', 'DB손해보험', '2027-05-27', ''),
    '143고9687': ('법인 소유', 'DB손해보험', '2027-05-18', ''),
    '204두8988': ('법인 소유', '삼성화재', '2027-09-10', ''),
    '260나7350': ('법인 소유(그룹)', 'KB손해보험', '2027-07-31', ''),
    '133가6633': ('법인 소유(플랜트)', 'DB손해보험', '2027-11-29', ''),
}

_DDL = [
    '''CREATE TABLE IF NOT EXISTS vehicle_info (
        plate TEXT PRIMARY KEY,
        ownership TEXT,
        insurer TEXT,
        insurance_end TEXT,
        inspection_end TEXT,
        note TEXT,
        updated_by TEXT,
        updated_at TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS vehicle_costs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        plate TEXT NOT NULL,
        day TEXT NOT NULL,
        kind TEXT NOT NULL,
        item TEXT,
        amount INTEGER NOT NULL DEFAULT 0,
        shared INTEGER NOT NULL DEFAULT 1,
        vendor TEXT,
        source TEXT NOT NULL,
        source_key TEXT NOT NULL UNIQUE,
        created_by TEXT,
        created_at TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_vehicle_costs_plate_day ON vehicle_costs(plate, day)',
    '''CREATE TABLE IF NOT EXISTS vehicle_odometer (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        plate TEXT NOT NULL,
        day TEXT NOT NULL,
        km REAL NOT NULL,
        note TEXT,
        created_by TEXT,
        created_at TEXT,
        UNIQUE (plate, day)
    )''',
]
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


@contextmanager
def _conn():
    global _table_ready
    with _db()._get_connection() as conn:
        if not _table_ready:
            for ddl in _DDL:
                conn.execute(ddl)
            for plate, (own, ins, end, note) in _INFO_SEED.items():
                conn.execute('INSERT OR IGNORE INTO vehicle_info (plate, ownership, insurer, insurance_end, inspection_end, '
                             'note, updated_by, updated_at) VALUES (?,?,?,?,?,?,?,?)',
                             (plate, own, ins, end, '', note, 'seed', _now()))
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _fleet() -> List[Dict]:
    from dashboard.services import vehicle_assign as va
    return va.fleet(True)


# ── 차량 번호 찾기 ───────────────────────────────────────────────────────
_UNIQUE_MODELS = {   # 한 대뿐인 차종 — 이름만 적어도 그 차
    '카니발': '143고9687', '스파크': '53우1783', '모닝': '60거5484', '파나메라': '204두8988',
    '벤츠': '260나7350', '랜드로버': '133가6633', '디펜더': '133가6633', '전기트럭': '90머9131',
}


def find_plates(text: str, fleet: Optional[List[Dict]] = None) -> List[str]:
    """글에서 차를 찾는다 — 끝 4자리('8021', '817루 4904', '81루 4904') 또는 한 대뿐인 차종 이름."""
    s = str(text or '')
    fleet = fleet if fleet is not None else _fleet()
    found: List[str] = []
    for c in fleet:
        last4 = c['plate'][-4:]
        # 앞뒤가 숫자가 아닌 4자리만(금액 '18,021'·카드 끝자리 같은 다른 숫자에 안 걸리게)
        if re.search(r'(?<![\d,])' + last4 + r'(?![\d,])', s) and c['plate'] not in found:
            found.append(c['plate'])
    for word, plate in _UNIQUE_MODELS.items():
        if word in s and plate not in found and any(c['plate'] == plate for c in fleet):
            found.append(plate)
    return found


_KIND_WORDS = [   # 위에서부터 먼저 맞는 것(11-22 실제 결재 116줄로 맞춤)
    ('과태료', ('과태료', '범칙금', '주정차위반', '주정차단속', '단속')),
    ('보험', ('보험',)),
    ('주차', ('주차',)),
    ('통행료', ('통행료', '톨게이트', '하이패스', '고속도로')),
    ('세차', ('세차',)),
    ('검사', ('정기검사', '자동차검사', '검사')),
    ('정비', ('정비', '수리', '타이어', '오일', '부품', '배터리', '베터리', '브레이크', '전조등', '콤프레샤', '점검',
              '공임', '판금', '도색', '얼라인먼트', '센서', '유리', '미션', '주유구', '엔진')),
    ('소모품', ('워셔', '와이퍼', '방향제', '블랙박스', '매트', '소모품')),
    ('유류', ('주유', '충전', '기름', 'LPG', '경유', '휘발유', '요소수')),
]
_REPAIR_SHOPS = ('하연', '타이어', '모터스', '카센터', '공업사', '정비', '부품', '자동차')   # 상호만 보고 정비로


def kind_of(text: str, vendor: str = '') -> str:
    s = str(text or '')
    for kind, words in _KIND_WORDS:
        if any(w in s for w in words):
            return kind
    if any(w in str(vendor or '') for w in _REPAIR_SHOPS) or '교체' in s:
        return '정비'
    return '기타'


def _amount(v) -> int:
    s = re.sub(r'[^\d]', '', str(v or ''))
    return int(s) if s else 0


# ── 결재에서 비용 모으기 ─────────────────────────────────────────────────
# 법인카드 지출결의 메모(표가 한 줄로 붙어 있음): '…12026.10.06SK쉴더스_주차료 ₩ 3,0001783 주차비52026.10.08…'
#   줄 번호 + 사용일자(2026.10.06 / 26.10.06 / 25.12.04.) + 가맹점 + ₩ 금액 + 사용내역
_CARD_ROW = re.compile(
    r'(?P<no>\d{1,3}?)(?P<date>(?:20)?\d{2}\.\d{1,2}\.\d{1,2})\.?\s*(?P<shop>[^₩]{0,60}?)\s*₩\s*(?P<amt>[\d,]+?)'
    r'(?P<desc>(?:(?!\d{1,3}(?:20)?\d{2}\.\d{1,2}\.\d{1,2}).)*?)'
    r'(?=\d{1,3}(?:20)?\d{2}\.\d{1,2}\.\d{1,2}|합\s*계|$)', re.S)


def _card_day(raw: str) -> str:
    y, m, d = raw.split('.')[:3]
    y = int(y) + (2000 if len(y) == 2 else 0)
    try:
        return date(y, int(m), int(d)).isoformat()
    except ValueError:
        return ''


def _split_amount(amt: str, desc: str) -> Tuple[int, str]:
    """'3,0001783 주차비' → 금액 3,000 / 내역 '1783 주차비'. 금액 뒤에 붙은 차 번호를 떼 낸다."""
    m = re.match(r'^(\d{1,3}(?:,\d{3})*)(.*)$', amt + desc, re.S)
    if not m:
        return _amount(amt), desc
    return _amount(m.group(1)), m.group(2).strip()


def parse_card_memo(memo: str, fleet: List[Dict]) -> List[Dict]:
    """법인카드 지출결의 메모 → 차량 줄만 [{day, shop, amount, desc, plates}]."""
    out = []
    text = str(memo or '').replace('\n', ' ')
    start = text.find('내역(상세기재)')
    text = text[start + 7:] if start >= 0 else text
    for m in _CARD_ROW.finditer(text):
        day = _card_day(m.group('date'))
        if not day:
            continue
        amount, desc = _split_amount(m.group('amt'), m.group('desc'))
        plates = find_plates(desc, fleet)
        if not plates:
            continue
        out.append({'day': day, 'shop': m.group('shop').strip()[:40], 'amount': amount, 'desc': desc.strip()[:80],
                    'plates': plates})
    return out


def _cost_rows_from_flex(row: Dict, fleet: List[Dict]) -> List[Dict]:
    form = row['form'] or ''
    try:
        f = json.loads(row['fields'] or '{}')
    except ValueError:
        f = {}
    created = (row['created_at'] or '')[:10]
    rows = []
    if form == '법인카드 지출결의':
        for i, r in enumerate(parse_card_memo(row['memo'], fleet)):
            rows.append({'day': r['day'], 'item': r['desc'], 'amount': r['amount'], 'vendor': r['shop'],
                         'plates': r['plates'], 'key': f"flex:{row['doc_no']}:{i}"})
    elif form == '개인경비청구서':
        desc = f.get('사유') or ''
        plates = find_plates(desc, fleet)
        if plates:
            rows.append({'day': (f.get('사용일자') or created)[:10], 'item': desc, 'amount': _amount(f.get('청구 금액')),
                         'vendor': '', 'plates': plates, 'key': f"flex:{row['doc_no']}"})
    elif form.startswith('자금집행요청'):
        desc = f.get('내용') or ''
        plates = find_plates(desc, fleet)
        if plates:
            rows.append({'day': created, 'item': desc, 'amount': _amount(f.get('총 금액(VAT포함)') or row.get('amount')),
                         'vendor': f.get('상호명') or '', 'plates': plates, 'key': f"flex:{row['doc_no']}"})
    elif form == '비품 신청':
        desc = f"{f.get('사유') or ''} {f.get('품명') or ''}".strip()
        plates = find_plates(desc, fleet)
        if plates:
            rows.append({'day': created, 'item': desc, 'amount': _amount(f.get('금액(VAT포함)')), 'vendor': '',
                         'plates': plates, 'key': f"flex:{row['doc_no']}", 'kind': '소모품'})
    return rows


def _cost_rows_from_approval(row: Dict, fleet: List[Dict]) -> List[Dict]:
    desc = row['content'] or ''
    plates = find_plates(desc, fleet)
    if not plates:
        return []
    kind = '보험' if row['form'] == 'insurance' else None
    return [{'day': (row['created_at'] or '')[:10], 'item': desc, 'amount': int(row['amount'] or 0),
             'vendor': row['partner'] or '', 'plates': plates, 'key': f"appr:{row['id']}", 'kind': kind}]


def collect_from_approvals() -> List[Dict]:
    """결재 두 곳(플렉스 보관함·슬랙 전자결재)에서 차량 비용 줄을 모은다. 완료·승인된 결재만."""
    fleet = _fleet()
    out: List[Dict] = []
    with _db()._get_connection() as conn:
        try:
            for r in conn.execute("SELECT doc_no, form, status, created_at, memo, fields, amount FROM flex_approval_archive "
                                  "WHERE status = '완료'").fetchall():
                out += _cost_rows_from_flex(dict(r), fleet)
        except Exception as e:      # 표가 없는 환경(시험 등)
            logger.debug(f'[차량관리] 플렉스 보관함 읽기 건너뜀: {e}')
        try:
            for r in conn.execute("SELECT id, form, status, content, amount, partner, created_at FROM approvals "
                                  "WHERE status IN ('done', 'approved')").fetchall():
                out += _cost_rows_from_approval(dict(r), fleet)
        except Exception as e:
            logger.debug(f'[차량관리] 전자결재 읽기 건너뜀: {e}')
    return out


_last_sync = 0.0


def sync_costs(force: bool = False, min_gap: int = 600) -> int:
    """결재에서 모은 비용을 장부에 넣는다(이미 있는 줄은 건너뜀). 10분에 한 번만. 새로 넣은 줄 수."""
    global _last_sync
    if not force and time.time() - _last_sync < min_gap:
        return 0
    _last_sync = time.time()
    added = 0
    now = _now()
    with _conn() as conn:
        for r in collect_from_approvals():
            n = len(r['plates'])
            for p in r['plates']:
                amount = round(r['amount'] / n) if n > 1 else r['amount']
                cur = conn.execute(
                    'INSERT OR IGNORE INTO vehicle_costs (plate, day, kind, item, amount, shared, vendor, source, source_key, '
                    'created_by, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)',
                    (p, r['day'], r.get('kind') or kind_of(r['item'], r['vendor']), r['item'][:120], amount, n, r['vendor'][:40],
                     'approval', f"{r['key']}:{p}", 'sync', now))
                added += cur.rowcount
    if added:
        logger.info(f'[차량관리] 결재에서 차량 비용 {added}줄 새로 넣음')
    return added


# ── 직접 입력 ────────────────────────────────────────────────────────────
def add_cost(plate: str, day: str, kind: str, item: str, amount, vendor: str, by: str) -> Dict:
    if plate not in {c['plate'] for c in _fleet()}:
        raise ValueError('없는 차량이에요.')
    try:
        date.fromisoformat(day)
    except ValueError:
        raise ValueError('날짜가 올바르지 않아요.')
    if kind not in KINDS:
        raise ValueError('종류를 골라 주세요.')
    amt = _amount(amount)
    item = (item or '').strip()[:120]
    if not item:
        raise ValueError('내용을 적어 주세요.')
    with _conn() as conn:
        cur = conn.execute(
            'INSERT INTO vehicle_costs (plate, day, kind, item, amount, shared, vendor, source, source_key, created_by, created_at) '
            'VALUES (?,?,?,?,?,1,?,?,?,?,?)',
            (plate, day, kind, item, amt, (vendor or '').strip()[:40], 'manual', f'manual:{time.time_ns()}', by, _now()))
        return {'id': cur.lastrowid}


def delete_cost(cost_id: int) -> None:
    """직접 넣은 줄만 지운다(결재에서 온 줄은 결재가 근거라 지우지 않음)."""
    with _conn() as conn:
        r = conn.execute('SELECT source FROM vehicle_costs WHERE id = ?', (cost_id,)).fetchone()
        if not r:
            raise ValueError('없는 줄이에요.')
        if r['source'] != 'manual':
            raise ValueError('결재에서 온 줄은 지울 수 없어요.')
        conn.execute('DELETE FROM vehicle_costs WHERE id = ?', (cost_id,))


def update_info(plate: str, values: Dict, by: str) -> Dict:
    allowed = {}
    for k in ('ownership', 'insurer', 'note'):
        if k in values:
            allowed[k] = (values.get(k) or '').strip()[:60]
    for k in ('insurance_end', 'inspection_end'):
        if k in values:
            v = (values.get(k) or '').strip()
            if v:
                try:
                    date.fromisoformat(v)
                except ValueError:
                    raise ValueError('날짜가 올바르지 않아요.')
            allowed[k] = v
    if not allowed:
        raise ValueError('고칠 내용이 없어요.')
    if plate not in {c['plate'] for c in _fleet()}:
        raise ValueError('없는 차량이에요.')
    with _conn() as conn:
        conn.execute('INSERT OR IGNORE INTO vehicle_info (plate) VALUES (?)', (plate,))
        sets = ', '.join(f'{k} = ?' for k in allowed)
        conn.execute(f'UPDATE vehicle_info SET {sets}, updated_by = ?, updated_at = ? WHERE plate = ?',
                     (*allowed.values(), by, _now(), plate))
    return info(plate)


def add_odometer(plate: str, day: str, km, by: str, note: str = '') -> Dict:
    if plate not in {c['plate'] for c in _fleet()}:
        raise ValueError('없는 차량이에요.')
    try:
        date.fromisoformat(day)
        km = float(str(km).replace(',', ''))
    except ValueError:
        raise ValueError('날짜·km 를 확인해 주세요.')
    prev = odometer(plate)
    if prev and km < prev[0]['km'] and day > prev[0]['day']:
        raise ValueError(f"앞서 넣은 km({prev[0]['km']:,.0f})보다 작아요.")
    with _conn() as conn:
        conn.execute('INSERT INTO vehicle_odometer (plate, day, km, note, created_by, created_at) VALUES (?,?,?,?,?,?) '
                     'ON CONFLICT(plate, day) DO UPDATE SET km = excluded.km, note = excluded.note, '
                     'created_by = excluded.created_by, created_at = excluded.created_at',
                     (plate, day, km, (note or '').strip()[:60], by, _now()))
    return {'plate': plate, 'readings': odometer(plate)}


# ── 읽기 ─────────────────────────────────────────────────────────────────
def info(plate: str) -> Dict:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM vehicle_info WHERE plate = ?', (plate,)).fetchone()
    return dict(r) if r else {'plate': plate}


def odometer(plate: str, limit: int = 24) -> List[Dict]:
    with _conn() as conn:
        rows = conn.execute('SELECT day, km, note FROM vehicle_odometer WHERE plate = ? ORDER BY day DESC LIMIT ?',
                            (plate, limit)).fetchall()
    out = [dict(r) for r in rows]
    for i, r in enumerate(out):        # 앞 기록과의 차이 = 그 기간 주행거리
        nxt = out[i + 1] if i + 1 < len(out) else None
        r['driven'] = round(r['km'] - nxt['km'], 1) if nxt else None
    return out


def costs(plate: str = '', year: str = '', limit: int = 300) -> List[Dict]:
    q, args = 'SELECT * FROM vehicle_costs WHERE 1=1', []
    if plate:
        q += ' AND plate = ?'
        args.append(plate)
    if year:
        q += ' AND substr(day,1,4) = ?'
        args.append(year)
    q += ' ORDER BY day DESC, id DESC LIMIT ?'
    args.append(limit)
    with _conn() as conn:
        return [dict(r) for r in conn.execute(q, args).fetchall()]


def _d_day(end: str, today: date) -> Optional[int]:
    try:
        return (date.fromisoformat(end) - today).days
    except (TypeError, ValueError):
        return None


def overview(today: Optional[date] = None) -> Dict:
    """차량 관리 첫 화면 — 차량 표 + 요약(만기 임박·지남, 이달·올해 비용)."""
    today = today or date.today()
    year, month = str(today.year), today.strftime('%Y-%m')
    with _conn() as conn:
        infos = {r['plate']: dict(r) for r in conn.execute('SELECT * FROM vehicle_info').fetchall()}
        sums = conn.execute('SELECT plate, kind, substr(day,1,7) AS ym, SUM(amount) AS amt, COUNT(*) AS n FROM vehicle_costs '
                            'WHERE substr(day,1,4) = ? GROUP BY plate, kind, ym', (year,)).fetchall()
        last_fix = {r['plate']: dict(r) for r in conn.execute(
            "SELECT plate, MAX(day) AS day, item FROM vehicle_costs WHERE kind IN ('정비','검사','소모품') GROUP BY plate")}
        last_km = {r['plate']: dict(r) for r in conn.execute(
            'SELECT plate, MAX(day) AS day, km FROM vehicle_odometer GROUP BY plate')}
    by_plate: Dict[str, Dict] = {}
    kinds_year: Dict[str, int] = {}
    month_total = year_total = 0
    for r in sums:
        p = by_plate.setdefault(r['plate'], {'year': 0, 'month': 0, 'fix_n': 0})
        p['year'] += r['amt'] or 0
        year_total += r['amt'] or 0
        kinds_year[r['kind']] = kinds_year.get(r['kind'], 0) + (r['amt'] or 0)
        if r['ym'] == month:
            p['month'] += r['amt'] or 0
            month_total += r['amt'] or 0
        if r['kind'] in ('정비', '검사', '소모품'):
            p['fix_n'] += r['n']
    cars, alerts = [], 0
    for c in _fleet():
        i = infos.get(c['plate'], {})
        ins_d = _d_day(i.get('insurance_end'), today)
        insp_d = _d_day(i.get('inspection_end'), today)
        warn = any(d is not None and d <= 30 for d in (ins_d, insp_d))
        alerts += warn
        s = by_plate.get(c['plate'], {})
        lf = last_fix.get(c['plate']) or {}
        lk = last_km.get(c['plate']) or {}
        cars.append({
            'plate': c['plate'], 'label': c['label'], 'model': c['model'], 'fuel': c['fuel'], 'corp': c['corp'],
            'usage': c['usage'], 'uplus': c['uplus'],
            'ownership': i.get('ownership') or '', 'insurer': i.get('insurer') or '',
            'insurance_end': i.get('insurance_end') or '', 'insurance_d': ins_d,
            'inspection_end': i.get('inspection_end') or '', 'inspection_d': insp_d, 'note': i.get('note') or '',
            'cost_year': s.get('year', 0), 'cost_month': s.get('month', 0), 'fix_n': s.get('fix_n', 0),
            'last_fix_day': lf.get('day') or '', 'last_fix_item': (lf.get('item') or '')[:40],
            'km': lk.get('km'), 'km_day': lk.get('day') or '', 'warn': warn,
        })
    cars.sort(key=lambda x: (not x['warn'], x['usage'] == 'off', x['label']))
    return {'today': today.isoformat(), 'year': year, 'month': month, 'cars': cars, 'alerts': alerts,
            'month_total': month_total, 'year_total': year_total,
            'kinds_year': [{'kind': k, 'amount': v} for k, v in sorted(kinds_year.items(), key=lambda kv: -kv[1])]}
