"""나라장터 입찰 공고 모으기·거르기 (2026-11-18, 부장님).

조달청 나라장터 입찰공고정보서비스(공공데이터포털)에서 냉난방 관련 '공사' 공고를 하루 몇 번 모아
사용자 DB(g2b_bids)에 저장한다. 그룹웨어 「입찰 공고」 화면은 저장된 것만 읽고, 아침 DM 은
새로 들어온 '들어갈 수 있는 공고'만 하루 한 번 묶어서 보낸다.

거르는 기준(11-18 부장님):
- 지역: 김포 업체라 참가가능지역에 '경기도'(도 전체) 또는 '경기도 김포시'가 있어야 한다.
  시 단위 제한(예: 군포·수원만)이 많아 '경기도 ○○시' 만으로는 안 된다. 지역 제한 없는 전국 공고는 경쟁이 세서 뺀다.
- 면허: 나라장터 면허제한은 같은 묶음(lmtGrpNo) 안 면허를 모두 가져야 하고, 묶음끼리는 하나만 맞으면 된다.
  보유 면허 = G2B_LICENSES(기본 기계설비ㆍ가스공사업). 냉매회수업은 준비 중(G2B_LICENSES_PLANNED) —
  그것만 없어서 못 들어가는 공고는 '면허만 없음'으로 따로 보여 준다.
- 신규 면허라 까다로운 조건 제외: 실적제한·PQ·지명경쟁, 추정가격 4억 이상 적격심사.

API 메모: 조회기간 최대 7일(26-11-10 축소), 개발계정 오퍼레이션당 하루 1,000건.
참가가능지역·면허제한 조회는 inqryDiv=2 + bidNtceNo + bidNtceOrd 셋 다 있어야 한다(ord 빠지면 08 에러).
"""
from __future__ import annotations

import json
import os
import re
import threading
from contextlib import contextmanager
from datetime import datetime, timedelta
from typing import Dict, Iterable, List, Optional

import requests

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

BASE = 'https://apis.data.go.kr/1230000/ad/BidPublicInfoService02/'
OP_LIST = 'getBidPblancListInfoCnstwkPPSSrch02'        # 나라장터 검색조건 공사 공고
OP_REGION = 'getBidPblancListInfoPrtcptPsblRgn02'      # 참가가능지역
OP_LICENSE = 'getBidPblancListInfoLicenseLimit02'      # 면허제한

KEYWORDS = ('냉난방', '시스템에어컨', '냉방기', '에어컨', 'EHP', '냉온수기', '공조', '항온항습')
REGIONS_OK = ('경기도', '경기도 김포시')
HARD_PRICE = 400_000_000      # 이 금액 이상 적격심사는 시공경험·경영상태 점수가 커서 신규 면허로는 어렵다

ELIGIBLE, LICENSE_ONLY, OTHER = 'eligible', 'license_only', 'other'
MARKS = {'': '', 'review': '검토 중', 'join': '참여', 'skip': '안 함'}

# 아침 DM 받는 곳 — 기본은 부장님 DM(슬랙 사용자 ID 로 보내면 봇과의 DM 으로 간다)
_DEFAULT_ALERT = 'U04UL2ZLJAX'

_DDL = [
    '''CREATE TABLE IF NOT EXISTS g2b_bids (
        bid_no TEXT PRIMARY KEY,
        bid_ord TEXT NOT NULL DEFAULT '000',
        title TEXT NOT NULL DEFAULT '',
        org TEXT NOT NULL DEFAULT '',
        demand_org TEXT NOT NULL DEFAULT '',
        contract_method TEXT NOT NULL DEFAULT '',
        award_method TEXT NOT NULL DEFAULT '',
        price INTEGER NOT NULL DEFAULT 0,
        budget INTEGER NOT NULL DEFAULT 0,
        posted_at TEXT NOT NULL DEFAULT '',
        close_at TEXT NOT NULL DEFAULT '',
        open_at TEXT NOT NULL DEFAULT '',
        site_region TEXT NOT NULL DEFAULT '',
        regions TEXT NOT NULL DEFAULT '[]',
        license_groups TEXT NOT NULL DEFAULT '[]',
        files TEXT NOT NULL DEFAULT '[]',
        detail_url TEXT NOT NULL DEFAULT '',
        contact TEXT NOT NULL DEFAULT '',
        hard_reason TEXT NOT NULL DEFAULT '',
        category TEXT NOT NULL DEFAULT 'other',
        reason TEXT NOT NULL DEFAULT '',
        missing_licenses TEXT NOT NULL DEFAULT '',
        keyword TEXT NOT NULL DEFAULT '',
        first_seen TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        notified_at TEXT NOT NULL DEFAULT '',
        mark TEXT NOT NULL DEFAULT '',
        mark_by TEXT NOT NULL DEFAULT '',
        mark_at TEXT NOT NULL DEFAULT ''
    )''',
    'CREATE INDEX IF NOT EXISTS ix_g2b_bids_close ON g2b_bids (close_at)',
    '''CREATE TABLE IF NOT EXISTS g2b_kv (
        key TEXT PRIMARY KEY,
        value TEXT
    )''',
]
_table_ready = False
_collect_lock = threading.Lock()


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
            # 11-18 펼친 상세용 — 나라장터가 준 공고 항목 전체(JSON). 처음 버전 줄은 비어 있다가 다음 모으기에서 채움
            cols = {r[1] for r in conn.execute('PRAGMA table_info(g2b_bids)').fetchall()}
            if 'raw' not in cols:
                conn.execute("ALTER TABLE g2b_bids ADD COLUMN raw TEXT NOT NULL DEFAULT ''")
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _key() -> str:
    return os.getenv('G2B_SERVICE_KEY', '').strip()


def enabled() -> bool:
    return bool(_key())


# ── 설정 ──────────────────────────────────────────────────────────────
def _norm_lic(name: str) -> str:
    """'기계설비ㆍ가스공사업/6202' → '기계설비가스공사업' (가운뎃점·띄어쓰기·업종코드 무시)."""
    name = (name or '').split('/')[0]
    return re.sub(r'[\sㆍ·・‧.,]', '', name)


def _env_list(name: str, default: str) -> List[str]:
    raw = os.getenv(name)
    raw = default if raw is None else raw
    return [x.strip() for x in raw.split(',') if x.strip()]


def held_licenses() -> List[str]:
    return _env_list('G2B_LICENSES', '기계설비ㆍ가스공사업')


def planned_licenses() -> List[str]:
    return _env_list('G2B_LICENSES_PLANNED', '냉매회수업')


def alert_channel() -> str:
    return os.getenv('G2B_ALERT_CHANNEL', '').strip() or _DEFAULT_ALERT


# ── 분류(순수 함수) ───────────────────────────────────────────────────
def region_ok(regions: List[str]) -> bool:
    return any((r or '').strip() in REGIONS_OK for r in regions)


def license_status(groups: List[List[str]], held: Iterable[str], planned: Iterable[str] = ()) -> tuple:
    """(되는가, 준비 중 면허로 채워지는 모자란 면허 목록).
    groups = 묶음마다 면허 이름 목록. 묶음이 없으면 면허 제한 없음."""
    have = {_norm_lic(x) for x in held}
    if not groups:
        return True, []
    for g in groups:
        if {_norm_lic(x) for x in g} <= have:
            return True, []
    more = have | {_norm_lic(x) for x in planned}
    best: Optional[List[str]] = None
    for g in groups:
        if {_norm_lic(x) for x in g} <= more:
            miss = [x.split('/')[0] for x in g if _norm_lic(x) not in have]
            if best is None or len(miss) < len(best):
                best = miss
    return False, (best or [])


def hard_reason(item: dict) -> str:
    """신규 면허로는 들어가기 어려운 조건이면 그 이유, 아니면 ''."""
    if (item.get('arsltCmptYn') or '') == 'Y':
        return '실적제한'
    if (item.get('pqEvalYn') or '') == 'Y':
        return 'PQ(사전심사)'
    if (item.get('dsgntCmptYn') or '') == 'Y' or '지명' in (item.get('cntrctCnclsMthdNm') or ''):
        return '지명경쟁'
    method = item.get('sucsfbidMthdNm') or ''
    if '적격심사' in method and _int(item.get('presmptPrce')) >= HARD_PRICE:
        return '4억 이상 적격심사'
    return ''


def classify(item: dict, regions: List[str], groups: List[List[str]],
             held: Iterable[str], planned: Iterable[str]) -> dict:
    hard = hard_reason(item)
    lic_ok, missing = license_status(groups, held, planned)
    if not regions:
        reason = '지역 제한 없는 전국 공고'
    elif not region_ok(regions):
        reason = '참가가능지역에 경기도·김포시 없음'
    elif hard:
        reason = hard
    elif lic_ok:
        return {'category': ELIGIBLE, 'reason': '', 'missing': []}
    elif missing:
        return {'category': LICENSE_ONLY, 'reason': '면허 없음: ' + ', '.join(missing), 'missing': missing}
    else:
        reason = '면허 안 맞음'
    return {'category': OTHER, 'reason': reason, 'missing': missing}


def _int(v) -> int:
    try:
        return int(float(str(v or '0').replace(',', '')))
    except ValueError:
        return 0


# ── 나라장터 부르기 ───────────────────────────────────────────────────
def _safe_err(e) -> str:
    """오류 글에 인증키가 주소째 찍히지 않게(requests 오류는 URL 전체를 담는다)."""
    s, k = str(e), _key()
    if k:
        from urllib.parse import quote
        s = s.replace(k, '***').replace(quote(k, safe=''), '***')
    return re.sub(r'(serviceKey=)[^&\s\'")]+', r'\1***', s)[:300]


def _get(op: str, params: dict) -> dict:
    p = {'serviceKey': _key(), 'type': 'json', **params}
    try:
        r = requests.get(BASE + op, params=p, timeout=(10, 30))
    except (requests.ConnectionError, requests.Timeout):   # 공공데이터포털이 가끔 연결을 놓침 — 한 번만 다시
        r = requests.get(BASE + op, params=p, timeout=(10, 30))
    r.raise_for_status()
    data = r.json()
    if 'response' not in data:   # 나라장터 오류 형식 {"nkoneps.com.response.ResponseError": {...}}
        raise RuntimeError(f'{op} 오류: {json.dumps(data, ensure_ascii=False)[:200]}')
    head = data['response'].get('header') or {}
    if str(head.get('resultCode', '00')) != '00':
        raise RuntimeError(f"{op} 오류: {head.get('resultCode')} {head.get('resultMsg')}")
    return data['response'].get('body') or {}


def _items(body: dict) -> List[dict]:
    items = body.get('items') or []
    if isinstance(items, dict):   # XML 변환 흔적 {"item": [...]}
        items = items.get('item') or []
    if isinstance(items, dict):
        items = [items]
    return list(items)


def search(keyword: str, start: datetime, end: datetime) -> List[dict]:
    out, page = [], 1
    while True:
        body = _get(OP_LIST, {'inqryDiv': 1, 'inqryBgnDt': start.strftime('%Y%m%d%H%M'),
                              'inqryEndDt': end.strftime('%Y%m%d%H%M'), 'bidNtceNm': keyword,
                              'numOfRows': 100, 'pageNo': page})
        items = _items(body)
        out.extend(items)
        total = _int(body.get('totalCount'))
        if not items or page * 100 >= total or page >= 10:
            return out
        page += 1


def fetch_limits(bid_no: str, bid_ord: str) -> tuple:
    q = {'inqryDiv': 2, 'bidNtceNo': bid_no, 'bidNtceOrd': bid_ord, 'numOfRows': 50, 'pageNo': 1}
    regions = sorted({(x.get('prtcptPsblRgnNm') or '').strip()
                      for x in _items(_get(OP_REGION, q)) if x.get('prtcptPsblRgnNm')})
    grp: Dict[str, List[str]] = {}
    for x in _items(_get(OP_LICENSE, q)):
        name = (x.get('lcnsLmtNm') or '').strip()
        if name:
            grp.setdefault(str(x.get('lmtGrpNo') or '1'), []).append(name.split('/')[0])
    groups = [grp[k] for k in sorted(grp)]
    return regions, groups


def _files(item: dict) -> List[dict]:
    files = []
    for i in range(1, 11):
        url, name = item.get(f'ntceSpecDocUrl{i}'), item.get(f'ntceSpecFileNm{i}')
        if url and name:
            files.append({'name': name, 'url': url})
    return files


def _contact(item: dict) -> str:
    name, tel = (item.get('ntceInsttOfclNm') or '').strip(), (item.get('ntceInsttOfclTelNo') or '').strip()
    return ' '.join(x for x in (name, tel) if x)


# ── 모으기 ────────────────────────────────────────────────────────────
def collect(days: int = 2, now: Optional[datetime] = None) -> dict:
    """최근 days 일 공고를 검색어마다 찾아 새것·정정공고만 지역·면허를 붙여 저장."""
    if not enabled():
        return {'skipped': 'G2B_SERVICE_KEY 없음'}
    if not _collect_lock.acquire(blocking=False):
        return {'skipped': '이미 모으는 중'}
    try:
        now = now or datetime.now()
        # 조회기간은 7일을 넘으면 07(입력범위값 초과) — days=7 이면 6일 전 00:00부터 지금까지
        start = datetime.combine(now.date() - timedelta(days=min(max(days, 1), 7) - 1), datetime.min.time())
        found: Dict[str, dict] = {}
        errors = 0
        for kw in KEYWORDS:
            try:
                for it in search(kw, start, now):
                    no = it.get('bidNtceNo')
                    if not no:
                        continue
                    prev = found.get(no)
                    if prev is None or (it.get('bidNtceOrd') or '') > (prev.get('bidNtceOrd') or ''):
                        it['_kw'] = kw
                        found[no] = it
            except Exception as e:
                errors += 1
                logger.warning(f'[G2B] 검색 실패 ({kw}): {_safe_err(e)}')
        with _conn() as conn:
            known, no_raw = {}, set()
            for r in conn.execute('SELECT bid_no, bid_ord, raw FROM g2b_bids'):
                known[r[0]] = r[1]
                if not r[2]:
                    no_raw.add(r[0])
        held, planned = held_licenses(), planned_licenses()
        new = changed = 0
        for no, it in found.items():
            ord_ = it.get('bidNtceOrd') or '000'
            if no in known and known[no] >= ord_:
                if no in no_raw and known[no] == ord_:   # 상세 칸만 채움(지역·면허는 다시 안 부름)
                    with _conn() as conn:
                        conn.execute('UPDATE g2b_bids SET raw=? WHERE bid_no=?', (_raw(it), no))
                continue
            try:
                regions, groups = fetch_limits(no, ord_)
            except Exception as e:
                errors += 1
                logger.warning(f'[G2B] 지역·면허 조회 실패 {no}: {_safe_err(e)}')
                continue
            _save(it, regions, groups, held, planned, is_new=no not in known)
            if no in known:
                changed += 1
            else:
                new += 1
        if errors < len(KEYWORDS):   # 검색이 하나라도 됐으면 '마지막으로 모은 때' 갱신
            mark_collected()
        result = {'found': len(found), 'new': new, 'changed': changed, 'errors': errors}
        logger.info(f'[G2B] 공고 모으기 — {result}')
        return result
    finally:
        _collect_lock.release()


def _raw(it: dict) -> str:
    return json.dumps({k: v for k, v in it.items() if not k.startswith('_') and v not in ('', None)},
                      ensure_ascii=False)


def _yn(v) -> str:
    return {'Y': '예', 'N': '아니오'}.get(str(v or '').strip(), '')


def _detail(raw: str) -> dict:
    """펼친 상세 — 나라장터 항목 중 입찰 판단에 쓰는 것만 골라 사람이 읽는 말로."""
    try:
        it = json.loads(raw) if raw else {}
    except ValueError:
        it = {}
    if not it:
        return {}
    t = lambda k: str(it.get(k) or '').strip()   # noqa: E731
    rate = t('sucsfbidLwltRate')
    vat = _int(it.get('VAT'))
    return {
        'ref_no': t('refNo'), 'kind': t('ntceKindNm'), 're_notice': _yn(it.get('reNtceYn')),
        'bid_method': t('bidMethdNm'),
        'begin_at': t('bidBeginDt')[:16], 'qual_at': t('bidQlfctRgstDt')[:16],
        'rebid_open_at': t('rbidOpengDt')[:16], 'open_place': t('opengPlce'),
        'budget': _int(it.get('bdgtAmt')), 'vat': vat,
        'main_license': t('mainCnsttyNm'), 'main_price': _int(it.get('mainCnsttyPresmptPrce')),
        'award_full': t('sucsfbidMthdNm'), 'lower_rate': f'{rate}%' if rate else '',
        'price_method': t('prearngPrceDcsnMthdNm'),
        # 복수예가일 때만(단일예가는 '1개 중 0개'로 와서 뺌)
        'price_count': (f"{t('totPrdprcNum')}개 중 {t('drwtPrdprcNum')}개 추첨"
                        if _int(it.get('drwtPrdprcNum')) > 0 else ''),
        'joint': t('cmmnSpldmdMethdNm'), 'joint_region': _yn(it.get('rgnDutyJntcontrctYn')),
        'site_region': t('cnstrtsiteRgnNm'),
        'contact_name': t('ntceInsttOfclNm') or t('exctvNm'), 'contact_tel': t('ntceInsttOfclTelNo'),
        'contact_email': t('ntceInsttOfclEmailAdrs'),
        'std_doc_url': t('stdNtceDocUrl'),
    }


def _save(it: dict, regions, groups, held, planned, is_new: bool) -> None:
    c = classify(it, regions, groups, held, planned)
    row = {
        'bid_no': it['bidNtceNo'], 'bid_ord': it.get('bidNtceOrd') or '000',
        'title': (it.get('bidNtceNm') or '').strip(), 'org': (it.get('ntceInsttNm') or '').strip(),
        'demand_org': (it.get('dminsttNm') or '').strip(),
        'contract_method': it.get('cntrctCnclsMthdNm') or '', 'award_method': it.get('sucsfbidMthdNm') or '',
        'price': _int(it.get('presmptPrce')), 'budget': _int(it.get('bdgtAmt')),
        # 마감이 비어 있는 공고(일부 수의)는 개찰 일시로 — 비면 '마감 전' 목록에서 빠지므로
        'posted_at': (it.get('bidNtceDt') or '')[:16],
        'close_at': (it.get('bidClseDt') or it.get('opengDt') or '')[:16],
        'open_at': (it.get('opengDt') or '')[:16], 'site_region': it.get('cnstrtsiteRgnNm') or '',
        'regions': json.dumps(regions, ensure_ascii=False), 'license_groups': json.dumps(groups, ensure_ascii=False),
        'files': json.dumps(_files(it), ensure_ascii=False), 'detail_url': it.get('bidNtceDtlUrl') or '',
        'contact': _contact(it), 'hard_reason': hard_reason(it), 'category': c['category'], 'reason': c['reason'],
        'missing_licenses': ', '.join(c['missing']), 'keyword': it.get('_kw', ''), 'updated_at': _now(),
        'raw': _raw(it),
    }
    cols = list(row)
    with _conn() as conn:
        if is_new:
            row['first_seen'] = row['updated_at']
            cols = list(row)
            conn.execute(f"INSERT OR REPLACE INTO g2b_bids ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                         [row[c] for c in cols])
        else:   # 정정공고 — 표시(검토 중 등)·처음 본 때·알림 보낸 때는 그대로
            conn.execute(f"UPDATE g2b_bids SET {','.join(c + '=?' for c in cols if c != 'bid_no')} WHERE bid_no=?",
                         [row[c] for c in cols if c != 'bid_no'] + [row['bid_no']])


def reclassify() -> int:
    """보유 면허를 바꾼 뒤(냉매회수업 취득 등) 저장된 공고를 다시 분류. API 는 부르지 않는다."""
    held, planned = held_licenses(), planned_licenses()
    n = 0
    with _conn() as conn:
        rows = [dict(r) for r in conn.execute('SELECT * FROM g2b_bids').fetchall()]
        for r in rows:
            item = {'arsltCmptYn': 'Y' if r['hard_reason'] == '실적제한' else 'N',
                    'pqEvalYn': 'Y' if r['hard_reason'].startswith('PQ') else 'N',
                    'dsgntCmptYn': 'Y' if r['hard_reason'] == '지명경쟁' else 'N',
                    'cntrctCnclsMthdNm': r['contract_method'], 'sucsfbidMthdNm': r['award_method'],
                    'presmptPrce': r['price']}
            c = classify(item, json.loads(r['regions']), json.loads(r['license_groups']), held, planned)
            if c['category'] != r['category'] or c['reason'] != r['reason']:
                conn.execute('UPDATE g2b_bids SET category=?, reason=?, missing_licenses=? WHERE bid_no=?',
                             (c['category'], c['reason'], ', '.join(c['missing']), r['bid_no']))
                n += 1
    return n


def maybe_bootstrap() -> None:
    """표가 비어 있으면(처음 켠 날) 7일치를 한 번 채운다 — 서버 시작 직후 백그라운드에서."""
    if not enabled():
        return
    try:
        with _conn() as conn:
            empty = conn.execute('SELECT COUNT(*) FROM g2b_bids').fetchone()[0] == 0
        if empty:
            collect(days=7)
    except Exception as e:
        logger.warning(f'[G2B] 첫 채우기 실패: {_safe_err(e)}')


# ── 화면용 ────────────────────────────────────────────────────────────
def _view(r: dict, now_s: str) -> dict:
    files = json.loads(r.get('files') or '[]')
    groups = json.loads(r.get('license_groups') or '[]')
    return {
        'bid_no': r['bid_no'], 'bid_ord': r['bid_ord'], 'title': r['title'], 'org': r['org'],
        'demand_org': r['demand_org'], 'contract_method': r['contract_method'],
        'award_method': (r['award_method'] or '').split('-')[0], 'price': r['price'],
        'posted_at': r['posted_at'], 'close_at': r['close_at'], 'open_at': r['open_at'],
        'closed': bool(r['close_at']) and r['close_at'] < now_s,
        'site_region': r['site_region'], 'regions': json.loads(r.get('regions') or '[]'),
        'licenses': [' + '.join(g) for g in groups], 'files': files, 'detail_url': r['detail_url'],
        'contact': r['contact'], 'category': r['category'], 'reason': r['reason'],
        'missing': r['missing_licenses'], 'mark': r['mark'], 'mark_label': MARKS.get(r['mark'], ''),
        'mark_by': r['mark_by'], 'mark_at': r['mark_at'], 'first_seen': r['first_seen'],
        'is_new': r['first_seen'] >= (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d %H:%M:%S'),
        'hard_reason': r.get('hard_reason', ''),
        'detail': _detail(r.get('raw') or ''),
    }


def list_bids(tab: str = ELIGIBLE, include_closed: bool = False, q: str = '', limit: int = 300) -> dict:
    now_s = datetime.now().strftime('%Y-%m-%d %H:%M')
    with _conn() as conn:
        counts = {k: 0 for k in (ELIGIBLE, LICENSE_ONLY, OTHER)}
        for cat, n in conn.execute('SELECT category, COUNT(*) FROM g2b_bids WHERE close_at >= ? GROUP BY category',
                                   (now_s,)):
            counts[cat] = n
        sql, args = 'SELECT * FROM g2b_bids WHERE category = ?', [tab if tab in counts else ELIGIBLE]
        if not include_closed:
            sql += ' AND close_at >= ?'
            args.append(now_s)
        if q:
            sql += ' AND (title LIKE ? OR org LIKE ? OR demand_org LIKE ?)'
            args += [f'%{q}%'] * 3
        sql += ' ORDER BY (close_at < ?) ASC, CASE WHEN close_at < ? THEN close_at END DESC, close_at ASC LIMIT ?'
        args += [now_s, now_s, limit]
        rows = [dict(r) for r in conn.execute(sql, args).fetchall()]
        last = conn.execute("SELECT value FROM g2b_kv WHERE key='last_collect'").fetchone()
    return {'items': [_view(r, now_s) for r in rows], 'counts': counts,
            'held': held_licenses(), 'planned': planned_licenses(),
            'last_collect': last[0] if last else '', 'enabled': enabled()}


def set_mark(bid_no: str, mark: str, who: str) -> dict:
    if mark not in MARKS:
        raise ValueError('표시 값이 올바르지 않습니다.')
    with _conn() as conn:
        cur = conn.execute('UPDATE g2b_bids SET mark=?, mark_by=?, mark_at=? WHERE bid_no=?',
                           (mark, who if mark else '', _now() if mark else '', bid_no))
        if cur.rowcount == 0:
            raise LookupError('공고를 찾지 못했습니다.')
        r = dict(conn.execute('SELECT * FROM g2b_bids WHERE bid_no=?', (bid_no,)).fetchone())
    return _view(r, datetime.now().strftime('%Y-%m-%d %H:%M'))


def home_summary() -> tuple:
    """그룹웨어 홈 한 줄 — (마감 전 '들어갈 수 있는 공고' 수, 바뀜 표시).
    바뀜 표시 = 공고 번호 묶음 → [닫기] 한 뒤 새 공고가 들어오면 다시 뜬다. '안 함' 표시한 공고는 뺀다."""
    try:
        with _conn() as conn:
            nos = [r[0] for r in conn.execute(
                "SELECT bid_no FROM g2b_bids WHERE category=? AND close_at >= ? AND mark != 'skip' ORDER BY bid_no",
                (ELIGIBLE, datetime.now().strftime('%Y-%m-%d %H:%M')))]
        return len(nos), ','.join(nos)[-200:]
    except Exception as e:
        logger.warning(f'[G2B] 홈 건수 읽기 실패: {e}')
        return 0, ''


def mark_collected() -> None:
    with _conn() as conn:
        conn.execute("INSERT OR REPLACE INTO g2b_kv (key, value) VALUES ('last_collect', ?)", (_now()[:16],))


# ── 아침 DM ───────────────────────────────────────────────────────────
def _won(n: int) -> str:
    if n >= 100_000_000:
        eok, man = divmod(n, 100_000_000)
        man //= 10_000
        return f'{eok}억 {man:,}만 원' if man else f'{eok}억 원'
    if n >= 10_000:
        return f'{n // 10_000:,}만 원'
    return f'{n:,}원' if n else '-'


_WD = '월화수목금토일'


def _when(s: str) -> str:
    try:
        d = datetime.strptime(s[:16], '%Y-%m-%d %H:%M')
        return f'{d.month}/{d.day}({_WD[d.weekday()]}) {d:%H:%M}'
    except ValueError:
        return s or '-'


def format_digest(rows: List[dict], license_only: int, page_url: str) -> str:
    lines = [f'*나라장터 새 입찰 공고 {len(rows)}건* (경기·김포 업체 참가 가능, 지금 면허로 들어갈 수 있는 것)']
    for r in rows:
        files = json.loads(r.get('files') or '[]')
        links = [f"<{r['detail_url']}|나라장터에서 보기>"] if r.get('detail_url') else []
        for f in files[:4]:
            name = f['name']
            kind = '공고문' if '공고' in name else '내역서' if '내역' in name else '시방서' if '시방' in name else ''
            if not kind:
                continue   # 알림에는 공고문·내역서·시방서만 (나머지는 그룹웨어에서)
            ext = name.rsplit('.', 1)[-1].upper() if '.' in name else ''
            links.append(f"<{f['url']}|{kind}{' ' + ext if ext else ''}>")
        lines.append(
            f"\n• *{r['title']}*\n"
            f"   {r['demand_org'] or r['org']} · {(r['award_method'] or r['contract_method']).split('-')[0]} · "
            f"추정가격 {_won(r['price'])}\n"
            f"   마감 {_when(r['close_at'])} · 참가가능지역: {', '.join(json.loads(r['regions']))}\n"
            f"   {' · '.join(links)}")
    if license_only:
        lines.append(f"\n냉매회수업 등 면허만 없어서 못 들어가는 새 공고도 {license_only}건 있어요.")
    lines.append(f"\n<{page_url}|그룹웨어에서 전체 보기>")
    return '\n'.join(lines)


def _post_slack(channel: str, text: str) -> bool:
    token = os.getenv('SLACK_BOT_TOKEN', '').strip()
    if not token or not channel:
        logger.warning('[G2B] 슬랙 토큰/받는 곳 없음 — DM skip')
        return False
    r = requests.post('https://slack.com/api/chat.postMessage', timeout=10,
                      headers={'Authorization': f'Bearer {token}'},
                      json={'channel': channel, 'text': text, 'unfurl_links': False, 'unfurl_media': False})
    ok = bool(r.json().get('ok'))
    if not ok:
        logger.warning(f"[G2B] DM 실패: {r.json().get('error')}")
    return ok


def send_morning_digest(force: bool = False, today=None) -> dict:
    """평일 아침 한 번 — 아직 알리지 않은 '들어갈 수 있는' 마감 전 공고가 있을 때만 묶어서 보낸다."""
    from dashboard.services.absent_remind import _is_business_day
    today = today or datetime.now().date()
    if not force and not _is_business_day(today):
        return {'skipped': '영업일 아님'}
    now_s = datetime.now().strftime('%Y-%m-%d %H:%M')
    with _conn() as conn:
        rows = [dict(r) for r in conn.execute(
            "SELECT * FROM g2b_bids WHERE category=? AND notified_at='' AND close_at >= ? ORDER BY close_at",
            (ELIGIBLE, now_s)).fetchall()]
        lic_only = conn.execute(
            "SELECT COUNT(*) FROM g2b_bids WHERE category=? AND notified_at='' AND close_at >= ?",
            (LICENSE_ONLY, now_s)).fetchone()[0]
    if not rows:
        logger.info(f'[G2B] 아침 알림 — 새 공고 없음(면허만 없음 {lic_only}건) → 안 보냄')
        return {'sent': False, 'eligible': 0, 'license_only': lic_only}
    base = os.getenv('GW_PUBLIC_URL', 'https://pm.itg-aircon.com').rstrip('/')
    ok = _post_slack(alert_channel(), format_digest(rows, lic_only, f'{base}/bids'))
    if ok:
        with _conn() as conn:
            conn.execute("UPDATE g2b_bids SET notified_at=? WHERE notified_at='' AND category IN (?, ?) AND close_at >= ?",
                         (_now(), ELIGIBLE, LICENSE_ONLY, now_s))
    return {'sent': ok, 'eligible': len(rows), 'license_only': lic_only}
