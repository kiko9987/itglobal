"""나라장터 입찰 공고 모으기·거르기 (2026-11-18, 부장님).

조달청 나라장터 입찰공고정보서비스(공공데이터포털)에서 냉난방 관련 '공사' 공고를 하루 몇 번 모아
사용자 DB(g2b_bids)에 저장한다. 그룹웨어 「입찰 공고」 화면은 저장된 것만 읽고, 아침 DM 은
새로 들어온 '들어갈 수 있는 공고'만 하루 한 번 묶어서 보낸다.

거르는 기준(11-18 부장님):
- 지역: 김포 업체라 참가가능지역에 '경기도'(도 전체) 또는 '경기도 김포시'가 있어야 한다.
  시 단위 제한(예: 군포·수원만)이 많아 '경기도 ○○시' 만으로는 안 된다. 참가가능지역은 회사 소재지 기준이라
  '서울만'·'인천만' 공고는 시공 지역이어도 못 들어간다.
  지역 제한 없는 전국 공고는 들어갈 수 있다(11-18 부장님 — 처음엔 경쟁이 세서 뺐다가 넣기로). 화면에 '전국' 표시.
- 업종: 면허 제한이 안 적힌 전국 공고라도 제목이 전기·통신 공사면 그 등록이 있어야 해서 뺀다(ELEC_WORDS).
- 면허: 나라장터 면허제한은 같은 묶음(lmtGrpNo) 안 면허를 모두 가져야 하고, 묶음끼리는 하나만 맞으면 된다.
  보유 면허 = G2B_LICENSES(기본 기계설비ㆍ가스공사업). 냉매회수업은 준비 중(G2B_LICENSES_PLANNED) —
  그것만 없어서 못 들어가는 공고는 '면허만 없음'으로 따로 보여 준다.
- 신규 면허라 까다로운 조건 제외: 실적제한·사전심사(PQ)·지명경쟁, 추정가격 4억 이상 적격심사.

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
# 누리장터 민간입찰(11-20) — 사립학교법인·병원·교회·호텔 등 '수요기관 외의 자' 공고. 같은 키로 됨.
#   검색어 조건이 안 먹어서(공고명 칸 이름이 ntceNm) 종류별 목록을 통째로 받아 제목으로 거름(주 ~40건이라 1~2번 호출)
#   참가자격·지역은 거의 다 '공고서 참조'라 자동 판정이 안 됨 → 입찰 가능에 두고 '확인할 것'에 공고서 확인 주의
NURI_BASE = 'https://apis.data.go.kr/1230000/ao/PrvtBidNtceService/'
NURI_KINDS = (('공사', 'getPrvtBidPblancListInfoCnstwk'), ('물품', 'getPrvtBidPblancListInfoThng'))
NURI_REGION = 'getPrvtBidPblancListInfoPrtcptPsblRgn'
# 공동주택 입찰(K-apt, 국토교통부, 11-20) — 아파트 관리주체의 공사·물품 업체 선정. 같은 키.
#   공고일 범위로 통째로 받아(주 ~1,300건, 1,000건씩) 제목으로 거름. 금액·면허 항목 없음 → 상세는 K-apt 공고 화면에서
APT_BASE = 'https://apis.data.go.kr/1613000/ApHusBidPblAncInfoOfferServiceV3/'
APT_OP = 'getPblAncDeSearchV3'
APT_KIND = {'01': '일반경쟁', '02': '제한경쟁', '03': '지명경쟁'}
APT_SUC = {'01': '최저낙찰', '02': '적격심사', '03': '최저가낙찰', '04': '최고가낙찰'}
APT_CLS = {'02': '공사', '03': '용역', '04': '물품'}
APT_AUTH = {'01': 'K-apt', '02': '조달청', '03': '아파트비드포유', '04': 'KG2B전자입찰', '05': '나이스아파트',
            '06': '이아파트', '10': '민간입찰', '99': '기타 민간업체'}
# 12 = K-apt '전남광주'(11-20 K-apt 지역 목록으로 확인)
APT_AREA = {11: '서울특별시', 26: '부산광역시', 27: '대구광역시', 28: '인천광역시', 29: '광주광역시', 30: '대전광역시',
            31: '울산광역시', 36: '세종특별자치시', 12: '전남광주통합특별시', 41: '경기도', 42: '강원도', 51: '강원특별자치도', 43: '충청북도',
            44: '충청남도', 45: '전라북도', 52: '전북특별자치도', 46: '전라남도', 47: '경상북도', 48: '경상남도',
            50: '제주특별자치도'}
NURI_LICENSE = 'getPrvtBidPblancListInfoLicenseLimit'
OP_LIST = 'getBidPblancListInfoCnstwkPPSSrch02'        # 나라장터 검색조건 공사 공고
OP_LIST_THNG = 'getBidPblancListInfoThngPPSSrch02'     # 물품 공고 — '구매 및 설치'만 씀(11-19 추산초 계기)
KINDS = (('공사', OP_LIST), ('물품', OP_LIST_THNG))
OP_REGION = 'getBidPblancListInfoPrtcptPsblRgn02'      # 참가가능지역
OP_LICENSE = 'getBidPblancListInfoLicenseLimit02'      # 면허제한

# 나라장터 공고명 검색은 글자 그대로 비교 — '냉난방'은 '냉·난방'·'냉.난방'을 놓쳐서 '난방'으로(11-18 실측).
#   '난방'에 딸려 오는 지역난방·보일러 등은 면허·지역 판정에서 '입찰 불가능'으로 걸러진다.
KEYWORDS = ('난방', '시스템에어컨', '냉방기', '에어컨', 'EHP', '냉온수기', '공조', '항온항습', '공기조화', '냉각탑')
REGIONS_OK = ('경기도', '경기도 김포시')
# 면허 제한이 공고에 없어도 법상 그 업 등록이 필요한 공사(제목으로 판단) — 우리 면허로는 못 함
ELEC_WORDS = (('전기공사', '전기공사업'), ('정보통신공사', '정보통신공사업'), ('통신공사', '정보통신공사업'),
              ('소방시설공사', '소방시설공사업'))
HARD_PRICE = 400_000_000      # 이 금액 이상 적격심사는 시공경험·경영상태 점수가 커서 신규 면허로는 어렵다

ELIGIBLE, LICENSE_ONLY, OTHER = 'eligible', 'license_only', 'other'
BLOCKED = 'blocked'   # 화면 탭 '입찰 불가능한 공고' = license_only + other
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
            if 'source' not in cols:   # 11-20 누리장터 추가 — 그 전 줄은 모두 나라장터
                conn.execute("ALTER TABLE g2b_bids ADD COLUMN source TEXT NOT NULL DEFAULT '나라장터'")
            if 'kind' not in cols:   # 11-19 물품(구매·설치) 추가 — 그 전 줄은 모두 공사
                conn.execute("ALTER TABLE g2b_bids ADD COLUMN kind TEXT NOT NULL DEFAULT '공사'")
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
    p = get_profile()
    return p['licenses'] if p.get('licenses') else _env_list('G2B_LICENSES', '기계설비ㆍ가스공사업')


def planned_licenses() -> List[str]:
    p = get_profile()
    if p.get('_saved'):   # 화면에서 저장한 뒤로는 저장값(빈 목록도 그대로)
        return p.get('planned') or []
    return _env_list('G2B_LICENSES_PLANNED', '냉매회수업')


# ── 우리 회사 정보(11-20 부장님: 추산초에서 신용등급 기준으로 탈락 — 금액보다 먼저 걸러야 할 조건) ──
#   그룹웨어 입찰 공고 화면 「우리 회사 정보」에서 고친다. 공고마다 대조: 면허·물품 분류 등록·적격심사 85점.
CREDIT_GRADES = ('AAA', 'AA+', 'AA0', 'AA-', 'A+', 'A0', 'A-', 'BBB+', 'BBB0', 'BBB-',
                 'BB+', 'BB0', 'BB-', 'B+', 'B0', 'B-', 'CCC+이하', '없음')
_PROFILE_DEFAULT = {'credit_grade': '', 'credit_agency': '', 'credit_until': '', 'licenses': [], 'planned': [],
                    'products': [], 'bonus': 0, 'startup': False, 'entity': '아이티글로벌그룹'}
_profile_cache = {'at': 0.0, 'val': None}


def get_profile() -> dict:
    import time
    if _profile_cache['val'] is not None and time.time() - _profile_cache['at'] < 30:
        return _profile_cache['val']
    val = dict(_PROFILE_DEFAULT)
    try:
        with _conn() as conn:
            row = conn.execute("SELECT value FROM g2b_kv WHERE key='profile'").fetchone()
        if row and row[0]:
            val.update(json.loads(row[0]))
            val['_saved'] = True
    except Exception as e:
        logger.warning(f'[G2B] 회사 정보 읽기 실패: {e}')
    _profile_cache.update(at=time.time(), val=val)
    return val


def _codes(text) -> List[dict]:
    """'4010178702 냉난방기' 줄들 → [{'code','name'}]. 10자리 물품분류번호만."""
    out, seen = [], set()
    lines = text if isinstance(text, list) else str(text or '').splitlines()
    for ln in lines:
        if isinstance(ln, dict):
            code, name = str(ln.get('code') or ''), str(ln.get('name') or '')
        else:
            m = re.match(r'\s*(\d{8,10})\s*(.*)', str(ln))
            if not m:
                continue
            code, name = m.group(1), m.group(2).strip()
        if code and code not in seen:
            seen.add(code)
            out.append({'code': code, 'name': name[:40]})
    return out


def save_profile(data: dict, who: str) -> dict:
    grade = str(data.get('credit_grade') or '')
    if grade and grade not in CREDIT_GRADES:
        raise ValueError('신용평가등급 값이 올바르지 않습니다.')
    until = str(data.get('credit_until') or '')[:10]
    if until and not re.match(r'^\d{4}-\d{2}-\d{2}$', until):
        raise ValueError('유효기간은 날짜(YYYY-MM-DD)로 넣어 주세요.')
    lic = lambda v: [x.strip() for x in (v if isinstance(v, list) else re.split(r'[,\n]', str(v or ''))) if x.strip()][:10]   # noqa: E731
    try:
        bonus = max(-2.0, min(2.0, float(data.get('bonus') or 0)))
    except (TypeError, ValueError):
        raise ValueError('신인도 가점은 -2 ~ +2 숫자로 넣어 주세요.')
    val = {'credit_grade': grade, 'credit_agency': str(data.get('credit_agency') or '')[:30], 'credit_until': until,
           'licenses': lic(data.get('licenses')), 'planned': lic(data.get('planned')),
           'products': _codes(data.get('products'))[:40], 'bonus': bonus, 'startup': bool(data.get('startup')),
           'entity': str(data.get('entity') or '아이티글로벌그룹')[:30],
           'updated_by': who, 'updated_at': _now()[:16]}
    with _conn() as conn:
        conn.execute("INSERT OR REPLACE INTO g2b_kv (key, value) VALUES ('profile', ?)", (json.dumps(val, ensure_ascii=False),))
    _profile_cache.update(at=0.0, val=None)
    n = reclassify()   # 면허가 바뀌었을 수 있어 저장된 공고 다시 분류
    logger.info(f'[G2B] 회사 정보 저장 by {who} — 다시 분류 {n}건')
    return get_profile()


# ── 물품 적격심사 85점 계산(지방자치단체 입찰시 낙찰자 결정기준 제2장의3, 2026-11-02 시행) ──
#   별표2(10억 미만 고시금액 이상): 기술능력 10(물품 구매는 만점) + 신용 30 + 가격 60(60-2|88-x|, 95.5~100%=45) + 신인도
#   별표3(고시금액 미만): 기술능력 10 + 신용 20(창업 7년 이내 만점) + 가격 70(70-4|88-x|, 91.75~100%=55) + 신인도
#   별표1(10억 이상)은 이행실적 30점이 있어 계산 안 함. 공고에 이행실적 평가가 따로 있으면 계산이 달라짐(공고문 확인).
_CREDIT_T2 = {'AAA': 30, 'AA+': 30, 'AA0': 30, 'AA-': 30, 'A+': 30, 'A0': 30, 'A-': 30, 'BBB+': 29.7, 'BBB0': 29.5,
              'BBB-': 29.3, 'BB+': 29.0, 'BB0': 29.0, 'BB-': 28.5, 'B+': 28.0, 'B0': 28.0, 'B-': 28.0,
              'CCC+이하': 25.0, '없음': 20.0}
_CREDIT_T3 = {**{g: 20.0 for g in ('AAA', 'AA+', 'AA0', 'AA-', 'A+', 'A0', 'A-', 'BBB+', 'BBB0', 'BBB-')},
              'BB+': 19.7, 'BB0': 19.7, 'BB-': 19.3, 'B+': 19.0, 'B0': 19.0, 'B-': 19.0, 'CCC+이하': 15.0, '없음': 10.0}


def _price_t2(x: float) -> float:
    return 45.0 if 95.5 <= x <= 100 else max(2.0, 60 - 2 * abs(88 - x))


def _price_t3(x: float) -> float:
    return 55.0 if 91.75 <= x <= 100 else max(2.0, 70 - 4 * abs(88 - x))


def qual_table(item: dict) -> str:
    """물품 적격심사 공고가 어느 별표인지 — 'T2'·'T3'·'T1'(계산 안 함)·''(해당 없음)."""
    method = item.get('sucsfbidMthdNm') or ''
    std = item.get('sucsfbidMthdAppStd') or ''
    if item.get('_kind') != '물품' or '적격심사' not in method or '지방자치단체' not in std:
        return ''
    if '10억원 미만' in method and '고시금액 이상' in method:
        return 'T2'
    if '고시금액 미만' in method:
        return 'T3'
    if '10억원 이상' in method:
        return 'T1'
    return ''


def qual_need(table: str, grade: str, bonus: float = 0.0, startup: bool = False, lower: float = 0.0) -> dict:
    """85점에 필요한 최소 '투찰금액/예정가격'(%) — 낙찰하한율 이상. 계산 못 하면 {}."""
    if table not in ('T2', 'T3') or not grade:
        return {}
    credit = (_CREDIT_T2 if table == 'T2' else _CREDIT_T3).get(grade)
    if credit is None:
        return {}
    if table == 'T3' and startup:
        credit = 20.0
    price = _price_t2 if table == 'T2' else _price_t3
    x = round(max(lower, 0.0) * 100) / 100 if lower else 80.0   # 0.01%p 단위로 올려 가며
    while x <= 100.0:
        if 10 + credit + price(round(x, 2)) + bonus >= 85 - 1e-9:
            return {'table': table, 'credit_pts': credit, 'need_pct': round(x, 2), 'lower_pct': lower,
                    'gap_pct': round(x - lower, 2) if lower else 0.0}
        x = round(x + 0.01, 2)
    return {'table': table, 'credit_pts': credit, 'need_pct': None, 'lower_pct': lower}


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
        return '사전심사(PQ)'   # 11-18 부장님: PQ 만으론 뜻을 몰라 우리말 먼저
    if (item.get('dsgntCmptYn') or '') == 'Y' or '지명' in (item.get('cntrctCnclsMthdNm') or ''):
        return '지명경쟁'
    if (item.get('mnfctYn') or '') == 'Y':
        return '제조업체만 참가'
    method = item.get('sucsfbidMthdNm') or ''
    # 4억 이상 적격심사 제외는 시공경험 점수 때문 — 물품(구매·설치) 적격심사엔 안 씀(11-19 추산초 8억 물품)
    if item.get('_kind') != '물품' and '적격심사' in method and _int(item.get('presmptPrce')) >= HARD_PRICE:
        return '4억 이상 적격심사'
    return ''


# 검색어 '난방'이 딸려 오는 다른 공사 — 이 말을 지운 뒤에도 검색어가 남아야 냉난방 공사로 본다(11-18 '지역난방설비동 채광창 교체')
NOT_HVAC_WORDS = ('지역난방', '재난방송', '재난방지', '피난방', '난방비')
# 냉난방 말이 들어가도 우리 일이 아닌 것(11-20 공동주택 실측: '난방배관 설비보호제 납품', '난방배관 교체공사 감리업자 선정')
NOT_OUR_WORK = ('설비보호제', '감리')


def is_relevant(title: str, kind: str = '공사') -> bool:
    """물품은 '구매 및 설치'만(단순 구매는 제조·공급업체 몫)."""
    t = title or ''
    if kind == '물품' and '설치' not in t:
        return False
    if any(w in t for w in NOT_OUR_WORK):
        return False
    for w in NOT_HVAC_WORDS:
        t = t.replace(w, '')
    return any(k in t for k in KEYWORDS)


def classify(item: dict, regions: List[str], groups: List[List[str]],
             held: Iterable[str], planned: Iterable[str]) -> dict:
    """입찰 불가능 사유는 걸리는 것 모두 적는다(11-18 부장님: 탭은 가능/불가능, 사유는 내용에).
    분류: eligible=가능 / license_only=준비 중 면허만 없으면 됨 / other=그 밖에 불가능."""
    hard = hard_reason(item)
    lic_ok, missing = license_status(groups, held, planned)
    reasons = []
    if not is_relevant(item.get('bidNtceNm') or '', item.get('_kind') or '공사'):
        reasons.append('설치 없는 물품 구매' if item.get('_kind') == '물품' else '냉난방 공사 아님(제목)')
    if regions and not region_ok(regions):      # 지역 제한 없는 전국 공고는 들어갈 수 있다(11-18)
        reasons.append('김포 업체 참가 불가(지역 제한)')
    if hard:
        reasons.append(f'{hard} 있음' if hard in ('실적제한', '사전심사(PQ)') else hard)   # 제조업체만 참가 포함
    title = item.get('bidNtceNm') or ''
    if not groups:   # 면허 제한이 없다고 아무나 되는 게 아님 — 전기·통신·소방 공사는 그 업 등록이 있어야
        need = next((lic for word, lic in ELEC_WORDS if word in title), '')
        if need and _norm_lic(need) not in {_norm_lic(x) for x in held}:
            reasons.append(f'면허 안 맞음({need})')
    if not lic_ok:
        if missing:
            reasons.append(f"{', '.join(missing)} 면허 없음")
        else:
            reasons.append('면허 안 맞음(' + ' 또는 '.join(' + '.join(g) for g in groups) + ')')
    if not reasons:
        return {'category': ELIGIBLE, 'reason': '', 'missing': []}
    only_license = len(reasons) == 1 and not lic_ok and bool(missing)
    return {'category': LICENSE_ONLY if only_license else OTHER, 'reason': ' · '.join(reasons), 'missing': missing}


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


def _get(op: str, params: dict, base: str = BASE) -> dict:
    p = {'serviceKey': _key(), 'type': 'json', **params}
    try:
        r = requests.get(base + op, params=p, timeout=(10, 30))
    except (requests.ConnectionError, requests.Timeout):   # 공공데이터포털이 가끔 연결을 놓침 — 한 번만 다시
        r = requests.get(base + op, params=p, timeout=(10, 30))
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


def search(keyword: str, start: datetime, end: datetime, op: str = OP_LIST) -> List[dict]:
    out, page = [], 1
    while True:
        body = _get(op, {'inqryDiv': 1, 'inqryBgnDt': start.strftime('%Y%m%d%H%M'),
                              'inqryEndDt': end.strftime('%Y%m%d%H%M'), 'bidNtceNm': keyword,
                              'numOfRows': 100, 'pageNo': page})
        items = _items(body)
        out.extend(items)
        total = _int(body.get('totalCount'))
        if not items or page * 100 >= total or page >= 10:
            return out
        page += 1


def fetch_limits(bid_no: str, bid_ord: str, source: str = '나라장터') -> tuple:
    if source == '공동주택':   # K-apt 는 면허·지역 제한 조회가 없음(공고 화면·첨부에만)
        return [], []
    q = {'inqryDiv': 2, 'bidNtceNo': bid_no, 'bidNtceOrd': bid_ord, 'numOfRows': 50, 'pageNo': 1}
    nuri = source == '누리장터'
    base, op_r, op_l = (NURI_BASE, NURI_REGION, NURI_LICENSE) if nuri else (BASE, OP_REGION, OP_LICENSE)
    regions = sorted({(x.get('prtcptPsblRgnNm') or '').strip()
                      for x in _items(_get(op_r, q, base)) if x.get('prtcptPsblRgnNm')})
    grp: Dict[str, List[str]] = {}
    for x in _items(_get(op_l, q, base)):
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
    return sort_files(files)


def sort_files(files: List[dict]) -> List[dict]:
    """첨부는 나라장터가 준 순서가 03·04·01… 처럼 섞여 온다 → 이름 앞 번호 순, 번호 없는 건 뒤에 받은 순서대로."""
    def key(x):
        i, f = x
        m = re.match(r'\s*[★*\[(]?\s*(\d+)', f.get('name') or '')
        return (0, int(m.group(1)), i) if m else (1, 0, i)
    return [f for _, f in sorted(enumerate(files), key=key)]


def _contact(item: dict) -> str:
    name, tel = (item.get('ntceInsttOfclNm') or '').strip(), (item.get('ntceInsttOfclTelNo') or '').strip()
    return ' '.join(x for x in (name, tel) if x)


def nuri_list(op: str, start: datetime, end: datetime) -> List[dict]:
    out, page = [], 1
    while True:
        body = _get(op, {'inqryDiv': 1, 'inqryBgnDt': start.strftime('%Y%m%d%H%M'),
                         'inqryEndDt': end.strftime('%Y%m%d%H%M'), 'numOfRows': 500, 'pageNo': page}, NURI_BASE)
        items = _items(body)
        out.extend(items)
        if not items or page * 500 >= _int(body.get('totalCount')) or page >= 5:
            return out
        page += 1


def nuri_normalize(it: dict) -> dict:
    """누리장터 항목 이름을 나라장터 이름으로 맞춤 — 나머지 코드(분류·저장·상세)를 그대로 쓰려고."""
    n = dict(it)
    pairs = (('ntceNm', 'bidNtceNm'), ('nticeDt', 'bidNtceDt'), ('ofclNm', 'ntceInsttOfclNm'),
             ('ofclTelNo', 'ntceInsttOfclTelNo'), ('ofclEmail', 'ntceInsttOfclEmailAdrs'),
             ('cntrctMthdNm', 'cntrctCnclsMthdNm'), ('ntceDivNm', 'ntceKindNm'),
             ('sptDscrptDt', 'dcmtgOprtnDt'), ('sptDscrptPlce', 'dcmtgOprtnPlce'))
    for a, b in pairs:
        if it.get(a) and not n.get(b):
            n[b] = it[a]
    for i in range(1, 11):
        if it.get(f'ntceSpecDocNm{i}'):
            n[f'ntceSpecFileNm{i}'] = it[f'ntceSpecDocNm{i}']
    ref, asign = _int(it.get('refAmt')), _int(it.get('asignBdgtAmt'))
    n['presmptPrce'] = str(ref or asign)          # 기준금액이 있으면 그것, 없으면 배정예산
    if asign:
        n['bdgtAmt'] = str(asign)
    # 공사 장소: '[1^공사명^제주시 탑동로 66 ...^2026-12-31]' 의 세 번째 칸
    m = re.match(r'\[[^\^\]]*\^[^\^\]]*\^([^\^\]]*)', it.get('cnstwkDtlList') or '')
    if m and m.group(1).strip():
        n['cnstrtsiteRgnNm'] = m.group(1).strip()
    n['_source'] = '누리장터'
    return n


def apt_list(start: datetime, end: datetime) -> List[dict]:
    out, page = [], 1
    while True:
        body = _get(APT_OP, {'_type': 'json', 'startDate': start.strftime('%Y%m%d'), 'endDate': end.strftime('%Y%m%d'),
                             'numOfRows': 1000, 'pageNo': page}, APT_BASE)
        items = _items(body)
        out.extend(items)
        if not items or page * 1000 >= _int(body.get('totalCount')) or page >= 5:
            return out
        page += 1


def apt_normalize(it: dict) -> dict:
    """K-apt 항목을 나라장터 이름으로. 공고번호는 'KAPT-' 를 붙여 나라장터 번호와 안 겹치게."""
    num = str(it.get('bidNum') or '')
    area = APT_AREA.get(_int(it.get('bidArea')), '')
    auth = APT_AUTH.get(str(it.get('codeAuth') or ''), '')
    n = {
        'bidNtceNo': f'KAPT-{num}', 'bidNtceOrd': '001' if str(it.get('bidState')) in ('2', '3') else '000',
        'bidNtceNm': (it.get('bidTitle') or '').strip(), 'ntceInsttNm': (it.get('bidKaptname') or '').strip(),
        'dminsttNm': (it.get('bidKaptname') or '').strip(),
        'bidNtceDt': str(it.get('bidRegDate') or '')[:10], 'bidClseDt': str(it.get('bidDeadline') or '')[:16],
        'cntrctCnclsMthdNm': APT_KIND.get(str(it.get('codeKind') or ''), ''),
        'sucsfbidMthdNm': APT_SUC.get(str(it.get('codeSucWay') or ''), ''),
        'cnstrtsiteRgnNm': area,
        'dcmtgOprtnDt': str(it.get('bidFieldDesDate') or '')[:16], 'dcmtgOprtnPlce': it.get('bidFieldDesLoc') or '',
        'bidNtceDtlUrl': f'https://www.k-apt.go.kr/bid/bidDetail.do?bidNum={num}' if num else '',
        'reNtceYn': 'Y' if str(it.get('bidState')) == '3' else 'N',
        'aptAuth': auth, 'aptPerfCert': it.get('bidMngCertYn') or '', 'aptCreditCert': it.get('bidCreditConfYn') or '',
        'aptContent': (it.get('bidContent') or '').strip(), 'aptPay': (it.get('bidTermsOfPayment') or '').strip(),
        '_source': '공동주택', '_kind': APT_CLS.get(str(it.get('codeClassifyType2') or ''), '공사'),
    }
    return {k: v for k, v in n.items() if v not in (None,)}


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
        for kind, op in KINDS:
            for kw in KEYWORDS:
                try:
                    for it in search(kw, start, now, op):
                        no = it.get('bidNtceNo')
                        if not no:
                            continue
                        if kind == '물품' and '설치' not in (it.get('bidNtceNm') or ''):
                            continue   # 단순 구매는 저장도 안 함(이번 주 '난방' 물품 87건 대부분)
                        prev = found.get(no)
                        if prev is None or (it.get('bidNtceOrd') or '') > (prev.get('bidNtceOrd') or ''):
                            it['_kw'] = kw
                            it['_kind'] = kind
                            found[no] = it
                except Exception as e:
                    errors += 1
                    logger.warning(f'[G2B] 검색 실패 ({kind} {kw}): {_safe_err(e)}')
        try:   # 공동주택(K-apt) — 공고일 범위로 통째로(물품도 '설치' 없어도 받음: 아파트 물품은 대개 납품·설치 한 묶음)
            for raw_it in apt_list(start, now):
                it = apt_normalize(raw_it)
                no = it.get('bidNtceNo')
                if no and no != 'KAPT-' and is_relevant(it.get('bidNtceNm') or ''):
                    it['_kw'] = '공동주택'
                    found[no] = it
        except Exception as e:
            errors += 1
            logger.warning(f'[G2B] 공동주택 목록 실패: {_safe_err(e)}')
        for kind, op in NURI_KINDS:   # 누리장터 — 목록 통째로 받아 제목으로
            try:
                for raw_it in nuri_list(op, start, now):
                    it = nuri_normalize(raw_it)
                    no, title = it.get('bidNtceNo'), it.get('bidNtceNm') or ''
                    if not no or not is_relevant(title, kind):
                        continue
                    prev = found.get(no)
                    if prev is None or (it.get('bidNtceOrd') or '') > (prev.get('bidNtceOrd') or ''):
                        it['_kw'] = '누리장터'
                        it['_kind'] = kind
                        found[no] = it
            except Exception as e:
                errors += 1
                logger.warning(f'[G2B] 누리장터 목록 실패 ({kind}): {_safe_err(e)}')
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
                regions, groups = (([], []) if it.get('_source') == '공동주택'   # K-apt 는 제한 조회 없음
                                   else fetch_limits(no, ord_, it.get('_source') or '나라장터'))
            except Exception as e:
                errors += 1
                logger.warning(f'[G2B] 지역·면허 조회 실패 {no}: {_safe_err(e)}')
                continue
            _save(it, regions, groups, held, planned, is_new=no not in known)
            if no in known:
                changed += 1
            else:
                new += 1
        if errors < len(KEYWORDS) * len(KINDS):   # 검색이 하나라도 됐으면 '마지막으로 모은 때' 갱신
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
        'budget': _int(it.get('bdgtAmt') or it.get('asignBdgtAmt')), 'vat': vat,
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
        # 현장설명회(11-18 용인예술과학대 — 당일에야 알았음). 참석 필수 여부는 API 에 없어 공고문 확인
        'briefing_at': t('dcmtgOprtnDt')[:16], 'briefing_place': t('dcmtgOprtnPlce'),
        # 물품(구매·설치) — 나라장터 물품 분류 등록이 있어야 투찰됨. 등록 여부는 API 로 알 수 없어 안내만
        'products': _products(t('purchsObjPrdctList')) or (
            [f"{t('dtilPrdctClsfcNoNm')}({t('dtilPrdctClsfcNo')})"] if t('dtilPrdctClsfcNo') else []),
        'product_limit': _yn(it.get('prdctClsfcLmtYn')), 'maker_only': _yn(it.get('mnfctYn')),
        'deliver_days': t('dlvrDaynum'), 'deliver_cond': t('dlvryCndtnNm'),
        'product_codes': [c for c in re.findall(r'\^(\d{8,10})\^', t('purchsObjPrdctList'))] or (
            [t('dtilPrdctClsfcNo')] if t('dtilPrdctClsfcNo') else []),
        'award_std': t('sucsfbidMthdAppStd'),
        # 공동주택(K-apt) — 입찰 시스템·실적증명서·신용평가 제출 여부·지급조건
        'apt_auth': t('aptAuth'), 'apt_perf_cert': _yn(it.get('aptPerfCert')),
        'apt_credit_cert': _yn(it.get('aptCreditCert')), 'apt_content': t('aptContent'), 'apt_pay': t('aptPay'),
        # 누리장터 — 참가자격·지역제한 문구(대개 '공고서 참조'), 부가세 포함 안내
        'nuri_qual': t('bidQlfctNm'), 'nuri_region': t('rgnLmtDivNm'), 'vat_note': t('vatInclsnYnNm'),
    }


def _products(raw: str) -> List[str]:
    """'[1^4010178702^냉난방기],[1^4010180601^전기히트펌프]' → ['냉난방기(4010178702)', ...]"""
    out = []
    for m in re.finditer(r'\[([^\]]*)\]', raw or ''):
        parts = m.group(1).split('^')
        if len(parts) >= 3 and parts[2]:
            out.append(f'{parts[2]}({parts[1]})')
    return out


def _save(it: dict, regions, groups, held, planned, is_new: bool) -> None:
    c = classify(it, regions, groups, held, planned)
    row = {
        'bid_no': it['bidNtceNo'], 'bid_ord': it.get('bidNtceOrd') or '000',
        'title': (it.get('bidNtceNm') or '').strip(), 'org': (it.get('ntceInsttNm') or '').strip(),
        'demand_org': (it.get('dminsttNm') or '').strip(),
        'contract_method': it.get('cntrctCnclsMthdNm') or '', 'award_method': it.get('sucsfbidMthdNm') or '',
        'price': _int(it.get('presmptPrce')), 'budget': _int(it.get('bdgtAmt') or it.get('asignBdgtAmt')),
        'kind': it.get('_kind') or '공사',
        'source': it.get('_source') or '나라장터',
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
            try:
                full = json.loads(r.get('raw') or '{}')
            except ValueError:
                full = {}
            item = {'arsltCmptYn': 'Y' if r['hard_reason'] == '실적제한' else 'N',
                    'pqEvalYn': 'Y' if 'PQ' in r['hard_reason'] else 'N',   # 옛 'PQ(사전심사)'도
                    'dsgntCmptYn': 'Y' if r['hard_reason'] == '지명경쟁' else 'N',
                    'cntrctCnclsMthdNm': r['contract_method'], 'sucsfbidMthdNm': r['award_method'],
                    'presmptPrce': r['price'], 'bidNtceNm': r['title']}
            if full:   # 원본 항목이 있으면 그걸로(제조업체 한정 등 새 기준도 반영)
                item = {**full}
            item['_kind'] = r.get('kind') or '공사'
            item['_source'] = r.get('source') or '나라장터'
            c = classify(item, json.loads(r['regions']), json.loads(r['license_groups']), held, planned)
            hr = hard_reason(item)
            if c['category'] != r['category'] or c['reason'] != r['reason'] or hr != r['hard_reason']:
                conn.execute('UPDATE g2b_bids SET category=?, reason=?, missing_licenses=?, hard_reason=? WHERE bid_no=?',
                             (c['category'], c['reason'], ', '.join(c['missing']), hr, r['bid_no']))
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
        else:   # 보유 면허·사유 문구가 바뀌었을 수 있어 재시작마다 다시 분류(API 안 부름)
            n = reclassify()
            if n:
                logger.info(f'[G2B] 다시 분류 {n}건')
    except Exception as e:
        logger.warning(f'[G2B] 첫 채우기 실패: {_safe_err(e)}')


# ── 화면용 ────────────────────────────────────────────────────────────
# 지역 고르기(11-20 부장님 — 목록이 길어 지역으로 좁혀 보기). 공사는 공사현장, 없으면 수요·공고기관 이름으로.
#   이름에 지역이 없는 기관(대학·공기업 일부)은 '기타'. 순서가 중요: '전남광주'는 호남, '경기'는 '경북'보다 먼저 볼 필요 없음(글자 다름)
AREAS = (
    ('수도권', ('서울', '경기', '인천')),
    ('충청', ('대전', '세종', '충북', '충남', '충청')),
    ('호남', ('광주', '전북', '전남', '전라')),
    ('영남', ('부산', '대구', '울산', '경북', '경남', '경상')),
    ('강원·제주', ('강원', '제주')),
)


def area_of(*names: str) -> str:
    for name in names:
        n = name or ''
        for area, words in AREAS:
            if any(w in n for w in words):
                return area
    return '기타'



def _view(r: dict, now_s: str) -> dict:
    files = sort_files(json.loads(r.get('files') or '[]'))   # 정렬 전에 저장된 줄도
    groups = json.loads(r.get('license_groups') or '[]')
    regions = json.loads(r.get('regions') or '[]')
    return {
        'bid_kind': r.get('kind') or '공사',   # 공사 / 물품(구매·설치)
        'source': r.get('source') or '나라장터',
        'area': area_of(r.get('site_region') or '', r.get('demand_org') or '', r.get('org') or ''),
        # 펼친 '확인할 것' 판정용 — 지역·면허 각각 되는지
        'region_ok': region_ok(regions), 'nationwide': not regions,
        'license_ok': license_status(groups, held_licenses())[0],
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
        # New 뱃지 = 나라장터 게시 24시간 안(11-18 부장님 — 처음 모은 때 기준이면 첫날 전부 New)
        'is_new': bool(r['posted_at']) and r['posted_at'] >= (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d %H:%M'),
        'hard_reason': r.get('hard_reason', ''),
        'detail': _detail(r.get('raw') or ''),
        'ours': _ours(r),
    }


def _ours(r: dict) -> dict:
    """우리 회사 정보와 대조 — 물품 분류 등록·적격심사 85점(11-20)."""
    p = get_profile()
    try:
        it = json.loads(r.get('raw') or '{}')
    except ValueError:
        it = {}
    it['_kind'] = r.get('kind') or '공사'
    out = {'grade': p.get('credit_grade') or '', 'entity': p.get('entity') or ''}
    codes = [c for c in re.findall(r'\^(\d{8,10})\^', it.get('purchsObjPrdctList') or '')] or (
        [it['dtilPrdctClsfcNo']] if it.get('dtilPrdctClsfcNo') else [])
    if codes and it['_kind'] == '물품':
        mine = {x['code'] for x in p.get('products') or []}
        out['products_missing'] = [c for c in codes if c not in mine]
        out['products_known'] = bool(mine)
    table = qual_table(it)
    if table:
        lower = 0.0
        try:
            lower = float(it.get('sucsfbidLwltRate') or 0)
        except ValueError:
            pass
        out['qual'] = qual_need(table, p.get('credit_grade') or '', float(p.get('bonus') or 0),
                                bool(p.get('startup')), lower) or {'table': table}
    return out


def list_bids(tab: str = ELIGIBLE, include_closed: bool = False, q: str = '', limit: int = 300) -> dict:
    now_s = datetime.now().strftime('%Y-%m-%d %H:%M')
    with _conn() as conn:
        # 탭 = 입찰 가능(eligible) / 입찰 불가능(blocked = 그 밖 전부, 사유는 줄마다)
        counts = {ELIGIBLE: 0, BLOCKED: 0, LICENSE_ONLY: 0}
        for cat, n in conn.execute('SELECT category, COUNT(*) FROM g2b_bids WHERE close_at >= ? GROUP BY category',
                                   (now_s,)):
            counts[ELIGIBLE if cat == ELIGIBLE else BLOCKED] += n
            if cat == LICENSE_ONLY:
                counts[LICENSE_ONLY] += n
        if tab == BLOCKED:
            # 준비 중 면허만 없는 공고를 먼저(면허 따면 바로 들어갈 수 있는 것)
            sql, args = 'SELECT * FROM g2b_bids WHERE category != ?', [ELIGIBLE]
        else:
            sql, args = 'SELECT * FROM g2b_bids WHERE category = ?', [ELIGIBLE]
        if not include_closed:
            sql += ' AND close_at >= ?'
            args.append(now_s)
        if q:
            sql += ' AND (title LIKE ? OR org LIKE ? OR demand_org LIKE ?)'
            args += [f'%{q}%'] * 3
        sql += (" ORDER BY (close_at < ?) ASC, (category != 'license_only') ASC,"
                ' CASE WHEN close_at < ? THEN close_at END DESC, close_at ASC LIMIT ?')
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
    if 0 < n < 1000:   # 발주처가 1원 등으로 넣은 금액 비공개 공고
        return '비공개(공고문 확인)'
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
        files = sort_files(json.loads(r.get('files') or '[]'))
        site = 'K-apt' if r.get('source') == '공동주택' else '나라장터'
        links = [f"<{r['detail_url']}|{site}에서 보기>"] if r.get('detail_url') else []
        for f in files[:4]:
            name = f['name']
            kind = '공고문' if '공고' in name else '내역서' if '내역' in name else '시방서' if '시방' in name else ''
            if not kind:
                continue   # 알림에는 공고문·내역서·시방서만 (나머지는 그룹웨어에서)
            ext = name.rsplit('.', 1)[-1].upper() if '.' in name else ''
            links.append(f"<{f['url']}|{kind}{' ' + ext if ext else ''}>")
        lines.append(
            f"\n• *{r['title']}*\n"
            f"   {r['demand_org'] or r['org']}"
            f"{' (' + r['source'] + ')' if r.get('source') in ('누리장터', '공동주택') else ''} · "
            f"{(r['award_method'] or r['contract_method']).split('-')[0]} · "
            f"추정가격 {_won(r['price'])}\n"
            f"   마감 {_when(r['close_at'])} · 참가가능지역: {', '.join(json.loads(r['regions']))}\n"
            + _briefing_line(r) +
            f"   {' · '.join(links)}")
    if license_only:
        lines.append(f"\n냉매회수업 등 면허만 없어서 못 들어가는 새 공고도 {license_only}건 있어요.")
    lines.append(f"\n<{page_url}|그룹웨어에서 전체 보기>")
    return '\n'.join(lines)


def _briefing_line(r: dict) -> str:
    d = _detail(r.get('raw') or '')
    at = d.get('briefing_at') or ''
    if not at:
        return ''
    gone = at < datetime.now().strftime('%Y-%m-%d %H:%M')
    place = f" · {d['briefing_place']}" if d.get('briefing_place') else ''
    warn = ' — 이미 지났어요(참석 필수인지 공고문 확인)' if gone else ''
    return f"   :warning: 현장설명회 {_when(at)}{place}{warn}\n"


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
