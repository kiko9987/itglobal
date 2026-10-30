"""정합성/불변식 점검 — "조용히 틀린 값" 자동 감지 (2026-09-11, 관측성 확장).

개념: 운영 데이터에 대해 '항상 참이어야 하는 단언'을 주기적으로 돌려, 현실이 어기면
관리자 슬랙으로 알림. 예외/크래시는 안 나지만 결과가 틀린 사각지대(유령/고아 코드·금액
이상치 등)를 잡는다. 과거 사고 하나하나를 영구 불변식으로 박는 문화.

설계
----
  * Redis 스캔은 build_context() 가 한 번에 수행해 집합으로 만들고, 각 check_* 는
    ctx(레코드 + 미리계산 집합)만 받는 **순수 함수** → 유닛 테스트 쉽고 오탐 잡기 좋음.
  * 각 검사는 [Violation] 반환. runner 가 Redis dedup(하루 1회) 후 신규만 슬랙 1건으로 묶어 발송.
  * 안전장치(reconciler 패턴): mutex · 검사별 예외 격리 · 검사당 위반 상한(systemic 폭주 방어).

환경변수
--------
  INVARIANT_CHECKS_ENABLED (기본 '1')
  INVARIANT_ALERT_CHANNEL  (기본 SLACK_ADMIN_CHANNEL)
"""
import hashlib
import json
import os
import re
import time
import urllib.request
from datetime import datetime

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_MUTEX_KEY = 'invariant_checks:running'
_MUTEX_TTL = 300
_SEEN_PREFIX = 'invariant_seen:'
_SEEN_TTL = 20 * 3600                 # 동일 위반 하루 1회 (20h)
_MAX_PER_CHECK = 25                   # 이 이상 = systemic → 상세 대신 요약(폭주 방어)
_AMOUNT_MAX = 99_999_999_999         # 9999억 (validate_amount 와 정렬)
# 금액(통화) 필드만. '부가세'는 통화가 아니라 부가세 포함여부 불리언(True/False)이라 제외.
_AMOUNT_FIELDS = ('총액 1', '계약금', '중도금', '잔금')

_NUM_CLEAN_RE = re.compile(r'[,\s원₩]')


# ─────────────────────────────────────────────────────────────
# 유틸
# ─────────────────────────────────────────────────────────────
def _dec(v):
    return v.decode() if isinstance(v, bytes) else v


def _parse_amount(raw):
    """금액 문자열 → float. 파싱 불가면 None, 빈값/NaN 이면 'empty'."""
    if raw is None:
        return 'empty'
    s = str(raw).strip()
    if s.lower() in ('', '-', 'nan', 'none'):   # 빈 숫자셀은 pandas NaN → 'nan'
        return 'empty'
    s = _NUM_CLEAN_RE.sub('', s)
    try:
        f = float(s)
    except (ValueError, TypeError):
        return None
    if f != f or f in (float('inf'), float('-inf')):  # NaN/Inf → 빈값 취급
        return 'empty'
    return f


class Violation(dict):
    """위반 1건. key(dedup용), check(분류), title(코드 등), detail."""
    def __init__(self, check, key, title, detail):
        super().__init__(check=check, key=key, title=title, detail=detail)


# ─────────────────────────────────────────────────────────────
# 불변식 (순수 함수 — ctx 만 받음)
# ─────────────────────────────────────────────────────────────
def check_phantom_codes(ctx):
    """카드/입금 기록이 공사현황에 없는 프로젝트 코드를 가리킴 (re-key 고아·유령 코드)."""
    valid = ctx['valid_codes']
    out = []
    for code in sorted(ctx['card_codes']):
        if code and code not in valid:
            out.append(Violation(
                'phantom_code', f'phantom:{code}', code,
                '슬랙 카드/입금 기록이 공사현황에 없는 코드를 가리킴 (re-key 고아 또는 유령 코드)'))
    return out


def check_amount_anomaly(ctx):
    """금액 필드가 파싱 불가·음수·범위 초과 (수식셀 행번호 오추출·컬럼 시프트 등)."""
    out = []
    for r in ctx['records']:
        code = (r.get('프로젝트 코드') or '').strip()
        if not code:
            continue
        for f in _AMOUNT_FIELDS:
            v = _parse_amount(r.get(f))
            if v == 'empty' or isinstance(v, float) and 0 <= v <= _AMOUNT_MAX:
                continue
            if v is None:
                detail = f"{f} 금액 파싱 불가: {str(r.get(f))[:40]!r}"
            elif v < 0:
                detail = f"{f} 음수 금액: {v:,.0f}"
            else:
                detail = f"{f} 범위 초과({v:,.0f}) — 수식셀 행번호 오추출 의심"
            out.append(Violation('amount_anomaly', f'amount:{code}:{f}', code, detail))
    return out


# Redis set: 처음 본 기존 이상 칸(알림 제외 기준). v2 = 키에 열 글자 대신 필드명 (2026-10-22) —
#   열을 끼워 넣어도 기존 이상 칸이 '새 이상'으로 쏟아지지 않게. v2 가 없으면 첫 실행이 다시 기준을 잡음.
_CELL_KNOWN_KEY = 'invariant_known_cell_anomalies_v2'
_CELL_KNOWN_KEY_V1 = 'invariant_known_cell_anomalies'   # 옛 키 'cell:AH4149:값' (열 글자는 10-22 이전 배치)


# v1 키가 쓰던 열 글자 — 10-22 거래처명 삽입 전 배치(고정). 기본 배치(_STATIC_COLUMN_MAPPING)는 시트를
#   따라 바뀌므로 옮길 때 쓰면 안 된다. v1 이 점검한 칸만.
_V1_CELL_LETTERS = {'T': '총액 2', 'X': '미수금', 'AD': '수금 확인', 'AE': '제품대', 'AF': '도급비', 'AG': '자재비',
                    'AH': '기타비', 'AI': '순익', 'AJ': '마진율'}


def _migrate_known_v1(keys):
    """v1 'cell:{열}{행}:{값}' → v2 'cell:{필드명}:{행}:{값}'. 열 글자는 10-22 이전 배치로 읽는다."""
    out = set()
    for k in keys:
        m = re.match(r'^cell:([A-Z]+)(\d+):(.*)$', str(k), re.S)
        if m and m.group(1) in _V1_CELL_LETTERS:
            out.add(f'cell:{_V1_CELL_LETTERS[m.group(1)]}:{m.group(2)}:{m.group(3)}')
    return out
_CELL_FIELDS = (('총액 2', 'err'), ('미수금', 'err'), ('수금 확인', 'bool'),
                ('제품대', 'num'), ('도급비', 'num'), ('자재비', 'num'), ('기타비', 'num'),
                ('순익', 'err'), ('마진율', 'err'))


def sheet_cell_anomalies(rows, f2l=None):
    """시트 원본 칸 이상 — 수식 오류(#REF! 등)·금액 칸 글자·수금 확인 체크박스 칸 글자.

    rows: [(행번호, 코드, {필드명: 값})] (UNFORMATTED). Returns [(key, 코드, 문구)].
    계기: 2026-10-19 14:58 브라우저에서 칸을 끌어 옮겨 AL 글자가 AH(기타비)·AD(수금 확인)로 가고
    AI(순익) 수식이 #REF! — 프로젝트 레코드는 값이 정리돼 들어와 기존 점검이 못 잡았다.
    문구의 열 글자는 헤더 지도(f2l)에서, 기억 키는 필드명 기준(열 이동 무관).
    """
    if f2l is None:
        from dashboard.utils.google_sheets import main_field_to_letter
        f2l = main_field_to_letter()
    out = []
    for row, code, cells in rows:
        for label, kind in _CELL_FIELDS:
            v = cells.get(label, '')
            letter = f2l.get(label, '?')
            msg = ''
            if isinstance(v, str) and v.startswith('#'):
                msg = f"{label}({letter}{row}) 수식 오류 {v.split(' ')[0]}"
            elif kind == 'bool' and not isinstance(v, bool) and v not in ('', None):
                msg = f"{label}({letter}{row}) 체크박스 칸에 값 {str(v)[:30]!r}"
            elif kind == 'num' and isinstance(v, str) and v.strip() not in ('', '-'):
                msg = f"{label}({letter}{row}) 금액 칸에 글자 {v[:30]!r}"
            if msg:
                out.append((f'cell:{label}:{row}:{str(v)[:40]}', code, msg))
    return out


def check_sheet_cells(ctx):
    """시트 원본 칸 이상 중 **새로 생긴 것만** (처음 실행 때 있던 106칸은 기준으로 기억 — 옛 #DIV/0! 등)."""
    known = ctx.get('known_cell_anomalies') or set()
    return [Violation('sheet_cell', key, code, msg)
            for key, code, msg in ctx.get('cell_anomalies') or [] if key not in known]


# 칸 참조(A1·$U$12) — 앞뒤가 글자·숫자면 아님(LOG10 같은 함수 이름), 뒤에 '(' 면 함수(ATAN2( 등)
_CELL_REF_RE = re.compile(r'(?<![A-Za-z0-9_$])\$?([A-Z]{1,3})\$?(\d+)(?![\dA-Za-z_(])')


def formula_row_mismatches(rows, f2l=None):
    """수식 칸이 자기 행이 아닌 다른 행을 계산 — [(key, 코드, 문구)].

    rows: [(행번호, 코드, {열 글자: 수식})] (FORMULA). 다른 탭('!')·전체 열(A:A) 참조는 대상 아님.
    계기: 2026-10-28 G4153(4154행) 총액2·미수금·순익·마진율이 4153행(G4152)을 가리켜 시트와 입금 카드가
    G4152 값(받을 돈 2,248,400)을 보였다. PM 은 서버가 따로 계산해 정상이라 아무도 몰랐다.
    """
    if f2l is None:
        from dashboard.utils.google_sheets import main_field_to_letter
        f2l = main_field_to_letter()
    l2f = {v: k for k, v in f2l.items()}
    out = []
    for row, code, cells in rows:
        for letter, f in cells.items():
            if not isinstance(f, str) or not f.startswith('=') or '!' in f:
                continue
            others = sorted({int(n) for _, n in _CELL_REF_RE.findall(re.sub(r'"[^"]*"', '', f))} - {row})
            if others:
                label = l2f.get(letter, letter)
                out.append((f'formula_row:{code}:{label}:{row}', code,
                            f"{label}({letter}{row}) 수식이 {', '.join(map(str, others))}행을 계산 — "
                            f"시트·입금 카드에 다른 공사 값이 보임"))
    return out


def check_formula_rows(ctx):
    """수식 칸이 다른 행을 가리킴 — 시트 전체, 자기 행으로 고칠 때까지 매일 알림."""
    return [Violation('formula_row', key, code, msg) for key, code, msg in ctx.get('formula_mismatches') or []]


# 채널별로 영업일 기준 이만큼 새 문의가 없으면 '유입 멈춤 의심' (2026-10-28 5차). 지난 1년 하루 평균:
#   당근 ~5건·전화 ~4건·홈페이지 ~2건. 아임웹 알림 메일 문구가 바뀌거나 당근 연동 시트가 멈추면 문의가 0건으로
#   조용히 이어지는데 알려 주는 곳이 없었다.
_INFLOW_RULES = {'당근': 1, '전화': 2, '홈페이지': 2}


def _is_business_day(d) -> bool:
    if d.weekday() >= 5:
        return False
    try:
        import holidays
        return d not in holidays.KR()
    except Exception:
        return True


def lead_inflow_gaps(leads, today=None):
    """[(채널, 마지막 문의 날짜, 그 뒤로 지난 영업일 수)] — 규칙 이상 비었을 때만. 오늘은 셈에서 뺌."""
    from datetime import date, timedelta
    from dashboard.services.lead_sync import _parse_consult_dt
    today = today or date.today()
    last = {}
    for l in leads or []:
        p = str(l.get('플랫폼') or '').strip()
        if p not in _INFLOW_RULES:
            continue
        dt = _parse_consult_dt(l.get('상담 시간'))
        if dt and (p not in last or dt.date() > last[p]):
            last[p] = dt.date()
    out = []
    for p, need in _INFLOW_RULES.items():
        d0 = last.get(p)
        if not d0:
            continue
        gap = sum(1 for k in range(1, (today - d0).days) if _is_business_day(d0 + timedelta(days=k)))
        if gap >= need:
            out.append((p, d0, gap))
    return out


def check_lead_inflow(ctx):
    """채널별 새 문의가 영업일 기준 일정 기간 없음 — 메일 문구 변경·연동 중단 의심. 영업일에만 알림."""
    from datetime import date
    if not _is_business_day(ctx.get('today') or date.today()):
        return []
    return [Violation('lead_inflow', f'inflow:{p}:{d0}', p,
                      f'{d0.month}/{d0.day} 이후 새 문의 없음 (영업일 {gap}일) — 알림 메일 문구 변경·연동 중단 확인')
            for p, d0, gap in ctx.get('lead_inflow_gaps') or []]


def check_overpay_lingering(ctx):
    """미수금이 마이너스(과입금)인 채 7일 넘게 남은 공사 — 반환했는데 시트 기록이 빠진 것일 수 있음 (2026-10-28 R4182).

    입금 메모의 '⚠️ 과입금' 표시 날짜(메모 첫 줄 날짜) 기준. 날짜를 못 읽으면 알리지 않음(오탐 방지)."""
    from datetime import date, datetime
    today = ctx.get('today') or date.today()
    out = []
    for r in ctx['records']:
        code = (r.get('프로젝트 코드') or '').strip()
        v = _parse_amount(r.get('미수금'))
        if not code or not isinstance(v, float) or v > -2:
            continue
        if re.search(r'공사\s*취소', str(r.get('수금 관련 특이사항') or '')):
            continue
        when = None
        for s in ('잔금', '중도금', '계약금'):
            note = str(r.get(f'{s}_메모') or '')
            if '과입금' not in note:
                continue
            ds = re.findall(r'(\d{4})/(\d{1,2})/(\d{1,2})', note)
            if ds:
                try:
                    when = max(datetime(int(y), int(m), int(d)).date() for y, m, d in ds)
                except ValueError:
                    when = None
            break
        if when and (today - when).days > 7:
            out.append(Violation('overpay_lingering', f'overpay:{code}:{int(v)}', code,
                                 f'과입금 {-int(v):,}원이 {(today - when).days}일째 남음 — 반환했으면 반환 카드로 기록, '
                                 f'안 돌려줬으면 처리 방법 결정'))
    return out


def backup_stale(backup_root, now=None, max_hours: int = 26) -> list:
    """일일 백업(03:15) 폴더별 가장 최근 파일이 max_hours 넘게 오래됐으면 [(종류, 시간)] (없으면 None 시간).

    03:15 에 서버가 꺼져 있으면 그날 백업은 조용히 건너뛰어짐(1분 유예) — 백업이 멈춰도 알 수 없었음
    (2026-10-29 점검)."""
    import time as _t
    from pathlib import Path
    now = now or _t.time()
    out = []
    for kind in ('redis', 'users_db', 'secrets'):
        d = Path(backup_root) / kind
        try:
            newest = max((p.stat().st_mtime for p in d.iterdir()), default=None)
        except OSError:
            newest = None
        if newest is None:
            out.append((kind, None))
        elif (now - newest) / 3600 > max_hours:
            out.append((kind, int((now - newest) // 3600)))
    return out


def check_backup_fresh(ctx):
    """일일 백업이 26시간 넘게 안 된 종류 — 백업 실패·건너뜀 감지."""
    from pathlib import Path
    root = ctx.get('backup_root') or (Path(__file__).resolve().parents[2] / 'backup')
    out = []
    for kind, hours in backup_stale(root, now=ctx.get('now_ts')):
        when = '백업 파일 없음' if hours is None else f'마지막 백업 {hours}시간 전'
        out.append(Violation('backup_stale', f'backup:{kind}:{hours if hours is None else hours // 24}', kind,
                             f'{when} — 일일 백업(03:15)이 안 되고 있음, 서버 로그 [BACKUP] 확인'))
    hours = offsite_stale_hours(root, now=ctx.get('now_ts'),
                                enabled=ctx.get('offsite_enabled', bool(os.getenv('BACKUP_ENCRYPT_PASSWORD'))))
    if hours is not False:
        when = '올린 기록 없음' if hours is None else f'마지막 성공 {hours}시간 전'
        out.append(Violation('backup_stale', f'backup:offsite:{hours if hours is None else hours // 24}', 'offsite',
                             f'외부 백업(경영지원관리Ⅱ) {when} — 서버 로그 [BACKUP] offsite 줄 확인'))
    return out


def offsite_stale_hours(backup_root, now=None, enabled=True, max_hours: int = 26):
    """외부 백업(scripts/offsite_backup.py 가 남기는 offsite_last.json)이 max_hours 넘게 성공 못 했으면 시간
    (기록 없으면 None), 괜찮거나 꺼져 있으면 False (2026-10-30)."""
    from pathlib import Path
    if not enabled:
        return False
    try:
        at = float(json.loads((Path(backup_root) / 'offsite_last.json').read_text(encoding='utf-8'))['at'])
    except Exception:
        return None
    age = ((now or time.time()) - at) / 3600
    return int(age) if age > max_hours else False


def check_sheet_write_failed(ctx):
    """시트 쓰기 데드레터(3회 실패한 편집)가 남아 있음 — 발생 순간 DM 한 번뿐이라 놓치면 영영 묻혔다 (2026-10-30).
    화면(캐시)엔 저장돼 보이지만 시트엔 안 써진 편집. /api/admin/sheet-write-queue 에서 재시도."""
    out = []
    for op in ctx.get('queue_failed') or []:
        p = op.get('payload') or {}
        tag = p.get('tag') or p.get('project_code') or ''
        when = datetime.fromtimestamp(float(op.get('created_at') or 0)).strftime('%m/%d %H:%M')
        day = (ctx.get('now') or datetime.now()).strftime('%Y-%m-%d')    # 남아 있는 동안 하루 한 번씩 다시
        out.append(Violation('sheet_write_failed', f"sheetq:{op.get('op_id', '')}:{day}", tag or op.get('op_type', ''),
                             f"{when} {op.get('op_type', '')} 시트 쓰기 실패 — `{str(op.get('last_error', ''))[:120]}`"))
    return out


def check_split_payers(ctx):
    """나눠 내는 공사(○○ 부담자 칸) — 칸 형식 오류·부담액 합계≠총액2·부담자 없는 칸에 입금 (2026-10-19)."""
    from dashboard.services.payer_split import split_problems
    out = []
    for r in ctx['records']:
        code = (r.get('프로젝트 코드') or '').strip()
        if not code:
            continue
        for msg in split_problems(r):
            out.append(Violation('split_payer', f'split:{code}:{msg}', code, msg))
    return out


# check_orphan_settlement(미수금 0인데 수금카드 없음)은 2026-09-14 제거.
#   전제("수금완료면 카드 있어야")가 거짓 — ①수기 입력 입금은 카드를 안 만듦 ②payment_slack
#   카드 키는 90일 TTL이라 카드를 보냈어도 만료돼 '없음'으로 잡힘. 즉 "카드 키 없음 ≠ 미발송"
#   이라 정밀화 불가(수금날짜든 메모날짜든 노이즈). 진짜 미발송은 payment_sync 폴러가 이미
#   복구하므로 가치 중복. 재추가 금지.
INVARIANTS = [
    check_phantom_codes,
    check_amount_anomaly,
    check_split_payers,
    check_sheet_cells,
    check_formula_rows,
    check_lead_inflow,
    check_overpay_lingering,
    check_backup_fresh,
    check_sheet_write_failed,
]

_CHECK_LABEL = {
    'phantom_code': '유령/고아 코드',
    'amount_anomaly': '금액 이상치',
    'split_payer': '나눠 내는 공사(부담자 칸)',
    'sheet_cell': '시트 칸 이상(수식 오류·칸 이동 의심)',
    'formula_row': '수식 행 어긋남(다른 행을 계산)',
    'lead_inflow': '문의 유입 멈춤 의심',
    'overpay_lingering': '과입금 7일 넘게 남음',
    'backup_stale': '일일 백업 멈춤',
    'sheet_write_failed': '시트에 안 써진 편집(쓰기 실패)',
}


# ─────────────────────────────────────────────────────────────
# Context 빌드 (Redis 스캔 격리)
# ─────────────────────────────────────────────────────────────
def build_context():
    """records + Redis 파생 집합을 한 번에 구성. 검사들은 이 ctx 만 사용(순수)."""
    from dashboard.services.project_service import get_project_records
    from dashboard.utils.redis_client import get_redis_client

    records = get_project_records() or []
    valid_codes = {(_dec(r.get('프로젝트 코드')) or '').strip()
                   for r in records if (r.get('프로젝트 코드') or '').strip()}

    rc = get_redis_client().redis
    card_codes = set()          # 모든 카드/입금 기록이 가리키는 코드 (phantom 검사용)
    try:
        for key in rc.scan_iter('payment_slack:ts:*', count=500):
            parts = _dec(key).split(':')
            if len(parts) >= 4 and parts[2]:
                card_codes.add(parts[2])
        for key in rc.scan_iter('project_card_msg:*', count=500):
            parts = _dec(key).split(':', 1)
            if len(parts) == 2 and parts[1]:
                card_codes.add(parts[1])
    except Exception as exc:
        logger.warning(f'[INVARIANT] Redis 스캔 일부 실패: {exc}')

    # 시트 원본 칸 이상(수식 오류·칸 이동 의심) — 처음 실행 때 있던 것은 기준으로 기억, 새로 생긴 것만 알림
    cell_anomalies, known = [], set()
    try:
        cell_anomalies = sheet_cell_anomalies(_read_sheet_cells())
        if not rc.exists(_CELL_KNOWN_KEY) and rc.exists(_CELL_KNOWN_KEY_V1):
            # v1(열 글자 키) → v2(필드명 키) 옮김 — 옮긴 뒤 새로 생긴 이상만 알림 (2026-10-22)
            migrated = _migrate_known_v1({_dec(x) for x in (rc.smembers(_CELL_KNOWN_KEY_V1) or set())})
            rc.sadd(_CELL_KNOWN_KEY, '__seeded__', *migrated)
            logger.info(f'[INVARIANT] 시트 칸 이상 기준 v1→v2 옮김: {len(migrated)}칸')
        if not rc.exists(_CELL_KNOWN_KEY):
            if cell_anomalies:
                rc.sadd(_CELL_KNOWN_KEY, *[k for k, _, _ in cell_anomalies])
            else:
                rc.sadd(_CELL_KNOWN_KEY, '__seeded__')
            logger.info(f'[INVARIANT] 시트 칸 이상 기준 저장: {len(cell_anomalies)}칸')
        known = {_dec(x) for x in (rc.smembers(_CELL_KNOWN_KEY) or set())}
    except Exception as exc:
        logger.warning(f'[INVARIANT] 시트 칸 점검 준비 실패(건너뜀): {exc}')
        cell_anomalies = []

    formula_mismatches = []
    try:
        formula_mismatches = formula_row_mismatches(_read_sheet_formulas())
    except Exception as exc:
        logger.warning(f'[INVARIANT] 수식 행 점검 준비 실패(건너뜀): {exc}')

    inflow_gaps = []
    try:
        from dashboard.services.lead_service import get_lead_records
        inflow_gaps = lead_inflow_gaps(get_lead_records() or [])
    except Exception as exc:
        logger.warning(f'[INVARIANT] 문의 유입 점검 준비 실패(건너뜀): {exc}')

    queue_failed = []
    try:
        from dashboard.services.sheet_write_queue import peek_failed
        queue_failed = peek_failed(50)
    except Exception as exc:
        logger.warning(f'[INVARIANT] 시트 쓰기 실패 목록 조회 실패(건너뜀): {exc}')

    return {
        'records': records,
        'queue_failed': queue_failed,
        'valid_codes': valid_codes,
        'card_codes': card_codes,
        'cell_anomalies': cell_anomalies,
        'known_cell_anomalies': known,
        'formula_mismatches': formula_mismatches,
        'lead_inflow_gaps': inflow_gaps,
        'now': datetime.now(),
    }


def _read_sheet_cells():
    """공사 현황 코드 + 점검 칸(_CELL_FIELDS) 원본(UNFORMATTED) → [(행, 코드, {필드명: 값})] — API 1회.

    열 위치는 헤더 지도에서 (예전 T·X·AD:AJ 고정, 2026-10-22)."""
    from dashboard.services.payment_sync import _get_payment_service
    from dashboard.utils.google_sheets import main_field_to_letter
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    if not sid or not sn:
        return []
    f2l = main_field_to_letter()
    fields = [f for f, _ in _CELL_FIELDS]
    ranges = [f"'{sn}'!A2:A6000"] + [f"'{sn}'!{f2l[f]}2:{f2l[f]}6000" for f in fields]
    vr = _get_payment_service().spreadsheets().values().batchGet(
        spreadsheetId=sid, valueRenderOption='UNFORMATTED_VALUE', ranges=ranges,
    ).execute().get('valueRanges', [])
    cols = [vr[k].get('values', []) if len(vr) > k else [] for k in range(len(ranges))]
    out = []
    for i, c in enumerate(cols[0]):
        code = str(c[0] if c else '').strip()
        if not code:
            continue
        cells = {f: (cols[j + 1][i][0] if i < len(cols[j + 1]) and cols[j + 1][i] else '')
                 for j, f in enumerate(fields)}
        out.append((i + 2, code, cells))
    return out


def _read_sheet_formulas():
    """공사 현황 전체 행의 수식(FORMULA) → [(행, 코드, {열 글자: 수식})] — API 1회. 수식 칸만 담는다."""
    from dashboard.services.payment_sync import _get_payment_service
    from dashboard.utils.google_sheets import main_field_to_letter
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    if not sid or not sn:
        return []

    def _idx(letter):
        n = 0
        for ch in letter:
            n = n * 26 + ord(ch) - 64
        return n - 1

    def _letter(i):
        s = ''
        i += 1
        while i:
            i, r = divmod(i - 1, 26)
            s = chr(65 + r) + s
        return s

    last = max(main_field_to_letter().values(), key=_idx)
    rows = _get_payment_service().spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{sn}'!A2:{last}6000", valueRenderOption='FORMULA',
    ).execute().get('values', [])
    out = []
    for i, r in enumerate(rows):
        code = str(r[0] if r else '').strip()
        cells = {_letter(j): v for j, v in enumerate(r) if isinstance(v, str) and v.startswith('=')}
        if code and cells:
            out.append((i + 2, code, cells))
    return out


# ─────────────────────────────────────────────────────────────
# Runner
# ─────────────────────────────────────────────────────────────
def run_invariant_checks():
    """전체 불변식 실행 → 신규 위반만 슬랙 1건으로 발송. 스케줄러 진입점."""
    result = {'checks': 0, 'violations': 0, 'new': 0, 'sent': False}

    if os.getenv('INVARIANT_CHECKS_ENABLED', '1').strip().lower() in ('0', 'false', 'no'):
        return result

    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception as exc:
        logger.warning(f'[INVARIANT] Redis 불가: {exc}')
        return result

    try:
        if not rc.set(_MUTEX_KEY, '1', nx=True, ex=_MUTEX_TTL):
            return result
    except Exception:
        pass

    try:
        ctx = build_context()
    except Exception as exc:
        logger.error(f'[INVARIANT] context 빌드 실패: {exc}', exc_info=True)
        return result

    groups = []  # (check_id, [violations], systemic_count)
    for check in INVARIANTS:
        result['checks'] += 1
        try:
            vios = check(ctx) or []
        except Exception as exc:
            logger.error(f'[INVARIANT] 검사 실패 {check.__name__}: {exc}', exc_info=True)
            continue
        result['violations'] += len(vios)
        if not vios:
            continue
        check_id = vios[0]['check']
        systemic = len(vios) if len(vios) > _MAX_PER_CHECK else 0
        groups.append((check_id, vios, systemic))

    # dedup — 신규 위반만
    new_groups = []
    for check_id, vios, systemic in groups:
        fresh = []
        for v in vios:
            h = hashlib.md5(v['key'].encode('utf-8')).hexdigest()[:16]
            seen_key = _SEEN_PREFIX + h
            try:
                if rc.set(seen_key, '1', nx=True, ex=_SEEN_TTL):
                    fresh.append(v)
            except Exception:
                fresh.append(v)
        if fresh:
            new_groups.append((check_id, fresh, systemic))
            result['new'] += len(fresh)

    if new_groups:
        text = _format_message(new_groups, result['new'])
        if _send_slack(text):
            result['sent'] = True

    logger.info(f'[INVARIANT] 점검 완료: {result}')
    return result


def _format_message(new_groups, total_new):
    lines = [f':mag: *정합성 점검 — 신규 {total_new}건*']
    for check_id, vios, systemic in new_groups:
        label = _CHECK_LABEL.get(check_id, check_id)
        if systemic:
            lines.append(f'*{label}* — :warning: {systemic}건 대량 발생(systemic 의심) — 상세 생략, 로그 확인')
            continue
        lines.append(f'*{label}* ({len(vios)})')
        for v in vios[:_MAX_PER_CHECK]:
            lines.append(f'  • `{v["title"]}` — {v["detail"]}')
    lines.append('_하루 1회 dedup · 오탐이면 알려주세요 (INVARIANT_CHECKS_ENABLED=0 으로 중지)_')
    return '\n'.join(lines)


def _send_slack(text):
    token = os.getenv('SLACK_BOT_TOKEN', '').strip()
    channel = (os.getenv('INVARIANT_ALERT_CHANNEL', '').strip()
               or os.getenv('SLACK_ADMIN_CHANNEL', '').strip())
    if not token or not channel:
        logger.warning('[INVARIANT] SLACK_BOT_TOKEN/채널 미설정 — 발송 skip')
        return False
    try:
        body = json.dumps({'channel': channel, 'text': text}).encode('utf-8')
        req = urllib.request.Request(
            'https://slack.com/api/chat.postMessage', data=body,
            headers={'Authorization': f'Bearer {token}',
                     'Content-Type': 'application/json; charset=utf-8'})
        urllib.request.urlopen(req, timeout=8).read()
        return True
    except Exception as exc:
        logger.warning(f'[INVARIANT] 슬랙 발송 실패: {exc}')
        return False
