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


# check_orphan_settlement(미수금 0인데 수금카드 없음)은 2026-09-14 제거.
#   전제("수금완료면 카드 있어야")가 거짓 — ①수기 입력 입금은 카드를 안 만듦 ②payment_slack
#   카드 키는 90일 TTL이라 카드를 보냈어도 만료돼 '없음'으로 잡힘. 즉 "카드 키 없음 ≠ 미발송"
#   이라 정밀화 불가(수금날짜든 메모날짜든 노이즈). 진짜 미발송은 payment_sync 폴러가 이미
#   복구하므로 가치 중복. 재추가 금지.
INVARIANTS = [
    check_phantom_codes,
    check_amount_anomaly,
]

_CHECK_LABEL = {
    'phantom_code': '유령/고아 코드',
    'amount_anomaly': '금액 이상치',
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

    return {
        'records': records,
        'valid_codes': valid_codes,
        'card_codes': card_codes,
        'now': datetime.now(),
    }


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
