"""세금계산서·수금 리마인드 (#영업_관리, 세금계산서 관리 알림 봇).

- 주간(매주 월 09시): 💰 수금 필요 — 세금계산서 발행완료인데 미수금 남은 건(선발행 포함).
  발행은 했는데 돈이 안 들어온 건을 놓치지 않게 회수 리마인드.
- 매월 10일 09시(계산서 마감): ① 수금완료·미발행 ② 부분입금·미발행 ③ 수금 필요.

billStatus.js(computeBillStagesFromColumns/isFullyCollected) 규칙을 파이썬으로 미러.
스타일은 pin_remind(미처리 카드) 준용 — 헤더 + 구분선 + 섹션 + 불릿.
"""
import os
import re

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_SEP = '--------------------------------------------'
_BLANK = '⠀'

# 제외(전체): 소송 진행 등 리마인드 부적합.
#   G0862-MW(소송중, 2026-09-10), R2676-JSH((주)매너마인드 소송건, 2026-09-11).
_EXCLUDE = {'G0862-MW', 'R2676-JSH'}

# 수금 리마인드만 제외 — 내부 처리 건(수금 데이터 신뢰 불가). 발행(①②)엔 영향 없음. 2026-09-11.
_COLLECT_EXCLUDE = {
    'G2217-SH', 'G2646-SH', 'G3288-SH', 'R3642-SH', 'G3645-SH',
    'G3646-SH', 'G3647-SH', 'G3766-SH', 'G3857-SH', 'G3939-SH',
}

_STAGES = ['계약금', '중도금', '잔금']
_COL = {'계약금': '계약금 계산서', '중도금': '중도금 계산서', '잔금': '잔금 계산서'}


def _client():
    from slack_sdk import WebClient
    tok = os.getenv('SLACK_INVOICE_BOT_TOKEN', '').strip()
    return WebClient(token=tok) if tok else None


def _channel() -> str:
    return os.getenv('SLACK_INVOICE_CHANNEL_ID', '').strip()


def _num(v):
    try:
        return float(str(v).replace(',', '').strip() or 0)
    except (ValueError, TypeError):
        return 0.0


def _ntok(t):
    s = str(t or '').strip()
    if s in ('', '-'):
        return ''
    if s == '카드결제':
        return '카드'
    if s == '일반':
        return '발행'
    if s in ('혼합', '확인필요'):
        return '기타'
    return s


def _collected(row) -> bool:
    t2 = _num(row.get('총액 2'))
    paid = sum(_num(row.get(s)) for s in _STAGES)
    unpaid = _num(row.get('미수금'))
    auto = t2 > 0 and paid > 0 and abs(unpaid) < 1
    return auto or str(row.get('수금 확인')).strip() in ('TRUE', 'true', '1', 'True')


def _bill_stages(row):
    """billStatus.js computeBillStagesFromColumns 미러 → {단계: 상태}."""
    res = {s: 'none' for s in _STAGES}
    uninv = '미발행' if _collected(row) else '발행예정'
    raw = lambda s: str(row.get(_COL[s]) or '').strip()
    hasIssued = any(_ntok(raw(s)) == '발행' for s in _STAGES)
    active = [s for s in _STAGES if _num(row.get(s)) > 0 or _ntok(raw(s))]
    anchor = active[-1] if active else None
    aInv = anchor and (_ntok(raw(anchor)) == '발행' or (raw(anchor) == '-' and hasIssued))
    full = _collected(row) and aInv
    for s in _STAGES:
        r = raw(s); v = _ntok(r)
        if v:
            res[s] = ('none' if (v == '미발행' and full) else (uninv if v == '미발행' else v))
        elif r == '-' and hasIssued:
            pass
        elif _num(row.get(s)) > 0:
            res[s] = ('none' if full else uninv)
        elif r == '-':
            pass
    return res


def _end_passed(row) -> bool:
    """공사 종료일이 지났는가(=완공). 종료일 없으면 True(불명은 노출), 미래면 False(진행중)."""
    from datetime import date
    s = str(row.get('공사 종료') or '').strip()
    m = re.search(r'(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})', s)
    if not m:
        return True
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3))) < date.today()
    except ValueError:
        return True


def _collection_needed(row) -> bool:
    """수금 필요 = 완공(공사 종료일 지남) + 발행한 단계인데 그 단계 입금 0 + 미수금 남음.

    '계약금 발행·입금 + 잔금 대기' 같은 정상 진행건은 제외(발행한 단계는 이미 입금됨).
    공사 종료 전이면 잔금이 아직 안 나올 시점 → 독촉 이르니 제외. 선발행(전액 발행
    했는데 입금 0)·발행 후 미입금만, 그것도 완공된 건만 걸린다. 전액발행은 잔금='발행'
    (앞단계 covered '-')이라 잔금 입금 0으로 잡힘.
    """
    if _num(row.get('미수금')) <= 0:
        return False
    if not _end_passed(row):
        return False
    raw = lambda s: str(row.get(_COL[s]) or '').strip()
    return any(_ntok(raw(s)) == '발행' and _num(row.get(s)) <= 0 for s in _STAGES)


_DATE = re.compile(r'\d{2,4}[.\-/]\d{1,2}[.\-/]\d{1,2}')


def _addr(row) -> str:
    a = str(row.get('현장 주소') or '').strip()
    a = a.split('최종수정')[0].strip()
    a = _DATE.sub('', a).strip(' \n·-')
    return a or '-'


def _won(n) -> str:
    return f'{int(round(n)):,}원'


def classify(recs):
    """프로젝트 → {issue_collected, issue_partial, collect} 3버킷."""
    b = {'issue_collected': [], 'issue_partial': [], 'collect': []}
    for r in recs:
        code = str(r.get('프로젝트 코드', '')).strip()
        if not code or code in _EXCLUDE:
            continue
        if _num(r.get('총액 2')) <= 0:
            continue
        bs = _bill_stages(r)
        unpaid = _num(r.get('미수금'))
        has_uninv = any(bs[s] == '미발행' for s in _STAGES)     # 수금완료+미발행
        has_pending = any(bs[s] == '발행예정' for s in _STAGES)  # 부분입금+미발행
        if has_uninv:
            b['issue_collected'].append(r)
        elif has_pending:
            b['issue_partial'].append(r)
        if code not in _COLLECT_EXCLUDE and _collection_needed(r):
            b['collect'].append(r)
    return b


def _mgr(row) -> str:
    return str(row.get('담당자') or '').strip()


def _initial(row) -> str:
    """담당자 이니셜 = 프로젝트 코드 접미(예: G3961-MS → MS). 정렬용."""
    code = str(row.get('프로젝트 코드', '') or '')
    return (code.rsplit('-', 1)[-1] if '-' in code else code).upper()


def _sort(items):
    return sorted(items, key=lambda r: (_initial(r), str(r.get('프로젝트 코드', ''))))


def _biz(row) -> str:
    b = str(row.get('사업자명') or '').strip()
    if b and b != '-':
        return b
    return str(row.get('프로젝트 코드') or '').strip()   # 사업자명 없으면 코드로


def _collect_line(r) -> str:
    return f'• {_biz(r)}  ·  {_addr(r)}  ·  미수금 {_won(_num(r.get("미수금")))}'


def _issue_line(r) -> str:
    return f'• {_biz(r)}  ·  {_addr(r)}'


def _section(header, items, line_fn):
    return '\n'.join([header] + [line_fn(r) for r in _sort(items)])


def build_collection_text(buckets) -> str:
    """주간 수금 필요 문안 (③ 만)."""
    coll = buckets.get('collect', [])
    if not coll:
        return ''
    body = _section(f':moneybag: *수금 필요 ({len(coll)}건)*', coll, _collect_line)
    return (
        f'{_BLANK}\n'
        f':moneybag: *수금 필요 — 세금계산서 발행 완료, 미수금 확인 부탁드립니다*\n'
        f'{_SEP}\n'
        f'{body}\n'
        f'{_SEP}\n'
        f'{_BLANK}'
    )


def build_monthly_text(buckets) -> str:
    """매월 10일 문안 (① 수금완료·미발행 ② 부분입금·미발행 ③ 수금 필요)."""
    a = buckets.get('issue_collected', [])
    b = buckets.get('issue_partial', [])
    c = buckets.get('collect', [])
    if not (a or b or c):
        return ''
    secs = []
    if a:
        secs.append(_section(f':receipt: *① 수금완료 · 미발행 ({len(a)}건)*', a, _issue_line))
    if b:
        secs.append(_section(f':receipt: *② 부분입금 · 미발행 ({len(b)}건)*', b, _issue_line))
    if c:
        secs.append(_section(f':moneybag: *③ 수금 필요 ({len(c)}건)*', c, _collect_line))
    body = f'\n{_BLANK}\n'.join(secs)
    return (
        f'{_BLANK}\n'
        f':receipt: *세금계산서 마감(매월 10일) — 발행/보류 여부를 경영지원실에 전달해주세요*\n'
        f'{_SEP}\n'
        f'{body}\n'
        f'{_SEP}\n'
        f'{_BLANK}'
    )


def _load_recs():
    from dashboard.services.project_service import get_project_records
    return get_project_records(force_refresh=True) or []


def _post(text: str, tag: str) -> dict:
    if not text:
        logger.info(f'[수금리마인드/{tag}] 대상 0건 — 발송 skip')
        return {'ok': True, 'total': 0}
    c = _client(); ch = _channel()
    if not c or not ch:
        logger.warning(f'[수금리마인드/{tag}] 봇/채널 미설정 — skip')
        return {'ok': False, 'reason': 'no_client'}
    try:
        r = c.chat_postMessage(channel=ch, text=text, unfurl_links=False, unfurl_media=False)
        logger.info(f'[수금리마인드/{tag}] 발송 완료 (ts={r.get("ts")})')
        return {'ok': True, 'ts': r.get('ts', '')}
    except Exception as exc:
        logger.error(f'[수금리마인드/{tag}] 발송 예외: {exc}', exc_info=True)
        return {'ok': False, 'reason': str(exc)}


def send_weekly_collection_remind() -> dict:
    """매주 월 09시 — 수금 필요 리마인드 진입점."""
    return _post(build_collection_text(classify(_load_recs())), '주간')


def send_monthly_invoice_remind() -> dict:
    """매월 10일 09시 — 계산서 마감 리마인드 진입점."""
    return _post(build_monthly_text(classify(_load_recs())), '월간')


if __name__ == '__main__':
    import sys
    from pathlib import Path
    _ROOT = Path(__file__).resolve().parents[2]
    sys.path.insert(0, str(_ROOT))
    try:
        from dotenv import load_dotenv; load_dotenv(_ROOT / '.env')
    except Exception:
        pass
    bk = classify(_load_recs())
    print(f"[버킷] 수금완료·미발행={len(bk['issue_collected'])}  "
          f"부분입금·미발행={len(bk['issue_partial'])}  수금필요={len(bk['collect'])}")
    print("\n===== 주간(수금 필요) 미리보기 =====")
    print(build_collection_text(bk) or "(대상 0건)")
    print("\n===== 매월 10일 미리보기 =====")
    print(build_monthly_text(bk) or "(대상 0건)")
