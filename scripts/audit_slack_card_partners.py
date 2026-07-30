"""#수금_관리 채널 카드 이력 라인에서 이상한 파싱 결과 리스트업.

이상 판단 기준:
- partner 가 '-' (없음)
- partner 가 숫자+원 형태 (파싱 오류)
- partner 가 계좌번호 (숫자만)
- partner 가 은행이름 단독
- partner 가 MM/DD 날짜
- date 가 '-' (없음)
- partner 매우 짧음 (1글자, 특수문자만)
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

try:
    from dotenv import load_dotenv
    load_dotenv(_ROOT / '.env')
except Exception:
    pass


CODE_RE = re.compile(r'\*([GRNP]{1,2}\d{4}(?:-[A-Z]{2,3})?)\*')
STAGE_TITLE_RE = re.compile(r'\*(계약금 입금|중도금 입금|잔금 입금|수금완료|통합 입금|수금완료 \(통합 입금\))\*')

# payment 이력 라인:
#   "계약금  MM/DD  bank_code  amount원  partner"
#   "잔금  02/23  기업 (G)  1,538,000원  고려/비즈이노베이"
#   "계약금  24/03  G  1,500,000원  이성민(이즈이노베)"
HISTORY_LINE_RE = re.compile(
    r'^(계약금|중도금|잔금)\s+'
    r'(\S+)\s+'          # date (MM/DD or MM/YY)
    r'(.+?)\s+'          # bank code
    r'([\d,]+)원\s+'     # amount
    r'(.*?)(?:\s*\(카드\))?$'  # partner (카드 suffix optional)
)

NOISE_PARTNER_RE = re.compile(
    r'^(?:'
    r'-'                              # 없음
    r'|[\d,]+\s*원?'                  # 숫자+원 (파싱 오류)
    r'|\d+[*\d\-]+'                   # 계좌번호
    r'|(?:기업|하나|국민|신한|우리|농협|카카오|토스)'  # 은행 단독
    r'|\d{1,2}[/.-]\d{1,2}'           # MM/DD
    r'|[GRNP]\s*[\d,]+원?'            # "G 1,000,000원"
    r')$'
)


def _stage_from_title(t):
    if '계약금' in t:
        return '계약금'
    if '중도금' in t:
        return '중도금'
    return '잔금'


def main() -> int:
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    from slack_sdk import WebClient
    slack = WebClient(token=token)

    print('[*] 슬랙 채널 히스토리 fetch...')
    all_msgs = []
    cursor = None
    LIMIT = 5000
    while len(all_msgs) < LIMIT:
        kwargs = {'channel': channel, 'limit': 200}
        if cursor:
            kwargs['cursor'] = cursor
        r = slack.conversations_history(**kwargs)
        msgs = r.get('messages', [])
        if not msgs:
            break
        all_msgs.extend(msgs)
        cursor = r.get('response_metadata', {}).get('next_cursor')
        if not cursor:
            break
    print(f'[*] {len(all_msgs)} messages fetched')
    print()

    issues = []  # (code, stage_title, line, reason)
    seen_codes = set()
    for m in all_msgs:
        text = m.get('text', '') or ''
        title_line = ''
        for ln in text.split('\n'):
            if STAGE_TITLE_RE.search(ln):
                title_line = ln
                break
        if not title_line:
            continue
        codes_in_title = CODE_RE.findall(title_line)
        if not codes_in_title:
            continue
        code = codes_in_title[0]
        # 같은 code 여러 개 카드면 최신만 (역순 fetch 라 첫 발견이 최신)
        if code in seen_codes:
            continue
        seen_codes.add(code)

        # 이력 라인 파싱 (`[입금 이력]` or `[누적 이력]` 블록)
        lines = text.split('\n')
        in_history = False
        for ln in lines:
            if '[입금 이력]' in ln or '[누적 이력]' in ln:
                in_history = True
                continue
            if '---' in ln or ln.startswith('총액') or ln.startswith('미수금'):
                in_history = False
                continue
            if not in_history:
                continue
            ln_stripped = ln.strip()
            if not ln_stripped:
                continue
            m2 = HISTORY_LINE_RE.match(ln_stripped)
            if not m2:
                continue
            stage, date, bank, amount, partner = m2.groups()
            partner = (partner or '').strip()

            reasons = []
            if not partner or NOISE_PARTNER_RE.match(partner):
                reasons.append(f'partner 이상: {partner!r}')
            if date == '-':
                reasons.append(f'date 없음')
            if reasons:
                issues.append({
                    'code': code, 'stage': stage, 'date': date,
                    'bank': bank, 'amount': amount, 'partner': partner,
                    'reasons': ' | '.join(reasons),
                    'ts': m.get('ts', ''),
                })

    # 각 이상 건 → 현재 시트 메모 재파싱으로 '올바른 입금자' 제안 (2026-07-29).
    #   카드 표시(과거 파싱) ↔ 현재 파서 결과 대조 → 정정 대상·값 한눈에.
    try:
        from dashboard.services.payment_sync import (
            _get_payment_service, _parse_notes, _to_int_won)
        svc = _get_payment_service()
        _sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
        _sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
        _vals = svc.spreadsheets().values().get(
            spreadsheetId=_sid, range=f"'{_sn}'!A2:AA10000",
            valueRenderOption='UNFORMATTED_VALUE').execute().get('values', [])
        _nr = svc.spreadsheets().get(
            spreadsheetId=_sid, ranges=[f"'{_sn}'!U2:W10000"],
            fields='sheets.data.rowData.values.note',
            includeGridData=True).execute()['sheets'][0]['data'][0].get('rowData', [])
        _c2 = {}
        for _off, _r in enumerate(_vals):
            _cc = str((_r[0] if _r else '')).strip()
            if _cc:
                _c2[_cc] = _off

        def _suggest(code, stage):
            off = _c2.get(code)
            if off is None:
                return '(코드없음)'
            _r = list(_vals[off]) + [''] * (27 - len(_vals[off]))
            u, v, w = _to_int_won(_r[20]), _to_int_won(_r[21]), _to_int_won(_r[22])  # U/V/W
            _notes = ['', '', '']
            if off < len(_nr):
                _vv = _nr[off].get('values', []) or []
                for _i in range(min(3, len(_vv))):
                    _notes[_i] = _vv[_i].get('note', '') or ''
            for p in (_parse_notes(_notes, {'계약금': u, '중도금': v, '잔금': w}) or []):
                if p.get('stage') == stage:
                    return p.get('partner', '') or '(빈값)'
            return '(파싱없음)'
        for it in issues:
            it['suggest'] = _suggest(it['code'], it['stage'])
    except Exception as exc:
        print(f'[!] 재파싱 제안 생략 (시트 접근 실패): {exc}')
        for it in issues:
            it['suggest'] = '?'

    print(f'[*] 이상 이력 라인: {len(issues)}건')
    print()
    print(f"{'code':<12} {'stage':<3}  {'date':<6}  {'amount':>10}  {'카드표시':<24}  →  {'재파싱 제안':<16}  |  이슈")
    print('-' * 120)
    for it in issues:
        print(f"{it['code']:<12} {it['stage']:<3}  {it['date']:<6}  {it['amount']:>10}원  "
              f"{it['partner']!r:<24}  →  {it.get('suggest','?')!r:<16}  |  {it['reasons']}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
