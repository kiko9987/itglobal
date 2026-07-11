"""최근 payment 알림 누락 감사 (슬랙 히스토리 대조).

payment_sync 도입일(2026-06-26) 이후 발생한 payment 중 슬랙 채널에 카드가 안
올라간 것을 리스트업.

방식:
1. Slack #수금_관리 채널 히스토리 200개 fetch → 발송된 (project, stage) set 구축
   (개별 카드 + 통합 카드 파싱)
2. 시트에서 각 행의 U/V/W 값 + 메모 fetch
3. 메모의 date_md 파싱 → cutoff 이후 payment 만 필터
4. sent set 에 없는 payment 리스트업
"""
from __future__ import annotations

import os
import re
import sys
from datetime import date
from pathlib import Path
from typing import Dict, List, Set, Tuple

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

try:
    from dotenv import load_dotenv
    load_dotenv(_ROOT / '.env')
except Exception:
    pass

from dashboard.services.payment_sync import (  # noqa: E402
    _get_payment_service,
    _fetch_row_notes,
    _parse_notes,
)


def _to_int_won(v) -> int:
    try:
        if isinstance(v, (int, float)):
            return int(v)
        s = str(v).replace(',', '').replace('￦', '').replace('₩', '').strip()
        if not s:
            return 0
        return int(float(s))
    except Exception:
        return 0


TITLE_STAGE_RE = re.compile(r'\*(계약금 입금|중도금 입금|잔금 입금|수금완료|통합 입금|수금완료 \(통합 입금\))\*')
CODE_RE = re.compile(r'\*([GRNP]\d{4}(?:-[A-Z]{2})?)\*')
LIST_CODE_RE = re.compile(r'([GRNP]\d{4}(?:-[A-Z]{2})?)')


def _stage_from_title(t: str) -> str:
    if '계약금' in t:
        return '계약금'
    if '중도금' in t:
        return '중도금'
    return '잔금'


def fetch_slack_sent_set(channel: str, token: str, limit: int = 1500) -> Set[Tuple[str, str]]:
    """슬랙 채널 title 라인 파싱 → (project_code, stage) 발송 세트."""
    from slack_sdk import WebClient
    from slack_sdk.errors import SlackApiError
    slack = WebClient(token=token)
    sent: Set[Tuple[str, str]] = set()
    cursor = None
    fetched = 0
    while fetched < limit:
        try:
            kwargs = {'channel': channel, 'limit': min(200, limit - fetched)}
            if cursor:
                kwargs['cursor'] = cursor
            resp = slack.conversations_history(**kwargs)
        except SlackApiError as exc:
            print(f'[!] 슬랙 히스토리 fetch 실패: {exc.response["error"]}')
            break
        msgs = resp.get('messages', [])
        for msg in msgs:
            text = msg.get('text', '') or ''
            title_line = ''
            for ln in text.split('\n'):
                if TITLE_STAGE_RE.search(ln):
                    title_line = ln
                    break
            if not title_line:
                continue
            sm = TITLE_STAGE_RE.search(title_line)
            if not sm:
                continue
            stage = _stage_from_title(sm.group(1))
            is_unified = '통합' in sm.group(1)
            if is_unified:
                id_part = title_line.split(':id:', 1)[-1] if ':id:' in title_line else title_line
                for c in LIST_CODE_RE.findall(id_part):
                    sent.add((c, stage))
            else:
                m2 = CODE_RE.search(title_line)
                if m2:
                    sent.add((m2.group(1), stage))
                else:
                    m3 = LIST_CODE_RE.search(title_line)
                    if m3:
                        sent.add((m3.group(1), stage))
        fetched += len(msgs)
        cursor = resp.get('response_metadata', {}).get('next_cursor')
        if not cursor:
            break
    return sent


def main() -> int:
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    if not all([sheet_id, sheet_name, channel, token]):
        print('필수 환경변수 미설정')
        return 1

    # cutoff: 오늘·어제만 (슬랙 히스토리 1000개로 커버 가능한 범위)
    CUTOFF = date(2026, 7, 10)
    print(f'[*] cutoff = {CUTOFF.isoformat()} (이 날짜 이후 발생한 payment 만 검사)')

    print('[*] Slack #수금_관리 채널 히스토리 fetch 중...')
    sent_set = fetch_slack_sent_set(channel, token, limit=1500)
    print(f'[*] 슬랙에서 발송 확인된 (project, stage): {len(sent_set)} 개')

    print('[*] 시트 값 fetch 중...')
    svc = _get_payment_service()
    if not svc:
        print('service 초기화 실패')
        return 1
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{sheet_name}'!A2:AA10000",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = resp.get('values', [])
    print(f'[*] 시트 rows: {len(rows)}')

    def col_idx(c: str) -> int:
        if len(c) == 1:
            return ord(c) - ord('A')
        return (ord(c[0]) - ord('A') + 1) * 26 + (ord(c[1]) - ord('A'))

    IDX_A = col_idx('A')
    IDX_F = col_idx('F')
    IDX_U = col_idx('U')
    IDX_V = col_idx('V')
    IDX_W = col_idx('W')

    def _payment_date_ok(pd_str: str) -> bool:
        """date_md 문자열을 date 로 변환. cutoff 이후면 True."""
        if not pd_str or pd_str == '-':
            return False
        m = re.match(r'(\d{1,2})/(\d{1,2})', pd_str)
        if not m:
            return False
        mm, dd = int(m.group(1)), int(m.group(2))
        # year 추정: 현재 date 기준 미래면 작년, 아니면 현재년.
        # 도입일 cutoff 이후 판단만 하면 되므로 단순화: 현재년 사용, 미래면 작년.
        today = date(2026, 7, 11)
        try:
            y = today.year
            candidate = date(y, mm, dd)
            if candidate > today:
                candidate = date(y - 1, mm, dd)
        except ValueError:
            return False
        return candidate >= CUTOFF

    missing: List[Dict] = []
    checked = 0

    for i, row in enumerate(rows):
        sheet_row = i + 2

        def _get(idx: int):
            return row[idx] if idx < len(row) else ''

        project = str(_get(IDX_A)).strip()
        if not project:
            continue
        u = _to_int_won(_get(IDX_U))
        v = _to_int_won(_get(IDX_V))
        w = _to_int_won(_get(IDX_W))
        if u == 0 and v == 0 and w == 0:
            continue
        stage_vals = {'계약금': u, '중도금': v, '잔금': w}

        try:
            notes = _fetch_row_notes(sheet_id, sheet_name, sheet_row)
        except Exception:
            continue
        payments = _parse_notes(notes, stage_vals=stage_vals)
        if not payments:
            continue

        for p in payments:
            stage = p.get('stage', '')
            date_md = p.get('date_md', '')
            if not _payment_date_ok(date_md):
                continue
            checked += 1
            if (project, stage) in sent_set:
                continue
            missing.append({
                'row': sheet_row,
                'project': project,
                'stage': stage,
                'date_md': date_md,
                'amount': int(p.get('amount', 0)),
                'partner': p.get('partner', ''),
                'address': str(_get(IDX_F)).strip(),
            })

    print(f'[*] cutoff 이후 payment 총 {checked} 건 검사')
    print()

    if not missing:
        print('[OK] 누락 없음')
        return 0

    print(f'[!] 누락 후보 {len(missing)} 건:')
    print()
    print(f"{'row':>5}  {'project':<12}  {'stage':<6}  {'date':>6}  {'amount':>12}  partner  address")
    print('-' * 120)
    for m in missing:
        print(
            f"{m['row']:>5}  {m['project']:<12}  {m['stage']:<6}  {m['date_md']:>6}  "
            f"{m['amount']:>12,}  {m['partner'][:20]:<20}  {m['address'][:40]}"
        )
    return 0


if __name__ == '__main__':
    sys.exit(main())
