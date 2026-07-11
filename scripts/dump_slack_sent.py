"""#수금_관리 채널 최근 1000개 메시지 → 발송된 (project, stage) 세트 dump.

각 메시지의 첫 line 만 파싱해서 (project, stage) 추출. 이력 섹션 참조는 제외.
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from typing import Set, Tuple

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

try:
    from dotenv import load_dotenv
    load_dotenv(_ROOT / '.env')
except Exception:
    pass


TITLE_STAGE_RE = re.compile(r'\*(계약금 입금|중도금 입금|잔금 입금|수금완료|통합 입금|수금완료 \(통합 입금\))\*')
CODE_RE = re.compile(r'\*([GRNP]\d{4}(?:-[A-Z]{2})?)\*')
# 통합: `G2855-SH, G2856-SH, G2857-SH, G2858-SH` 형태
LIST_CODE_RE = re.compile(r'([GRNP]\d{4}(?:-[A-Z]{2})?)')


def stage_from_title(t: str) -> str:
    if '계약금' in t:
        return '계약금'
    if '중도금' in t:
        return '중도금'
    if '잔금' in t or '수금완료' in t or '통합' in t:
        return '잔금'
    return ''


def main() -> int:
    token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    from slack_sdk import WebClient
    from slack_sdk.errors import SlackApiError
    slack = WebClient(token=token)

    sent: Set[Tuple[str, str]] = set()
    cursor = None
    fetched = 0
    LIMIT = 1000
    oldest_ts = None
    newest_ts = None

    while fetched < LIMIT:
        try:
            kwargs = {'channel': channel, 'limit': min(200, LIMIT - fetched)}
            if cursor:
                kwargs['cursor'] = cursor
            resp = slack.conversations_history(**kwargs)
        except SlackApiError as exc:
            print(f'[!] fetch 실패: {exc.response.get("error")}')
            break

        msgs = resp.get('messages', [])
        for m in msgs:
            text = m.get('text', '') or ''
            ts = m.get('ts', '')
            if newest_ts is None or float(ts) > float(newest_ts):
                newest_ts = ts
            if oldest_ts is None or float(ts) < float(oldest_ts):
                oldest_ts = ts

            # title 라인만 뽑기
            lines = text.split('\n')
            title_line = ''
            for ln in lines:
                if TITLE_STAGE_RE.search(ln):
                    title_line = ln
                    break
            if not title_line:
                continue
            stage_m = TITLE_STAGE_RE.search(title_line)
            if not stage_m:
                continue
            stage = stage_from_title(stage_m.group(1))
            if not stage:
                continue

            # 통합 카드 여부
            is_unified = ('통합' in title_line) or ('수금완료 (통합' in title_line)

            if is_unified:
                # title 라인에서 :id: 뒤 코드 리스트 모두
                # :id: *G2855-SH, G2856-SH, G2857-SH, G2858-SH* (총 4건)
                id_part = title_line.split(':id:', 1)[-1] if ':id:' in title_line else title_line
                codes = LIST_CODE_RE.findall(id_part)
                for c in codes:
                    sent.add((c, stage))
            else:
                # title 라인의 *CODE* 하나
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

    print(f'[*] fetched msgs: {fetched}')
    if oldest_ts and newest_ts:
        from datetime import datetime
        old_dt = datetime.fromtimestamp(float(oldest_ts))
        new_dt = datetime.fromtimestamp(float(newest_ts))
        print(f'[*] oldest: {old_dt}, newest: {new_dt}')
    print(f'[*] unique (project, stage) sent: {len(sent)}')
    print()
    # 청크 프린트
    from collections import defaultdict
    by_stage = defaultdict(list)
    for (p, s) in sent:
        by_stage[s].append(p)
    for stage in ('계약금', '중도금', '잔금'):
        codes = sorted(by_stage[stage])
        print(f'== {stage} ({len(codes)}) ==')
        for c in codes:
            print(f'  {c}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
