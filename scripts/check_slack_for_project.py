"""각 프로젝트 코드가 #수금_관리 채널 최근 3000개 메시지 안에 있는지 검색."""
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


TARGETS = [
    'G2583-MW', 'G2704-SH', 'G3717-SH', 'G3803-SJ',
    'R3826-MJ', 'R3845-JSH', 'R3853-SJ', 'R3854-SJ',
]


def main() -> int:
    token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()

    from slack_sdk import WebClient
    slack = WebClient(token=token)

    hits = {t: [] for t in TARGETS}
    cursor = None
    fetched = 0
    LIMIT = 3000
    while fetched < LIMIT:
        kwargs = {'channel': channel, 'limit': 200}
        if cursor:
            kwargs['cursor'] = cursor
        resp = slack.conversations_history(**kwargs)
        msgs = resp.get('messages', [])
        for m in msgs:
            text = m.get('text', '') or ''
            ts = m.get('ts', '')
            for target in TARGETS:
                if target in text:
                    # 첫 라인만 저장 (title)
                    title_line = text.split('\n', 5)[:3]
                    hits[target].append((ts, ' | '.join([l.strip() for l in title_line if l.strip()])[:200]))
        fetched += len(msgs)
        cursor = resp.get('response_metadata', {}).get('next_cursor')
        if not cursor:
            break

    from datetime import datetime
    print(f'[*] fetched msgs: {fetched}')
    print()
    for target in TARGETS:
        hs = hits[target]
        if not hs:
            print(f'[!] {target}: 카드 없음 (실제 누락 가능성)')
        else:
            print(f'[OK] {target}: {len(hs)} 개 카드')
            for ts, line in hs[:3]:
                dt = datetime.fromtimestamp(float(ts))
                print(f'    {dt}: {line[:150]}')
        print()
    return 0


if __name__ == '__main__':
    sys.exit(main())
