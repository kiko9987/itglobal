"""#수금_관리 채널 최근 카드 3개 실제 텍스트 확인."""
from __future__ import annotations

import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

try:
    from dotenv import load_dotenv
    load_dotenv(_ROOT / '.env')
except Exception:
    pass


def main() -> int:
    token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    if not token or not channel:
        print('env 미설정')
        return 1

    from slack_sdk import WebClient
    slack = WebClient(token=token)
    resp = slack.conversations_history(channel=channel, limit=15)
    msgs = resp.get('messages', [])
    for i, m in enumerate(msgs):
        text = m.get('text', '') or ''
        ts = m.get('ts', '')
        print(f'--- msg #{i} (ts={ts}) ---')
        print(text[:500])
        print()
    return 0


if __name__ == '__main__':
    sys.exit(main())
