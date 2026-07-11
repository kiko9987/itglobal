"""daily scan dry-run — 실제 슬랙 발송 없이 감지된 alerts 리스트만 출력.

send_daily_summary 를 monkey-patch 로 capture 해서 alerts 뽑아냄.
"""
from __future__ import annotations

import os
import sys
from collections import Counter
from pathlib import Path
from typing import List, Dict

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

try:
    from dotenv import load_dotenv
    load_dotenv(_ROOT / '.env')
except Exception:
    pass


def main() -> int:
    # monkey-patch: send_daily_summary 를 alerts capture 로 교체
    import dashboard.services.payment_alert as pa
    captured: List[Dict] = []

    def _capture(alerts):
        captured.extend(alerts)
        return True

    pa.send_daily_summary = _capture

    from dashboard.services.payment_alert_daily import run_daily_scan_and_summary
    result = run_daily_scan_and_summary()

    print(f'[*] 스캔 결과: {result}')
    print(f'[*] captured alerts: {len(captured)}')
    print()

    # 카테고리별 그룹
    by_cat: Dict[str, List[Dict]] = {}
    for a in captured:
        by_cat.setdefault(a['category'], []).append(a)

    counts = Counter(a['category'] for a in captured)
    print('=== 카테고리별 건수 ===')
    for cat, n in sorted(counts.items(), key=lambda x: -x[1]):
        print(f'  {n:>4}  {cat}')
    print()

    # 카테고리별 리스트
    for cat, items in by_cat.items():
        print(f'=== {cat} ({len(items)}건) ===')
        for a in items[:100]:
            addr = (a.get('address') or '')[:40]
            body = (a.get('body') or '')[:120]
            print(f'  {a["code"]:<15} {addr:<42} | {body}')
        if len(items) > 100:
            print(f'  ... ({len(items) - 100}건 더)')
        print()
    return 0


if __name__ == '__main__':
    sys.exit(main())
