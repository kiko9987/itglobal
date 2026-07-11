"""audit3 후보 30건에 대해 Redis payment_slack:ts 매핑 확인.

payment_slack:ts 매핑이 있으면 실제로 발송됨 (오탐), 없으면 진짜 누락 후보.
"""
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

from dashboard.utils.redis_client import get_redis_client  # noqa: E402


CANDIDATES = [
    ('G0392-YG', '잔금', '07/10', 4510000),
    ('G0415-MW', '잔금', '07/10', 3500000),
    ('G0425-GH', '잔금', '07/10', 1401400),
    ('G0439-SH', '잔금', '07/10', 2530000),
    ('G0451-YG', '잔금', '07/10', 2970000),
    ('G1113-YG', '잔금', '07/10', 15840000),
    ('G1317-YG', '잔금', '07/10', 4015000),
    ('G1324-SH', '잔금', '07/10', 6160000),
    ('G1357-DN', '잔금', '07/10', 1430000),
    ('G1367-SH', '잔금', '07/11', 4840000),
    ('G1385-SH', '잔금', '07/10', 6820000),
    ('G1417-JW', '잔금', '07/10', 2400000),
    ('GG1419-JW', '잔금', '07/10', 12320000),
    ('G1423-JW', '잔금', '07/10', 440000),
    ('G1435-JW', '계약금', '07/10', 5400000),
    ('G1441-YG', '잔금', '07/10', 3850000),
    ('G1444-MW', '계약금', '07/11', 4686000),
    ('G1449-YG', '계약금', '07/11', 3036000),
    ('G1450-SH', '계약금', '07/10', 300000),
    ('G1452-YG', '계약금', '07/10', 286000),
    ('G2583-MW', '계약금', '07/10', 26400000),
    ('G2616-MW', '계약금', '07/11', 7975000),
    ('R2618-MJ', '계약금', '07/11', 869000),
    ('G2704-SH', '계약금', '07/10', 3300000),
    ('G3717-SH', '계약금', '07/10', 10120000),
    ('G3803-SJ', '계약금', '07/10', 2079000),
    ('R3826-MJ', '잔금', '07/10', 5390000),
    ('R3845-JSH', '계약금', '07/10', 24519000),
    ('R3853-SJ', '계약금', '07/10', 1350000),
    ('R3854-SJ', '계약금', '07/10', 1380000),
]


def main() -> int:
    rc = get_redis_client().redis
    sent, not_sent = [], []
    for project, stage, date_md, amount in CANDIDATES:
        ts = rc.get(f'payment_slack:ts:{project}:{stage}')
        if ts:
            if isinstance(ts, bytes):
                ts = ts.decode('utf-8')
            sent.append((project, stage, date_md, amount, ts))
        else:
            not_sent.append((project, stage, date_md, amount))

    print(f'=== 실제 발송됨 (Redis ts 있음): {len(sent)} 건 ===')
    for p, s, d, a, ts in sent:
        print(f'  {p:<12} {s:<6} {d}  {a:>12,}  ts={ts}')

    print()
    print(f'=== 진짜 누락 후보 (Redis ts 없음): {len(not_sent)} 건 ===')
    for p, s, d, a in not_sent:
        print(f'  {p:<12} {s:<6} {d}  {a:>12,}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
