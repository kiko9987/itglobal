"""Redis 에 payment_slack:ts 저장된 카드 대상 T열 float 오차 카운트.

이미 발송된 카드들 중 총액 T 가 float 오차(예: X.999999)로 저장돼 카드에
1원 부족 표시된 것들 파악. 대상: 잔금 카드 (수금완료 = T 표시).
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


def main() -> int:
    from dashboard.utils.redis_client import get_redis_client
    from dashboard.services.payment_sync import _get_payment_service
    rc = get_redis_client().redis

    # 시트 fetch (T값 조회)
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    svc = _get_payment_service()
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{sheet_name}'!A2:AA10000",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = resp.get('values', [])

    def col(c):
        return ord(c) - ord('A')

    IDX_A = col('A')
    IDX_T = col('T')
    IDX_F = col('F')

    # code → T raw value 매핑
    t_by_code = {}
    for r in rows:
        code = str(r[IDX_A] if len(r) > IDX_A else '').strip()
        if not code:
            continue
        raw = r[IDX_T] if len(r) > IDX_T else None
        addr = str(r[IDX_F] if len(r) > IDX_F else '')[:40]
        t_by_code[code] = (raw, addr)

    # Redis payment_slack:ts 조회 — 잔금 stage 필터
    all_keys = list(rc.scan_iter('payment_slack:ts:*', count=1000))
    keys = [k for k in all_keys if (k if isinstance(k, str) else k.decode()).endswith(':잔금')]
    print(f'[*] 전체 카드 ts 저장: {len(all_keys)}, 잔금 카드: {len(keys)}')

    float_diff = []  # T 가 정수 아닌 것
    for k in keys:
        k_str = k if isinstance(k, str) else k.decode()
        # payment_slack:ts:{code}:잔금
        code = k_str.split(':')[2] if k_str.count(':') >= 3 else ''
        if not code:
            continue
        raw, addr = t_by_code.get(code, (None, ''))
        if raw is None:
            continue
        try:
            f = float(raw)
        except Exception:
            continue
        rounded = round(f)
        floor_v = int(f)
        # int() 와 round() 결과가 다르면 float 오차로 표시 1원 다름
        if floor_v != rounded:
            float_diff.append((code, raw, floor_v, rounded, addr))

    print(f'[*] T열 float 오차로 카드 표시 1원 부족·초과 건: {len(float_diff)}')
    print()
    print(f"{'code':<15} {'raw T':<26} {'int()':<15} {'round()':<15} 주소")
    print('-' * 110)
    for code, raw, floor_v, rounded, addr in float_diff[:50]:
        print(f'{code:<15} {raw!r:<26} {floor_v:>10,}    {rounded:>10,}    {addr}')
    if len(float_diff) > 50:
        print(f'... ({len(float_diff) - 50} 건 더)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
