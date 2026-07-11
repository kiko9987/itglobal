"""시트 X열 (미수금) float 오차 잔여 건 카운트.

시트 X값이 0 이 아니지만 절대값 100원 미만인 프로젝트 리스트업.
Google Sheets UNFORMATTED_VALUE 는 float 저장이라 시트 수식 결과가
-9.3e-10, 0.199... 같은 반올림 오차로 남는 케이스.
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
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    from dashboard.services.payment_sync import _get_payment_service
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
    IDX_F = col('F')
    IDX_X = col('X')
    IDX_AA = col('A') + 1  # AA = 26

    residue = []       # 절대값 0 < |X| < 1 (숨겨진 오차)
    small_residue = [] # 1 <= |X| < 100 (알림 안 가지만 시트에 남음)

    for i, r in enumerate(rows):
        code = str(r[IDX_A] if len(r) > IDX_A else '').strip()
        if not code:
            continue
        raw = r[IDX_X] if len(r) > IDX_X else None
        if raw is None or raw == '':
            continue
        try:
            x = float(raw)
        except Exception:
            continue
        if x == 0:
            continue
        abs_x = abs(x)
        addr = str(r[IDX_F] if len(r) > IDX_F else '')[:40]
        entry = (code, addr, raw)
        if abs_x < 1:
            residue.append(entry)
        elif abs_x < 100:
            small_residue.append(entry)

    print(f'=== 절대값 0 < |X| < 1 (숨겨진 float 오차): {len(residue)} 건 ===')
    for code, addr, raw in residue[:30]:
        print(f'  {code:<12}  X={raw!r:<30}  {addr}')
    if len(residue) > 30:
        print(f'  ... ({len(residue) - 30} 건 더)')

    print()
    print(f'=== 1 <= |X| < 100 (알림 skip 임계값 안, 시트에 남음): {len(small_residue)} 건 ===')
    for code, addr, raw in small_residue[:30]:
        print(f'  {code:<12}  X={raw!r:<30}  {addr}')
    if len(small_residue) > 30:
        print(f'  ... ({len(small_residue) - 30} 건 더)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
