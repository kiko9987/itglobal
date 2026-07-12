"""sheet_note_mismatch 12건 각 셀 원본 노트 조회 + 파서 결과 비교."""
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


TARGETS = [
    # sheet_note_mismatch (7건)
    'G0278-SH', 'G0482-YG', 'G1019-MW', 'G1884-SH', 'G2259-YG', 'G2587-SJ', 'R3611-YM',
    # required_missing (5건)
    'G0323-YG', 'G1865-YG', 'G1897-MW', 'R3688-TH', 'R3692-MJ',
    # reference_mismatch (2건)
    'G2016-YG', 'G2224-YM',
]


def _int(v):
    try:
        return int(float(str(v).replace(',', '').strip() or 0))
    except (ValueError, TypeError):
        return 0


def main() -> int:
    from dashboard.services.payment_sync import (
        _get_payment_service, _fetch_row_notes, _parse_notes,
    )

    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    svc = _get_payment_service()

    # 시트 전체 fetch (row 찾기)
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{sheet_name}'!A2:AA10000",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = resp.get('values', [])

    def col_idx(c):
        return ord(c) - ord('A')

    for code in TARGETS:
        row_i = None
        for i, r in enumerate(rows):
            if len(r) > 0 and str(r[0]).strip() == code:
                row_i = i + 2
                r_data = r
                break
        if row_i is None:
            print(f'[!] {code}: 못 찾음')
            continue

        def _get(idx):
            return r_data[idx] if idx < len(r_data) else ''
        u = _int(_get(20))  # U
        v = _int(_get(21))  # V
        w = _int(_get(22))  # W
        print(f'\n### {code} (row {row_i}) ###')
        print(f'  시트값: U={u:,}  V={v:,}  W={w:,}')

        notes = _fetch_row_notes(sheet_id, sheet_name, row_i)
        for j, (n, stage) in enumerate(zip(notes, ['계약금(U)', '중도금(V)', '잔금(W)'])):
            if not n:
                continue
            print(f'  --- {stage} 셀 노트 ---')
            for ln in n.splitlines():
                print(f'    │ {ln}')

        payments = _parse_notes(notes, stage_vals={'계약금': u, '중도금': v, '잔금': w})
        print(f'  파싱 결과 ({len(payments)}):')
        for p in payments:
            print(f'    stage={p.get("stage")}  amount={p.get("amount"):,}  date={p.get("date_md")}  partner={p.get("partner")}')

    return 0


if __name__ == '__main__':
    sys.exit(main())
