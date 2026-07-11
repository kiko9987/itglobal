"""각 프로젝트의 해당 stage 시트 노트 원본 조회.

카드에 partner='-' 로 표시된 케이스 → 시트 노트 상태 파악.
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


TARGETS = [
    ('R3779-YM', '잔금'),
    ('R3520-YM', '잔금'),
    ('G3662-MS', '잔금'),
    ('G3442-JW', '잔금'),
    ('R3348-MJ', '잔금'),
    ('R3347-MJ', '잔금'),
    ('R3346-MW', '계약금'),
    ('G2998-MJ', '잔금'),
    ('G3129-SH', '잔금'),
    ('R3070-TH', '잔금'),
    ('G3213-SH', '잔금'),
    ('R3166-YM', '잔금'),
    ('R3155-MJ', '잔금'),
    ('G2977-MJ', '잔금'),
    ('R3085-SH', '잔금'),
    ('R2983-YM', '잔금'),
    ('G2587-SJ', '잔금'),
    ('R2589-TH', '잔금'),
    ('R2560-MJ', '잔금'),
    ('G1685-YG', '계약금'),
    ('G1865-YG', '계약금'),
]


def main() -> int:
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    from dashboard.services.payment_sync import _get_payment_service, _fetch_row_notes
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

    code_to_row = {}
    for i, r in enumerate(rows):
        code = str(r[IDX_A] if len(r) > IDX_A else '').strip()
        if code:
            code_to_row[code] = i + 2

    stage_to_col = {'계약금': 'U', '중도금': 'V', '잔금': 'W'}

    for code, stage in TARGETS:
        row_i = code_to_row.get(code)
        if not row_i:
            print(f'=== {code} {stage}: 시트에 없음 ===')
            continue
        notes = _fetch_row_notes(sheet_id, sheet_name, row_i)
        stage_idx = {'계약금': 0, '중도금': 1, '잔금': 2}[stage]
        note = notes[stage_idx] if stage_idx < len(notes) else ''
        print(f'=== {code} {stage} (row {row_i}) ===')
        if note:
            print(note)
        else:
            print('(empty)')
        print()
    return 0


if __name__ == '__main__':
    sys.exit(main())
