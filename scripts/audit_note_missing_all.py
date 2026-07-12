"""AA=True 포함 모든 note_missing 프로젝트 리스트 감사.

시트값 U/V/W > 0 이지만 노트 파싱 fallback (partner='-', date='-') 인 케이스.
Daily 알림은 AA skip 이라 안 뜨지만 회계·이력 정확성 위해 정리 대상.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from collections import Counter

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

try:
    from dotenv import load_dotenv
    load_dotenv(_ROOT / '.env')
except Exception:
    pass


def _int(v):
    try:
        return int(round(float(str(v).replace(',','').strip() or 0)))
    except Exception:
        return 0


def main() -> int:
    from dashboard.services.payment_sync import (
        _get_payment_service, _fetch_row_notes, _parse_notes,
    )
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

    IDX_A, IDX_F, IDX_U, IDX_V, IDX_W, IDX_AA = (
        col('A'), col('F'), col('U'), col('V'), col('W'), 26,
    )

    # 노트 batch fetch
    resp_notes = svc.spreadsheets().get(
        spreadsheetId=sheet_id,
        ranges=[f"'{sheet_name}'!U2:W{len(rows) + 1}"],
        fields='sheets.data.rowData.values.note',
        includeGridData=True,
    ).execute()
    row_notes = {}
    if resp_notes.get('sheets'):
        row_data_list = resp_notes['sheets'][0]['data'][0].get('rowData', [])
        for offset, rd in enumerate(row_data_list):
            vals = rd.get('values', [])
            n3 = [v.get('note', '') or '' for v in vals]
            while len(n3) < 3:
                n3.append('')
            row_notes[offset + 2] = n3[:3]

    missing = []  # (code, addr, stage, val, aa)
    for i, r in enumerate(rows, start=2):
        code = str(r[IDX_A] if len(r) > IDX_A else '').strip()
        if not code:
            continue
        u = _int(r[IDX_U] if len(r) > IDX_U else 0)
        v = _int(r[IDX_V] if len(r) > IDX_V else 0)
        w = _int(r[IDX_W] if len(r) > IDX_W else 0)
        stages_with_val = [(s, val) for s, val in [('계약금', u), ('중도금', v), ('잔금', w)] if val > 0]
        if not stages_with_val:
            continue
        aa_raw = r[IDX_AA] if len(r) > IDX_AA else ''
        aa = aa_raw is True or (isinstance(aa_raw, str) and aa_raw.strip().upper() in ('TRUE', 'Y'))
        addr = str(r[IDX_F] if len(r) > IDX_F else '')[:60]

        notes = row_notes.get(i, ['', '', ''])
        payments = _parse_notes(notes, stage_vals={'계약금': u, '중도금': v, '잔금': w})
        # fallback 인 stage 만 (date='-' + partner='-')
        for stage_name, val in stages_with_val:
            sp = [p for p in payments if p.get('stage') == stage_name]
            if sp and sp[-1].get('date_md') == '-' and sp[-1].get('partner') == '-':
                missing.append((code, addr, stage_name, val, aa))

    print(f'[*] 전체 note_missing 감지: {len(missing)} 건')
    print()

    # AA 별 분리
    aa_true = [m for m in missing if m[4]]
    aa_false = [m for m in missing if not m[4]]

    print(f'=== AA=True (완결 처리, 노트 미기입 옛 프로젝트): {len(aa_true)} 건 ===')
    # 프로젝트 번호 대략적 분포 표시
    prefix_count = Counter()
    for code, addr, stage, val, aa in aa_true:
        # 프로젝트 번호 앞자리
        try:
            num = int(''.join(c for c in code if c.isdigit())[:4])
            bucket = (num // 500) * 500
            prefix_count[f'{bucket}-{bucket+499}'] += 1
        except Exception:
            pass
    print('   번호대 분포:')
    for bucket, cnt in sorted(prefix_count.items()):
        print(f'     {bucket}: {cnt}건')
    print()

    print(f'=== AA=False (진행중 or 미완결 - Daily 알림 대상): {len(aa_false)} 건 ===')
    for code, addr, stage, val, aa in aa_false:
        print(f'  {code:<12}  {stage:<3}  {val:>12,}원  {addr}')
    print()

    print('[AA=True 상위 30건 샘플]')
    for code, addr, stage, val, aa in aa_true[:30]:
        print(f'  {code:<12}  {stage:<3}  {val:>12,}원  {addr}')
    if len(aa_true) > 30:
        print(f'  ... ({len(aa_true) - 30}건 더)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
