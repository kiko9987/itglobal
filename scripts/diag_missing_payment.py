"""누락 3건 진단 — 시트값 / 메모 파싱 / phash / 시트 modifiedTime 확인."""
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

from dashboard.services.payment_sync import (  # noqa: E402
    _get_payment_service,
    _fetch_row_notes,
    _parse_notes,
    _hash_payments,
)
from dashboard.utils.redis_client import get_redis_client  # noqa: E402


TARGETS = [
    (1368, 'G1367-SH'),
    (1445, 'G1444-MW'),
    (1450, 'G1449-YG'),
]


def _to_int_won(v) -> int:
    try:
        if isinstance(v, (int, float)):
            return int(v)
        s = str(v).replace(',', '').replace('￦', '').replace('₩', '').strip()
        if not s:
            return 0
        return int(float(s))
    except Exception:
        return 0


def main() -> int:
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    svc = _get_payment_service()
    if not svc:
        return 1
    rc = get_redis_client().redis

    for row_num, code in TARGETS:
        print(f'\n=== {code} (row {row_num}) ===')
        # 시트 값
        resp = svc.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A{row_num}:AA{row_num}",
            valueRenderOption='UNFORMATTED_VALUE',
        ).execute()
        row = (resp.get('values') or [[]])[0]
        def _get(idx):
            return row[idx] if idx < len(row) else ''
        u = _to_int_won(_get(20))  # U
        v = _to_int_won(_get(21))  # V
        w = _to_int_won(_get(22))  # W
        print(f'  시트값: U(계약금)={u:,}  V(중도금)={v:,}  W(잔금)={w:,}')

        # Redis 저장 값
        key = f'payment_sync:row:{row_num}'
        h = rc.hgetall(key)
        # decode if bytes
        h_dec = {}
        for k, val in (h or {}).items():
            k_str = k.decode('utf-8') if isinstance(k, bytes) else k
            v_str = val.decode('utf-8') if isinstance(val, bytes) else val
            h_dec[k_str] = v_str
        print(f'  Redis 저장값: {h_dec}')

        # 노트 파싱
        notes = _fetch_row_notes(sheet_id, sheet_name, row_num)
        payments = _parse_notes(notes, stage_vals={'계약금': u, '중도금': v, '잔금': w})
        print(f'  노트 파싱 payments ({len(payments)} 개):')
        for p in payments:
            print(f'    stage={p.get("stage")}  amount={p.get("amount"):,}  date={p.get("date_md")}  partner={p.get("partner")}')
        print(f'  new_phash: {_hash_payments(payments)}')

        # 슬랙 ts 매핑
        for stage in ('계약금', '중도금', '잔금'):
            ts = rc.get(f'payment_slack:ts:{code}:{stage}')
            if ts:
                if isinstance(ts, bytes):
                    ts = ts.decode('utf-8')
                print(f'  payment_slack:ts:{code}:{stage} = {ts}')

    return 0


if __name__ == '__main__':
    sys.exit(main())
