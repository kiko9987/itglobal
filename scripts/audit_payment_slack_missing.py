"""Payment 알림 누락 감사.

시트에서 U/V/W(계약금/중도금/잔금) > 0 이면서 메모에 payment 블록이 있는데
Redis `payment_slack:ts:{project}:{stage}` 매핑이 없는 케이스 리스트업.

기존 도입 전에 이미 완료된 프로젝트는 제외 (메모에 payment 블록 자체 없음).
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# Repo root 부트스트랩
_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

# .env 로드
try:
    from dotenv import load_dotenv
    load_dotenv(_ROOT / '.env')
except Exception:
    pass

from dashboard.services.payment_sync import (  # noqa: E402
    _get_payment_service,
    _fetch_row_notes,
    _parse_notes,
)
from dashboard.utils.redis_client import get_redis_client  # noqa: E402


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
    if not sheet_id or not sheet_name:
        print('환경변수 GOOGLE_SHEET_ID / GOOGLE_SHEET_NAME 미설정')
        return 1

    svc = _get_payment_service()
    if not svc:
        print('payment service 초기화 실패')
        return 1

    resp = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{sheet_name}'!A2:AA10000",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = resp.get('values', [])
    if not rows:
        print('시트 rows 없음')
        return 0

    def col_idx(c: str) -> int:
        if len(c) == 1:
            return ord(c) - ord('A')
        return (ord(c[0]) - ord('A') + 1) * 26 + (ord(c[1]) - ord('A'))

    IDX_A = col_idx('A')
    IDX_F = col_idx('F')
    IDX_U = col_idx('U')
    IDX_V = col_idx('V')
    IDX_W = col_idx('W')
    IDX_AA = col_idx('AA')

    rc = get_redis_client().redis

    missing = []  # {row, project, stage, sheet_val, memo_amount, memo_partner, address}

    for i, row in enumerate(rows):
        sheet_row = i + 2

        def _get(idx: int):
            return row[idx] if idx < len(row) else ''

        project = str(_get(IDX_A)).strip()
        if not project:
            continue
        u = _to_int_won(_get(IDX_U))
        v = _to_int_won(_get(IDX_V))
        w = _to_int_won(_get(IDX_W))
        if u == 0 and v == 0 and w == 0:
            continue

        stage_vals = {'계약금': u, '중도금': v, '잔금': w}

        # payment_slack:ts:{project}:{stage} 확인
        stages_to_check = [s for s, val in stage_vals.items() if val > 0]
        missing_stages = []
        for s in stages_to_check:
            ts = rc.get(f'payment_slack:ts:{project}:{s}')
            if not ts:
                missing_stages.append(s)
        if not missing_stages:
            continue

        # 노트 fetch → payment 블록 있는지
        try:
            notes = _fetch_row_notes(sheet_id, sheet_name, sheet_row)
        except Exception as exc:
            print(f'[!]  row {sheet_row} ({project}) 노트 fetch 실패: {exc}')
            continue

        payments = _parse_notes(notes, stage_vals=stage_vals)
        if not payments:
            # 메모 자체 없음 → 도입 전 이미 완료된 프로젝트, 정상
            continue

        # 실제로 payment 블록이 있는데 매핑이 없다면 → 누락 후보
        for s in missing_stages:
            stage_payments = [p for p in payments if p.get('stage') == s]
            if not stage_payments:
                # 시트값은 있는데 메모에 이 stage 의 payment 없음 → 도입 전 데이터
                continue
            memo_sum = sum(int(p.get('amount', 0)) for p in stage_payments)
            missing.append({
                'row': sheet_row,
                'project': project,
                'stage': s,
                'sheet_val': stage_vals[s],
                'memo_sum': memo_sum,
                'memo_count': len(stage_payments),
                'address': str(_get(IDX_F)).strip(),
            })

    if not missing:
        print('[OK] 누락 없음')
        return 0

    print(f'[!] 누락 후보 {len(missing)} 건:')
    print()
    print(f"{'row':>5}  {'project':<12}  {'stage':<6}  {'sheet_val':>12}  {'memo_sum':>12}  {'cnt':>3}  address")
    print('-' * 100)
    for m in missing:
        print(
            f"{m['row']:>5}  {m['project']:<12}  {m['stage']:<6}  "
            f"{m['sheet_val']:>12,}  {m['memo_sum']:>12,}  {m['memo_count']:>3}  "
            f"{m['address'][:40]}"
        )
    return 0


if __name__ == '__main__':
    sys.exit(main())
