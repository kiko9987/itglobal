"""시트 T열, X열 수식 일괄 업데이트.

T열: =if($S:S=TRUE, ROUND($R:R*1.1, 0), $R:R)
X열: =IF(ABS($T:T-$U:U-$V:V-$W:W)<2, 0, $T:T-$U:U-$V:V-$W:W)

각 행에 동일 수식 입력. 실행 전 현재 값 백업 → 새 수식 적용 → 검증.
"""
from __future__ import annotations

import json
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


T_FORMULA = '=if($S:S=TRUE, ROUND($R:R*1.1, 0), $R:R)'
X_FORMULA = '=IF(ABS($T:T-$U:U-$V:V-$W:W)<2, 0, $T:T-$U:U-$V:V-$W:W)'


def main() -> int:
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    from dashboard.services.payment_sync import _get_payment_service
    svc = _get_payment_service()

    # 1. 백업 — T, X 현재 값 저장
    print('[1] 백업 조회 중...')
    resp = svc.spreadsheets().values().batchGet(
        spreadsheetId=sheet_id,
        ranges=[f"'{sheet_name}'!T2:T3855", f"'{sheet_name}'!X2:X3855"],
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    ranges = resp.get('valueRanges', [])
    t_vals = [(r or [''])[0] if r else '' for r in ranges[0].get('values', [])]
    x_vals = [(r or [''])[0] if r else '' for r in ranges[1].get('values', [])]
    backup = {
        'sheet': sheet_name,
        't_column': t_vals,
        'x_column': x_vals,
    }
    backup_path = _ROOT / 'scratchpad_backup_t_x.json'
    backup_path.parent.mkdir(exist_ok=True)
    with open(backup_path, 'w', encoding='utf-8') as f:
        json.dump(backup, f, ensure_ascii=False)
    print(f'    T {len(t_vals)}행, X {len(x_vals)}행 백업 → {backup_path}')

    # 2. 새 수식 업데이트
    print('[2] T열 수식 업데이트...')
    n_rows = 3854
    t_values = [[T_FORMULA] for _ in range(n_rows)]
    x_values = [[X_FORMULA] for _ in range(n_rows)]
    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': [
            {
                'range': f"'{sheet_name}'!T2:T{n_rows + 1}",
                'values': t_values,
            },
            {
                'range': f"'{sheet_name}'!X2:X{n_rows + 1}",
                'values': x_values,
            },
        ],
    }
    resp = svc.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id, body=body,
    ).execute()
    print(f'    응답: {resp.get("totalUpdatedCells", 0)} 셀 업데이트됨')

    # 3. 검증 — T, X 값 재조회 후 이상 케이스 찾기
    print('[3] 검증 중...')
    resp = svc.spreadsheets().values().batchGet(
        spreadsheetId=sheet_id,
        ranges=[
            f"'{sheet_name}'!A2:A{n_rows + 1}",
            f"'{sheet_name}'!F2:F{n_rows + 1}",
            f"'{sheet_name}'!R2:R{n_rows + 1}",
            f"'{sheet_name}'!T2:T{n_rows + 1}",
            f"'{sheet_name}'!U2:U{n_rows + 1}",
            f"'{sheet_name}'!V2:V{n_rows + 1}",
            f"'{sheet_name}'!W2:W{n_rows + 1}",
            f"'{sheet_name}'!X2:X{n_rows + 1}",
        ],
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    ranges = resp.get('valueRanges', [])

    def _flat(idx):
        return [(v or [''])[0] if v else '' for v in (ranges[idx].get('values') or [])]

    codes = _flat(0)
    addrs = _flat(1)
    rs = _flat(2)
    ts = _flat(3)
    us = _flat(4)
    vs = _flat(5)
    ws = _flat(6)
    xs = _flat(7)

    def _int(v):
        try:
            return int(round(float(v))) if v not in ('', None) else 0
        except Exception:
            return 0

    # 이상 케이스: X != 0 이면서 T-U-V-W 부호와 다른 것, T 값 이상, X > 100 등
    issues = []
    for i in range(len(codes)):
        code = str(codes[i]).strip() if i < len(codes) else ''
        if not code:
            continue
        R = _int(rs[i]) if i < len(rs) else 0
        T = _int(ts[i]) if i < len(ts) else 0
        U = _int(us[i]) if i < len(us) else 0
        V = _int(vs[i]) if i < len(vs) else 0
        W = _int(ws[i]) if i < len(ws) else 0
        X = _int(xs[i]) if i < len(xs) else 0

        # 예상 X = T - U - V - W (임계값 2 미만은 0)
        raw_diff = T - U - V - W
        expected_X = 0 if abs(raw_diff) < 2 else raw_diff

        # 이상 감지
        if X != expected_X:
            issues.append((code, 'X 계산 불일치', T, U, V, W, X, expected_X, addrs[i][:40]))
        # 옛 X 값이 큰데 지금 0 이 된 것 (오탐 감소)
        if abs(X) >= 100:
            issues.append((code, '큰 미수금', T, U, V, W, X, expected_X, addrs[i][:40]))

    print(f'[*] 검증 결과: 총 {len(issues)} 건 확인 대상')
    print()
    for code, tag, T, U, V, W, X, expected_X, addr in issues[:50]:
        print(f'  {code:<12} [{tag}]  T={T:,} U={U:,} V={V:,} W={W:,} X={X:,} (예상 {expected_X:,})  {addr}')
    if len(issues) > 50:
        print(f'  ... ({len(issues) - 50} 건 더)')

    return 0


if __name__ == '__main__':
    sys.exit(main())
