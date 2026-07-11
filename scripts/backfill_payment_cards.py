"""누락 카드 backfill — 각 프로젝트의 최신 payment stage 카드 발송.

사용자가 슬랙에서 수동으로 지운 카드들을 다시 발송.
발송 후 Redis phash + ts 저장으로 재발송 방지.

usage: python scripts/backfill_payment_cards.py
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


# (project_code, sheet_row, target_stage)
# target_stage 는 발송할 stage (계약금/중도금/잔금)
TARGETS = [
    ('R3692-MJ', '중도금'),  # 07/10 1,000,000 + 2,000,000 두 블록 (파서 개선 후 3M 다 파싱)
    ('R3854-SJ', '계약금'),  # 재발송 — 사용자가 슬랙에서 이전 카드 지운 후
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
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    bot_token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    if not all([sheet_id, sheet_name, channel, bot_token]):
        print('필수 환경변수 미설정')
        return 1

    from dashboard.services.payment_sync import (
        _get_payment_service,
        _fetch_row_notes,
        _parse_notes,
        _hash_payments,
        _build_stage_message,
        _build_complete_message,
        _build_stage_with_history_message,
    )
    from dashboard.utils.redis_client import get_redis_client
    from slack_sdk import WebClient

    svc = _get_payment_service()
    slack = WebClient(token=bot_token)
    rc = get_redis_client().redis

    # 시트 전체 fetch (row 찾기 위해)
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{sheet_name}'!A2:AA10000",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = resp.get('values', [])

    def col_idx(c):
        if len(c) == 1:
            return ord(c) - ord('A')
        return (ord(c[0]) - ord('A') + 1) * 26 + (ord(c[1]) - ord('A'))
    IDX = {c: col_idx(c) for c in 'AFLRTUVWXY'}
    IDX_AA = col_idx('AA')

    for code, target_stage in TARGETS:
        # row 찾기
        row_idx = None
        for i, r in enumerate(rows):
            if len(r) > IDX['A'] and str(r[IDX['A']]).strip() == code:
                row_idx = i
                break
        if row_idx is None:
            print(f'[!] {code}: 시트에서 row 못 찾음 → skip')
            continue
        sheet_row = row_idx + 2  # header + 1-based
        row = rows[row_idx]

        def _get(idx):
            return row[idx] if idx < len(row) else ''

        u = _to_int_won(_get(IDX['U']))
        v = _to_int_won(_get(IDX['V']))
        w = _to_int_won(_get(IDX['W']))
        unpaid = _to_int_won(_get(IDX['X']))
        total_r = _to_int_won(_get(IDX['R']))
        total_t = _to_int_won(_get(IDX['T']))
        address = str(_get(IDX['F'])).strip()
        construction = str(_get(IDX['L'])).strip()
        invoice = str(_get(IDX['Y'])).strip()

        # 노트 파싱
        stage_vals = {'계약금': u, '중도금': v, '잔금': w}
        notes = _fetch_row_notes(sheet_id, sheet_name, sheet_row)
        payments = _parse_notes(notes, stage_vals=stage_vals)

        stage_payments = [p for p in payments if p.get('stage') == target_stage]
        if not stage_payments:
            print(f'[!] {code} {target_stage}: 노트에 이 stage payment 없음 → skip')
            continue
        last_payment = stage_payments[-1]

        # 카드 build
        if target_stage == '잔금' and unpaid == 0:
            text = _build_complete_message(
                project=code, address=address,
                payments=payments, invoice_value=invoice,
                total_t=total_t,
                stage_sheet_vals=stage_vals,
                construction=construction,
            )
        elif target_stage in ('중도금', '잔금'):
            text = _build_stage_with_history_message(
                stage=target_stage, project=code, address=address,
                last_payment=last_payment, all_payments=payments,
                invoice_value=invoice,
                total_r=total_r, total_t=total_t,
                unpaid=unpaid,
                stage_sheet_vals=stage_vals,
                construction=construction,
            )
        else:  # 계약금
            text = _build_stage_message(
                stage=target_stage, project=code, address=address,
                payment=last_payment, invoice_value=invoice,
                total_r=total_r, total_t=total_t,
                unpaid=unpaid,
                stage_sheet_val=stage_vals.get(target_stage, 0),
                construction=construction,
            )

        # 스레드 연결 (이전 stage 카드 있으면)
        thread_ts = ''
        for prev_stage in ('계약금', '중도금'):
            if prev_stage == target_stage:
                break
            prev_ts = rc.get(f'payment_slack:ts:{code}:{prev_stage}')
            if prev_ts:
                thread_ts = prev_ts if isinstance(prev_ts, str) else prev_ts.decode()
                break

        # 발송
        post_kwargs = {'channel': channel, 'text': text}
        if thread_ts:
            post_kwargs['thread_ts'] = thread_ts
            post_kwargs['reply_broadcast'] = True

        try:
            resp = slack.chat_postMessage(**post_kwargs)
        except Exception as exc:
            print(f'[!] {code} {target_stage}: 슬랙 발송 실패 — {exc}')
            continue

        if not resp.get('ok'):
            print(f'[!] {code} {target_stage}: 슬랙 응답 실패 — {resp.get("error")}')
            continue

        ts = resp.get('ts', '')
        # Redis ts 저장 (90일)
        if ts:
            rc.set(f'payment_slack:ts:{code}:{target_stage}', ts, ex=60 * 60 * 24 * 90)

        # phash 저장 (재발송 방지)
        new_phash = _hash_payments(payments)
        row_key = f'payment_sync:row:{sheet_row}'
        rc.hset(row_key, 'phash', new_phash)

        print(f'[OK] {code} {target_stage} 발송 완료 (row {sheet_row}, ts={ts})')

    print()
    print('완료.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
