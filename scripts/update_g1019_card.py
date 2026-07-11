"""G1019-MW 기존 슬랙 수금완료 카드에 특이사항 라인 추가 chat.update.

payment_slack:ts:G1019-MW:잔금 매핑 조회 → 카드 build (특이사항 포함) →
slack.chat_update 로 기존 카드 수정.
"""
from __future__ import annotations

import os
import re
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

try:
    from dotenv import load_dotenv
    load_dotenv(_ROOT / '.env')
except Exception:
    pass


def _int(v):
    try:
        return int(round(float(str(v).replace(',', '').strip() or 0)))
    except (ValueError, TypeError):
        return 0


def main() -> int:
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()

    from dashboard.services.payment_sync import (
        _get_payment_service, _fetch_row_notes, _parse_notes,
        _build_complete_message,
    )
    from dashboard.utils.redis_client import get_redis_client
    from slack_sdk import WebClient

    rc = get_redis_client().redis
    # 슬랙 링크에서 ts 직접 지정 (07-10 15:07 이전 발송분이라 Redis 없음)
    ts = '1783653646.065319'
    print(f'[*] 기존 카드 ts: {ts}')

    svc = _get_payment_service()
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{sheet_name}'!A1020:AA1020",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    r = resp.get('values', [[]])[0]

    def _get(idx):
        return r[idx] if idx < len(r) else ''

    def col(c):
        return ord(c) - ord('A')

    address = str(_get(col('F'))).strip()
    construction = str(_get(col('L'))).strip()
    invoice = str(_get(col('Y'))).strip()
    total_t = _int(_get(col('T')))
    u = _int(_get(col('U')))
    v = _int(_get(col('V')))
    w = _int(_get(col('W')))

    notes = _fetch_row_notes(sheet_id, sheet_name, 1020)
    payments = _parse_notes(notes, stage_vals={'계약금': u, '중도금': v, '잔금': w})

    text = _build_complete_message(
        project='G1019-MW', address=address,
        payments=payments, invoice_value=invoice,
        total_t=total_t,
        stage_sheet_vals={'계약금': u, '중도금': v, '잔금': w},
        construction=construction,
    )

    # 특이사항 라인 (payment_sync 로직과 동일)
    _SEP = '--------------------------------------------'
    _stage_by_idx = ('계약금', '중도금', '잔금')
    _date_prefix_re = re.compile(
        r'^(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2})\s*'
        r'(?:\d{1,2}:\d{2}\s*)?'
    )
    _special_re = re.compile(r'(?:제외|차감|채권추심|반환|안분)')
    _special_entries = []
    for _idx, _note in enumerate(notes):
        if not _note:
            continue
        _stage_name = _stage_by_idx[_idx] if _idx < 3 else ''
        for _ln in _note.splitlines():
            _ln = _ln.strip()
            if not _ln or _ln.startswith('입금 '):
                continue
            if not _special_re.search(_ln):
                continue
            _cleaned = _date_prefix_re.sub('', _ln).strip()
            if not _cleaned:
                continue
            _entry = (_stage_name, _cleaned)
            if _entry not in _special_entries:
                _special_entries.append(_entry)
    if _special_entries:
        _special_block = ['', '[특이사항]']
        for _stg, _cln in _special_entries[:5]:
            _prefix = f'{_stg} ' if _stg else ''
            _special_block.append(f'{_prefix}{_cln}')
        _special_text = '\n'.join(_special_block)
        _parts = text.rsplit(_SEP, 1)
        if len(_parts) == 2:
            text = _parts[0].rstrip() + '\n' + _special_text + '\n' + _SEP + _parts[1]
        print(f'[*] 특이사항 {len(_special_entries)}개:')
        for _stg, _cln in _special_entries:
            print(f'  {_stg} {_cln}')

    print()
    print('=== 새 카드 텍스트 ===')
    print(text)
    print()

    slack = WebClient(token=token)
    resp = slack.chat_update(channel=channel, ts=ts, text=text)
    if resp.get('ok'):
        print('[OK] chat.update 성공')
    else:
        print(f'[!] chat.update 실패: {resp.get("error")}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
