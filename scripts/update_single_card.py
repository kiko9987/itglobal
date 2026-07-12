"""단일 프로젝트 카드 chat.update.

usage: python scripts/update_single_card.py <CODE> [<STAGE>]
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


CODE_RE = re.compile(r'\*([GRNP]{1,2}\d{4}(?:-[A-Z]{2,3})?)\*')
STAGE_TITLE_RE = re.compile(r'\*(계약금 입금|중도금 입금|잔금 입금|수금완료|통합 입금|수금완료 \(통합 입금\))\*')


def _stage_from_title(t):
    if '계약금' in t:
        return '계약금'
    if '중도금' in t:
        return '중도금'
    return '잔금'


def _int(v):
    try:
        return int(round(float(v))) if v not in ('', None) else 0
    except Exception:
        return 0


def main() -> int:
    if len(sys.argv) < 2:
        print('usage: update_single_card.py <CODE> [<STAGE>]')
        return 1
    target_code = sys.argv[1]
    target_stage = sys.argv[2] if len(sys.argv) > 2 else None

    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()

    from dashboard.services.payment_sync import (
        _get_payment_service, _fetch_row_notes, _parse_notes,
        _build_complete_message, _build_stage_message,
        _build_stage_with_history_message,
        _TRANSFER_RE,
    )
    from slack_sdk import WebClient

    svc = _get_payment_service()
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{sheet_name}'!A2:AA10000",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = resp.get('values', [])

    def col(c):
        return ord(c) - ord('A')

    IDX = {c: col(c) for c in 'AFLRSTUVWXY'}

    info = None
    for i, r in enumerate(rows):
        code = str(r[IDX['A']] if len(r) > IDX['A'] else '').strip()
        if code == target_code:
            def _g(k):
                return r[IDX[k]] if len(r) > IDX[k] else ''
            info = {
                'row': i + 2,
                'address': str(_g('F')).strip(),
                'construction': str(_g('L')).strip(),
                'invoice': str(_g('Y')).strip(),
                'total_t': _int(_g('T')),
                'total_r': _int(_g('R')),
                'u': _int(_g('U')),
                'v': _int(_g('V')),
                'w': _int(_g('W')),
                'unpaid': _int(_g('X')),
            }
            break
    if not info:
        print(f'[!] {target_code} 시트에 없음')
        return 1

    # 슬랙 히스토리에서 카드 찾기
    slack = WebClient(token=token)
    card_ts = None
    card_stage = None
    cursor = None
    for _ in range(30):  # 최대 6000 messages
        kwargs = {'channel': channel, 'limit': 200}
        if cursor:
            kwargs['cursor'] = cursor
        resp = slack.conversations_history(**kwargs)
        for m in resp.get('messages', []):
            text = m.get('text', '') or ''
            for ln in text.split('\n'):
                if STAGE_TITLE_RE.search(ln) and target_code in ln:
                    sm = STAGE_TITLE_RE.search(ln)
                    stage = _stage_from_title(sm.group(1))
                    if target_stage and stage != target_stage:
                        continue
                    if card_ts is None:  # 첫 발견 (가장 최신)
                        card_ts = m.get('ts')
                        card_stage = stage
                    break
        if card_ts:
            break
        cursor = resp.get('response_metadata', {}).get('next_cursor')
        if not cursor:
            break

    if not card_ts:
        print(f'[!] {target_code} 슬랙 카드 못 찾음')
        return 1

    print(f'[*] {target_code} {card_stage} ts={card_ts}')

    # 노트 + 카드 build
    notes = _fetch_row_notes(sheet_id, sheet_name, info['row'])
    stage_vals = {'계약금': info['u'], '중도금': info['v'], '잔금': info['w']}
    payments = _parse_notes(notes, stage_vals=stage_vals)
    print(f'[*] 파싱 payments ({len(payments)}):')
    for p in payments:
        print(f'  {p.get("stage")} amt={p.get("amount"):,} date={p.get("date_md")} partner={p.get("partner")}')

    if card_stage == '잔금' and info['unpaid'] == 0:
        # 수금완료 카드 — 전체 history 취합, stage 필터 불필요
        text = _build_complete_message(
            project=target_code, address=info['address'],
            payments=payments, invoice_value=info['invoice'],
            total_t=info['total_t'],
            stage_sheet_vals=stage_vals,
            construction=info['construction'],
        )
    else:
        stage_payments = [p for p in payments if p.get('stage') == card_stage]
        if not stage_payments:
            print(f'[!] {card_stage} payment 없음')
            return 1
        last_payment = stage_payments[-1]
        if card_stage in ('중도금', '잔금'):
            text = _build_stage_with_history_message(
                stage=card_stage, project=target_code, address=info['address'],
                last_payment=last_payment, all_payments=payments,
                invoice_value=info['invoice'],
                total_r=info['total_r'], total_t=info['total_t'],
                unpaid=info['unpaid'],
                stage_sheet_vals=stage_vals,
                construction=info['construction'],
            )
        else:
            text = _build_stage_message(
                stage=card_stage, project=target_code, address=info['address'],
                payment=last_payment, invoice_value=info['invoice'],
                total_r=info['total_r'], total_t=info['total_t'],
                unpaid=info['unpaid'],
                stage_sheet_val=stage_vals.get(card_stage, 0),
                construction=info['construction'],
            )

    # 특이사항 append
    _SEP = '--------------------------------------------'
    _stage_by_idx = ('계약금', '중도금', '잔금')
    _date_prefix_re = re.compile(
        r'^(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2})\s*'
        r'(?:\d{1,2}:\d{2}\s*)?'
    )
    _special_re = re.compile(r'(?:제외|차감|채권추심|반환|안분|상계|매입)')
    _transfer_md_re = re.compile(r'^\d{4}-(\d{1,2})-(\d{1,2})')
    _special_entries = []
    for _idx, _note in enumerate(notes):
        if not _note:
            continue
        _stage_name = _stage_by_idx[_idx] if _idx < 3 else ''
        for _ln in _note.splitlines():
            _ln = _ln.strip()
            if not _ln or _ln.startswith('입금 '):
                continue
            _tr = _TRANSFER_RE.search(_ln)
            if _tr:
                _md_m = _transfer_md_re.match(_ln)
                _md = f'{int(_md_m.group(1)):02d}/{int(_md_m.group(2)):02d}' if _md_m else ''
                _a, _b = _tr.group(1).upper(), _tr.group(2).upper()
                _cleaned = f'{_md} {_a}>{_b} 매출이동'.strip()
            elif _special_re.search(_ln):
                _cleaned = _date_prefix_re.sub('', _ln).strip()
            else:
                continue
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

    print()
    print('=== 새 카드 텍스트 ===')
    print(text)
    print()

    r = slack.chat_update(channel=channel, ts=card_ts, text=text)
    if r.get('ok'):
        print(f'[OK] chat.update 성공')
    else:
        print(f'[!] {r.get("error")}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
