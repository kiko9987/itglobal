"""21건 카드 partner/date 개선 후 chat.update.

각 카드의 이전 payment 이력 vs 새 이력 비교 리스트업.
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

CODE_RE = re.compile(r'\*([GRNP]{1,2}\d{4}(?:-[A-Z]{2,3})?)\*')
STAGE_TITLE_RE = re.compile(r'\*(계약금 입금|중도금 입금|잔금 입금|수금완료|통합 입금|수금완료 \(통합 입금\))\*')
HISTORY_LINE_RE = re.compile(
    r'^(계약금|중도금|잔금)\s+(\S+)\s+(.+?)\s+([\d,]+)원\s+(.*?)(?:\s*\(카드\))?$'
)


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
    channel = os.getenv('SLACK_PAYMENT_CHANNEL', '').strip()
    token = os.getenv('SLACK_PAYMENT_BOT_TOKEN', '').strip()
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()

    from dashboard.services.payment_sync import (
        _get_payment_service, _fetch_row_notes, _parse_notes,
        _build_complete_message, _build_stage_message,
        _build_stage_with_history_message,
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

    code_to_row = {}
    for i, r in enumerate(rows):
        code = str(r[IDX['A']] if len(r) > IDX['A'] else '').strip()
        if code:
            code_to_row[code] = (i + 2, r)

    # 슬랙 fetch (넉넉히)
    slack = WebClient(token=token)
    print('[*] 슬랙 히스토리 fetch (최대 5000)...')
    all_msgs = []
    cursor = None
    LIMIT = 5000
    while len(all_msgs) < LIMIT:
        kwargs = {'channel': channel, 'limit': 200}
        if cursor:
            kwargs['cursor'] = cursor
        r = slack.conversations_history(**kwargs)
        msgs = r.get('messages', [])
        if not msgs:
            break
        all_msgs.extend(msgs)
        cursor = r.get('response_metadata', {}).get('next_cursor')
        if not cursor:
            break
    print(f'[*] {len(all_msgs)} messages')
    print()

    # (code, stage) → (ts, message text)
    card_map = {}
    for m in all_msgs:
        text = m.get('text', '') or ''
        title_line = ''
        for ln in text.split('\n'):
            if STAGE_TITLE_RE.search(ln):
                title_line = ln
                break
        if not title_line:
            continue
        sm = STAGE_TITLE_RE.search(title_line)
        stage = _stage_from_title(sm.group(1))
        for c in CODE_RE.findall(title_line):
            key = (c, stage)
            if key not in card_map:
                card_map[key] = (m.get('ts'), text)

    results = []  # 카드 update 결과
    for code, stage in TARGETS:
        rowinfo = code_to_row.get(code)
        if not rowinfo:
            results.append((code, stage, '시트 없음', '', ''))
            continue
        row_i, r = rowinfo

        def _g(k):
            return r[IDX[k]] if len(r) > IDX[k] else ''

        info = {
            'row': row_i,
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

        ts_text = card_map.get((code, stage))
        if not ts_text:
            results.append((code, stage, '슬랙 카드 없음', '', ''))
            continue
        ts, old_text = ts_text

        # 이전 이력 라인 파싱
        old_lines = []
        in_hist = False
        for ln in old_text.split('\n'):
            if '[입금 이력]' in ln or '[누적 이력]' in ln:
                in_hist = True
                continue
            if '---' in ln or ln.startswith('총액') or ln.startswith('미수금'):
                in_hist = False
                continue
            if not in_hist:
                continue
            m = HISTORY_LINE_RE.match(ln.strip())
            if m:
                old_lines.append(f'{m.group(1)} {m.group(2)} {m.group(4)}원 {m.group(5) or "-"}')

        # 새 카드 build
        notes = _fetch_row_notes(sheet_id, sheet_name, info['row'])
        stage_vals = {'계약금': info['u'], '중도금': info['v'], '잔금': info['w']}
        payments = _parse_notes(notes, stage_vals=stage_vals)
        stage_payments = [p for p in payments if p.get('stage') == stage]
        if not stage_payments:
            results.append((code, stage, 'payment 없음', '\n    '.join(old_lines), ''))
            continue
        last_payment = stage_payments[-1]

        if stage == '잔금' and info['unpaid'] == 0:
            text = _build_complete_message(
                project=code, address=info['address'],
                payments=payments, invoice_value=info['invoice'],
                total_t=info['total_t'],
                stage_sheet_vals=stage_vals,
                construction=info['construction'],
            )
        elif stage in ('중도금', '잔금'):
            text = _build_stage_with_history_message(
                stage=stage, project=code, address=info['address'],
                last_payment=last_payment, all_payments=payments,
                invoice_value=info['invoice'],
                total_r=info['total_r'], total_t=info['total_t'],
                unpaid=info['unpaid'],
                stage_sheet_vals=stage_vals,
                construction=info['construction'],
            )
        else:
            text = _build_stage_message(
                stage=stage, project=code, address=info['address'],
                payment=last_payment, invoice_value=info['invoice'],
                total_r=info['total_r'], total_t=info['total_t'],
                unpaid=info['unpaid'],
                stage_sheet_val=stage_vals.get(stage, 0),
                construction=info['construction'],
            )

        # 특이사항 append (동일 로직)
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

        # 새 이력 라인 파싱
        new_lines = []
        in_hist = False
        for ln in text.split('\n'):
            if '[입금 이력]' in ln or '[누적 이력]' in ln:
                in_hist = True
                continue
            if '---' in ln or ln.startswith('총액') or ln.startswith('미수금') or ln.startswith('['):
                in_hist = False
                continue
            if not in_hist:
                continue
            m = HISTORY_LINE_RE.match(ln.strip())
            if m:
                new_lines.append(f'{m.group(1)} {m.group(2)} {m.group(4)}원 {m.group(5) or "-"}')

        try:
            r_up = slack.chat_update(channel=channel, ts=ts, text=text)
            if r_up.get('ok'):
                status = 'OK'
            else:
                status = f'실패 {r_up.get("error")}'
        except Exception as exc:
            status = f'예외 {exc}'

        results.append((code, stage, status, ' / '.join(old_lines), ' / '.join(new_lines)))

    print(f'{"code":<12}  {"stage":<3}  status  |  이전  →  이후')
    print('-' * 130)
    for code, stage, status, old, new in results:
        print(f'{code:<12}  {stage:<3}  {status:<10}')
        if old:
            print(f'  이전: {old}')
        if new:
            print(f'  이후: {new}')
        print()
    return 0


if __name__ == '__main__':
    sys.exit(main())
