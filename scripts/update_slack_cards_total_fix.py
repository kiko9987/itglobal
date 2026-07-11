"""시트 T열 float 오차 수정 후 이미 발송된 카드 chat.update.

전략:
1. 시트 fetch → 각 프로젝트 T/U/V/W 값
2. 백업 파일에서 이전 T raw 로드 → int(prev) != new_T 인 대상 리스트
3. 슬랙 채널 히스토리 fetch (충분히 오래까지)
4. 각 카드 title 에서 project code 매칭
5. 대상 카드에 대해 새 텍스트 build → chat.update
"""
from __future__ import annotations

import json
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


def _int_round(v):
    try:
        return int(round(float(v))) if v not in ('', None) else 0
    except Exception:
        return 0


def _int_floor(v):
    try:
        return int(float(v)) if v not in ('', None) else 0
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
    from slack_sdk.errors import SlackApiError

    # 백업 파일에서 이전 T raw
    backup_path = _ROOT / 'scratchpad_backup_t_x.json'
    with open(backup_path, encoding='utf-8') as f:
        backup = json.load(f)
    prev_t = backup['t_column']

    # 현재 시트 fetch
    svc = _get_payment_service()
    resp = svc.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{sheet_name}'!A2:AA3855",
        valueRenderOption='UNFORMATTED_VALUE',
    ).execute()
    rows = resp.get('values', [])

    def col(c):
        return ord(c) - ord('A')

    IDX = {c: col(c) for c in 'AFLRSTUVWXY'}
    IDX_AA = 26

    # 대상 리스트 (int_floor(prev_T) != round(prev_T)) 이면서 code 있음
    targets = {}  # code → {row info}
    for i, r in enumerate(rows):
        code = str(r[IDX['A']] if len(r) > IDX['A'] else '').strip()
        if not code:
            continue
        old_t_raw = prev_t[i] if i < len(prev_t) else ''
        if old_t_raw in ('', None):
            continue
        try:
            old_t_f = float(old_t_raw)
        except Exception:
            continue
        if _int_floor(old_t_f) == round(old_t_f):
            continue
        # 대상
        def _g(k):
            return r[IDX[k]] if len(r) > IDX[k] else ''
        targets[code] = {
            'row': i + 2,
            'address': str(_g('F')).strip(),
            'construction': str(_g('L')).strip(),
            'invoice': str(_g('Y')).strip(),
            'total_t': _int_round(_g('T')),
            'total_r': _int_round(_g('R')),
            'u': _int_round(_g('U')),
            'v': _int_round(_g('V')),
            'w': _int_round(_g('W')),
            'unpaid': _int_round(_g('X')),
            'old_t': _int_floor(old_t_f),
            'new_t': _int_round(old_t_f),
        }

    print(f'[*] 카드 update 대상: {len(targets)} 건')
    print()

    # 슬랙 히스토리 fetch (충분히)
    slack = WebClient(token=token)
    print('[*] 슬랙 채널 히스토리 fetch 중 (최대 5000)...')
    all_msgs = []
    cursor = None
    LIMIT = 5000
    fetched = 0
    while fetched < LIMIT:
        try:
            kwargs = {'channel': channel, 'limit': min(200, LIMIT - fetched)}
            if cursor:
                kwargs['cursor'] = cursor
            r = slack.conversations_history(**kwargs)
        except SlackApiError as exc:
            print(f'[!] history fetch 실패: {exc.response.get("error")}')
            break
        msgs = r.get('messages', [])
        all_msgs.extend(msgs)
        fetched += len(msgs)
        cursor = r.get('response_metadata', {}).get('next_cursor')
        if not cursor:
            break
    print(f'[*] 총 {len(all_msgs)} 메시지 fetch 됨')
    print()

    # 각 카드에서 code + stage 파싱 → 매핑
    # 최신 카드만 사용 (같은 code+stage 여러 개면 첫 것 = 가장 최신)
    card_map = {}  # (code, stage) → ts
    for m in all_msgs:
        text = m.get('text', '') or ''
        ts = m.get('ts', '')
        if not ts:
            continue
        title_line = ''
        for ln in text.split('\n'):
            if STAGE_TITLE_RE.search(ln):
                title_line = ln
                break
        if not title_line:
            continue
        sm = STAGE_TITLE_RE.search(title_line)
        if not sm:
            continue
        stage = _stage_from_title(sm.group(1))
        codes = CODE_RE.findall(title_line)
        for c in codes:
            key = (c, stage)
            if key not in card_map:
                card_map[key] = ts

    # 대상 대응 시나리오:
    #   각 target 마다 잔금 카드가 있음 (가장 대표적으로 수금완료)
    #   찾아서 update
    updated = 0
    not_found = []
    for code, info in targets.items():
        # 잔금 stage 카드 우선 (수금완료)
        ts = card_map.get((code, '잔금'))
        # 없으면 다른 stage
        if not ts:
            for s in ('중도금', '계약금'):
                ts = card_map.get((code, s))
                if ts:
                    break
        if not ts:
            not_found.append(code)
            continue

        # 카드 텍스트 재구성 - 원래 어느 stage 카드인지 알 수 없으니 잔금 우선
        stage = None
        for s in ('잔금', '중도금', '계약금'):
            if card_map.get((code, s)) == ts:
                stage = s
                break
        if stage is None:
            not_found.append(code)
            continue

        # 노트 파싱
        notes = _fetch_row_notes(sheet_id, sheet_name, info['row'])
        stage_vals = {'계약금': info['u'], '중도금': info['v'], '잔금': info['w']}
        payments = _parse_notes(notes, stage_vals=stage_vals)
        stage_payments = [p for p in payments if p.get('stage') == stage]
        if not stage_payments:
            not_found.append(f'{code} (stage {stage} payment 없음)')
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

        # 특이사항 append (payment_sync 로직과 동일)
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

        try:
            resp = slack.chat_update(channel=channel, ts=ts, text=text)
            if resp.get('ok'):
                updated += 1
                print(f'  ✓ {code} {stage} (T {info["old_t"]:,} → {info["new_t"]:,})')
            else:
                print(f'  ✗ {code} {stage}: {resp.get("error")}')
                not_found.append(f'{code} ({resp.get("error")})')
        except Exception as exc:
            print(f'  ✗ {code} {stage}: {exc}')
            not_found.append(f'{code} ({exc})')

    print()
    print(f'[*] update 성공: {updated}')
    print(f'[*] 실패/못 찾음: {len(not_found)}')
    for x in not_found[:20]:
        print(f'  - {x}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
