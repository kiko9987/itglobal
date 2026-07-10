"""수금 관리 매니저 실수 일일 스캔 + 요약 발송 (2026-07-10).

매일 오전 9시 (평일만) 실행. 전체 시트 순회 → 6개 카테고리 감지 →
매니저 친화적 요약을 #수금_관리 채널에 공개 발송.

폴링 훅 (payment_sync 내부) 은 U/V/W 변화 시점만 감지 → 즉시 알림.
이 일일 스캔은 변화 없이 방치된 실수도 모두 잡음.
"""
from __future__ import annotations

import os
from typing import Dict, List

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


def _int(v) -> int:
    try:
        return int(float(str(v).replace(',', '').strip() or 0))
    except (ValueError, TypeError):
        return 0


def _col_idx(col: str) -> int:
    if len(col) == 1:
        return ord(col) - ord('A')
    return (ord(col[0]) - ord('A') + 1) * 26 + (ord(col[1]) - ord('A'))


def run_daily_scan_and_summary() -> Dict:
    """전체 시트 스캔 → 감지 → 매니저 친화적 요약 채널 발송."""
    from dashboard.services.payment_sync import _get_payment_service, _parse_notes
    from dashboard.services.payment_alert import (
        detect_note_missing, detect_amount_typo, detect_required_missing,
        detect_complete_untick, detect_unpaid_invalid, detect_unknown_initial,
        send_daily_summary, CATEGORY_META,
    )
    from dashboard.blueprints.slack_helpers import _load_initials_from_config

    result = {'scanned': 0, 'issues': 0, 'sent': False}
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    if not (sheet_id and sheet_name):
        logger.debug('[PAYMENT_ALERT_DAILY] 환경변수 미설정 — skip')
        return result

    service = _get_payment_service()
    if service is None:
        logger.warning('[PAYMENT_ALERT_DAILY] Google Sheets service 초기화 실패')
        return result

    try:
        resp = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A2:AA10000",
            valueRenderOption='UNFORMATTED_VALUE',
        ).execute()
    except Exception as exc:
        logger.error(f'[PAYMENT_ALERT_DAILY] 시트 fetch 실패: {exc}', exc_info=True)
        return result
    rows = resp.get('values', [])

    try:
        resp_notes = service.spreadsheets().get(
            spreadsheetId=sheet_id,
            ranges=[f"'{sheet_name}'!U2:W10000"],
            fields='sheets.data.rowData.values.note',
            includeGridData=True,
        ).execute()
    except Exception as exc:
        logger.error(f'[PAYMENT_ALERT_DAILY] 노트 fetch 실패: {exc}', exc_info=True)
        return result
    row_notes = {}
    if resp_notes.get('sheets'):
        row_data_list = resp_notes['sheets'][0]['data'][0].get('rowData', [])
        for offset_n, rd in enumerate(row_data_list):
            vals = rd.get('values', [])
            n3 = [v.get('note', '') or '' for v in vals]
            while len(n3) < 3:
                n3.append('')
            row_notes[offset_n + 2] = n3[:3]

    IDX = {c: _col_idx(c) for c in ['A', 'F', 'T', 'U', 'V', 'W', 'X', 'AA']}
    known_initials = set(_load_initials_from_config().values())

    # 참조 정합성 검증용 인덱스 (2026-07-10)
    # 프로젝트 코드 prefix (예: G2172) → (row_idx, code, payments) 매핑
    import re as _re
    _CODE_PREFIX_RE = _re.compile(r'^([GRN]\d{4})-')
    project_by_prefix: Dict[str, Dict] = {}

    # 1차 순회: 모든 프로젝트의 payments 파싱해 인덱스 구축
    all_parsed = {}  # (i) → {'code', 'payments', 'row'}
    for i, row in enumerate(rows, start=2):
        code = (row[IDX['A']] if len(row) > IDX['A'] else '').strip()
        if not code:
            continue
        u = _int(row[IDX['U']] if len(row) > IDX['U'] else 0)
        v = _int(row[IDX['V']] if len(row) > IDX['V'] else 0)
        w = _int(row[IDX['W']] if len(row) > IDX['W'] else 0)
        notes = row_notes.get(i, ['', '', ''])
        payments = _parse_notes(notes, stage_vals={'계약금': u, '중도금': v, '잔금': w})
        all_parsed[i] = {'code': code, 'payments': payments, 'row': row, 'u': u, 'v': v, 'w': w}
        m = _CODE_PREFIX_RE.match(code)
        if m:
            project_by_prefix[m.group(1)] = all_parsed[i]

    def _check_reference_mismatch(code: str, payments: List[Dict]) -> Optional[Dict]:
        """참조된 프로젝트에 상응 금액 기록 없으면 감지."""
        _ref_re = _re.compile(r'^참조 ([GRN]\d{4}(?:/\d{4})?)$')
        for p in payments:
            partner = p.get('partner') or ''
            m = _ref_re.match(partner)
            if not m:
                continue
            amount = p.get('amount', 0)
            # 참조 코드 리스트 (예: G2172/2173 → ['G2172', 'G2173'])
            ref_codes = []
            parts = m.group(1).split('/')
            first_letter = parts[0][0]
            ref_codes.append(parts[0])
            for extra in parts[1:]:
                ref_codes.append(f'{first_letter}{extra}')
            # 각 참조 프로젝트에 해당 amount 언급 있는지 확인
            unmatched = []
            for rc in ref_codes:
                target = project_by_prefix.get(rc)
                if not target:
                    unmatched.append(f'{rc} (프로젝트 없음)')
                    continue
                # target 의 payments 중 amount 매칭 확인 (10% 오차 허용)
                found = any(
                    abs(tp.get('amount', 0) - amount) <= max(amount * 0.1, 1000)
                    for tp in target['payments']
                )
                if not found:
                    # target 의 노트 텍스트에도 amount 언급 있는지
                    target_notes_text = '\n'.join(row_notes.get(
                        [k for k, v in all_parsed.items() if v['code'] == target['code']][0], []
                    ))
                    if str(amount // 1000) not in target_notes_text and f'{amount:,}' not in target_notes_text:
                        unmatched.append(rc)
            if unmatched:
                return {
                    'body': (
                        f"'{m.group(1)}' 참조로 {amount:,}원 표기했는데 "
                        f"{', '.join(unmatched)} 노트에 해당 기록이 없어요."
                    ),
                }
        return None

    def _check_sheet_note_mismatch(row: Dict, payments: List[Dict], u: int, v: int, w: int, notes_text: str = '') -> Optional[Dict]:
        """각 stage 별 노트 amount 합 vs 시트 값 비교.

        필터:
          - 시트 음수 값 (반환 처리) skip
          - 카드 수수료·반올림 오차 (5% 미만) skip
          - 시트=0 or 노트=0 인 미기입 케이스는 다른 감지에서 처리
          - 노트에 '반환/안분' 언급 있으면 skip (매니저 특수 계산)
        """
        # 노트 특수 키워드 있으면 매니저가 개별 계산한 것으로 간주 → skip
        if any(kw in notes_text for kw in ('반환', '안분', '오입금', '매출반환', '차감')):
            return None
        stage_totals = {'계약금': 0, '중도금': 0, '잔금': 0}
        for p in payments:
            stage = p.get('stage')
            if stage in stage_totals:
                stage_totals[stage] += p.get('amount', 0)
        sheet_vals = {'계약금': u, '중도금': v, '잔금': w}
        mismatches = []
        for stage in ('계약금', '중도금', '잔금'):
            note_sum = stage_totals[stage]
            sheet_val = sheet_vals[stage]
            # 시트 음수 (오입금 반환 처리) skip
            if sheet_val < 0:
                continue
            # 한쪽 0 이면 다른 감지에서 처리
            if sheet_val == 0 or note_sum == 0:
                continue
            diff = abs(note_sum - sheet_val)
            # 5% 미만 오차 (카드 수수료·반올림) skip. 최소 1,000원 오차.
            ratio = diff / max(sheet_val, 1)
            if ratio < 0.05 or diff < 1000:
                continue
            mismatches.append(
                f"{stage}: 노트 {note_sum:,}원 ↔ 시트 {sheet_val:,}원"
            )
        if not mismatches:
            return None
        return {
            'body': f"노트 금액과 시트 값이 달라요: {'; '.join(mismatches)}",
        }

    alerts: List[Dict] = []
    for i, row in enumerate(rows, start=2):
        code = (row[IDX['A']] if len(row) > IDX['A'] else '').strip()
        if not code:
            continue
        # 담당자 퇴사한 옛 프로젝트도 우리가 시공한 건이라 감지 대상 (2026-07-10 사용자 결정).
        # AA(수금완료) 된 건은 감지 함수 안에서 각자 skip.
        result['scanned'] += 1
        address = (row[IDX['F']] if len(row) > IDX['F'] else '').strip()
        total_t = _int(row[IDX['T']] if len(row) > IDX['T'] else 0)
        u = _int(row[IDX['U']] if len(row) > IDX['U'] else 0)
        v = _int(row[IDX['V']] if len(row) > IDX['V'] else 0)
        w = _int(row[IDX['W']] if len(row) > IDX['W'] else 0)
        unpaid = _int(row[IDX['X']] if len(row) > IDX['X'] else 0)
        aa_raw = row[IDX['AA']] if len(row) > IDX['AA'] else ''
        aa_chk = (
            aa_raw is True
            or (isinstance(aa_raw, str) and aa_raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
            or aa_raw == 1
        )
        notes = row_notes.get(i, ['', '', ''])
        payments = _parse_notes(notes, stage_vals={'계약금': u, '중도금': v, '잔금': w})

        alert_row = {
            'code': code, 'address': address, 'total_t': total_t,
            'unpaid': unpaid, 'aa': aa_chk,
            '계약금': u, '중도금': v, '잔금': w,
        }

        # 8개 감지 순회 (2026-07-10 참조 정합성/노트-시트 불일치 추가)
        for category, fn in [
            ('note_missing',        lambda: detect_note_missing(alert_row, payments)),
            ('amount_typo',         lambda: detect_amount_typo(alert_row, payments)),
            ('required_missing',    lambda: detect_required_missing(alert_row, payments)),
            ('complete_untick',     lambda: detect_complete_untick(alert_row, payments)),
            ('unpaid_invalid',      lambda: detect_unpaid_invalid(alert_row)),
            ('unknown_initial',     lambda: detect_unknown_initial(alert_row, known_initials)),
            ('reference_mismatch',  lambda: _check_reference_mismatch(code, payments)),
            ('sheet_note_mismatch', lambda: _check_sheet_note_mismatch(alert_row, payments, u, v, w, '\n'.join(notes))),
        ]:
            try:
                detected = fn()
            except Exception:
                continue
            if detected is None:
                continue
            alerts.append({
                'code': code,
                'address': address,
                'category': category,
                'body': detected.get('body', ''),
            })

    result['issues'] = len(alerts)
    if send_daily_summary(alerts):
        result['sent'] = True
    logger.info(
        f"[PAYMENT_ALERT_DAILY] 스캔 완료: {result['scanned']}행, "
        f"이슈 {result['issues']}건, 발송 {result['sent']}"
    )
    return result
