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
        logger.error(f'[PAYMENT_ALERT_DAILY] 메모 fetch 실패: {exc}', exc_info=True)
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

    # ─────────────────────────────────────────────────────────────
    # 통합 입금 그룹 감지 (2026-07-11)
    #   같은 stage 에 완전히 동일한 payments 시그니처가 여러 프로젝트에 반복되고
    #   그 메모합이 관련 프로젝트들의 시트 개별값 합과 일치하면 "통합 입금".
    #   예: SK텔레콤 4지점 (G2855~G2858) 잔금 메모 모두 5,621,000원 = 시트 개별
    #   (1,045 + 1,826 + 1,045 + 1,705) 합계.
    #   그룹 소속 프로젝트+stage 는 sheet_note_mismatch 감지에서 skip.
    # ─────────────────────────────────────────────────────────────
    unified_groups: List[Dict] = []  # 최종 그룹 리스트
    _unified_skip: set = set()  # (row_i, stage) 튜플 — sheet_note_mismatch 스킵 대상

    def _payment_signature(payments: List[Dict], stage: str) -> tuple:
        """해당 stage payments 를 정렬된 튜플로 시그니처화 (동등 비교용)."""
        rows_ = sorted(
            ((p.get('amount', 0), (p.get('partner') or '').strip(),
              (p.get('date_md') or '').strip())
             for p in payments if p.get('stage') == stage),
        )
        return tuple(rows_)

    for _stage, _sheet_col in (('계약금', 'u'), ('중도금', 'v'), ('잔금', 'w')):
        # 시그니처별 그룹핑
        _by_sig: Dict[tuple, List[int]] = {}
        for _i, _p in all_parsed.items():
            sig = _payment_signature(_p['payments'], _stage)
            if not sig:
                continue
            # 시그니처 합 = 메모합
            note_sum = sum(a for a, _, _ in sig)
            if note_sum <= 0:
                continue
            _by_sig.setdefault(sig, []).append(_i)
        # 같은 시그니처가 2건 이상이고 시트합 = 메모합이면 통합 입금 그룹
        for sig, indices in _by_sig.items():
            if len(indices) < 2:
                continue
            note_sum = sum(a for a, _, _ in sig)
            sheet_sum = sum(all_parsed[_i][_sheet_col] for _i in indices)
            if sheet_sum <= 0:
                continue
            # 5% 이내 오차 허용
            if abs(note_sum - sheet_sum) / max(sheet_sum, 1) > 0.05:
                continue
            # 확정 — 그룹 등록
            codes = [all_parsed[_i]['code'] for _i in indices]
            addrs = [
                (all_parsed[_i]['row'][IDX['F']] if len(all_parsed[_i]['row']) > IDX['F'] else '')[:40]
                for _i in indices
            ]
            partners = sorted({p for _, p, _ in sig if p and p != '-'})
            unified_groups.append({
                'stage': _stage,
                'codes': codes,
                'addrs': addrs,
                'note_sum': note_sum,
                'sheet_sum': sheet_sum,
                'partners': partners,
            })
            for _i in indices:
                _unified_skip.add((_i, _stage))

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
                    # target 의 메모 텍스트에도 amount 언급 있는지
                    target_notes_text = '\n'.join(row_notes.get(
                        [k for k, v in all_parsed.items() if v['code'] == target['code']][0], []
                    ))
                    if str(amount // 1000) not in target_notes_text and f'{amount:,}' not in target_notes_text:
                        unmatched.append(rc)
            if unmatched:
                return {
                    'body': (
                        f"'{m.group(1)}' 참조로 {amount:,}원 표기했는데 "
                        f"{', '.join(unmatched)} 메모에 해당 기록이 없어요."
                    ),
                }
        return None

    def _check_sheet_note_mismatch(row: Dict, payments: List[Dict], u: int, v: int, w: int, notes_text: str = '', row_i: int = 0) -> Optional[Dict]:
        """각 stage 별 메모 amount 합 vs 시트 값 비교.

        필터:
          - AA(수금완료) 되지 않은 진행중 프로젝트는 skip — 아직 매니저가 노트
            채우는 중이라 자연스러운 불일치 (2026-07-12 사용자 결정)
          - 시트 음수 값 (반환 처리) skip
          - 카드 수수료·반올림 오차 (5% 미만) skip
          - 시트=0 or 메모=0 인 미기입 케이스는 다른 감지에서 처리
          - 메모에 '반환/안분/통합' 등 특수 계산 언급 있으면 skip
          - 참조 프로젝트 표기(partner='참조 GXXXX') 있으면 skip — 다른 프로젝트에서
            통합 입금 처리된 케이스 (예: SK텔레콤 4지점 합산 입금)
        """
        # AA=False (진행중) 이면 노트 미완성이 정상 → skip
        if not row.get('aa', False):
            return None
        # 참조 표기가 payments 에 하나라도 있으면 이 프로젝트의 메모-시트 비교는 무의미.
        #   2026-07-10 SK텔레콤 케이스 대응 — 여러 지점을 한 프로젝트에서 통합 입금하고
        #   나머지 지점 메모에는 참조만 표기하는 실사용 패턴. 개별 검증은 참조 프로젝트
        #   대응 감지(_check_reference_mismatch) 가 담당.
        for p in payments:
            partner = str(p.get('partner') or '').strip()
            if partner.startswith('참조 '):
                return None

        # 메모 특수 키워드 있으면 매니저가 개별 계산한 것으로 간주 → skip
        #   2026-07-10 확장 — '통합/합산/합쳐/지점' 도 SK텔레콤 등 다지점 통합 입금 대응.
        #   2026-07-11 확장 — '채권추심/수수료/제외' 도 매니저가 시트값 차액을 명시적으로
        #   표시한 케이스 (G1019-MW 관측).
        #   2026-07-12 확장 — '상계/매입' 도 기존제품 매입가에서 공사비 차감한 상계처리
        #   케이스 (G0273-SH 관측: "제품매입90만원"으로 80만원 공사비 상계).
        if any(kw in notes_text for kw in (
            '반환', '안분', '오입금', '매출반환', '차감',
            '통합', '합산', '합쳐', '지점 통합',
            '채권추심', '수수료', '제외',
            '상계', '매입',
        )):
            return None
        stage_totals = {'계약금': 0, '중도금': 0, '잔금': 0}
        for p in payments:
            stage = p.get('stage')
            if stage in stage_totals:
                stage_totals[stage] += p.get('amount', 0)
        sheet_vals = {'계약금': u, '중도금': v, '잔금': w}
        mismatches = []
        for stage in ('계약금', '중도금', '잔금'):
            # 2026-07-11 통합 입금 그룹 소속이면 skip (그룹 별도 알림에서 처리)
            if (row_i, stage) in _unified_skip:
                continue
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
                f"{stage}: 메모 {note_sum:,}원 ↔ 시트 {sheet_val:,}원"
            )
        if not mismatches:
            return None
        return {
            'body': f"메모 금액과 시트 값이 달라요: {'; '.join(mismatches)}",
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

        # 8개 감지 순회 (2026-07-10 참조 정합성/메모-시트 불일치 추가)
        for category, fn in [
            ('note_missing',        lambda: detect_note_missing(alert_row, payments)),
            ('amount_typo',         lambda: detect_amount_typo(alert_row, payments)),
            ('required_missing',    lambda: detect_required_missing(alert_row, payments)),
            ('complete_untick',     lambda: detect_complete_untick(alert_row, payments)),
            ('unpaid_invalid',      lambda: detect_unpaid_invalid(alert_row)),
            ('unknown_initial',     lambda: detect_unknown_initial(alert_row, known_initials)),
            ('reference_mismatch',  lambda: _check_reference_mismatch(code, payments)),
            ('sheet_note_mismatch', lambda: _check_sheet_note_mismatch(alert_row, payments, u, v, w, '\n'.join(notes), row_i=i)),
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

    # 2026-07-11 통합 입금 그룹을 informational alert 로 추가.
    #   매니저는 그룹 요약만 확인 (검토 불필요). 개별 프로젝트는 sheet_note_mismatch
    #   에서 이미 skip 됨.
    for g in unified_groups:
        codes_str = ', '.join(f'`{c}`' for c in g['codes'])
        partners = ', '.join(g['partners']) if g['partners'] else '입금자 미상'
        alerts.append({
            'code': g['codes'][0],
            'address': g['addrs'][0] if g['addrs'] else '',
            'category': 'unified_payment',
            'body': (
                f"{g['stage']} 통합 입금 {g['note_sum']:,}원 = 시트 개별 합계 "
                f"{g['sheet_sum']:,}원 — {len(g['codes'])}개 프로젝트 ({codes_str}), "
                f"입금자: {partners}"
            ),
        })

    result['issues'] = len(alerts)
    if send_daily_summary(alerts):
        result['sent'] = True
    logger.info(
        f"[PAYMENT_ALERT_DAILY] 스캔 완료: {result['scanned']}행, "
        f"이슈 {result['issues']}건, 발송 {result['sent']}"
    )
    return result
