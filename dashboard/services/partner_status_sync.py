"""거래처 탭 사업자등록 상태 갱신 — 국세청 상태조회 → J(상태)/K(최종확인일) 기록.

거래처 탭이 '마지막 세금계산서 발행 시점 스냅샷'이라 상대방 폐업/신설을 못 따라가는
문제의 안전망. 폐업 번호로 세금계산서 발행하는 사고 방지 (2026-07-28 도입).

- 소스: 메인 시트(GOOGLE_SHEET_ID) '거래처' 탭 A열 등록번호
- 조회: dashboard.services.nts_status.check_business_status (국세청 odcloud)
- 기록: J열 = 상태(계속사업자/휴업자/폐업자/조회안됨, 폐업은 폐업일 병기), K열 = 최종확인일
"""
from __future__ import annotations

import logging
import os
from datetime import date
from typing import List, Optional, Tuple

from dashboard.services.nts_status import check_business_status, normalize_bno

logger = logging.getLogger(__name__)

_TAB = '거래처'
_COL_STATUS = 'J'
_COL_CHECKED = 'K'
_WRITE_CHUNK = 500  # batchUpdate 1회 range 수


def _sheet():
    from dashboard.utils.google_sheets import GoogleSheetsManager
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    if not sid:
        raise RuntimeError('GOOGLE_SHEET_ID 미설정')
    return GoogleSheetsManager(), sid


def _fmt_end(end: str) -> str:
    e = (end or '').strip()
    return f'{e[:4]}-{e[4:6]}-{e[6:8]}' if len(e) == 8 and e.isdigit() else e


def refresh_partner_status(dry_run: bool = True,
                            today_str: Optional[str] = None) -> dict:
    """거래처 탭 등록번호 전체 상태 조회 후 J/K 기록.

    dry_run=True: API 조회·집계만, 시트 미기록. (폐업 건수 확인용)
    dry_run=False: J(상태)/K(최종확인일) 실제 기록 + 헤더 세팅.

    Returns: {total_rows, queried, resolved, summary{상태:건수}, closed[(row,bno,end)], ...}
    """
    m, sid = _sheet()
    today = today_str or date.today().strftime('%Y-%m-%d')

    # A열 등록번호 로드 (행 번호 보존)
    a_vals = m.service.spreadsheets().values().get(
        spreadsheetId=sid, range=f'{_TAB}!A1:A',
    ).execute().get('values', [])
    rows: List[Tuple[int, str]] = []  # (sheet_row_1base, bno_norm)
    for i, r in enumerate(a_vals):
        nb = normalize_bno(r[0] if r else '')
        if nb:
            rows.append((i + 1, nb))

    bnos = list({nb for _, nb in rows})
    status_map = check_business_status(bnos)

    summary = {'계속사업자': 0, '휴업자': 0, '폐업자': 0, '조회안됨': 0}
    closed: List[Tuple[int, str, str]] = []
    updates: List[Tuple[int, str, str]] = []  # (row, status_text, checked)
    for row, nb in rows:
        d = status_map.get(nb)
        stt = (d or {}).get('b_stt', '').strip() if d else ''
        if stt == '폐업자':
            end_fmt = _fmt_end((d or {}).get('end_dt', ''))
            status_text = f'폐업자 (폐업 {end_fmt})' if end_fmt else '폐업자'
            closed.append((row, nb, end_fmt))
            summary['폐업자'] += 1
        elif stt in ('계속사업자', '휴업자'):
            status_text = stt
            summary[stt] += 1
        else:
            status_text = '조회안됨'
            summary['조회안됨'] += 1
        updates.append((row, status_text, today))

    result = {
        'total_rows': len(rows), 'queried': len(bnos),
        'resolved': len(status_map), 'summary': summary,
        'closed': closed, 'dry_run': dry_run, 'today': today,
    }
    if dry_run:
        return result

    # 실제 기록 — 헤더 + 행별 J:K (비데이터 행 보호 위해 행별 range)
    data = [{'range': f'{_TAB}!J1:K1', 'values': [['상태', '최종확인일']]}]
    for row, stt, chk in updates:
        data.append({'range': f'{_TAB}!J{row}:K{row}', 'values': [[stt, chk]]})

    written = 0
    for i in range(0, len(data), _WRITE_CHUNK):
        chunk = data[i:i + _WRITE_CHUNK]
        m.service.spreadsheets().values().batchUpdate(
            spreadsheetId=sid,
            body={'valueInputOption': 'USER_ENTERED', 'data': chunk},
        ).execute()
        written += len(chunk)
        logger.info(f'[거래처상태] 기록 진행 {written}/{len(data)}')
    result['written_ranges'] = written
    logger.info(
        f"[거래처상태] 완료 — 총 {len(rows)}행, 폐업 {summary['폐업자']}, "
        f"휴업 {summary['휴업자']}, 조회안됨 {summary['조회안됨']}"
    )
    return result
