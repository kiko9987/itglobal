"""A/S (사후 관리) 서비스 레이어.

- Google Sheets 'A/S 관리' 시트 read/write
- AS-XXXX 자동 발번
- 프로젝트 자동완성 (공사 확정된 것만)
- 방문 예정자 후보 (시공자 + 영업 담당자 + 서비스 기사)
"""
from __future__ import annotations

import os
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


# ─────────────────────────────────────────────────────────────
# 시트 컬럼 매핑 (2026-07-09)
# ─────────────────────────────────────────────────────────────
COL_NO = 'A'
COL_PROJECT_CODE = 'B'
COL_ADDRESS = 'C'
COL_WORK_CONTENT = 'D'
COL_WORK_END = 'E'
COL_REQUEST_CONTENT = 'F'
COL_REQUESTER = 'G'
COL_ACCEPTER = 'H'
COL_ACCEPT_DATE = 'I'
COL_VISITOR = 'J'
COL_VISIT_DATE = 'K'
COL_STATUS = 'L'
COL_RESOLUTION = 'M'

STATUS_REQUESTED = '요청됨'
STATUS_ACCEPTED = '접수 완료'
STATUS_COMPLETED = '처리 완료'


def _get_sheet() -> Tuple[Any, str, str]:
    """(manager, sheet_id, sheet_name) 반환. 환경변수 미설정 시 예외."""
    from dashboard.utils.google_sheets import GoogleSheetsManager
    sheet_id = os.getenv('AS_SHEET_ID', '').strip()
    sheet_name = os.getenv('AS_SHEET_NAME', 'A/S 관리').strip()
    if not sheet_id:
        raise RuntimeError('AS_SHEET_ID 미설정')
    return GoogleSheetsManager(), sheet_id, sheet_name


# ─────────────────────────────────────────────────────────────
# 발번 / 행 조회
# ─────────────────────────────────────────────────────────────
def get_next_as_no() -> str:
    """시트 A열에서 최대 AS-XXXX 번호+1 반환. 초기값 AS-0001."""
    manager, sheet_id, sheet_name = _get_sheet()
    try:
        resp = manager.service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f'{sheet_name}!A:A',
            valueRenderOption='FORMATTED_VALUE',
        ).execute()
        values = resp.get('values', [])
    except Exception as exc:
        logger.error(f'[AS] A열 조회 실패: {exc}', exc_info=True)
        return 'AS-0001'
    max_n = 0
    for row in values[1:]:  # header skip
        if not row:
            continue
        m = re.match(r'^AS-(\d+)', str(row[0]).strip())
        if m:
            try:
                max_n = max(max_n, int(m.group(1)))
            except ValueError:
                pass
    return f'AS-{max_n + 1:04d}'


def find_row_by_as_no(as_no: str) -> Optional[int]:
    """A열에서 as_no 행 번호 조회 (1-based)."""
    manager, sheet_id, sheet_name = _get_sheet()
    try:
        resp = manager.service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f'{sheet_name}!A:A',
            valueRenderOption='FORMATTED_VALUE',
        ).execute()
        values = resp.get('values', [])
        for i, row in enumerate(values, start=1):
            if row and str(row[0]).strip() == as_no:
                return i
    except Exception as exc:
        logger.warning(f'[AS] 행 번호 조회 실패 ({as_no}): {exc}')
    return None


def get_as_data(as_no: str) -> Optional[Dict[str, Any]]:
    """A/S 행 전체 데이터 반환."""
    manager, sheet_id, sheet_name = _get_sheet()
    row_num = find_row_by_as_no(as_no)
    if not row_num:
        return None
    try:
        resp = manager.service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f'{sheet_name}!A{row_num}:M{row_num}',
            valueRenderOption='FORMATTED_VALUE',
        ).execute()
        values = resp.get('values', [])
        if not values:
            return None
        row = list(values[0]) + [''] * (13 - len(values[0]))
        headers = [
            'No', '프로젝트 코드', '현장주소', '공사내용', '공사 종료일',
            '요청 내용', '요청자', '접수자', '접수 일자',
            '방문 예정자', '방문 예정일', '진행 상태', '조치 내용',
        ]
        return dict(zip(headers, row))
    except Exception as exc:
        logger.warning(f'[AS] 데이터 조회 실패 ({as_no}): {exc}')
        return None


_AS_HEADERS = [
    'No', '프로젝트 코드', '현장주소', '공사내용', '공사 종료일',
    '요청 내용', '요청자', '접수자', '접수 일자',
    '방문 예정자', '방문 예정일', '진행 상태', '조치 내용',
]


def list_as(status_filter: Optional[str] = None) -> List[Dict[str, Any]]:
    """A/S 전체 목록 반환 (AS번호 내림차순, 최신 먼저). PM 대시보드 A/S 모드용.

    status_filter: 지정 시 해당 진행 상태만 (예: STATUS_REQUESTED). None=전체.
    각 항목: _AS_HEADERS 키 dict + '_row'(시트 행번호).
    """
    manager, sheet_id, sheet_name = _get_sheet()
    try:
        resp = manager.service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f'{sheet_name}!A2:M',
            valueRenderOption='FORMATTED_VALUE',
        ).execute()
        rows = resp.get('values', [])
    except Exception as exc:
        logger.error(f'[AS] 목록 조회 실패: {exc}', exc_info=True)
        return []
    items: List[Dict[str, Any]] = []
    for i, r in enumerate(rows, start=2):  # 시트 행번호 (헤더=1)
        if not r or not str(r[0]).strip():
            continue
        r = list(r) + [''] * (13 - len(r))
        d = dict(zip(_AS_HEADERS, r))
        if status_filter and str(d.get('진행 상태', '')).strip() != status_filter:
            continue
        d['_row'] = i
        items.append(d)

    def _num(d: Dict[str, Any]) -> int:
        m = re.match(r'^AS-(\d+)', str(d.get('No', '')).strip())
        return int(m.group(1)) if m else 0

    items.sort(key=_num, reverse=True)
    return items


# ─────────────────────────────────────────────────────────────
# 시트 mutation
# ─────────────────────────────────────────────────────────────
def create_as_row(
    project_code: str, address: str, work_content: str, work_end: str,
    request_content: str, requester: str,
) -> Tuple[str, Optional[int]]:
    """새 A/S 행 생성. (as_no, row_number) 반환. 실패 시 (as_no, None)."""
    manager, sheet_id, sheet_name = _get_sheet()
    as_no = get_next_as_no()
    # 요청 시각 — 접수 완료 시점이 아니라 요청 발생 시점을 I열 대신 별도 저장하지 않음.
    # 시트 스키마상 요청 일자 컬럼이 없어 진행 상태에만 요청됨 기록.
    row_values = [
        as_no,               # A
        project_code,        # B
        address,             # C
        work_content,        # D
        work_end,            # E
        request_content,     # F
        requester,           # G
        '',                  # H (접수자)
        '',                  # I (접수 일자)
        '',                  # J (방문 예정자)
        '',                  # K (방문 예정일)
        STATUS_REQUESTED,    # L
        '',                  # M (조치 내용)
    ]
    try:
        resp = manager.service.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range=f'{sheet_name}!A:M',
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body={'values': [row_values]},
        ).execute()
        # 갱신 range 예: 'A/S 관리'!A3:M3 → 행 번호 추출
        updated_range = resp.get('updates', {}).get('updatedRange', '')
        m = re.search(r'!A(\d+):', updated_range)
        row_number = int(m.group(1)) if m else None
        logger.info(f'[AS] 신규 요청 등록: {as_no} (row={row_number})')
        return as_no, row_number
    except Exception as exc:
        logger.error(f'[AS] 행 추가 실패: {exc}', exc_info=True)
        return as_no, None


def update_as_row(as_no: str, updates: Dict[str, Any]) -> bool:
    """부분 필드 갱신. updates 키는 컬럼 letter (H/I/J/K/L/M 등)."""
    manager, sheet_id, sheet_name = _get_sheet()
    row_num = find_row_by_as_no(as_no)
    if not row_num:
        logger.warning(f'[AS] 갱신 대상 행 없음: {as_no}')
        return False
    batch = []
    for col, val in updates.items():
        batch.append({
            'range': f'{sheet_name}!{col}{row_num}',
            'values': [[val if val is not None else '']],
        })
    if not batch:
        return False
    try:
        # 2026-07-09 write-behind: A/S 시트 갱신을 큐로 위임
        from dashboard.services.sheet_write_queue import enqueue
        enqueue('sheet_batch_write', {
            'sheet_id': sheet_id, 'updates': batch, 'tag': f'as_update:{as_no}',
        })
        logger.info(f'[AS] 갱신 큐 위임: {as_no} row={row_num} fields={list(updates.keys())}')
        return True
    except Exception as exc:
        logger.error(f'[AS] 갱신 큐 실패 ({as_no}): {exc}', exc_info=True)
        return False


def append_as_log(as_no: str, content: str, initial: str = '') -> Optional[str]:
    """M열(메모/이력)에 '[MM.DD HH:MM 이니셜] 내용' 한 줄 append (온라인 리드 누적 방식).

    접수 메모·조치 내용 등 A/S 코멘트를 새 컬럼 없이 M열 한 셀에 시간순 누적한다.
    반환: append 후의 M열 전체 값 (실패 시 None). content 가 비면 현재 M값 그대로 반환.
    ※ write-behind 큐 경유라 시트 반영엔 약간의 지연이 있을 수 있음(반환값으로 즉시 표시 가능).
    """
    content = (content or '').strip()
    data = get_as_data(as_no) or {}
    cur = str(data.get('조치 내용', '') or '')
    if not content:
        return cur
    stamp = datetime.now().strftime('%m.%d %H:%M')
    entry = f'[{stamp} {initial}] {content}' if initial else f'[{stamp}] {content}'
    new_val = (cur.rstrip() + '\n' + entry) if cur.strip() else entry
    ok = update_as_row(as_no, {COL_RESOLUTION: new_val})
    return new_val if ok else None


# ─────────────────────────────────────────────────────────────
# 프로젝트 자동완성 (공사 확정된 것만)
# ─────────────────────────────────────────────────────────────
def search_confirmed_projects(query: str, limit: int = 100) -> List[Dict[str, str]]:
    """공사 확정된 프로젝트 검색. 프로젝트 코드 + 사업자명 부분 매칭. 최신순 정렬.

    각 결과: {'code', 'biz', 'address', 'work_content', 'work_end'}
    """
    if not query:
        return []
    try:
        import pandas as pd
        from dashboard.services.project_service import load_data
        df = load_data()
        if df is None or df.empty or '프로젝트 코드' not in df.columns:
            return []
        # 공사 확정된 것만 — '공사 확정' 컬럼에 날짜 있는 행
        confirmed_col = '공사 확정' if '공사 확정' in df.columns else None
        if confirmed_col:
            mask = df[confirmed_col].astype(str).str.strip().replace('', None).notna()
            mask &= df[confirmed_col].astype(str).str.strip() != ''
            filtered = df[mask].copy()
            # 최신순 정렬 — 공사 확정 날짜 descending
            filtered['_confirmed_dt'] = pd.to_datetime(
                filtered[confirmed_col], errors='coerce',
            )
            filtered = filtered.sort_values('_confirmed_dt', ascending=False, na_position='last')
        else:
            filtered = df

        q = query.lower()
        matched = []
        for _, r in filtered.iterrows():
            code = str(r.get('프로젝트 코드', '') or '').strip()
            biz = str(r.get('사업자명', '') or '').strip()
            if not code:
                continue
            address = str(r.get('현장 주소', '') or '').strip()
            # 코드 OR 사업자명 OR 주소 부분 매칭 (주소로도 좁히기 — 반복거래처 대응)
            if not (q in code.lower() or (biz and q in biz.lower())
                    or (address and q in address.lower())):
                continue
            matched.append({
                'code': code,
                'biz': biz,
                'address': address,
                'work_content': str(r.get('공사 내용', '') or '').strip(),
                'work_end': str(r.get('공사 종료', '') or '').strip()[:10],
            })
            if len(matched) >= limit:
                break
        return matched
    except Exception as exc:
        logger.warning(f'[AS] 프로젝트 검색 실패: {exc}')
        return []


def get_project_details(code: str) -> Optional[Dict[str, str]]:
    """프로젝트 코드로 상세 조회 (모달 pre-fill 용).

    공사 확정 카드와 동일한 정보량 반환.
    """
    try:
        from dashboard.services.project_service import get_project_records
        records = get_project_records() or []
        r = next(
            (rec for rec in records if (rec.get('프로젝트 코드') or '').strip() == code),
            None,
        )
        if not r:
            return None

        def _s(k: str) -> str:
            return str(r.get(k, '') or '').strip()

        # 금액 표기 (부가세 반영)
        amt_raw = r.get('총액 1', '')
        try:
            amt_int = int(float(str(amt_raw).replace(',', '').strip() or 0))
            amt_disp = f'{amt_int:,}원' if amt_int else '-'
        except (ValueError, TypeError):
            amt_int = 0
            amt_disp = '-'
        vat_raw = r.get('부가세')
        vat_sep = (
            vat_raw is True
            or (isinstance(vat_raw, str) and vat_raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
            or vat_raw == 1
        )
        if amt_disp != '-':
            amt_disp = f"{amt_disp} ({'VAT 별도' if vat_sep else 'VAT 없음'})"

        # 유입 구분 표시 정규화 (공사 확정 카드 _format_inflow 와 동일)
        raw_inflow = _s('유입 구분')
        online_platforms = {'홈페이지', '카카오톡', '당근', '기타'}
        if raw_inflow in online_platforms:
            inflow_disp = f'온라인 ({raw_inflow})'
        else:
            inflow_disp = raw_inflow or '-'

        return {
            'code': code,
            'inflow': inflow_disp,
            'manager': _s('담당자') or '-',  # 영업 담당자 (AS 알림 DM 대상)
            'biz': _s('사업자명') or '-',
            'address': _s('현장 주소') or '-',
            'client_manager': _s('발주처 담당자') or '-',
            'client_phone': _s('발주처 연락처') or '-',
            'client_email': _s('발주처 이메일') or '-',
            'work_content': _s('공사 내용') or '-',
            'contract_type': _s('도급 구분') or '-',
            'contractor': _s('시공자') or '-',
            'amount': amt_disp,
            'amount_raw': str(amt_int) if amt_int else '',  # 계산서 모달 입력칸용 raw 숫자
            'work_start': _s('공사 시작')[:10] or '-',
            'work_end': _s('공사 종료')[:10] or '-',
        }
    except Exception as exc:
        logger.warning(f'[AS] 프로젝트 상세 조회 실패 ({code}): {exc}')
        return None


# ─────────────────────────────────────────────────────────────
# 담당자 식별 (알림 DM 대상)
# ─────────────────────────────────────────────────────────────
def _email_from_manager_name(name: str) -> str:
    """영업 담당자 한국 이름 → 이메일. users.db 매핑."""
    name = (name or '').strip()
    if not name or name == '-':
        return ''
    try:
        from dashboard.utils.user_database import UserDatabase
        db = UserDatabase()
        for u in db.get_all_users():
            if (u.get('name') or '').strip() == name:
                return (u.get('email') or '').strip()
    except Exception as exc:
        logger.warning(f'[AS] 담당자 이메일 조회 실패 ({name}): {exc}')
    return ''


def resolve_project_manager(project_code: str,
                            proj: Optional[Dict[str, str]] = None) -> Tuple[str, str]:
    """프로젝트 코드 → (담당자 이름, 이메일).

    우선순위:
      1. 프로젝트 '담당자' 필드 (proj 미전달 시 조회)
      2. 코드 접미어 이니셜 (`G1038-JW` → JW) → users.db 로컬 파트 매칭 fallback

    Returns: (name, email). 못 찾으면 ('', '').
    """
    # 1) 담당자 필드
    if proj is None:
        proj = get_project_details(project_code) or {}
    mgr_name = (proj.get('manager', '') or '').strip()
    if mgr_name and mgr_name != '-':
        email = _email_from_manager_name(mgr_name)
        if email:
            return mgr_name, email

    # 2) 코드 접미어 이니셜 fallback (G1038-JW → JW)
    m = re.search(r'-([A-Za-z]{2,4})\s*$', str(project_code or '').strip())
    if m:
        ini = m.group(1).upper()
        try:
            from dashboard.utils.user_database import UserDatabase
            db = UserDatabase()
            for u in db.get_all_users():
                email = (u.get('email') or '').strip()
                if email and email.split('@')[0].strip().upper() == ini:
                    return (u.get('name') or ini).strip(), email
        except Exception as exc:
            logger.warning(f'[AS] 코드 접미어 담당자 조회 실패 ({project_code}): {exc}')
    return mgr_name if mgr_name and mgr_name != '-' else '', ''


# ─────────────────────────────────────────────────────────────
# 방문 예정자 후보 (시공자 + 영업 담당자 + 서비스 기사)
# ─────────────────────────────────────────────────────────────
def list_visitor_candidates() -> List[str]:
    """방문 예정자 후보 리스트.

    - 활성 시공자 (카테고리 flat, 이름만)
    - 영업 담당자 (SALES_INITIALS 키의 한국 이름)
    - '서비스 기사' 정적 옵션
    중복 제거. 순서 유지.
    """
    names: List[str] = []
    seen = set()

    def _push(n: str) -> None:
        n = (n or '').strip()
        if not n or n in seen:
            return
        seen.add(n)
        names.append(n)

    # 활성 시공자
    try:
        from dashboard.utils.user_database import get_constructor_repository
        grouped = get_constructor_repository().get_grouped(active_only=True)
        for cat_items in grouped.values():
            for c in cat_items:
                _push(c.get('name') or '')
    except Exception as exc:
        logger.warning(f'[AS] 시공자 로드 실패: {exc}')

    # 영업 담당자
    try:
        from dashboard.blueprints.slack_helpers import SALES_INITIALS
        for korean_name in SALES_INITIALS.keys():
            _push(korean_name)
    except Exception as exc:
        logger.warning(f'[AS] 영업 담당자 로드 실패: {exc}')

    # 서비스 기사 (정적)
    _push('서비스 기사')

    return names
