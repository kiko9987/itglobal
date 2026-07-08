"""공사 확정 카드 슬랙 액션 실행자 (내용 수정 / 공사 취소).

- Flask HTTP 라우트를 우회하고 내부 헬퍼를 직접 호출한다.
- 슬랙 이벤트는 이미 signing_secret 로 검증됐으므로 별도 로그인 세션이 필요 없다.
- 관리 사이트의 편집·취소 API와 동일한 시트 write / 캐시 무효화 / 감사 로그 / 슬랙 답글
  경로를 재사용해 UX·데이터가 일관되도록 한다.
"""
from __future__ import annotations

import os
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


# ─────────────────────────────────────────────────────────────
# 헬퍼 (공통)
# ─────────────────────────────────────────────────────────────
def _load_project(code: str) -> Optional[Dict[str, Any]]:
    from dashboard.services.project_service import get_project_records
    records = get_project_records() or []
    return next(
        (r for r in records if (r.get('프로젝트 코드') or '').strip() == code),
        None,
    )


def _get_sheet_context() -> Tuple[Any, str, str]:
    from dashboard.services.project_service import get_sheets_manager
    manager = get_sheets_manager()
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황')
    if not sheet_id:
        raise RuntimeError('GOOGLE_SHEET_ID 미설정')
    return manager, sheet_id, sheet_name


def _find_row_number(manager, sheet_id: str, sheet_name: str, code: str) -> Optional[int]:
    """A열 프로젝트 코드로 행 번호 조회. 관리 사이트 _find_project_and_row 와 동등."""
    try:
        return manager.find_row_by_project_code(sheet_id, code, f'{sheet_name}!A:A')
    except Exception as exc:
        logger.warning(f'[SLACK/action] 행 번호 조회 실패 ({code}): {exc}')
        return None


def _audit_log(**kwargs) -> None:
    """audit_repo.log_action 얇은 래퍼. 예외는 삼킴."""
    try:
        from dashboard.utils.user_database import get_audit_repository
        get_audit_repository().log_action(**kwargs)
    except Exception as exc:
        logger.warning(f'[SLACK/action] 감사 로그 실패: {exc}')


def _update_row_bg(manager, sheet_id: str, sheet_name: str, row_number: int, color_type: str) -> None:
    """행 배경색 갱신 — 취소 시 dark_grey, 재개 시 normal(흰색).

    관리 사이트 _update_project_background_color 와 동등. 실패해도 예외 안 던짐.
    """
    try:
        ok = manager.update_row_background_color(
            spreadsheet_id=sheet_id,
            sheet_name=sheet_name,
            row_number=row_number,
            color_type=color_type,
        )
        if ok:
            desc = '진한 회색' if color_type == 'dark_grey' else '흰색'
            logger.info(f'[SLACK/action] 행 배경색 갱신 완료: row={row_number} → {desc}')
        else:
            logger.warning(f'[SLACK/action] 행 배경색 갱신 실패: row={row_number}, color={color_type}')
    except Exception as exc:
        logger.warning(f'[SLACK/action] 행 배경색 예외: {exc}')


# ─────────────────────────────────────────────────────────────
# 공사 취소
# ─────────────────────────────────────────────────────────────
def perform_cancel(code: str, by_display_name: str) -> Dict[str, Any]:
    """관리 사이트 [공사 취소] 버튼과 동일 효과.

    수금 관련 특이사항='공사 취소', 수금 확인=FALSE, 공사 확정=''.
    감사 로그와 캐시 무효화까지 처리하고 결과 dict 반환.
    """
    project = _load_project(code)
    if not project:
        return {'ok': False, 'reason': 'not_found'}

    if re.search(r'공사\s*취소', project.get('수금 관련 특이사항', '') or ''):
        return {'ok': False, 'reason': 'already_cancelled', 'project': project}

    manager, sheet_id, sheet_name = _get_sheet_context()
    row_number = _find_row_number(manager, sheet_id, sheet_name, code)
    if not row_number:
        return {'ok': False, 'reason': 'row_not_found'}

    # 관리 사이트 _prepare_cancel_updates 와 동일 (2026-07 컬럼 시프트 반영)
    updates = [
        {'range': f'{sheet_name}!AH{row_number}', 'values': [['공사 취소']]},
        {'range': f'{sheet_name}!AA{row_number}', 'values': [['FALSE']]},
        {'range': f'{sheet_name}!AM{row_number}', 'values': [['']]},
    ]

    try:
        manager.batch_update_cells(sheet_id, updates)
    except Exception as exc:
        logger.error(f'[SLACK/취소] 시트 write 실패 ({code}): {exc}', exc_info=True)
        return {'ok': False, 'reason': 'sheet_write_failed'}

    # 캐시 부분 갱신 (실패 시 전체 무효화 fallback)
    try:
        from dashboard.services.project_service import (
            update_project_in_cache,
            invalidate_project_cache,
        )
        cache_updated = update_project_in_cache(code, {
            '수금 관련 특이사항': '공사 취소',
            '수금 확인': False,
            '공사 확정': '',
        })
        if not cache_updated:
            invalidate_project_cache(code)
    except Exception as exc:
        logger.warning(f'[SLACK/취소] 캐시 갱신 실패 ({code}): {exc}')

    # 감사 로그 (관리 사이트 _log_project_status_change 와 동일 스키마)
    _audit_log(
        user_email=f'slack:{by_display_name}',
        action='CANCEL_PROJECT',
        details=f'프로젝트 공사 취소: {code} (수금확인=FALSE, 공사확정일 초기화)',
        project_code=code,
        field_name='수금 관련 특이사항',
        old_value=project.get('수금 관련 특이사항', '-') or '-',
        new_value='공사 취소',
        ip_address='slack-bot',
    )

    # 행 배경색 → 진한 회색 (관리 사이트와 동일 UX)
    _update_row_bg(manager, sheet_id, sheet_name, row_number, 'dark_grey')

    logger.info(f'[SLACK/취소] 완료: {code} by slack:{by_display_name}')
    return {
        'ok': True,
        'project': project,  # 취소 전 스냅샷 (카드 UI 재렌더링용)
    }


def perform_uncancel(code: str, by_display_name: str) -> Dict[str, Any]:
    """공사 취소 되돌리기 (관리 사이트 [공사 재개] 와 동일)."""
    project = _load_project(code)
    if not project:
        return {'ok': False, 'reason': 'not_found'}

    if not re.search(r'공사\s*취소', project.get('수금 관련 특이사항', '') or ''):
        return {'ok': False, 'reason': 'already_active', 'project': project}

    manager, sheet_id, sheet_name = _get_sheet_context()
    row_number = _find_row_number(manager, sheet_id, sheet_name, code)
    if not row_number:
        return {'ok': False, 'reason': 'row_not_found'}

    today = datetime.now().strftime('%Y-%m-%d')
    updates = [
        {'range': f'{sheet_name}!AH{row_number}', 'values': [['']]},
        {'range': f'{sheet_name}!AM{row_number}', 'values': [[today]]},
    ]
    try:
        manager.batch_update_cells(sheet_id, updates)
    except Exception as exc:
        logger.error(f'[SLACK/재개] 시트 write 실패 ({code}): {exc}', exc_info=True)
        return {'ok': False, 'reason': 'sheet_write_failed'}

    try:
        from dashboard.services.project_service import (
            update_project_in_cache,
            invalidate_project_cache,
        )
        cache_updated = update_project_in_cache(code, {
            '수금 관련 특이사항': '',
            '공사 확정': today,
        })
        if not cache_updated:
            invalidate_project_cache(code)
    except Exception as exc:
        logger.warning(f'[SLACK/재개] 캐시 갱신 실패 ({code}): {exc}')

    _audit_log(
        user_email=f'slack:{by_display_name}',
        action='RESUME_PROJECT',
        details=f'프로젝트 공사 재개: {code}',
        project_code=code,
        field_name='수금 관련 특이사항',
        old_value='공사 취소',
        new_value='',
        ip_address='slack-bot',
    )

    # 행 배경색 → 흰색 복원
    _update_row_bg(manager, sheet_id, sheet_name, row_number, 'normal')

    logger.info(f'[SLACK/재개] 완료: {code} by slack:{by_display_name}')
    return {'ok': True}


# ─────────────────────────────────────────────────────────────
# 내용 수정
# ─────────────────────────────────────────────────────────────
EDITABLE_FIELDS = [
    '공사 내용', '도급 구분', '시공자',
    '총액 1', '부가세',
    '공사 시작', '공사 종료',
]

# 시트 컬럼 매핑 (2026-07 시프트 반영). 관리 사이트 update_project 와 동일 매핑.
# 정확한 컬럼 letter는 get_column_mapping()에서 조회하므로 여기서 하드코드하지 않음.


def perform_edit(
    code: str,
    updates: Dict[str, Any],
    reason: str,
    by_display_name: str,
) -> Dict[str, Any]:
    """관리 사이트 PUT /api/projects/<code> 와 동일 효과.

    updates: 편집할 필드 dict (편집 가능 필드만). 예: {'총액 1': 15000000, '공사 종료': '2026-08-10'}.
    reason: 필수 수정 사유 (감사 로그와 스레드 답글에 포함).
    """
    if not reason or not reason.strip():
        return {'ok': False, 'reason': 'reason_required'}

    project = _load_project(code)
    if not project:
        return {'ok': False, 'reason': 'not_found'}

    manager, sheet_id, sheet_name = _get_sheet_context()
    row_number = _find_row_number(manager, sheet_id, sheet_name, code)
    if not row_number:
        return {'ok': False, 'reason': 'row_not_found'}

    # 컬럼 매핑 조회 후 변경 field 만 batch update로 씀
    column_mapping = manager.get_column_mapping()  # {'A': '프로젝트 코드', ...}
    field_to_col = {name: letter for letter, name in column_mapping.items()}

    batch = []
    field_changes = []
    for field, new_value in updates.items():
        if field not in EDITABLE_FIELDS:
            continue
        col = field_to_col.get(field)
        if not col:
            logger.warning(f'[SLACK/편집] 컬럼 매핑 없음: {field}')
            continue
        old_value = project.get(field, '')
        # bool은 시트에 대문자 TRUE/FALSE로 (관리 사이트 관례)
        if isinstance(new_value, bool):
            sheet_value = 'TRUE' if new_value else 'FALSE'
        elif new_value is None:
            sheet_value = ''
        else:
            sheet_value = str(new_value)
        batch.append({
            'range': f'{sheet_name}!{col}{row_number}',
            'values': [[sheet_value]],
        })
        field_changes.append({
            'field_name': field,
            'old_value': old_value,
            'new_value': new_value,
        })

    if not batch:
        return {'ok': False, 'reason': 'no_changes'}

    try:
        manager.batch_update_cells(sheet_id, batch)
    except Exception as exc:
        logger.error(f'[SLACK/편집] 시트 write 실패 ({code}): {exc}', exc_info=True)
        return {'ok': False, 'reason': 'sheet_write_failed'}

    # 캐시 부분 갱신
    try:
        from dashboard.services.project_service import (
            update_project_in_cache,
            invalidate_project_cache,
        )
        cache_updated = update_project_in_cache(code, {
            fc['field_name']: fc['new_value'] for fc in field_changes
        })
        if not cache_updated:
            invalidate_project_cache(code)
    except Exception as exc:
        logger.warning(f'[SLACK/편집] 캐시 갱신 실패 ({code}): {exc}')

    # 감사 로그 — 필드별 개별 기록 + 수정 사유 별도 로그.
    for fc in field_changes:
        _audit_log(
            user_email=f'slack:{by_display_name}',
            action='UPDATE_PROJECT',
            details=f'프로젝트 필드 수정: {code}.{fc["field_name"]} (사유: {reason.strip()[:200]})',
            project_code=code,
            field_name=fc['field_name'],
            old_value=str(fc['old_value']) if fc['old_value'] not in (None, '') else '-',
            new_value=str(fc['new_value']) if fc['new_value'] not in (None, '') else '-',
            ip_address='slack-bot',
        )

    # 스레드 답글 + 원본 카드 재렌더링 (기존 로직 재사용)
    try:
        from dashboard.services.project_slack_notifier import notify_project_field_changes
        # 편집 후 최신 데이터 스냅샷
        latest = dict(project)
        for fc in field_changes:
            latest[fc['field_name']] = fc['new_value']
        # 수정 사유를 field_changes 앞에 삽입해 답글에 표시
        annotated_changes = [
            {'field_name': '수정 사유', 'old_value': '', 'new_value': reason.strip()},
        ] + field_changes
        notify_project_field_changes(code, annotated_changes, latest_data=latest)
    except Exception as exc:
        logger.warning(f'[SLACK/편집] 알림 발송 실패 ({code}): {exc}')

    logger.info(
        f'[SLACK/편집] 완료: {code} by slack:{by_display_name} '
        f'({len(field_changes)}개 필드, 사유="{reason.strip()[:40]}")'
    )
    return {'ok': True, 'field_changes': field_changes}
