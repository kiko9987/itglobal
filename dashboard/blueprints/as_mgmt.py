"""A/S 관리 블루프린트 — PM 대시보드 A/S 모드 (목록/요청/접수/완료).

슬랙 A/S 봇과 동일 시트(as_service) + 동일 카드(as_refresh_card)를 공유한다.
PM에서 요청/접수/완료해도 슬랙 카드·담당자 DM이 자동 동기화된다 (슬랙=서브 유지).
- 조회: login_required (전체)
- 요청/접수/완료: editor_required (편집자+)
"""
from datetime import datetime

from flask import Blueprint, request, session

from ..auth import login_required, editor_required
from ..api.responses import APIResponse, APIErrorCode
from ..services import as_service
from ..utils.logging_config import get_logger

logger = get_logger(__name__)

as_bp = Blueprint('as_mgmt', __name__, url_prefix='/as')


def _current_initial() -> str:
    """현재 로그인 사용자 이메일 → 이니셜 (요청자/접수자 기록용). 슬랙과 동일 매핑."""
    email = (session.get('user', {}) or {}).get('email', '') or ''
    name = ''
    try:
        from ..utils.user_database import get_user_database
        u = get_user_database().get_user_by_email(email) if email else None
        name = (u or {}).get('name', '') or ''
    except Exception:
        pass
    try:
        from .slack_helpers import _to_initial
        return _to_initial(name) or name or '-'
    except Exception:
        return name or '-'


def _find_open_as(project_code: str):
    """해당 프로젝트의 진행 중(요청됨/접수완료) A/S 반환 (없으면 None). 중복 요청 방지용."""
    code = (project_code or '').strip()
    if not code:
        return None
    open_states = {as_service.STATUS_REQUESTED, as_service.STATUS_ACCEPTED}
    try:
        for x in as_service.list_as():  # AS 번호 내림차순 → 첫 매치가 최신 진행건
            if (str(x.get('프로젝트 코드', '')).strip() == code
                    and str(x.get('진행 상태', '')).strip() in open_states):
                return x
    except Exception as exc:
        logger.warning(f'[AS] 진행 중 A/S 확인 실패 ({code}): {exc}')
    return None


def _sync_slack(as_no: str, send_dm: bool = False, dm_override=None, override=None) -> None:
    """슬랙 A/S 카드/DM 동기화 (실패해도 시트 반영은 유지).

    override: 방금 write-behind 큐에 넣은 값(진행 상태·방문 예정자 등)을 카드에 즉시 반영.
      큐 flush 전에 as_refresh_card 가 시트를 직접 읽으면 이전 상태가 나오므로 우회.
    """
    try:
        from .slack_bot import as_refresh_card
        as_refresh_card(as_no, send_dm=send_dm, dm_override=dm_override, override=override)
    except Exception as exc:
        logger.warning(f'[AS] 슬랙 동기화 실패 ({as_no}): {exc}')


@as_bp.route('/api/list', methods=['GET'])
@login_required
def api_list_as():
    """A/S 목록. 프로젝트 조인으로 담당자/유입 구분/사업자명 채움. ?status=<진행상태> 필터."""
    try:
        status = (request.args.get('status', '') or '').strip() or None
        items = as_service.list_as(status_filter=status)
        # 프로젝트 조인 — 왼쪽 컬럼(담당자/유입/사업자명)은 연결 프로젝트에서
        pidx = {}
        try:
            from ..services.project_service import get_project_records
            for r in (get_project_records() or []):
                c = str(r.get('프로젝트 코드', '') or '').strip()
                if c:
                    pidx[c] = r
        except Exception as exc:
            logger.warning(f'[AS] 프로젝트 조인 로드 실패 (계속): {exc}')
        for it in items:
            p = pidx.get(str(it.get('프로젝트 코드', '') or '').strip()) or {}
            it['담당자'] = p.get('담당자', '') or ''
            it['유입 구분'] = p.get('유입 구분', '') or ''
            it['사업자명'] = p.get('사업자명', '') or ''
        return APIResponse.success(data={'items': items, 'count': len(items)})
    except Exception as exc:
        logger.error(f'[AS] 목록 조회 실패: {exc}', exc_info=True)
        return APIResponse.error(
            message='A/S 목록 조회 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR, status_code=500,
        )


@as_bp.route('/api/open-check/<project_code>', methods=['GET'])
@login_required
def api_open_check(project_code):
    """프로젝트의 진행 중(요청됨/접수완료) A/S 존재 여부 — 'A/S 요청' 버튼 단계 차단용."""
    try:
        existing = _find_open_as(project_code)
        if existing:
            return APIResponse.success(data={
                'open': True, 'as_no': existing.get('No'), 'status': existing.get('진행 상태'),
            })
        return APIResponse.success(data={'open': False})
    except Exception as exc:
        logger.error(f'[AS] open-check 실패 ({project_code}): {exc}', exc_info=True)
        return APIResponse.error(
            message='진행 중 A/S 확인 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR, status_code=500,
        )


@as_bp.route('/api/request', methods=['POST'])
@editor_required
def api_request_as():
    """A/S 요청 생성.
    body: {project_code?, request_content, manual?: {address, work_content, manager_name, manager_email}}
    project_code 있으면 프로젝트 연결, 없으면 manual(코드 이전 공사) 모드.
    """
    try:
        data = request.get_json(silent=True) or {}
        request_content = (data.get('request_content') or '').strip()
        project_code = (data.get('project_code') or '').strip()
        requester = _current_initial()

        if project_code:
            # 중복 방지(하드 블록) — 진행 중(요청됨/접수완료) A/S 가 있으면 새 요청 거부.
            #   미완료 A/S 가 있는데 또 요청하는 것은 완료 후 처리해야 하므로, 버튼 단계
            #   (open-check)에서 먼저 막고 서버에서도 최종 차단한다.
            existing = _find_open_as(project_code)
            if existing:
                return APIResponse.error(
                    message=f"이미 진행 중인 A/S 가 있습니다 ({existing.get('No')} · {existing.get('진행 상태')}). 완료 후 다시 요청해주세요.",
                    error_code='AS_ALREADY_OPEN', status_code=409,
                    details={'as_no': existing.get('No'), 'status': existing.get('진행 상태')},
                )
            det = as_service.get_project_details(project_code) or {}
            as_no, _row = as_service.create_as_row(
                project_code=project_code,
                address=det.get('address', ''),
                work_content=det.get('work_content', ''),
                work_end=det.get('work_end', ''),
                request_content=request_content, requester=requester,
            )
            dm_override = None
        else:
            manual = data.get('manual') or {}
            address = (manual.get('address') or '').strip()
            if not address:
                return APIResponse.error(
                    message='현장 주소는 필수입니다 (코드 없는 A/S)',
                    error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
                )
            as_no, _row = as_service.create_as_row(
                project_code='', address=address,
                work_content=(manual.get('work_content') or '').strip(),
                work_end='', request_content=request_content, requester=requester,
            )
            dm_override = {
                'name': (manual.get('manager_name') or '').strip(),
                'email': (manual.get('manager_email') or '').strip(),
            }

        _sync_slack(as_no, send_dm=True, dm_override=dm_override)
        logger.info(f'[AS] PM 요청 생성: {as_no} (code={project_code or "수동"}, by={requester})')
        return APIResponse.success(data={'as_no': as_no})
    except Exception as exc:
        logger.error(f'[AS] 요청 생성 실패: {exc}', exc_info=True)
        return APIResponse.error(
            message='A/S 요청 생성 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR, status_code=500,
        )


@as_bp.route('/api/accept/<as_no>', methods=['POST'])
@editor_required
def api_accept_as(as_no):
    """A/S 접수. body: {visitor_type(서비스 기사/내부/외주), visitor_name, visit_date_start, visit_date_end?}"""
    try:
        data = request.get_json(silent=True) or {}
        vtype = (data.get('visitor_type') or '').strip()
        vname = (data.get('visitor_name') or '').strip()
        start = (data.get('visit_date_start') or '').strip()
        end = (data.get('visit_date_end') or '').strip()

        if not vtype:
            return APIResponse.error(
                message='방문자 유형을 선택해주세요',
                error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
            )
        if vtype in ('내부', '외주') and not vname:
            return APIResponse.error(
                message='내부/외주는 방문자 이름이 필수입니다',
                error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
            )
        if not start:
            return APIResponse.error(
                message='방문 예정일은 필수입니다',
                error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
            )

        visitor = '서비스 기사' if vtype == '서비스 기사' else vname
        try:
            from .slack_helpers import _format_visit_date_range
            visit_date = _format_visit_date_range(start, end) if end else start
        except Exception:
            visit_date = f'{start}~{end}' if end and end != start else start

        initial = _current_initial()
        accept_dt = datetime.now().strftime('%Y.%m.%d. %H:%M')
        as_service.update_as_row(as_no, {
            as_service.COL_ACCEPTER: initial,
            as_service.COL_ACCEPT_DATE: accept_dt,
            as_service.COL_VISITOR: visitor,
            as_service.COL_VISIT_DATE: visit_date,
            as_service.COL_STATUS: as_service.STATUS_ACCEPTED,
        })
        # 접수 메모(선택) — 새 컬럼 없이 M열(메모/이력)에 누적
        memo = (data.get('memo') or '').strip()
        new_log = as_service.append_as_log(as_no, memo, initial) if memo else None
        # 슬랙 카드에 방금 쓴 값 즉시 반영 (write-behind flush 전 stale read 우회)
        overr = {
            '접수자': initial, '접수 일자': accept_dt,
            '방문 예정자': visitor, '방문 예정일': visit_date,
            '진행 상태': as_service.STATUS_ACCEPTED,
        }
        if new_log is not None:
            overr['조치 내용'] = new_log
        _sync_slack(as_no, override=overr)
        logger.info(f'[AS] PM 접수: {as_no} (visitor={visitor}, date={visit_date}, memo={"Y" if memo else "N"})')
        return APIResponse.success(data={'as_no': as_no})
    except Exception as exc:
        logger.error(f'[AS] 접수 실패 ({as_no}): {exc}', exc_info=True)
        return APIResponse.error(
            message='A/S 접수 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR, status_code=500,
        )


@as_bp.route('/api/complete/<as_no>', methods=['POST'])
@editor_required
def api_complete_as(as_no):
    """조치 완료. body: {resolution}"""
    try:
        data = request.get_json(silent=True) or {}
        resolution = (data.get('resolution') or '').strip()
        if not resolution:
            return APIResponse.error(
                message='조치 내용은 필수입니다',
                error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
            )
        as_service.update_as_row(as_no, {
            as_service.COL_STATUS: as_service.STATUS_COMPLETED,
        })
        # 조치 내용도 M열(메모/이력)에 누적 (덮어쓰지 않음)
        new_log = as_service.append_as_log(as_no, resolution, _current_initial())
        # 슬랙 카드에 방금 쓴 값 즉시 반영 (write-behind flush 전 stale read 우회)
        overr = {'진행 상태': as_service.STATUS_COMPLETED}
        if new_log is not None:
            overr['조치 내용'] = new_log
        _sync_slack(as_no, override=overr)
        logger.info(f'[AS] PM 조치완료: {as_no}')
        return APIResponse.success(data={'as_no': as_no})
    except Exception as exc:
        logger.error(f'[AS] 조치완료 실패 ({as_no}): {exc}', exc_info=True)
        return APIResponse.error(
            message='조치 완료 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR, status_code=500,
        )
