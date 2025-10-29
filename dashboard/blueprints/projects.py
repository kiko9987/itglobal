import logging
import os
import json
import re
import time
import threading
from datetime import datetime
from collections import Counter, defaultdict

import pandas as pd
from flask import Blueprint, render_template, redirect, url_for, request, session, jsonify, current_app

from ..auth import login_required, get_user_role, admin_required
from ..services.project_service import (
    can_user_edit_project,
    check_overdue_status,
    get_project_records,
    invalidate_project_cache,
    get_sheets_manager,
    load_data,
    get_project_config,
    _build_company_prefix_map,
    _build_owner_suffix_map,
    _next_running_number,
    _safe_next_running_number_with_retry,
    _auto_project_code,
)
from ..services.calendar_service import (
    create_project_calendar_event,
    update_project_calendar_event,
    delete_project_calendar_event,
)
from ..services.calendar_sync_scheduler import CalendarSyncScheduler
from ..utils.error_handler import handle_error, ErrorCategory
from ..utils.api_response import APIResponse, APIErrorCode, api_response
from ..utils.smart_cache_manager import smart_invalidate, smart_get, CacheStrategy
from ..utils.logging_config import get_logger
from ..utils.request_middleware import track_business_operation, log_external_api_call

# Marshmallow 검증 스키마 임포트
from ..schemas import ProjectCreateSchema, ProjectUpdateSchema, CellMemoSchema
from ..schemas.project_schemas import (
    validate_request_data,
    format_validation_errors,
    ProjectAutoCreateSchema
)

# 프로젝트 전역 상수 임포트
from ..constants import PAYMENT_FIELD_TO_COLUMN, MEMOABLE_FIELDS, ERROR_MESSAGES

logger = get_logger(__name__)


def convert_excel_serial_to_date(value):
    """
    Excel/Google Sheets 시리얼 날짜를 YYYY-MM-DD 형식으로 변환

    Args:
        value: Excel 시리얼 숫자 또는 날짜 문자열

    Returns:
        str: YYYY-MM-DD 형식의 날짜 문자열, 실패 시 원본 값 반환

    Examples:
        convert_excel_serial_to_date(45943) -> '2025-10-13'
        convert_excel_serial_to_date('2025-10-13') -> '2025-10-13'
        convert_excel_serial_to_date('invalid') -> 'invalid'
    """
    if not value or value == '-':
        return value

    # 이미 YYYY-MM-DD 형식인 경우
    if isinstance(value, str) and re.match(r'^\d{4}-\d{2}-\d{2}$', value):
        return value

    try:
        # 숫자(시리얼 날짜)인 경우 변환
        serial_number = float(str(value).strip())

        # Excel 시리얼 날짜는 1900-01-01부터 시작 (실제로는 1899-12-30)
        # 1900년 2월 29일 버그 보정
        if serial_number < 60:
            base_date = datetime(1899, 12, 31)
            converted_date = base_date + pd.Timedelta(days=serial_number - 1)
        else:
            base_date = datetime(1899, 12, 30)
            converted_date = base_date + pd.Timedelta(days=serial_number)

        return converted_date.strftime('%Y-%m-%d')
    except (ValueError, TypeError, AttributeError):
        # 변환 실패 시 원본 값 반환
        return str(value) if value else value


def safe_parse_currency(value):
    """
    안전한 통화 파싱 함수
    다양한 통화 기호, 공백, 특수 문자를 처리하여 숫자만 추출
    부호(+/-)와 소수점을 보존하여 Google Sheets 값과 일치

    Args:
        value: 파싱할 값 (str, int, float, None)

    Returns:
        float: 파싱된 숫자 (실패 시 0)

    Examples:
        safe_parse_currency("1,234원") -> 1234.0
        safe_parse_currency("$1,000.50") -> 1000.5
        safe_parse_currency("-500") -> -500.0
        safe_parse_currency("  ¥ 500  ") -> 500.0
        safe_parse_currency("1,000원 + VAT") -> 1000.0
        safe_parse_currency("") -> 0
        safe_parse_currency(None) -> 0
    """
    if value is None:
        return 0

    # 이미 숫자인 경우
    if isinstance(value, (int, float)):
        return float(value)  # 부호 보존

    # 문자열로 변환
    str_value = str(value).strip()

    if not str_value:
        return 0

    try:
        # 정규식으로 숫자만 추출 (쉼표, 소수점, 부호 포함)
        # 통화 기호: ₩, 원, $, €, ¥, £ 등 제거
        # 공백, 특수문자 제거
        # 첫 번째로 발견되는 연속된 숫자 그룹 추출 (부호 포함)

        # 1단계: 통화 기호와 불필요한 문자 제거
        currency_symbols = r'[₩원$€¥£＄￦]'
        cleaned = re.sub(currency_symbols, '', str_value)

        # 2단계: 부호, 숫자, 쉼표, 소수점만 추출
        number_pattern = r'-?[\d,.]+'
        matches = re.findall(number_pattern, cleaned)

        if not matches:
            return 0

        # 첫 번째 매치된 숫자 사용
        number_str = matches[0]

        # 3단계: 쉼표 제거하고 소수점 처리
        number_str = number_str.replace(',', '')

        # 빈 문자열 체크 (마이너스만 있는 경우)
        if not number_str or number_str == '-':
            return 0

        # 4단계: 실수로 변환 (부호와 소수점 보존)
        result = float(number_str)
        return result

    except (ValueError, IndexError) as e:
        logger.warning(f"통화 파싱 실패: '{value}' -> 0 (오류: {e})")
        return 0


def sanitize_project_for_json(project):
    """
    프로젝트 데이터를 JSON 직렬화 가능하도록 변환

    pandas의 NaT, Timestamp, numpy 타입을 Python 네이티브 타입으로 변환

    Args:
        project: 프로젝트 dict 데이터

    Returns:
        dict: JSON 직렬화 가능한 프로젝트 데이터
    """
    if not project:
        return project

    sanitized = {}
    for key, value in project.items():
        # pandas NaT 처리
        if pd.isna(value):
            sanitized[key] = None
        # pandas Timestamp 처리
        elif isinstance(value, pd.Timestamp):
            try:
                sanitized[key] = value.strftime('%Y-%m-%d') if not pd.isna(value) else None
            except Exception as e:
                sanitized[key] = None
        # numpy 타입 처리
        elif hasattr(value, 'item'):  # numpy scalar
            sanitized[key] = value.item()
        else:
            sanitized[key] = value

    return sanitized


projects_bp = Blueprint('projects', __name__)


@projects_bp.route('/projects')
@login_required
def project_list():
    """
    전문가 리뷰: "초기 페이지에서는 최소 메타 정보만 제공하고 상세 데이터는 API로만 받도록"
    템플릿 렌더링과 API 데이터 로딩을 완전히 분리
    """
    from ..utils.frontend_helpers import asset_manager

    user_role = get_user_role()
    user_email = session.get('user', {}).get('email', '')
    user_name = user_email.split('@')[0] if '@' in user_email else user_email
    user_display_name = session.get('user', {}).get('name', '')  # 실제 사용자 이름

    # 최소한의 메타데이터만 제공 (통계는 별도 API로 로드)
    page_metadata = {
        'user_role': user_role,
        'user_email': user_email,
        'user_name': user_name,
        'user_display_name': user_display_name,
        'user_can_create': user_role.lower() in ['admin', 'editor'],
        'csrf_token': session.get('csrf_token', ''),
        'app_version': '1.0.0',
        'api_endpoints': {
            'projects_list': '/api/projects/list',
            'projects_statistics': '/api/projects/statistics'
        }
    }

    # JS 번들 URL 가져오기
    js_url = asset_manager.get_js_bundle('project-list')

    return render_template(
        'modern_project_list.html',
        page_metadata=page_metadata,
        # 통계는 제거 - API로만 로드
        user_role=user_role,
        user_email=user_email,
        user_name=user_name,
        user_display_name=user_display_name,
        user_can_create=user_role.lower() in ['admin', 'editor'],
        js_url=js_url
    )


# ===== API 엔드포인트들 =====

@projects_bp.route('/api/projects/list')
@login_required
@track_business_operation("api_projects_list")
def get_projects_list():
    """프로젝트 목록 API"""
    try:
        # refresh 파라미터가 있으면 강제로 구글시트에서 최신 데이터 로드
        refresh = request.args.get('refresh', 'false').lower() == 'true'

        logger.info(f'[API][PID:{os.getpid()}] 프로젝트 목록 요청 (refresh={refresh})')

        if refresh:
            logger.info(f'[API][PID:{os.getpid()}] 강제 새로고침 - 구글시트 로드')
            df = load_data(force_refresh=True)
            # set_current_data() 제거 - load_data()가 캐시에 저장함
        else:
            df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
            logger.info(f'[API][PID:{os.getpid()}] 캐시에서 데이터 조회 - {"있음" if df is not None and not df.empty else "없음"}')
            if df is None or (df is not None and len(df) == 0):
                logger.info(f'[API][PID:{os.getpid()}] 캐시 없음 - 구글시트에서 로드')
                df = load_data(force_refresh=False)

        if df is None or (df is not None and len(df) == 0):
            return jsonify({
                'success': False,
                'error': '데이터를 불러올 수 없습니다.',
                'data': None,
                'meta': {
                    'timestamp': datetime.now().isoformat()
                }
            }), 500

        # DataFrame을 dict 리스트로 변환 (NaN 값 처리)
        df = df.fillna('')  # NaN 값을 빈 문자열로 변환

        # 날짜 컬럼들을 올바르게 처리
        date_columns = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
        for col in date_columns:
            if col in df.columns:
                logger.info(f"날짜 컬럼 처리 시작: {col}")

                # 원본 데이터 샘플 확인 (최근 5개)
                recent_samples = df.tail(5)[col].tolist()
                logger.info(f"{col} 원본 데이터 샘플: {recent_samples}")

                # 다양한 날짜 형식 지원 (YYYY-MM-DD, YYYY/MM/DD 등)
                df[col] = pd.to_datetime(df[col], errors='coerce')

                # 변환 후 샘플 확인
                converted_samples = df.tail(5)[col].tolist()
                logger.info(f"{col} 변환 후 샘플: {converted_samples}")

                # 유효한 날짜는 YYYY-MM-DD 형식으로, 무효한 날짜는 빈 문자열로
                df[col] = df[col].apply(lambda x:
                    x.strftime('%Y-%m-%d') if pd.notna(x) and x is not pd.NaT else ''
                )

                # 최종 결과 샘플
                final_samples = df.tail(5)[col].tolist()
                logger.info(f"{col} 최종 결과 샘플: {final_samples}")

        projects = df.to_dict('records')

        # Google Sheets에서 이미 계산된 순익, 마진율 값을 그대로 사용
        # 재계산하지 않음 - 시트의 수식이 정확한 단일 소스(Single Source of Truth)

        # 참고: viewer 권한은 "읽기 전용" (신입 교육용)
        # - 전체 프로젝트 조회 가능 (검색, 학습 목적)
        # - 수정/삭제는 불가 (각 API 엔드포인트에서 권한 체크)

        response = jsonify({
            'success': True,
            'data': projects,
            'meta': {
                'total': len(projects),
                'timestamp': datetime.now().isoformat()
            }
        })

        # 캐시 헤더 설정 (성능 최적화)
        if not refresh:
            # 일반 조회: 30초 동안 브라우저 캐시 허용
            response.headers['Cache-Control'] = 'private, max-age=30'
        else:
            # 강제 새로고침: 캐시 사용 안함
            response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'

        return response

    except Exception as e:
        logger.error(f"프로젝트 목록 API 오류: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e),
            'data': None,
            'meta': {
                'timestamp': datetime.now().isoformat()
            }
        }), 500


@projects_bp.route('/api/next-project-code')
def get_next_project_code():
    """다음 프로젝트 코드 생성 API"""
    try:
        region_code = request.args.get('region', 'IT')

        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500

        manager = get_sheets_manager()
        project_code = manager.get_next_project_code(sheet_id, region_code)

        return jsonify({'project_code': project_code})

    except Exception as e:
        logger.error(f"프로젝트 코드 생성 API 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500


# ============================================================================
# 헬퍼 함수: update_project() 리팩토링
# ============================================================================

def _validate_update_request(data, project_code):
    """요청 데이터 검증 (Marshmallow 스키마)

    Returns:
        JsonResponse or None: 검증 실패 시 에러 응답, 성공 시 None
    """
    # 디버그 로깅
    logger.info(f"[PUT] 받은 데이터 타입: {type(data)}")
    logger.info(f"[PUT] 받은 데이터: {data}")

    if not project_code:
        return jsonify({'success': False, 'error': '프로젝트 코드가 필요합니다.'}), 400

    if not data:
        return jsonify({'success': False, 'error': '업데이트할 데이터가 없습니다.'}), 400

    if not isinstance(data, dict):
        logger.error(f"[PUT] 데이터가 dict가 아닙니다: {type(data)}")
        return jsonify({'success': False, 'error': '잘못된 데이터 형식입니다.'}), 400

    # Marshmallow 스키마로 데이터 검증 (업데이트는 모든 필드 선택)
    # 한글 필드명을 영문 스키마 필드로 매핑
    validation_data = {
        'project_code': data.get('프로젝트 코드', project_code),
        'manager': data.get('담당자'),
        'company': data.get('사업자'),
        'address': data.get('현장 주소'),
        'work_content': data.get('공사 내용'),
        'start_date': data.get('공사 시작'),
        'end_date': data.get('공사 종료'),
        'down_payment': data.get('계약금'),
        'mid_payment': data.get('중도금'),
        'final_payment': data.get('잔금'),
        'total_amount': data.get('총액 1'),
    }

    # None 값 제거 (업데이트되지 않는 필드)
    validation_data = {k: v for k, v in validation_data.items() if v is not None}

    validated_data, errors = validate_request_data(ProjectUpdateSchema, validation_data)

    if errors:
        error_messages = format_validation_errors(errors)
        logger.warning(f"[VALIDATION] 프로젝트 업데이트 검증 실패: {error_messages}")
        return jsonify({
            'success': False,
            'error': '입력 데이터 검증 실패',
            'validation_errors': error_messages
        }), 400

    # 검증 성공
    return None


def _load_project_row(manager, sheet_id, sheet_name, project_code):
    """프로젝트 행 조회 및 현재 값 로드

    Returns:
        tuple: (row_number, current_values, error_response)
        - success: (int, list, None)
        - failure: (None, None, JsonResponse)
    """
    # 프로젝트가 있는 행 찾기
    row_number = manager.find_row_by_project_code(sheet_id, project_code, f'{sheet_name}!A:A')

    if not row_number:
        logger.error(f"[PUT] 프로젝트 코드 '{project_code}'를 시트에서 찾을 수 없습니다. Sheet: {sheet_name}, ID: {sheet_id}")
        return None, None, (jsonify({'success': False, 'error': f'프로젝트 {project_code}를 찾을 수 없습니다.'}), 404)

    logger.info(f"[PUT] 프로젝트 업데이트 시작: {project_code}, 행번호: {row_number}")

    # 현재 행의 전체 데이터 조회 (수식 포함하여 읽기)
    # get_row_values 헬퍼 메서드 사용 (재시도/로깅/모니터링 포함)
    current_values = manager.get_row_values(
        sheet_id=sheet_id,
        sheet_name=sheet_name,
        row_number=row_number,
        start_col='A',
        end_col='AN',  # AN 컬럼까지 읽기 (_version 포함)
        value_render_option='FORMULA'  # 수식을 그대로 가져옴 (계산 필드 보존)
    )

    # 현재 값을 리스트로 확장 (40개 컬럼, AN까지)
    while len(current_values) < 40:
        current_values.append('')

    return row_number, current_values, None


def _check_optimistic_lock_update(manager, sheet_id, sheet_name, project_code, row_number, current_values, expected_version):
    """Optimistic Lock 버전 검증

    Returns:
        tuple: (new_version, error_response)
        - success: (int, None)
        - failure: (None, JsonResponse with 409)
    """
    current_version = current_values[39]  # AN 컬럼 (index 39)

    # 버전 값 정규화 (빈 문자열/None → '0')
    if not current_version or current_version == '':
        current_version = '0'
    if not expected_version or expected_version == '':
        expected_version = '0'

    # 버전 불일치 → 409 Conflict
    if str(current_version) != str(expected_version):
        logger.warning(f"[OPTIMISTIC_LOCK] 버전 충돌 감지: {project_code}, expected={expected_version}, current={current_version}")

        # 현재 최신 데이터를 반환하여 클라이언트가 병합할 수 있도록 함
        # 최신 데이터 다시 읽기 (FORMATTED_VALUE로 사용자가 보는 형식으로)
        fresh_values = manager.get_row_values(
            sheet_id=sheet_id,
            sheet_name=sheet_name,
            row_number=row_number,
            start_col='A',
            end_col='AN',
            value_render_option='FORMATTED_VALUE'
        )

        # 컬럼 매핑 임시 생성
        temp_column_mapping = manager.get_column_mapping()
        temp_field_to_index = {}
        for col_letter, field_name in temp_column_mapping.items():
            if len(col_letter) == 1:
                column_index = ord(col_letter) - ord('A')
            else:
                column_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
            temp_field_to_index[field_name] = column_index

        # 최신 데이터를 딕셔너리로 변환
        current_data = {}
        for field_name, index in temp_field_to_index.items():
            if index < len(fresh_values):
                value = fresh_values[index]
                current_data[field_name] = '' if value is None or value == '' else str(value)

        return None, (jsonify({
            'success': False,
            'error': 'conflict',
            'message': '다른 사용자가 먼저 수정했습니다. 페이지를 새로고침한 후 다시 시도하세요.',
            'current_data': current_data,
            'current_version': str(current_version)
        }), 409)

    # 버전 일치 → 버전 증가
    new_version = int(current_version) + 1
    logger.info(f"[OPTIMISTIC_LOCK] 버전 업데이트: {project_code}, {current_version} → {new_version}")

    return new_version, None


def _update_project_code_if_needed(data, current_values, row_number, manager, project_code, field_to_index):
    """사업자/담당자 변경 시 프로젝트 코드 자동 재산출

    Returns:
        tuple: (final_project_code, code_change_dict or None)
    """
    from dashboard.utils.project_code_generator import generate_project_code, should_update_project_code

    company = data.get('사업자') or current_values[field_to_index.get('사업자', 1)]
    manager_field = data.get('담당자') or current_values[field_to_index.get('담당자', 2)]

    # 사업자 또는 담당자가 변경되었는지 확인하고 프로젝트 코드 재생성
    if should_update_project_code(project_code, row_number, company, manager_field):
        new_project_code = generate_project_code(row_number, company, manager_field)
        if new_project_code and new_project_code != project_code:
            logger.info(f"[PUT] 프로젝트 코드 자동 업데이트: {project_code} → {new_project_code}")
            # 프로젝트 코드 업데이트
            current_values[0] = new_project_code
            code_change = {
                'field_name': '프로젝트 코드',
                'old_value': project_code,
                'new_value': new_project_code
            }
            # 클라이언트에도 반환할 수 있도록 data 업데이트
            data['프로젝트 코드'] = new_project_code

            return new_project_code, code_change

    return project_code, None


def _apply_field_updates(data, current_values, field_to_index, project_code):
    """필드 업데이트 적용 및 변경사항 추적

    Returns:
        list: field_changes (list of dict)
    """
    field_changes = []

    # 계산 필드 목록 (프론트엔드에서 계산, 백엔드는 그대로 저장)
    calculated_fields = ['총액 2', '총액2', '미수금', '마진율', '순익']

    # 업데이트할 필드만 변경
    for field_name, new_value in data.items():
        if field_name in ['projectCode', '프로젝트 코드', '_version']:
            continue  # 프로젝트 코드와 버전은 이미 처리됨

        # 계산 필드는 건너뛰기 (구글 시트 수식 보존)
        if field_name in calculated_fields:
            logger.info(f"[PUT] 계산 필드 건너뛰기: {field_name} (구글 시트 수식 사용)")
            continue

        if field_name in field_to_index:
            column_index = field_to_index[field_name]
            old_value = current_values[column_index] if column_index < len(current_values) else ''

            # 값이 실제로 변경된 경우에만 기록
            if str(old_value) != str(new_value):
                field_changes.append({
                    'field_name': field_name,
                    'old_value': str(old_value),
                    'new_value': str(new_value)
                })
                current_values[column_index] = str(new_value)
        else:
            logger.warning(f"[PUT] 알 수 없는 필드명: {field_name}")

    return field_changes


def _process_payment_field_comments(manager, sheet_id, sheet_name, row_number, field_changes):
    """금액 필드 변경 시 자동 댓글 생성/삭제 (Apps Script onEdit 로직 재현)"""
    payment_fields_to_check = MEMOABLE_FIELDS
    field_to_column_map = PAYMENT_FIELD_TO_COLUMN

    for change in field_changes:
        field_name = change['field_name']
        new_value = change['new_value']

        if field_name in payment_fields_to_check:
            try:
                # 금액 파싱
                amount = safe_parse_currency(new_value)
                column = field_to_column_map[field_name]
                cell_address = f"{column}{row_number}"

                logger.info(f"[AUTO_COMMENT] {field_name} 변경 감지: {amount}원 → 셀 {cell_address}")

                if amount > 0:
                    # 자동 템플릿 생성 제거: 사용자가 직접 메모를 작성하도록 유도
                    logger.debug(f"[AUTO_COMMENT] {cell_address} 금액 입력됨: {amount}원 (메모는 사용자가 직접 작성)")

                else:
                    # 금액이 0 이하면 댓글 삭제 (Apps Script와 동일)
                    manager.update_cell_note(sheet_id, sheet_name, cell_address, None)
                    logger.info(f"[AUTO_COMMENT] {cell_address} 댓글 삭제 완료 (금액 0 이하)")

            except Exception as comment_error:
                logger.warning(f"[AUTO_COMMENT] {field_name} 자동 댓글 처리 실패: {comment_error}")


def _record_update_audit_logs(field_changes, project_code):
    """감사 로그 배치 기록 (변경된 필드만)"""
    if not field_changes:
        return

    try:
        from dashboard.utils.user_database import get_audit_repository
        audit_repo = get_audit_repository()
        user_email = session.get('user', {}).get('email', 'unknown')

        # 날짜 필드 목록
        date_fields = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']

        # 모든 감사 로그를 리스트에 수집 (배치 저장용)
        audit_actions = []
        for change in field_changes:
            # 날짜 필드인 경우 Excel 시리얼 번호를 YYYY-MM-DD 형식으로 변환
            old_value_display = change['old_value']
            new_value_display = change['new_value']

            if change['field_name'] in date_fields:
                old_value_display = convert_excel_serial_to_date(old_value_display)
                new_value_display = convert_excel_serial_to_date(new_value_display)

            audit_actions.append({
                'user_email': user_email,
                'action': 'UPDATE_FIELD',
                'details': f"프로젝트 {project_code}의 {change['field_name']} 필드를 '{old_value_display}'에서 '{new_value_display}'로 변경",
                'project_code': project_code,
                'field_name': change['field_name'],
                'old_value': old_value_display,
                'new_value': new_value_display,
                'ip_address': request.remote_addr
            })

        # 배치로 한 번에 저장 (1회 트랜잭션)
        success, count = audit_repo.log_actions_batch(audit_actions)
        if success:
            logger.info(f"[PUT] 감사 로그 배치 기록 완료: {count}개 변경사항 (1회 트랜잭션)")
        else:
            logger.warning(f"[PUT] 감사 로그 배치 기록 실패")

    except Exception as log_error:
        logger.warning(f"[PUT] 감사 로그 기록 실패: {log_error}")


def _fetch_and_calculate_updated_project(manager, sheet_id, sheet_name, row_number, field_to_index):
    """저장된 행을 다시 조회하고 계산 필드 재계산

    Returns:
        dict: updated_project (계산 필드 포함)
    """
    import pandas as pd
    from dashboard.services.project_service import (
        _safe_parse_amount,
        _parse_vat_flag,
        _calculate_total2,
        _calculate_outstanding_amount,
        _calculate_net_profit,
        _calculate_margin_rate
    )

    # 저장된 행을 FORMATTED_VALUE로 다시 조회 (계산 필드의 값을 가져오기 위해)
    fresh_values = manager.get_row_values(
        sheet_id=sheet_id,
        sheet_name=sheet_name,
        row_number=row_number,
        start_col='A',
        end_col='AM',
        value_render_option='FORMATTED_VALUE'  # 계산된 값 가져오기
    )

    # fresh_values를 39개 컬럼으로 패딩 (마지막 열이 비어있을 경우 대비)
    while len(fresh_values) < 39:
        fresh_values.append('')

    # [디버깅] fresh_values 내용 확인
    logger.info(f"[PUT] fresh_values 길이: {len(fresh_values)}")
    logger.info(f"[PUT] 수금 관련 필드 값들 - 중도금[U/20]:{fresh_values[20] if len(fresh_values) > 20 else 'N/A'}, 잔금[V/21]:{fresh_values[21] if len(fresh_values) > 21 else 'N/A'}, 총액2[S/18]:{fresh_values[18] if len(fresh_values) > 18 else 'N/A'}, 미수금[W/22]:{fresh_values[22] if len(fresh_values) > 22 else 'N/A'}, 마진율[AF/31]:{fresh_values[31] if len(fresh_values) > 31 else 'N/A'}")

    # 업데이트된 행을 딕셔너리로 직접 변환
    updated_project = {}
    if fresh_values:
        # fresh_values를 딕셔너리로 변환
        for field_name, index in field_to_index.items():
            if index < len(fresh_values):
                value = fresh_values[index]
                # 빈 값 처리
                if value is None or (isinstance(value, float) and pd.isna(value)):
                    updated_project[field_name] = ''
                else:
                    updated_project[field_name] = str(value)

        # 날짜 필드 형식 변환 (Excel serial → YYYY-MM-DD)
        date_fields = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
        for field_name in date_fields:
            if field_name in updated_project:
                date_value = updated_project[field_name]
                if date_value and date_value != '':
                    try:
                        # pd.to_datetime으로 다양한 형식 지원
                        parsed_date = pd.to_datetime(date_value, errors='coerce')
                        if pd.notna(parsed_date):
                            updated_project[field_name] = parsed_date.strftime('%Y-%m-%d')
                        else:
                            updated_project[field_name] = ''
                    except:
                        updated_project[field_name] = ''
                else:
                    updated_project[field_name] = ''

        logger.info(f"[PUT] 업데이트된 프로젝트 데이터 생성")
        logger.info(f"[PUT] updated_project 수금 필드 - 중도금:{updated_project.get('중도금', 'N/A')}, 잔금:{updated_project.get('잔금', 'N/A')}, 총액2:{updated_project.get('총액 2', 'N/A')}, 미수금:{updated_project.get('미수금', 'N/A')}, 마진율:{updated_project.get('마진율', 'N/A')}")

        # ✨ 백엔드에서 계산 필드 재계산
        # 저장 직후 fresh_values는 수식 재계산 전이라 잘못된 값일 수 있음
        row_series = pd.Series(updated_project)

        # 총액2 계산 (부가세 플래그 기반)
        total1 = _safe_parse_amount(updated_project.get('총액 1', 0))
        vat_flag = _parse_vat_flag(updated_project.get('부가세'))
        total2 = _calculate_total2(total1, vat_flag)

        # 미수금, 순익, 마진율 계산
        outstanding = _calculate_outstanding_amount(row_series)
        net_profit = _calculate_net_profit(row_series)
        margin_rate = _calculate_margin_rate(row_series, net_profit)

        # 계산된 값을 updated_project에 덮어쓰기
        updated_project['총액 2'] = str(int(total2)) if total2 != 0 else ""
        updated_project['미수금'] = str(int(outstanding)) if outstanding != 0 else ""
        updated_project['순익'] = str(int(net_profit)) if net_profit != 0 else ""
        updated_project['마진율'] = str(round(margin_rate, 1)) if margin_rate != 0 else "0"

        logger.info(f"[PUT] 백엔드 계산 완료 - 부가세:{vat_flag}, 총액2:{updated_project['총액 2']}, 미수금:{updated_project['미수금']}, 순익:{updated_project['순익']}, 마진율:{updated_project['마진율']}")
    else:
        logger.warning(f"[PUT] fresh_values가 비어있음 (저장 후 조회 실패)")

    return updated_project


def _update_calendar_if_needed(updated_project, project_code, final_project_code):
    """구글 캘린더 이벤트 업데이트"""
    if updated_project:
        try:
            update_project_calendar_event(
                updated_project,
                old_project_code=project_code if project_code != final_project_code else None
            )
        except Exception as calendar_error:
            logger.warning(f"[CALENDAR] 이벤트 업데이트 실패 ({final_project_code}): {calendar_error}")


# ============================================================================
# 메인 함수: update_project() - 리팩토링됨
# ============================================================================

@projects_bp.route('/api/projects/<project_code>', methods=['PUT'])
@login_required
@track_business_operation("api_project_update")
def update_project(project_code):
    """프로젝트 업데이트 API (통합 편집 방식 - 전체 행 업데이트) - 리팩토링됨"""
    try:
        data = request.get_json()

        # 1. 요청 검증
        error_response = _validate_update_request(data, project_code)
        if error_response:
            return error_response

        # 2. 환경 설정
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'success': False, 'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500

        sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황의 사본')
        manager = get_sheets_manager()

        # 3. 프로젝트 행 로드
        row_number, current_values, error_response = _load_project_row(
            manager, sheet_id, sheet_name, project_code
        )
        if error_response:
            return error_response

        # 4. Optimistic Lock 버전 검증
        expected_version = data.get('_version')
        new_version, error_response = _check_optimistic_lock_update(
            manager, sheet_id, sheet_name, project_code, row_number,
            current_values, expected_version
        )
        if error_response:
            return error_response

        # 버전 업데이트
        current_values[39] = str(new_version)

        # 5. 필드 매핑 생성
        column_mapping = manager.get_column_mapping()
        field_to_index = {}
        for col_letter, field_name in column_mapping.items():
            if len(col_letter) == 1:
                column_index = ord(col_letter) - ord('A')
            else:
                column_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
            field_to_index[field_name] = column_index

        # 6. 프로젝트 코드 자동 업데이트 (사업자/담당자 변경 시)
        final_project_code, code_change = _update_project_code_if_needed(
            data, current_values, row_number, manager, project_code, field_to_index
        )

        # 프로젝트 코드 변경사항 추적
        field_changes = []
        if code_change:
            field_changes.append(code_change)

        # 7. 필드 업데이트 적용
        update_changes = _apply_field_updates(data, current_values, field_to_index, project_code)
        field_changes.extend(update_changes)

        # 8. Google Sheets 업데이트
        range_name = f'{sheet_name}!A{row_number}:AN{row_number}'
        manager.update_row(sheet_id, row_number, current_values, range_name)
        logger.info(f"[PUT] 전체 행 업데이트 완료: {project_code}, {len(field_changes)}개 필드 변경")

        # 9. 금액 필드 자동 댓글 처리
        _process_payment_field_comments(manager, sheet_id, sheet_name, row_number, field_changes)

        # 10. 감사 로그 배치 기록
        _record_update_audit_logs(field_changes, project_code)

        # 11. 캐시 무효화
        invalidate_project_cache(final_project_code)
        logger.info(f"[PUT] 프로젝트 업데이트 완료: {final_project_code}")

        # 12. 최신 데이터 조회 및 계산 필드 재계산
        updated_project = _fetch_and_calculate_updated_project(
            manager, sheet_id, sheet_name, row_number, field_to_index
        )

        # 13. 구글 캘린더 이벤트 업데이트
        _update_calendar_if_needed(updated_project, project_code, final_project_code)

        # 14. 최종 응답 반환
        logger.info(f"[PUT] 최종 응답 - project 필드 존재: {updated_project is not None}, 필드 개수: {len(updated_project) if updated_project else 0}")

        return jsonify({
            'success': True,
            'message': '성공적으로 업데이트되었습니다.',
            'project_code': final_project_code,
            'old_project_code': project_code if final_project_code != project_code else None,
            'project': updated_project
        })

    except Exception as e:
        import traceback
        error_traceback = traceback.format_exc()
        logger.error(f"[PUT] 프로젝트 업데이트 오류: {project_code}, {str(e)}")
        logger.error(f"[PUT] 스택 트레이스:\n{error_traceback}")
        return jsonify({'success': False, 'error': str(e)}), 500

# ============================================================================
# 헬퍼 함수: update_project_inline() 리팩토링
# ============================================================================

def _validate_inline_update_request(data):
    """
    인라인 업데이트 요청 데이터 검증

    Args:
        data: 요청 JSON 데이터

    Returns:
        tuple: (project_code, validated_data) 또는 (None, error_response)
    """
    project_code = data.get('projectCode') or data.get('프로젝트 코드')

    if not project_code:
        return None, (jsonify({'success': False, 'error': '프로젝트 코드가 필요합니다.'}), 400)

    # Marshmallow 스키마로 데이터 검증
    validation_data = {
        'project_code': project_code,
        'manager': data.get('담당자'),
        'company': data.get('사업자'),
        'address': data.get('현장 주소'),
        'work_content': data.get('공사 내용'),
        'start_date': data.get('공사 시작'),
        'end_date': data.get('공사 종료'),
        'down_payment': data.get('계약금'),
        'mid_payment': data.get('중도금'),
        'final_payment': data.get('잔금'),
        'total_amount': data.get('총액 1'),
    }

    # None 값 제거 (업데이트되지 않는 필드)
    validation_data = {k: v for k, v in validation_data.items() if v is not None}

    validated_data, errors = validate_request_data(ProjectUpdateSchema, validation_data)

    if errors:
        error_messages = format_validation_errors(errors)
        logger.warning(f"[VALIDATION] 인라인 업데이트 검증 실패: {error_messages}")
        return None, (jsonify({
            'success': False,
            'error': '입력 데이터 검증 실패',
            'validation_errors': error_messages
        }), 400)

    return project_code, None


def _load_row_data_for_inline_update(manager, sheet_id, sheet_name, project_code):
    """
    인라인 업데이트를 위한 행 데이터 로드

    Args:
        manager: GoogleSheetsManager 인스턴스
        sheet_id: 구글 시트 ID
        sheet_name: 시트 이름
        project_code: 프로젝트 코드

    Returns:
        tuple: (row_number, current_values) 또는 (None, error_response)
    """
    # 프로젝트가 있는 행 찾기
    row_number = manager.find_row_by_project_code(sheet_id, project_code)

    if not row_number:
        return None, (jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404)

    # 현재 행의 데이터를 가져오기 (AN 컬럼까지 - _version 포함)
    current_values = manager.get_row_values(
        sheet_id=sheet_id,
        sheet_name=sheet_name,
        row_number=row_number,
        start_col='A',
        end_col='AN'
    )

    # 현재 값을 리스트로 확장 (40개 컬럼, AN까지)
    while len(current_values) < 40:
        current_values.append('')

    return row_number, current_values


def _check_optimistic_lock_inline(data, current_values, project_code, manager, sheet_id, sheet_name, row_number):
    """
    Optimistic Lock 검증 (버전 충돌 확인)

    Args:
        data: 요청 데이터
        current_values: 현재 행 값
        project_code: 프로젝트 코드
        manager: GoogleSheetsManager 인스턴스
        sheet_id: 구글 시트 ID
        sheet_name: 시트 이름
        row_number: 행 번호

    Returns:
        tuple: (new_version, None) 성공 시, (None, error_response) 충돌 시
    """
    expected_version = data.get('_version')
    current_version = current_values[39]  # AN 컬럼 (index 39)

    # 버전 값 정규화 (빈 문자열/None → '0')
    if not current_version or current_version == '':
        current_version = '0'
    if not expected_version or expected_version == '':
        expected_version = '0'

    # 버전 불일치 → 409 Conflict
    if str(current_version) != str(expected_version):
        logger.warning(f"[OPTIMISTIC_LOCK][INLINE] 버전 충돌 감지: {project_code}, expected={expected_version}, current={current_version}")

        # 현재 최신 데이터를 반환하여 클라이언트가 병합할 수 있도록 함
        fresh_values = manager.get_row_values(
            sheet_id=sheet_id,
            sheet_name=sheet_name,
            row_number=row_number,
            start_col='A',
            end_col='AN',
            value_render_option='FORMATTED_VALUE'
        )

        # 컬럼 매핑 임시 생성
        temp_column_mapping = manager.get_column_mapping()
        temp_field_to_index = {}
        for col_letter, field_name in temp_column_mapping.items():
            if len(col_letter) == 1:
                column_index = ord(col_letter) - ord('A')
            else:
                column_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
            temp_field_to_index[field_name] = column_index

        # 최신 데이터를 딕셔너리로 변환
        current_data = {}
        for field_name, index in temp_field_to_index.items():
            if index < len(fresh_values):
                value = fresh_values[index]
                current_data[field_name] = '' if value is None or value == '' else str(value)

        return None, (jsonify({
            'success': False,
            'error': 'conflict',
            'message': '다른 사용자가 먼저 수정했습니다. 페이지를 새로고침한 후 다시 시도하세요.',
            'current_data': current_data,
            'current_version': str(current_version)
        }), 409)

    # 버전 일치 → 버전 증가
    new_version = int(current_version) + 1
    logger.info(f"[OPTIMISTIC_LOCK][INLINE] 버전 업데이트: {project_code}, {current_version} → {new_version}")

    return new_version, None


def _check_and_update_project_code(data, current_values, row_number, manager, project_code):
    """
    사업자/담당자 변경 시 프로젝트 코드 자동 재산출

    Args:
        data: 요청 데이터
        current_values: 현재 행 값 (수정됨)
        row_number: 행 번호
        manager: GoogleSheetsManager 인스턴스
        project_code: 원본 프로젝트 코드

    Returns:
        tuple: (updated_project_code, field_change_dict or None)
    """
    from dashboard.utils.project_code_generator import generate_project_code, should_update_project_code

    # 컬럼 매핑에서 field_to_index 생성
    column_mapping = manager.get_column_mapping()
    field_to_index = {}
    for col_letter, field_name in column_mapping.items():
        if len(col_letter) == 1:
            column_index = ord(col_letter) - ord('A')
        else:
            column_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
        field_to_index[field_name] = column_index

    # 사업자/담당자 값 가져오기 (변경된 값 또는 기존 값)
    company = data.get('사업자') or current_values[field_to_index.get('사업자', 1)]
    manager_field = data.get('담당자') or current_values[field_to_index.get('담당자', 2)]

    # 프로젝트 코드 재산출 필요 여부 확인
    if should_update_project_code(project_code, row_number, company, manager_field):
        new_project_code = generate_project_code(row_number, company, manager_field)
        if new_project_code and new_project_code != project_code:
            logger.info(f"[인라인] 프로젝트 코드 자동 업데이트: {project_code} → {new_project_code}")

            # 프로젝트 코드 업데이트
            current_values[0] = new_project_code

            # field_changes에 추가할 딕셔너리 반환
            return new_project_code, {
                'field_name': '프로젝트 코드',
                'old_value': project_code,
                'new_value': new_project_code
            }

    return project_code, None


def _apply_inline_field_changes(data, current_values, original_values, manager, project_code):
    """
    필드 변경사항을 current_values에 적용하고 감사 로그 기록

    Args:
        data: 요청 데이터
        current_values: 현재 행 값 (수정됨)
        original_values: 원본 행 값 (감사 로그용)
        manager: GoogleSheetsManager 인스턴스
        project_code: 프로젝트 코드

    Returns:
        tuple: (field_changes, final_project_code)
            - field_changes: 감사 로그용 변경 목록
            - final_project_code: 최종 프로젝트 코드 (자동 변경된 경우 새 코드)
    """
    column_mapping = manager.get_column_mapping()
    field_changes = []

    # 계산 필드 목록 (구글 시트에 수식이 있으므로 업데이트하지 않음)
    calculated_fields = ['총액 2', '총액2', '미수금', '마진율', '순익']

    # 날짜 필드 목록 (감사 로그 기록 시 Excel 시리얼 변환 필요)
    date_fields = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']

    # 업데이트할 필드만 변경
    for field_name, new_value in data.items():
        if field_name in ['프로젝트 코드', '_version', 'projectCode']:
            continue  # 프로젝트 코드와 버전은 변경하지 않음

        # 계산 필드는 건너뛰기 (구글 시트 수식 보존)
        if field_name in calculated_fields:
            logger.info(f"[인라인] 계산 필드 건너뛰기: {field_name} (구글 시트 수식 사용)")
            continue

        # 필드명에 해당하는 컬럼 인덱스 찾기
        column_index = None
        for col_letter, col_name in column_mapping.items():
            if col_name == field_name:
                # 컬럼 문자를 인덱스로 변환
                if len(col_letter) == 1:
                    column_index = ord(col_letter) - ord('A')
                else:
                    column_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
                break

        if column_index is not None and column_index < len(current_values):
            # 기존 값 저장 (감사 로그용 - 원본 값 사용)
            old_value = original_values[column_index] if column_index < len(original_values) else ''

            # 새 값 준비
            if new_value == '-' or new_value == '':
                new_value_processed = ''
            else:
                new_value_processed = str(new_value)

            # 값이 실제로 변경되었는지 확인
            if old_value != new_value_processed:
                field_changes.append({
                    'field_name': field_name,
                    'old_value': old_value,
                    'new_value': new_value_processed
                })

            # 값 업데이트
            current_values[column_index] = new_value_processed

    # 감사 로그 기록
    if field_changes:
        try:
            from dashboard.utils.user_database import get_audit_repository
            audit_repo = get_audit_repository()
            user_email = session.get('user', {}).get('email', 'unknown')

            for change in field_changes:
                # 날짜 필드인 경우 Excel 시리얼 번호를 YYYY-MM-DD 형식으로 변환
                old_value_display = change['old_value']
                new_value_display = change['new_value']

                if change['field_name'] in date_fields:
                    old_value_display = convert_excel_serial_to_date(old_value_display)
                    new_value_display = convert_excel_serial_to_date(new_value_display)

                audit_repo.log_action(
                    user_email=user_email,
                    action='UPDATE_FIELD',
                    details=f"프로젝트 {project_code}의 {change['field_name']} 필드를 '{old_value_display}'에서 '{new_value_display}'로 변경",
                    project_code=project_code,
                    field_name=change['field_name'],
                    old_value=old_value_display,
                    new_value=new_value_display,
                    ip_address=request.remote_addr
                )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")

    # 최종 프로젝트 코드 반환 (변경되지 않았으면 원본 그대로)
    return field_changes, project_code


def _save_and_return_inline_result(manager, sheet_id, sheet_name, row_number, current_values, project_code):
    """
    구글 시트 업데이트 및 결과 반환

    Args:
        manager: GoogleSheetsManager 인스턴스
        sheet_id: 구글 시트 ID
        sheet_name: 시트 이름
        row_number: 행 번호
        current_values: 업데이트할 행 값
        project_code: 프로젝트 코드

    Returns:
        tuple: (jsonify response, status_code)
    """
    # 구글 시트 업데이트 (AN 컬럼까지 포함)
    range_name = f'{sheet_name}!A{row_number}:AN{row_number}'
    update_result = manager.update_row(sheet_id, row_number, current_values, range_name)

    logger.info(f"[인라인] 프로젝트 업데이트 완료: {project_code}")

    # 업데이트된 행만 다시 읽기 (계산 필드가 구글 시트 수식으로 계산되므로)
    updated_row_values = manager.get_row_values(
        sheet_id=sheet_id,
        sheet_name=sheet_name,
        row_number=row_number,
        start_col='A',
        end_col='AM'
    )

    # 행 데이터를 딕셔너리로 변환
    column_mapping = manager.get_column_mapping()
    date_fields = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']

    updated_project = {}
    for col_letter, field_name in column_mapping.items():
        # 컬럼 문자를 인덱스로 변환
        if len(col_letter) == 1:
            col_index = ord(col_letter) - ord('A')
        else:
            col_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))

        if col_index < len(updated_row_values):
            value = updated_row_values[col_index]
            # 날짜 필드는 Excel 시리얼 번호를 YYYY-MM-DD로 변환
            if field_name in date_fields and value:
                value = convert_excel_serial_to_date(value)
            updated_project[field_name] = value if value else ''
        else:
            updated_project[field_name] = ''

    logger.info(f"[인라인] 업데이트된 행 데이터 로드 완료: {project_code}")

    # 캐시 무효화
    invalidate_project_cache(project_code)

    # 업데이트된 프로젝트 데이터 반환
    return jsonify({
        'success': True,
        'message': '성공적으로 업데이트되었습니다.',
        'project_code': project_code,
        'project': updated_project
    }), 200


# ============================================================================
# 메인 API 엔드포인트 (리팩토링됨)
# ============================================================================

@projects_bp.route('/api/update-project-inline', methods=['POST'])
@login_required
@track_business_operation("api_project_inline_update")
def update_project_inline():
    """프로젝트 인라인 편집 API - 구글 시트 직접 업데이트 (리팩토링됨)"""
    try:
        data = request.get_json()

        # 1. 요청 데이터 검증
        project_code, error_response = _validate_inline_update_request(data)
        if error_response:
            return error_response

        # 2. 환경 설정
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'success': False, 'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500

        sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황의 사본')
        manager = get_sheets_manager()

        # 3. 현재 행 데이터 로드
        row_number, current_values_or_error = _load_row_data_for_inline_update(
            manager, sheet_id, sheet_name, project_code
        )
        if row_number is None:
            return current_values_or_error  # 에러 응답 반환

        current_values = current_values_or_error

        # 4. Optimistic Lock 검증
        new_version, lock_error = _check_optimistic_lock_inline(
            data, current_values, project_code,
            manager, sheet_id, sheet_name, row_number
        )
        if lock_error:
            return lock_error  # 409 Conflict 반환

        # 버전 업데이트
        current_values[39] = str(new_version)

        # 5. 프로젝트 코드 자동 재산출 (사업자/담당자 변경 시)
        original_values = current_values.copy()
        final_project_code, code_change = _check_and_update_project_code(
            data, current_values, row_number, manager, project_code
        )

        # 6. 필드 변경사항 적용 및 감사 로그
        field_changes, _ = _apply_inline_field_changes(
            data, current_values, original_values, manager, final_project_code
        )

        # 프로젝트 코드가 자동 변경된 경우 field_changes 및 감사 로그 추가
        if code_change:
            field_changes.insert(0, code_change)  # 맨 앞에 추가

            # 프로젝트 코드 변경 감사 로그 기록
            try:
                from dashboard.utils.user_database import get_audit_repository
                audit_repo = get_audit_repository()
                user_email = session.get('user', {}).get('email', 'unknown')

                audit_repo.log_action(
                    user_email=user_email,
                    action='UPDATE_FIELD',
                    details=f"프로젝트 {project_code}의 프로젝트 코드 필드를 '{code_change['old_value']}'에서 '{code_change['new_value']}'로 변경 (자동)",
                    project_code=final_project_code,
                    field_name='프로젝트 코드',
                    old_value=code_change['old_value'],
                    new_value=code_change['new_value'],
                    ip_address=request.remote_addr
                )
            except Exception as log_error:
                logger.warning(f"프로젝트 코드 변경 감사 로그 기록 실패: {log_error}")

        # 7. 구글 시트 업데이트 및 결과 반환
        return _save_and_return_inline_result(
            manager, sheet_id, sheet_name, row_number, current_values, final_project_code
        )

    except Exception as e:
        logger.error(f"인라인 업데이트 오류: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

# 서비스 워커 라우트
@projects_bp.route('/sw.js')
def service_worker():
    """동적 서비스 워커 생성 - Vite 빌드 결과에 맞는 자산 경로 포함"""
    from ..utils.frontend_helpers import asset_manager
    from flask import Response

    # 현재 빌드된 자산들의 실제 경로 가져오기
    assets = asset_manager.get_service_worker_assets()

    # 안전한 중복 제거를 위한 URL 목록 생성
    raw_urls = [
        '/',  # 루트 경로 명시적 포함
        *assets['css_files'],
        *assets['js_files'],
        *assets['static_files']
    ]

    # 중복 제거 및 정렬 (None/빈 문자열 필터링 포함)
    all_urls = sorted(list(set(filter(None, raw_urls))))

    # 동적 서비스 워커 생성
    sw_content = f"""/**
 * Service Worker - 동적 PWA 기능
 * Vite 빌드 결과에 맞는 자산 경로로 자동 업데이트
 */

const CACHE_NAME = 'itg-dashboard-v4';
const STATIC_CACHE_URLS = {json.dumps(all_urls, indent=2)};

// 설치 이벤트
self.addEventListener('install', event => {{
    console.log('[SW] 서비스 워커 설치 중...');
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(STATIC_CACHE_URLS))
            .then(() => self.skipWaiting())
    );
}});

// 활성화 이벤트
self.addEventListener('activate', event => {{
    console.log('[SW] 서비스 워커 활성화');
    event.waitUntil(
        caches.keys().then(cacheNames => {{
            return Promise.all(
                cacheNames.map(cacheName => {{
                    if (cacheName !== CACHE_NAME) {{
                        console.log('[SW] 이전 캐시 삭제:', cacheName);
                        return caches.delete(cacheName);
                    }}
                }})
            );
        }}).then(() => self.clients.claim())
    );
}});

// 네트워크 요청 가로채기
self.addEventListener('fetch', event => {{
    if (event.request.method !== 'GET') return;

    event.respondWith(
        caches.match(event.request)
            .then(response => {{
                if (response) {{
                    return response;
                }}
                return fetch(event.request);
            }})
    );
}});
"""

    return Response(sw_content, mimetype='application/javascript')

# 캐시 상태 API (프로젝트 페이지용)
@projects_bp.route('/api/cache/status')
@login_required
def get_cache_status():
    """캐시 상태 조회 (프로젝트 페이지용)"""
    try:
        # 프런트엔드 CacheStatusManager 기대치에 맞춰 응답
        return jsonify({
            'cache_health': 'healthy',  # 'healthy', 'warning', 'error' 중 하나
            'prefetch_running': True,   # 백그라운드 갱신 상태
            'last_successful_update': datetime.now().isoformat(),  # 마지막 성공 업데이트 시간
            'cache_enabled': True,      # 호환성을 위해 유지
            'message': 'Cache status retrieved successfully'
        })
    except Exception as e:
        logger.error(f"캐시 상태 조회 오류: {str(e)}")
        return jsonify({'error': str(e)}), 500

# 전문가 리뷰: "초기 페이지에서는 최소 메타 정보만 제공하고 상세 데이터는 API로만 받도록"
@projects_bp.route('/api/projects/statistics')
@login_required
def get_project_statistics():
    """
    프로젝트 통계 정보 API
    템플릿 렌더링과 분리하여 데이터 계약 단일화
    """
    try:
        # 통계 계산을 위한 최소한의 데이터만 로드
        projects_data = get_project_records()

        if not projects_data:
            return jsonify({
                'success': True,
                'data': {
                    'total_projects': 0,
                    'total_amount': 0,
                    'pending_amount': 0,
                    'completed_projects': 0,
                    'last_updated': datetime.now().isoformat()
                }
            })

        # 통계 계산
        import math

        total_amount = sum(
            safe_parse_currency(p.get('총액 1', 0)) + safe_parse_currency(p.get('총액 2', 0))
            for p in projects_data
        )
        pending_amount = sum(safe_parse_currency(p.get('미수금', 0)) for p in projects_data)

        # NaN 값을 0으로 변환
        if math.isnan(total_amount):
            total_amount = 0
        if math.isnan(pending_amount):
            pending_amount = 0

        statistics = {
            'total_projects': len(projects_data),
            'total_amount': total_amount,
            'pending_amount': pending_amount,
            'completed_projects': len([p for p in projects_data if p.get('상태') == '완료']),
            'last_updated': datetime.now().isoformat()
        }

        return jsonify({
            'success': True,
            'data': statistics
        })

    except Exception as e:
        logger.error(f"프로젝트 통계 로드 실패: {e}")
        return jsonify({
            'success': False,
            'error': '통계 정보를 불러올 수 없습니다.',
            'message': str(e)
        }), 500


@projects_bp.route('/api/projects/field-memo', methods=['POST'])
@login_required
@track_business_operation("api_field_memo_save")
def save_field_memo():
    """
    필드 메모 저장 API (구글 시트 셀 노트로 저장)

    Request Body:
        {
            "project_code": "G0001",
            "field_name": "계약금" | "중도금" | "잔금",
            "memo": "입금일: 2025-01-10\\n입금자: 홍길동" | null
        }
    """
    try:
        data = request.get_json()
        project_code = data.get('project_code')
        field_name = data.get('field_name')  # '계약금', '중도금', '잔금'
        memo = data.get('memo')  # None이면 메모 삭제

        # 1. 기본 필수 필드 검증
        if not project_code or not field_name:
            return jsonify({
                'success': False,
                'message': '프로젝트 코드와 필드명이 필요합니다.'
            }), 400

        # 2. 필드명 유효성 검증 (상수 사용)
        if field_name not in MEMOABLE_FIELDS:
            return jsonify({
                'success': False,
                'message': ERROR_MESSAGES['memo']['invalid_field']
            }), 400

        # 3. 컬럼 매핑 (상수 사용)
        column = PAYMENT_FIELD_TO_COLUMN.get(field_name)
        if not column:
            return jsonify({
                'success': False,
                'message': f'컬럼 매핑을 찾을 수 없습니다: {field_name}'
            }), 500

        # 4. Google Sheets Manager 및 설정 초기화 (한 번만)
        manager = get_sheets_manager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황의 사본')

        # 5. 프로젝트 행 번호 조회 (한 번만 - 중복 제거)
        row_number = manager.find_row_by_project_code(sheet_id, project_code, f'{sheet_name}!A:A')

        if not row_number:
            return jsonify({
                'success': False,
                'message': ERROR_MESSAGES['memo']['project_not_found']
            }), 404

        # 6. Marshmallow 스키마로 상세 검증 (메모가 있는 경우에만)
        if memo:
            # 이미 조회된 row_number 재사용
            validation_data = {
                'row': row_number,
                'column': column,
                'memo': memo,
                'project_code': project_code,
                'field_name': field_name
            }

            validated_data, errors = validate_request_data(CellMemoSchema, validation_data)

            if errors:
                error_messages = format_validation_errors(errors)
                logger.warning(f"[VALIDATION] 필드 메모 검증 실패: {error_messages}")
                return jsonify({
                    'success': False,
                    'message': ERROR_MESSAGES['memo']['validation_failed'],
                    'validation_errors': error_messages
                }), 400

        # 셀 주소 구성 (예: 'T5')
        cell_address = f"{column}{row_number}"

        logger.info(f"[FIELD_MEMO] 셀 노트 저장 시작: {project_code} / {cell_address}")

        # 기존 메모 값 조회 (old_value 기록용)
        old_memo = manager.get_cell_note(sheet_id, sheet_name, cell_address) or ''

        # 구글 시트 셀 노트에 저장
        success = manager.update_cell_note(
            sheet_id,
            sheet_name,
            cell_address,
            memo if memo else None  # 빈 메모는 삭제
        )

        if success:
            # 셀 노트 캐시만 무효화 (성능 최적화)
            # 메모는 셀 노트에만 저장되고 프로젝트 데이터에 영향 없음
            # 따라서 전체 프로젝트 캐시 무효화는 불필요
            notes_cache_key = f"cell_notes_{sheet_id}"
            invalidated_notes = smart_invalidate(notes_cache_key)

            logger.info(f"[FIELD_MEMO][PID:{os.getpid()}] 셀 노트 캐시 무효화 완료:")
            logger.info(f"  - {notes_cache_key}: {invalidated_notes}개 항목")

            # 감사 로그 기록
            try:
                from dashboard.utils.user_database import get_audit_repository
                audit_repo = get_audit_repository()
                user_email = session.get('user', {}).get('email', 'unknown')

                action_desc = "삭제" if not memo else "저장"
                # 이전 값 포맷팅
                old_value_display = old_memo[:50] + '...' if old_memo and len(old_memo) > 50 else (old_memo or '-')
                new_value_display = memo[:50] + '...' if memo and len(memo) > 50 else (memo or '삭제')

                audit_repo.log_action(
                    user_email=user_email,
                    action='UPDATE_FIELD_MEMO',
                    details=f"프로젝트 {project_code}의 {field_name} 메모 {action_desc}",
                    project_code=project_code,
                    field_name=f"{field_name}_메모",
                    old_value=old_value_display,
                    new_value=new_value_display,
                    ip_address=request.remote_addr
                )
            except Exception as log_error:
                logger.warning(f"감사 로그 기록 실패: {log_error}")

            logger.info(f"[SUCCESS] 필드 메모 저장: {project_code} / {field_name}")

            return jsonify({
                'success': True,
                'message': '메모가 저장되었습니다.',
                'data': {
                    'project_code': project_code,
                    'field_name': field_name,
                    'memo': memo,
                    'cell_address': cell_address
                }
            })
        else:
            return jsonify({
                'success': False,
                'message': '메모 저장에 실패했습니다.'
            }), 500

    except Exception as e:
        logger.error(f"[ERROR] 필드 메모 저장 오류: {str(e)}")
        import traceback
        logger.error(f"스택 트레이스:\n{traceback.format_exc()}")
        return jsonify({
            'success': False,
            'message': f'메모 저장 중 오류가 발생했습니다: {str(e)}'
        }), 500


# ===== 헬퍼 함수 1: 기본 검증 및 행 조회 =====
def _validate_and_find_memo_batch_row(project_code, memos):
    """
    메모 배치 요청 검증 및 프로젝트 행 조회

    Returns:
        tuple: (manager, sheet_id, sheet_name, row_number) or (None, error_response, status_code)
    """
    # 기본 검증
    if not project_code:
        return None, jsonify({
            'success': False,
            'message': '프로젝트 코드가 필요합니다.'
        }), 400

    if not memos or not isinstance(memos, list):
        return None, jsonify({
            'success': False,
            'message': '메모 배열이 필요합니다.'
        }), 400

    # Google Sheets Manager 초기화
    manager = get_sheets_manager()
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황의 사본')

    # 프로젝트 행 번호 조회
    row_number = manager.find_row_by_project_code(sheet_id, project_code, f'{sheet_name}!A:A')

    if not row_number:
        return None, jsonify({
            'success': False,
            'message': ERROR_MESSAGES['memo']['project_not_found']
        }), 404

    logger.info(f"[BATCH_MEMO] 일괄 저장 시작: {project_code}, {len(memos)}개 메모")
    return manager, sheet_id, sheet_name, row_number


# ===== 헬퍼 함수 2: 메모 업데이트 준비 =====
def _prepare_batch_memo_updates(memos, row_number, manager, sheet_id, sheet_name, project_code):
    """
    메모 업데이트 준비: 검증, 매핑, 기존 메모 조회

    Returns:
        tuple: (results, failed_count, cell_notes_to_update, old_memos)
    """
    results = []
    failed_count = 0
    cell_notes_to_update = {}  # {cell_address: note_text}
    old_memos = {}  # {cell_address: old_memo}

    for memo_data in memos:
        field_name = memo_data.get('field_name')
        memo = memo_data.get('memo')

        try:
            # 필드명 검증
            if field_name not in MEMOABLE_FIELDS:
                results.append({
                    'field_name': field_name,
                    'success': False,
                    'message': f'잘못된 필드명: {field_name}'
                })
                failed_count += 1
                continue

            # 컬럼 매핑
            column = PAYMENT_FIELD_TO_COLUMN.get(field_name)
            if not column:
                results.append({
                    'field_name': field_name,
                    'success': False,
                    'message': f'컬럼 매핑 없음: {field_name}'
                })
                failed_count += 1
                continue

            # 셀 주소 구성
            cell_address = f"{column}{row_number}"

            # 기존 메모 값 조회 (감사 로그용)
            old_memo = manager.get_cell_note(sheet_id, sheet_name, cell_address) or ''
            old_memos[cell_address] = old_memo

            # Batch 업데이트 대상에 추가
            cell_notes_to_update[cell_address] = memo if memo else None

            # 결과 목록에 추가 (아직 성공은 아님)
            results.append({
                'field_name': field_name,
                'success': None,  # 나중에 업데이트
                'message': '대기 중',
                'cell_address': cell_address,
                'memo_index': len(results)  # 원본 memos 배열의 인덱스 추적
            })

        except Exception as memo_error:
            logger.error(f"[BATCH_MEMO] 메모 준비 오류 ({field_name}): {str(memo_error)}")
            results.append({
                'field_name': field_name,
                'success': False,
                'message': f'오류: {str(memo_error)}'
            })
            failed_count += 1

    return results, failed_count, cell_notes_to_update, old_memos


# ===== 헬퍼 함수 3: Batch Update 실행 =====
def _execute_batch_memo_update(cell_notes_to_update, manager, sheet_id, sheet_name):
    """
    Google Sheets Batch Update API로 메모 일괄 업데이트

    Returns:
        bool: 성공 여부
    """
    if not cell_notes_to_update:
        return True

    try:
        # 시트 ID 조회
        numeric_sheet_id = manager.get_sheet_id_by_name(sheet_id, sheet_name)

        # batchUpdate 요청 구성
        batch_requests = []
        for cell_address, note_text in cell_notes_to_update.items():
            col_letter = ''.join(filter(str.isalpha, cell_address))
            row_num = int(''.join(filter(str.isdigit, cell_address)))
            col_index = manager._column_letter_to_number(col_letter)
            row_index = row_num - 1  # 0-based

            batch_requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': numeric_sheet_id,
                        'startRowIndex': row_index,
                        'endRowIndex': row_index + 1,
                        'startColumnIndex': col_index,
                        'endColumnIndex': col_index + 1
                    },
                    'cell': {
                        'note': note_text if note_text else None
                    },
                    'fields': 'note'
                }
            })

        # Batch Update 실행
        body = {'requests': batch_requests}
        request_obj = manager.service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body=body
        )
        batch_result = manager._execute_with_retry(
            request_obj,
            f"batch_update_memos({len(batch_requests)} notes)"
        )

        # API 절감 효과 로깅
        individual_api_calls = len(batch_requests) * 2
        saved_api_calls = individual_api_calls - 1
        logger.info(
            f"[BATCH_MEMO] Batch Update 완료: {len(batch_requests)}개 메모 → "
            f"Batch 1회 (개별 대비 {saved_api_calls}회 절감)"
        )

        return True

    except Exception as batch_error:
        logger.error(f"[BATCH_MEMO] Batch Update 실패: {str(batch_error)}")
        return False


# ===== 헬퍼 함수 4: 결과 업데이트 및 감사 로그 준비 =====
def _build_batch_audit_logs(results, memos, old_memos, project_code, batch_success):
    """
    Batch Update 결과에 따라 results 업데이트 및 감사 로그 준비

    Returns:
        tuple: (success_count, failed_count, audit_actions)
    """
    success_count = 0
    failed_count = 0
    audit_actions = []

    for result_item in results:
        if result_item.get('success') is None:  # 대기 중이던 항목
            if batch_success:
                result_item['success'] = True
                result_item['message'] = '저장 완료'
                success_count += 1

                # 감사 로그 데이터 수집
                cell_address = result_item['cell_address']
                memo_index = result_item['memo_index']
                field_name = memos[memo_index]['field_name']
                memo = memos[memo_index]['memo']
                old_memo = old_memos.get(cell_address, '')

                action_desc = "삭제" if not memo else "저장"
                old_value_display = (
                    old_memo[:50] + '...' if old_memo and len(old_memo) > 50
                    else (old_memo or '-')
                )
                new_value_display = (
                    memo[:50] + '...' if memo and len(memo) > 50
                    else (memo or '삭제')
                )

                audit_actions.append({
                    'user_email': session.get('user', {}).get('email', 'unknown'),
                    'action': 'UPDATE_FIELD_MEMO',
                    'details': f"프로젝트 {project_code}의 {field_name} 메모 {action_desc}",
                    'project_code': project_code,
                    'field_name': f"{field_name}_메모",
                    'old_value': old_value_display,
                    'new_value': new_value_display,
                    'ip_address': request.remote_addr
                })
            else:
                result_item['success'] = False
                result_item['message'] = 'Batch 업데이트 실패'
                failed_count += 1
        elif result_item.get('success') is False:
            # 이미 실패 처리된 항목 (검증 실패 등)
            failed_count += 1

    return success_count, failed_count, audit_actions


# ===== 헬퍼 함수 5: 감사 로그 저장 및 캐시 무효화 =====
def _finalize_batch_memo_save(audit_actions, sheet_id, success_count):
    """
    감사 로그 배치 저장 및 캐시 무효화
    """
    # 감사 로그 배치 저장
    if audit_actions:
        try:
            from dashboard.utils.user_database import get_audit_repository
            audit_repo = get_audit_repository()
            batch_success, batch_count = audit_repo.log_actions_batch(audit_actions)
            if batch_success:
                logger.info(
                    f"[BATCH_MEMO] 감사 로그 배치 기록 완료: "
                    f"{batch_count}개 메모 (1회 트랜잭션)"
                )
            else:
                logger.warning(f"[BATCH_MEMO] 감사 로그 배치 기록 실패")
        except Exception as log_error:
            logger.warning(f"[BATCH_MEMO] 감사 로그 배치 기록 실패: {log_error}")

    # 캐시 무효화
    if success_count > 0:
        notes_cache_key = f"cell_notes_{sheet_id}"
        invalidated_notes = smart_invalidate(notes_cache_key)
        invalidated_data = smart_invalidate("current_sheet_data")

        logger.info(
            f"[BATCH_MEMO] 캐시 무효화 완료: "
            f"notes={invalidated_notes}개, data={invalidated_data}개"
        )


# ===== 메인 함수: save_field_memos_batch (리팩토링됨) =====
@projects_bp.route('/api/projects/field-memos/batch', methods=['POST'])
@login_required
@track_business_operation("api_field_memo_batch_save")
def save_field_memos_batch():
    """
    여러 필드 메모 일괄 저장 API (Batch Update로 성능 최적화)

    Request Body:
        {
            "project_code": "G0001-IT",
            "memos": [
                {"field_name": "계약금", "memo": "입금일: 2025-01-10\\n입금자: 홍길동"},
                {"field_name": "중도금", "memo": "입금일: 2025-02-15\\n입금자: 김철수"},
                {"field_name": "잔금", "memo": null}
            ]
        }

    Response:
        {
            "success": true,
            "message": "3개 중 3개 메모 저장 성공",
            "results": [...],
            "failed_count": 0,
            "success_count": 3
        }
    """
    try:
        data = request.get_json()
        project_code = data.get('project_code')
        memos = data.get('memos', [])

        # 1. 기본 검증 및 행 조회
        result = _validate_and_find_memo_batch_row(project_code, memos)
        if result[0] is None:
            return result[1], result[2]  # 에러 응답 반환

        manager, sheet_id, sheet_name, row_number = result

        # 2. 메모 업데이트 준비
        results, failed_count, cell_notes_to_update, old_memos = _prepare_batch_memo_updates(
            memos, row_number, manager, sheet_id, sheet_name, project_code
        )

        # 3. Batch Update 실행
        batch_success = _execute_batch_memo_update(
            cell_notes_to_update, manager, sheet_id, sheet_name
        )

        # 4. 결과 업데이트 및 감사 로그 준비
        success_count, failed_count, audit_actions = _build_batch_audit_logs(
            results, memos, old_memos, project_code, batch_success
        )

        # 5. 감사 로그 저장 및 캐시 무효화
        _finalize_batch_memo_save(audit_actions, sheet_id, success_count)

        # 6. 결과 반환
        overall_success = failed_count == 0
        message = f"{len(memos)}개 중 {success_count}개 메모 저장 성공"
        if failed_count > 0:
            message += f", {failed_count}개 실패"

        logger.info(f"[BATCH_MEMO] 완료: {project_code}, {message}")

        return jsonify({
            'success': overall_success,
            'message': message,
            'results': results,
            'success_count': success_count,
            'failed_count': failed_count,
            'total_count': len(memos)
        }), 200 if overall_success else 207  # 207 Multi-Status

    except Exception as e:
        logger.error(f"[BATCH_MEMO] 일괄 저장 오류: {str(e)}")
        import traceback
        logger.error(f"스택 트레이스:\\n{traceback.format_exc()}")
        return jsonify({
            'success': False,
            'message': f'일괄 메모 저장 중 오류 발생: {str(e)}'
        }), 500



# ===== 헬퍼 함수 1: 데이터 검증 =====
def _validate_project_auto_data(data):
    """
    프로젝트 자동 생성 데이터 검증 (Marshmallow + 기본 체크)

    Returns:
        tuple: (validated_data, None) or (None, error_response_tuple)
    """
    # Marshmallow 스키마 검증
    validation_data = {
        'manager': data.get('담당자', ''),
        'company': data.get('사업자', ''),
        'address': data.get('현장 주소'),
        'work_content': data.get('공사 내용'),
        'start_date': data.get('공사 시작'),
        'end_date': data.get('공사 종료'),
        'down_payment': data.get('계약금'),
        'mid_payment': data.get('중도금'),
        'final_payment': data.get('잔금'),
        'total_amount': data.get('총액 1'),
    }

    validated_data, errors = validate_request_data(ProjectAutoCreateSchema, validation_data)

    if errors:
        error_messages = format_validation_errors(errors)
        logger.warning(f"[VALIDATION] 프로젝트 생성 검증 실패: {error_messages}")
        return None, (jsonify({
            "success": False,
            "error": "입력 데이터 검증 실패",
            "validation_errors": error_messages
        }), 400)

    # 기본 검증 (호환성 유지)
    company = str(data.get("사업자", "")).strip()
    owner = str(data.get("담당자", "")).strip()

    if not company or not owner:
        return None, (jsonify({
            "success": False,
            "error": "사업자/담당자는 필수입니다",
            "code": "REQUIRED_FIELDS_MISSING"
        }), 400)

    return validated_data, None


# ===== 헬퍼 함수 2: 프로젝트 코드 생성 =====
def _generate_project_code(data):
    """
    현재 데이터를 로드하고 프로젝트 코드 생성

    Returns:
        tuple: (code, df) or (None, error_response_tuple)
    """
    # 현재 데이터 로드
    df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
    if df is None:
        df = load_data()
    if df is None:
        return None, (jsonify({
            "success": False,
            "error": "데이터를 불러올 수 없습니다",
            "code": "DATA_LOAD_FAILED"
        }), 500)

    # 프로젝트 코드 생성
    company = str(data.get("사업자", "")).strip()
    owner = str(data.get("담당자", "")).strip()
    code = _auto_project_code(df, company, owner)
    logger.info(f"자동 생성된 프로젝트 코드: {code}")

    return code, df


# ===== 헬퍼 함수 3: Sheets Manager 초기화 =====
def _initialize_sheets_manager():
    """
    Google Sheets Manager 초기화 및 검증

    Returns:
        tuple: (manager, sheet_id) or (None, error_response_tuple)
    """
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        return None, (jsonify({
            "success": False,
            "error": "GOOGLE_SHEET_ID가 설정되지 않았습니다",
            "code": "CONFIG_ERROR"
        }), 500)

    manager = get_sheets_manager()
    if manager is None:
        logger.error("Google Sheets Manager를 초기화할 수 없습니다")
        return None, (jsonify({
            "success": False,
            "error": "Google Sheets 연동이 비활성화되어 있습니다. 관리자에게 문의하세요."
        }), 503)

    return manager, sheet_id


# ===== 헬퍼 함수 4: 기본값 설정 =====
def _prepare_project_defaults(data):
    """
    프로젝트 데이터에 기본값 설정 (금액, 수식, 날짜, VAT 등)

    Args:
        data: 원본 데이터 딕셔너리 (in-place 수정됨)
    """
    from datetime import datetime

    # 공사 종료일 기본값
    if '공사 시작' in data and data['공사 시작']:
        if not data.get('공사 종료') or data.get('공사 종료').strip() == '':
            data['공사 종료'] = data['공사 시작']

    # 금액 필드 기본값
    money_fields = ['총액 1', '제품대', '도급비', '자재비', '기타비', '계약금', '중도금', '잔금']
    for field in money_fields:
        value = data.get(field, '')
        if value is None or str(value).strip() == '':
            data[field] = '₩0'

    # 수식 필드 자동 삽입
    data['총액 2'] = '=IF(R:R=TRUE, Q:Q+ROUND(Q:Q*0.1,0), Q:Q)'
    data['미수금'] = '=0-($S:S-($T:T+$U:U+$V:V))'
    data['순익'] = '=$Q:Q-(($AA:AA)+($AB:AB)+($AC:AC)+($AD:AD))'
    data['마진율'] = '=IF(OR($Q:Q=0, $AE:AE=0), 0, ($AE:AE/$Q:Q))'

    # 기타 필드 기본값
    if '계산서' not in data or not data.get('계산서'):
        data['계산서'] = '미발행'
    if '수금 확인' not in data:
        data['수금 확인'] = 'FALSE'

    # 부가세 포함 여부 (boolean → TRUE/FALSE 문자열)
    if '부가세' in data:
        vat_included = data.get('부가세')
        if isinstance(vat_included, bool):
            data['부가세'] = 'TRUE' if vat_included else 'FALSE'
        elif isinstance(vat_included, str):
            data['부가세'] = vat_included.upper()
    else:
        data['부가세'] = 'FALSE'

    # 공사 확정 날짜 기본값
    if '공사 확정' not in data or not data.get('공사 확정'):
        data['공사 확정'] = datetime.now().strftime('%Y-%m-%d')

    # Optimistic Lock 버전 초기화
    data['_version'] = '0'

    # 총액 1 포맷팅
    if '총액 1' in data and data['총액 1']:
        total_amount = str(data['총액 1']).strip()
        if total_amount.isdigit():
            data['총액 1'] = f"₩{int(total_amount):,}"
        elif total_amount.replace('₩', '').replace(',', '').isdigit():
            clean_number = total_amount.replace('₩', '').replace(',', '').strip()
            if clean_number:
                data['총액 1'] = f"₩{int(clean_number):,}"


# ===== 헬퍼 함수 5: 행 값 배열 구성 =====
def _build_row_values(data, manager):
    """
    데이터를 Google Sheets 행 배열로 변환 (40 컬럼)

    Returns:
        list: 40개 요소의 값 배열
    """
    column_mapping = manager.get_column_mapping()
    values = [''] * 40  # AN 컬럼까지

    # 컬럼 매핑에 따라 값 채우기
    for col_letter, field_name in column_mapping.items():
        if len(col_letter) == 1:
            column_index = ord(col_letter) - ord('A')
        else:
            column_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))

        if field_name in data:
            value = data[field_name]
            if value is None:
                values[column_index] = ''
            elif isinstance(value, str) and value.strip() == '':
                values[column_index] = ''
            else:
                values[column_index] = str(value)

    # 수식 필드 강제 삽입 (컬럼 매핑에 없어도 삽입)
    formula_fields = {
        'S': data.get('총액 2', '=IF(R:R=TRUE, Q:Q+ROUND(Q:Q*0.1,0), Q:Q)'),
        'W': data.get('미수금', '=0-($S:S-($T:T+$U:U+$V:V))'),
        'AE': data.get('순익', '=$Q:Q-(($AA:AA)+($AB:AB)+($AC:AC)+($AD:AD))'),
        'AF': data.get('마진율', '=IF(OR($Q:Q=0, $AE:AE=0), 0, ($AE:AE/$Q:Q))')
    }

    for col_letter, formula in formula_fields.items():
        if len(col_letter) == 1:
            column_index = ord(col_letter) - ord('A')
        else:
            column_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
        values[column_index] = formula

    return values


# ===== 헬퍼 함수 6: 생성 후처리 (캐시, 로그, 캘린더, 리드) =====
def _finalize_project_creation(code, data, project_data):
    """
    프로젝트 생성 후 후처리 작업
    - 캐시 무효화
    - 감사 로그 기록
    - 캘린더 이벤트 생성
    - 리드 연동

    Returns:
        dict: 추가 응답 데이터 {'lead_linked': bool}
    """
    # 1. 캐시 무효화
    invalidate_project_cache(code)
    logger.info(f"[CREATE_PROJECT] 서버 캐시 무효화 완료: {code}")

    # 2. 감사 로그 기록
    try:
        from ..utils.user_database import get_audit_repository
        audit_repo = get_audit_repository()
        user_email = session.get('user', {}).get('email', 'unknown')
        audit_repo.log_action(
            user_email=user_email,
            action='CREATE_PROJECT',
            details=f'새 프로젝트 등록: {code}',
            project_code=code,
            field_name='전체',
            old_value='-',
            new_value='새 프로젝트 생성',
            ip_address=request.remote_addr
        )
    except Exception as log_error:
        logger.warning(f"감사 로그 기록 실패: {log_error}")

    # 3. 캘린더 이벤트 생성
    try:
        logger.info(f"[CALENDAR] 이벤트 생성 시도: {code}")
        event_id = create_project_calendar_event(project_data)
        if event_id:
            logger.info(f"[CALENDAR] 이벤트 생성 성공: {code} -> {event_id}")
        else:
            logger.info(f"[CALENDAR] 이벤트 생성 건너뜀 (조건 미충족): {code}")
    except Exception as calendar_error:
        logger.warning(f"[CALENDAR] 이벤트 생성 실패 ({code}): {calendar_error}", exc_info=True)

    # 4. 리드 연동
    lead_no = data.get('lead_no')
    if lead_no:
        try:
            from ..services.lead_service import update_lead_status
            lead_update_result = update_lead_status(lead_no, '공사 확정')
            if lead_update_result.get('success'):
                logger.info(f"[LEAD_LINKED] 리드 {lead_no} 상태를 '공사 확정'으로 업데이트 완료")
            else:
                logger.warning(f"[LEAD_LINKED] 리드 {lead_no} 상태 업데이트 실패: {lead_update_result.get('error')}")
        except Exception as lead_error:
            logger.warning(f"[LEAD_LINKED] 리드 상태 업데이트 중 오류 (프로젝트 생성은 성공): {lead_error}")

    return {'lead_linked': bool(lead_no)}


# ===== 헬퍼 함수 7: 응답 데이터 구성 (금액 계산 포함) =====
def _build_project_response_data(code, data):
    """
    프로젝트 생성 응답 데이터 구성 (금액 계산 포함)

    Returns:
        dict: 프로젝트 데이터
    """
    from datetime import datetime

    # 금액 파싱 함수
    def parse_amount(value):
        """₩1,000 형식의 문자열을 숫자로 변환"""
        if not value:
            return 0
        clean = str(value).replace('₩', '').replace(',', '').strip()
        try:
            return int(clean) if clean else 0
        except ValueError:
            return 0

    # 총액 1 파싱
    total_1 = parse_amount(data.get('총액 1', '₩0'))

    # 총액 2 계산 (부가세 포함 여부)
    vat_included = data.get('부가세', 'FALSE')
    if vat_included == 'TRUE':
        total_2 = total_1 + round(total_1 * 0.1)
    else:
        total_2 = total_1

    # 입금액 계산
    contract = parse_amount(data.get('계약금', '₩0'))
    interim = parse_amount(data.get('중도금', '₩0'))
    final = parse_amount(data.get('잔금', '₩0'))
    total_paid = contract + interim + final

    # 미수금 계산
    receivable = total_2 - total_paid

    # 비용 필드
    product_cost = parse_amount(data.get('제품대', '₩0'))
    contract_cost = parse_amount(data.get('도급비', '₩0'))
    material_cost = parse_amount(data.get('자재비', '₩0'))
    other_cost = parse_amount(data.get('기타비', '₩0'))
    total_cost = product_cost + contract_cost + material_cost + other_cost

    # 순익 계산
    net_profit = total_1 - total_cost

    # 마진율 계산
    margin_rate = (net_profit / total_1 * 100) if total_1 > 0 else 0

    project_data = {
        '프로젝트 코드': code,
        '사업자': data.get('사업자', ''),
        '거래처': data.get('거래처', ''),
        '담당자': data.get('담당자', ''),
        '담당자 연락처': data.get('담당자 연락처', ''),
        '담당자 이메일': data.get('담당자 이메일', ''),
        '현장 주소': data.get('현장 주소', ''),
        '시공자': data.get('시공자', ''),
        '현장 담당자': data.get('현장 담당자', ''),
        '공사 내용': data.get('공사 내용', ''),
        '공사 시작': data.get('공사 시작', ''),
        '공사 종료': data.get('공사 종료', ''),
        '공사 확정': data.get('공사 확정', datetime.now().strftime('%Y-%m-%d')),
        '총액 1': total_1,
        '부가세': data.get('부가세', 'FALSE'),
        '총액 2': total_2,
        '계약금': contract,
        '중도금': interim,
        '잔금': final,
        '미수금': receivable,
        '수금 확인': data.get('수금 확인', 'FALSE'),
        '계산서': data.get('계산서', '미발행'),
        '공사 구분': data.get('공사 구분', ''),
        '기계 분류': data.get('기계 분류', ''),
        '브랜드': data.get('브랜드', ''),
        '도급 구분': data.get('도급 구분', ''),
        '제품대': product_cost,
        '도급비': contract_cost,
        '자재비': material_cost,
        '기타비': other_cost,
        '순익': net_profit,
        '마진율': margin_rate,
        '견적서 및 계약서 폴더 경로': data.get('견적서 및 계약서 폴더 경로', ''),
    }

    logger.info(f"프로젝트 데이터 구성 완료: {code}")
    return project_data



@projects_bp.route('/api/projects/auto', methods=['POST'])
@login_required
@track_business_operation("api_project_create_auto")
def add_project_auto():
    """신규 프로젝트 자동 코드 생성 및 추가"""
    try:
        data = request.get_json()
        logger.info(f"[POST /api/projects/auto] 받은 데이터: {data}")

        # 1. 데이터 검증
        validated_data, error = _validate_project_auto_data(data)
        if error:
            return error

        # 2. 프로젝트 코드 생성
        code, df = _generate_project_code(data)
        if not code:
            return df  # 에러 응답
        data["프로젝트 코드"] = code

        # 3. Sheets Manager 초기화
        manager, sheet_id = _initialize_sheets_manager()
        if not manager:
            return sheet_id  # 에러 응답

        # 4. 기본값 설정
        _prepare_project_defaults(data)

        # 5. 행 값 배열 구성
        values = _build_row_values(data, manager)

        # 6. Google Sheets에 추가
        result = manager.append_row(sheet_id, values)

        if result:
            # 7. 응답 데이터 구성 (금액 계산)
            project_data = _build_project_response_data(code, data)

            # 8. 후처리 (캐시, 로그, 캘린더, 리드)
            finalization_data = _finalize_project_creation(code, data, project_data)

            logger.info(f"[CREATE_PROJECT] 프로젝트 생성 완료: {code}")

            return jsonify({
                "success": True,
                "project_code": code,
                "project_data": project_data,
                "lead_linked": finalization_data['lead_linked']
            })
        else:
            return jsonify({
                "success": False,
                "error": "Google Sheets 추가 실패",
                "code": "SHEETS_APPEND_FAILED"
            }), 500

    except Exception as e:
        logger.error(f"프로젝트 자동 생성 API 오류: {str(e)}")
        import traceback
        logger.error(f"스택 트레이스:\n{traceback.format_exc()}")
        return jsonify({
            "success": False,
            "error": str(e),
            "code": "INTERNAL_ERROR"
        }), 500



# 임시 디버깅 라우트 (추후 제거 예정)
@projects_bp.route('/debug-frontend')
@login_required
def debug_frontend():
    """Template Context 디버깅용 임시 라우트"""
    return render_template('debug_template.html')


# 공사 취소/재개 JSON API
@projects_bp.route('/api/project/cancel', methods=['POST'])
@login_required
@track_business_operation("api_project_cancel")
def cancel_project_api():
    """공사 취소 API - JSON 응답"""
    try:
        data = request.get_json()
        project_code = data.get('projectCode')

        if not project_code:
            return jsonify({'success': False, 'error': '프로젝트 코드가 필요합니다.'}), 400

        user_email = session.get('user', {}).get('email', '')
        user_name = session.get('user', {}).get('name', '')

        # 프로젝트 존재 여부 확인
        projects_data = get_project_records()
        project = next((p for p in projects_data if p.get('프로젝트 코드') == project_code), None)

        if not project:
            return jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404

        # 이미 취소된 프로젝트인지 확인 (띄어쓰기 무시)
        if re.search(r'공사\s*취소', project.get('수금 관련 특이사항', '')):
            return jsonify({'success': False, 'error': '이미 취소된 프로젝트입니다.'}), 400

        # Google Sheets 업데이트
        manager = get_sheets_manager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황의 사본')

        # 프로젝트 코드로 행 찾기
        row_number = manager.find_row_by_project_code(sheet_id, project_code, f'{sheet_name}!A:A')

        if not row_number:
            return jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404

        # 배치 업데이트: AG(수금 관련 특이사항), Z(수금 확인), AL(공사 확정일)
        updates = [
            {
                'range': f'{sheet_name}!AG{row_number}',  # AG: 수금 관련 특이사항
                'values': [['공사 취소']]
            },
            {
                'range': f'{sheet_name}!Z{row_number}',  # Z: 수금 확인
                'values': [['FALSE']]
            },
            {
                'range': f'{sheet_name}!AL{row_number}',  # AL: 공사 확정일
                'values': [['']]
            }
        ]

        try:
            # 배치 업데이트 실행 (Thread-Safe: batch_update_cells는 _execute_with_retry 사용)
            batch_result = manager.batch_update_cells(sheet_id, updates)

            updated_cells = batch_result.get('totalUpdatedCells', 0)
            logger.info(f"프로젝트 취소 - {updated_cells}개 셀 업데이트 완료")

            # Google Sheets 행 배경색을 진한 회색으로 변경
            try:
                success = manager.update_row_background_color(
                    spreadsheet_id=sheet_id,
                    sheet_name=sheet_name,
                    row_number=row_number,
                    color_type='dark_grey'
                )
                if success:
                    logger.info(f"행 {row_number} 배경색을 진한 회색으로 변경 완료")
                else:
                    logger.warning(f"행 {row_number} 배경색 변경 실패")
            except Exception as color_error:
                logger.error(f"배경색 변경 오류: {str(color_error)}")
                # 색상 변경 실패해도 공사 취소는 진행

            # 캐시 무효화 (invalidate_project_cache 내부에서 current_sheet_data도 무효화)
            invalidate_project_cache(project_code)

            # 감사 로그 기록
            try:
                from dashboard.utils.user_database import get_audit_repository
                audit_repo = get_audit_repository()
                audit_repo.log_action(
                    user_email=user_email,
                    action='CANCEL_PROJECT',
                    details=f'프로젝트 공사 취소: {project_code} (수금확인=FALSE, 공사확정일 초기화)',
                    project_code=project_code,
                    field_name='수금 관련 특이사항',
                    old_value=project.get('수금 관련 특이사항', '-'),
                    new_value='공사 취소',
                    ip_address=request.remote_addr
                )
            except Exception as log_error:
                logger.warning(f"감사 로그 기록 실패: {log_error}")

            # 업데이트된 프로젝트 데이터 가져오기
            updated_projects = get_project_records()
            updated_project = next((p for p in updated_projects if p.get('프로젝트 코드') == project_code), None)

            # JSON 직렬화 가능하도록 변환
            sanitized_project = sanitize_project_for_json(updated_project) if updated_project else None

            try:
                delete_project_calendar_event(project_code)
            except Exception as calendar_error:
                logger.warning(f"[CALENDAR] 이벤트 삭제 실패 ({project_code}): {calendar_error}")

            # 실시간 알림 (SocketIO)
            try:
                socketio = current_app.extensions.get('socketio')
                if socketio:
                    socketio.emit('project_cancelled', {
                        'message': f'프로젝트 공사가 취소되었습니다: {project_code}',
                        'timestamp': datetime.now().isoformat(),
                        'action': 'cancel_project',
                        'project_code': project_code,
                        'user': user_name,
                        'updated_project': sanitized_project
                    })
                else:
                    logger.warning("SocketIO 인스턴스를 찾을 수 없습니다.")
            except Exception as socket_error:
                logger.warning(f"SocketIO 알림 실패: {socket_error}")

            logger.info(f"프로젝트 취소 완료: {project_code} by {user_name}")

            return jsonify({
                'success': True,
                'message': '공사가 취소되었습니다.',
                'project_code': project_code,
                'updated_cells': updated_cells,
                'updated_project': sanitized_project
            })

        except Exception as e:
            logger.error(f"Google Sheets 업데이트 실패: {str(e)}")
            return jsonify({'success': False, 'error': '프로젝트 취소에 실패했습니다.'}), 500

    except Exception as e:
        logger.error(f"프로젝트 취소 처리 오류: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500


@projects_bp.route('/api/project/resume', methods=['POST'])
@login_required
@track_business_operation("api_project_resume")
def resume_project_api():
    """공사 재개 API - JSON 응답"""
    try:
        data = request.get_json()
        project_code = data.get('projectCode')

        if not project_code:
            return jsonify({'success': False, 'error': '프로젝트 코드가 필요합니다.'}), 400

        user_email = session.get('user', {}).get('email', '')
        user_name = session.get('user', {}).get('name', '')

        # 프로젝트 존재 여부 확인
        projects_data = get_project_records()
        project = next((p for p in projects_data if p.get('프로젝트 코드') == project_code), None)

        if not project:
            return jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404

        # 취소된 프로젝트인지 확인 (띄어쓰기 무시)
        if not re.search(r'공사\s*취소', project.get('수금 관련 특이사항', '')):
            logger.info(f"프로젝트 재개 시도 - 이미 정상 상태: {project_code}")
            return jsonify({
                'success': True,  # 이미 정상 상태이므로 성공으로 처리
                'message': '이미 정상 상태입니다. (취소되지 않은 프로젝트)',
                'project_code': project_code,
                'already_active': True
            }), 200

        # Google Sheets 업데이트
        manager = get_sheets_manager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황의 사본')

        # 프로젝트 코드로 행 찾기
        row_number = manager.find_row_by_project_code(sheet_id, project_code, f'{sheet_name}!A:A')

        if not row_number:
            return jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404

        # AG 칼럼(수금 관련 특이사항)을 빈 문자열로 초기화하고 공사 확정일을 현재 날짜로 재설정
        updates = [
            {
                'range': f'{sheet_name}!AG{row_number}',
                'values': [['']]
            },
            {
                'range': f'{sheet_name}!AL{row_number}',
                'values': [[datetime.now().strftime('%Y-%m-%d')]]
            }
        ]

        try:
            batch_result = manager.batch_update_cells(sheet_id, updates)
            updated_cells = batch_result.get('totalUpdatedCells', 0)
            logger.info(f"프로젝트 재개 - 배치 업데이트 완료 ({updated_cells}개 셀)")

            # Google Sheets 행 배경색을 흰색으로 복원
            try:
                success = manager.update_row_background_color(
                    spreadsheet_id=sheet_id,
                    sheet_name=sheet_name,
                    row_number=row_number,
                    color_type='normal'
                )
                if success:
                    logger.info(f"행 {row_number} 배경색을 흰색으로 복원 완료")
                else:
                    logger.warning(f"행 {row_number} 배경색 복원 실패")
            except Exception as color_error:
                logger.error(f"배경색 복원 오류: {str(color_error)}")
                # 색상 복원 실패해도 공사 재개는 진행

            # 캐시 무효화 (invalidate_project_cache 내부에서 current_sheet_data도 무효화)
            invalidate_project_cache(project_code)

            # 감사 로그 기록
            try:
                from dashboard.utils.user_database import get_audit_repository
                audit_repo = get_audit_repository()
                audit_repo.log_action(
                    user_email=user_email,
                    action='RESUME_PROJECT',
                    details=f'프로젝트 공사 재개: {project_code}',
                    project_code=project_code,
                    field_name='수금 관련 특이사항',
                    old_value=project.get('수금 관련 특이사항', '-'),
                    new_value='',
                    ip_address=request.remote_addr
                )
            except Exception as log_error:
                logger.warning(f"감사 로그 기록 실패: {log_error}")

            # 업데이트된 프로젝트 데이터 가져오기
            updated_projects = get_project_records()
            updated_project = next((p for p in updated_projects if p.get('프로젝트 코드') == project_code), None)

            # JSON 직렬화 가능하도록 변환
            sanitized_project = sanitize_project_for_json(updated_project) if updated_project else None

            # 실시간 알림 (SocketIO)
            try:
                socketio = current_app.extensions.get('socketio')
                if socketio:
                    socketio.emit('project_resumed', {
                        'message': f'프로젝트 공사가 재개되었습니다: {project_code}',
                        'timestamp': datetime.now().isoformat(),
                        'action': 'resume_project',
                        'project_code': project_code,
                        'user': user_name,
                        'updated_project': sanitized_project
                    })
                else:
                    logger.warning("SocketIO 인스턴스를 찾을 수 없습니다.")
            except Exception as socket_error:
                logger.warning(f"SocketIO 알림 실패: {socket_error}")

            logger.info(f"프로젝트 재개 완료: {project_code} by {user_name}")

            return jsonify({
                'success': True,
                'message': '공사가 재개되었습니다.',
                'project_code': project_code,
                'updated_cells': updated_cells,
                'updated_project': sanitized_project
            })

        except Exception as e:
            logger.error(f"Google Sheets 업데이트 실패: {str(e)}")
            return jsonify({'success': False, 'error': '프로젝트 재개에 실패했습니다.'}), 500

    except Exception as e:
        logger.error(f"프로젝트 재개 처리 오류: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)}), 500


@projects_bp.route('/api/projects/calendar/sync', methods=['POST'])
@login_required
def manual_calendar_sync():
    """수동 캘린더 동기화 API"""
    try:
        logger.info(f"[CALENDAR_SYNC] 수동 동기화 시작 - 사용자: {session.get('user', {}).get('email', 'unknown')}")

        # 전역 CalendarSyncScheduler 인스턴스에서 동기화 실행
        from dashboard.services.calendar_sync_scheduler import get_calendar_sync_scheduler
        scheduler = get_calendar_sync_scheduler()
        scheduler.run_sync_once()

        logger.info("[CALENDAR_SYNC] 수동 동기화 완료")

        return jsonify({
            'success': True,
            'message': '캘린더 동기화가 완료되었습니다.'
        })

    except Exception as e:
        logger.error(f"[CALENDAR_SYNC] 수동 동기화 실패: {e}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캘린더 동기화에 실패했습니다.'
        }), 500
