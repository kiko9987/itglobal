import logging
import os
import json
import re
import time
import threading
import uuid
from datetime import datetime
from collections import Counter, defaultdict
from typing import Optional

import pandas as pd
from flask import Blueprint, render_template, redirect, url_for, request, session, jsonify, current_app

from ..auth import login_required, get_user_role, admin_required, editor_required
from ..services.project_service import (
    can_user_edit_project,
    check_overdue_status,
    get_project_records,
    invalidate_project_cache,
    update_project_in_cache,
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
from ..api.responses import APIResponse, APIErrorCode, api_response
from ..utils.smart_cache_manager import smart_invalidate, smart_get, CacheStrategy
from ..utils.logging_config import get_logger
from ..utils.request_middleware import track_business_operation, log_external_api_call
from ..utils.error_helpers import generate_error_id

# Marshmallow 검증 스키마 임포트
from ..schemas import ProjectCreateSchema, ProjectUpdateSchema, CellMemoSchema
from ..schemas.project_schemas import (
    validate_request_data,
    format_validation_errors,
    ProjectAutoCreateSchema
)

# ─────────────────────────────────────────────────────────────
# 컬럼 인덱스 상수 (google_sheets.py column_mapping 과 반드시 일치)
# 시트 컬럼 시프트 시 여기 하나만 수정하면 됨.
# `AP: '_version'` → index 41 (2026-07 시프트로 옛 AO 에서 이동)
# ─────────────────────────────────────────────────────────────
VERSION_COL_INDEX = 41  # AP 열


def _verify_version_col_index(manager) -> None:
    """부팅·최초 호출 시 상수 vs 실제 매핑 검증. 어긋나면 조용히 로그로 알림.

    회귀 감지용 — 실제 오류는 다음 편집 시도 시 튀지만, 로그에서 조기 발견 가능.
    """
    try:
        column_mapping = manager.get_column_mapping()
        for col_letter, field_name in column_mapping.items():
            if field_name == '_version':
                if len(col_letter) == 1:
                    actual = ord(col_letter) - ord('A')
                else:
                    actual = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
                if actual != VERSION_COL_INDEX:
                    logger.error(
                        f'[VERSION_COL] 상수 불일치! VERSION_COL_INDEX={VERSION_COL_INDEX} '
                        f'실제 _version 컬럼={col_letter}(idx={actual}) — projects.py 수정 필요'
                    )
                return
        logger.warning('[VERSION_COL] column_mapping 에 _version 필드 없음 — 시트 스키마 확인 필요')
    except Exception as exc:
        logger.debug(f'[VERSION_COL] 검증 skip: {exc}')

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

def _load_projects_data(refresh):
    """
    데이터 로드 전략 (refresh 파라미터에 따라 캐시 또는 강제 새로고침)

    Args:
        refresh: True면 강제 새로고침, False면 캐시 사용

    Returns:
        DataFrame or None
    """
    logger.info(f'[API][PID:{os.getpid()}] 프로젝트 목록 요청 (refresh={refresh})')

    if refresh:
        logger.info(f'[API][PID:{os.getpid()}] 강제 새로고침 - 구글시트 로드')
        return load_data(force_refresh=True)
    else:
        df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
        logger.info(f'[API][PID:{os.getpid()}] 캐시에서 데이터 조회 - {"있음" if df is not None and not df.empty else "없음"}')
        if df is None or (df is not None and len(df) == 0):
            logger.info(f'[API][PID:{os.getpid()}] 캐시 없음 - 구글시트에서 로드')
            return load_data(force_refresh=False)
        return df


def _process_date_columns(df):
    """
    DataFrame의 날짜 컬럼들을 YYYY-MM-DD 형식으로 변환

    잘못된 날짜 형식을 감지하고 경고 로그를 출력하여
    사용자 입력 오류를 추적할 수 있도록 개선

    Args:
        df: DataFrame

    Returns:
        tuple: (DataFrame (날짜 컬럼 변환 완료), invalid_dates 리스트)
    """
    date_columns = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
    all_invalid_dates = []  # 모든 잘못된 날짜 수집

    for col in date_columns:
        if col not in df.columns:
            continue

        logger.info(f"날짜 컬럼 처리 시작: {col}")

        # 원본 데이터 샘플 확인 (최근 5개)
        recent_samples = df.tail(5)[col].tolist()
        logger.info(f"{col} 원본 데이터 샘플: {recent_samples}")

        # 잘못된 날짜 추적
        invalid_dates_in_column = []

        for idx in df.index:
            value = df.at[idx, col]

            # 빈 값은 건너뛰기
            if pd.isna(value) or value == '' or value == '-':
                df.at[idx, col] = ''
                continue

            try:
                # 명시적 날짜 파싱 시도
                parsed_date = pd.to_datetime(value, errors='raise')
                df.at[idx, col] = parsed_date.strftime('%Y-%m-%d')
            except Exception as parse_error:
                # 파싱 실패 시 기록
                invalid_dates_in_column.append({
                    'row': idx,
                    'column': col,
                    'value': str(value),
                    'error': str(parse_error)
                })
                df.at[idx, col] = ''  # 빈 문자열로 설정

        # 잘못된 날짜가 있으면 경고 로그
        if invalid_dates_in_column:
            logger.warning(
                f"⚠️ '{col}' 컬럼에서 잘못된 날짜 형식 발견: {len(invalid_dates_in_column)}개"
            )
            # 처음 3개만 로그에 출력 (너무 많으면 로그 비대화)
            for invalid in invalid_dates_in_column[:3]:
                logger.warning(
                    f"   행 {invalid['row']}: '{invalid['value']}' → 빈 문자열로 변환"
                )
            if len(invalid_dates_in_column) > 3:
                logger.warning(f"   (외 {len(invalid_dates_in_column) - 3}개)")

            all_invalid_dates.extend(invalid_dates_in_column)

        # 최종 결과 샘플
        final_samples = df.tail(5)[col].tolist()
        logger.info(f"{col} 최종 결과 샘플: {final_samples}")

    # 전체 잘못된 날짜 요약
    if all_invalid_dates:
        logger.warning(
            f"⚠️ 총 {len(all_invalid_dates)}개의 잘못된 날짜가 빈 문자열로 변환되었습니다"
        )

    return df, all_invalid_dates


def _build_projects_response(projects, invalid_dates=None):
    """
    프로젝트 목록 응답 구조 생성

    Args:
        projects: 프로젝트 dict 리스트
        invalid_dates: 잘못된 날짜 목록 (선택사항)

    Returns:
        dict: 응답 구조
    """
    response = {
        'success': True,
        'data': projects,
        'meta': {
            'total': len(projects),
            'timestamp': datetime.now().isoformat()
        }
    }

    # 잘못된 날짜 정보가 있으면 메타데이터에 추가
    if invalid_dates:
        response['meta']['date_parsing_warnings'] = {
            'count': len(invalid_dates),
            'message': f'{len(invalid_dates)}개의 날짜가 잘못된 형식으로 빈 문자열로 변환되었습니다',
            'details': [
                {
                    'row': invalid['row'],
                    'column': invalid['column'],
                    'value': invalid['value'],
                    'error': invalid['error']
                }
                for invalid in invalid_dates[:10]  # 처음 10개만 (응답 크기 제한)
            ]
        }

        # 더 많은 오류가 있으면 표시
        if len(invalid_dates) > 10:
            response['meta']['date_parsing_warnings']['has_more'] = True
            response['meta']['date_parsing_warnings']['total_hidden'] = len(invalid_dates) - 10

    return response


def _set_cache_headers(response, refresh):
    """
    응답에 캐시 헤더 설정

    Args:
        response: Flask Response 객체
        refresh: refresh 파라미터 여부

    Returns:
        Response: 헤더가 설정된 응답
    """
    if not refresh:
        # 일반 조회: 30초 동안 브라우저 캐시 허용
        response.headers['Cache-Control'] = 'private, max-age=30'
    else:
        # 강제 새로고침: 캐시 사용 안함
        response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
    return response


@projects_bp.route('/api/projects/list')
@login_required
@track_business_operation("api_projects_list")
def get_projects_list():
    """프로젝트 목록 API"""
    try:
        # 1. 파라미터 파싱
        refresh = request.args.get('refresh', 'false').lower() == 'true'

        # 2. 데이터 로드 (캐시 또는 강제 새로고침)
        df = _load_projects_data(refresh)

        # 3. 데이터 없을 때 에러 반환
        if df is None or (df is not None and len(df) == 0):
            return jsonify({
                'success': False,
                'error': '데이터를 불러올 수 없습니다.',
                'data': None,
                'meta': {
                    'timestamp': datetime.now().isoformat()
                }
            }), 500

        # 4. DataFrame 전처리 (NaN 처리 + 날짜 변환)
        df = df.fillna('')
        df, invalid_dates = _process_date_columns(df)

        # 잘못된 날짜가 있으면 로그에 기록 (이미 함수 내부에서 경고 출력됨)
        if invalid_dates:
            logger.info(f"[DATA] 잘못된 날짜 {len(invalid_dates)}개가 감지되었습니다")

        # 5. dict 변환
        projects = df.to_dict('records')

        # 5-0. 날짜 필드 후처리: Timestamp/NaT 객체를 문자열로 변환
        # pandas의 to_dict()는 datetime64 컬럼을 Timestamp 객체로 변환하므로
        # JSON 직렬화 가능한 문자열로 최종 변환
        date_columns = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
        for project in projects:
            for date_col in date_columns:
                if date_col in project:
                    value = project[date_col]
                    # Timestamp 객체인 경우 문자열로 변환
                    if pd.notna(value) and hasattr(value, 'strftime'):
                        try:
                            project[date_col] = value.strftime('%Y-%m-%d')
                        except:
                            project[date_col] = ''
                    # NaT, None, 빈 문자열은 모두 빈 문자열로 통일
                    elif pd.isna(value) or value == '':
                        project[date_col] = ''

        # 5-1. 금액 필드 정규화 (₩300,000 → 300000)
        # FORMATTED_VALUE로 읽은 금액이 통화 기호 포함되어 있으면 프론트엔드 parseFloat()가 NaN 반환
        # 마진율/순익 계산에 사용되는 비용 필드도 포함 (ProjectRowAccordion.js 참조)
        currency_fields = ['총액 1', '총액 2', '계약금', '중도금', '잔금', '미수금', '제품대', '도급비', '자재비', '기타비']
        for project in projects:
            for field_name in currency_fields:
                if field_name in project:
                    currency_value = project[field_name]
                    # 0 값은 유효한 금액이므로 None과 빈 문자열만 필터링
                    if currency_value is not None and currency_value != '':
                        parsed_amount = safe_parse_currency(currency_value)
                        # 정수로 표현 가능하면 정수로, 아니면 float로
                        if parsed_amount == int(parsed_amount):
                            project[field_name] = str(int(parsed_amount))
                        else:
                            project[field_name] = str(parsed_amount)
                    else:
                        project[field_name] = ''

        # Google Sheets에서 이미 계산된 순익, 마진율 값을 그대로 사용
        # 재계산하지 않음 - 시트의 수식이 정확한 단일 소스(Single Source of Truth)

        # 참고: viewer 권한은 "읽기 전용" (신입 교육용)
        # - 전체 프로젝트 조회 가능 (검색, 학습 목적)
        # - 수정/삭제는 불가 (각 API 엔드포인트에서 권한 체크)

        # 6. 응답 생성 및 캐시 헤더 설정 (날짜 파싱 오류 정보 포함)
        response_data = _build_projects_response(projects, invalid_dates)
        response = jsonify(response_data)
        response = _set_cache_headers(response, refresh)

        return response

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 프로젝트 목록 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '프로젝트 목록을 불러올 수 없습니다.',
            'error_id': error_id,
            'data': None,
            'meta': {
                'timestamp': datetime.now().isoformat()
            }
        }), 500


@projects_bp.route('/api/inflow-options')
def get_inflow_options():
    """시트 D열(유입 구분)의 데이터 유효성 검사 드롭다운 값을 반환.

    사용자가 시트에서 드롭다운 목록을 수정하면 자동으로 반영됨.
    """
    try:
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500

        sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황')
        manager = get_sheets_manager()
        meta = manager.get_column_dropdown_values(sheet_id, sheet_name, 'D', scan_rows=200)
        return jsonify({
            'success': True,
            'options': meta.get('values', []),
            'debug': {
                'source_row': meta.get('source_row'),
                'condition_type': meta.get('condition_type'),
                'raw': meta.get('raw'),
                'sheet_name': sheet_name,
            },
        })
    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 유입 구분 옵션 조회 실패: {e}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '유입 구분 옵션 조회 중 오류가 발생했습니다.',
            'error_id': error_id,
            'options': [],
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
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 프로젝트 코드 생성 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'error': '프로젝트 코드 생성 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.',
            'error_id': error_id
        }), 500


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
        end_col='AP',  # AP 컬럼까지 읽기 (_version 포함, 2026-07 컬럼 시프트)
        value_render_option='FORMULA'  # 수식을 그대로 가져옴 (계산 필드 보존)
    )

    # 현재 값을 리스트로 확장 (42개 컬럼, AP까지)
    while len(current_values) < 42:
        current_values.append('')

    return row_number, current_values, None


def _check_optimistic_lock_update(manager, sheet_id, sheet_name, project_code, row_number, current_values, expected_version):
    """Optimistic Lock 버전 검증

    Returns:
        tuple: (new_version, error_response)
        - success: (int, None)
        - failure: (None, JsonResponse with 409)
    """
    current_version = current_values[VERSION_COL_INDEX]  # AP 컬럼 — _version

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
            end_col='AP',
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
        end_col='AP',  # AP 컬럼까지 조회 (_version 포함, 2026-07 시프트)
        value_render_option='FORMATTED_VALUE'  # 계산된 값 가져오기
    )

    # fresh_values를 40개 컬럼으로 패딩 (AN 컬럼까지, 마지막 열이 비어있을 경우 대비)
    while len(fresh_values) < 40:
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

        # 금액 필드 정규화 (₩300,000 → 300000)
        # FORMATTED_VALUE로 읽은 금액이 통화 기호 포함되어 있으면 프론트엔드 parseFloat()가 NaN 반환
        # 마진율/순익 계산에 사용되는 비용 필드도 포함 (ProjectRowAccordion.js 참조)
        currency_fields = ['총액 1', '총액 2', '계약금', '중도금', '잔금', '미수금', '제품대', '도급비', '자재비', '기타비']
        for field_name in currency_fields:
            if field_name in updated_project:
                currency_value = updated_project[field_name]
                # 0 값은 유효한 금액이므로 None과 빈 문자열만 필터링
                if currency_value is not None and currency_value != '':
                    parsed_amount = safe_parse_currency(currency_value)
                    # 정수로 표현 가능하면 정수로, 아니면 float로
                    if parsed_amount == int(parsed_amount):
                        updated_project[field_name] = str(int(parsed_amount))
                    else:
                        updated_project[field_name] = str(parsed_amount)
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
@editor_required
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

        sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황')
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
        current_values[VERSION_COL_INDEX] = str(new_version)  # AP 컬럼 — _version

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

        # 8. 시트 write 를 큐로 위임 (2026-07-09 write-behind)
        range_name = f'{sheet_name}!A{row_number}:AP{row_number}'
        from ..services.sheet_write_queue import enqueue as _q_enqueue
        _q_enqueue('project_update_sheet', {
            'sheet_id': sheet_id,
            'sheet_name': sheet_name,
            'row_number': row_number,
            'range_name': range_name,
            'current_values': current_values,
            'field_changes': field_changes,
            'project_code': final_project_code,
        }, meta={'user_email': session.get('user', {}).get('email', 'unknown')})
        logger.info(
            f"[PUT] 시트 write 큐 위임 완료: {final_project_code}, {len(field_changes)}개 필드"
        )

        # 10. 감사 로그 배치 기록 (동기 — 이력 즉시 확정)
        _record_update_audit_logs(field_changes, project_code)

        # 11. updated_project 계산 (시트 재조회 없이 로컬 계산)
        updated_project = _build_updated_project_from_values(current_values, field_to_index)

        # 12. 캐시 부분 갱신 시도 (Google Sheets API 호출 없이 즉시 반영)
        # - 프로젝트 코드 변경 없고 updated_project 있으면 캐시된 DataFrame에서 해당 row만 in-place 교체
        # - 실패 시(캐시 미스·타입 mismatch·못 찾음) 기존 전체 무효화 fallback
        code_unchanged = (project_code == final_project_code)
        cache_updated = False
        if code_unchanged and updated_project:
            cache_updated = update_project_in_cache(final_project_code, updated_project)
        if not cache_updated:
            invalidate_project_cache(final_project_code)
        logger.info(f"[PUT] 프로젝트 업데이트 완료: {final_project_code} (cache_partial={cache_updated})")

        # 13. 캘린더 + 슬랙 편집 알림 백그라운드 처리 (사용자 응답 대기 안 함, ~1s 절약)
        # 재시작 순간 진행 중이던 알림은 유실될 수 있지만 시트 저장은 이미 완료 상태 → 데이터 유실 아님.
        # 프로젝트 표준 패턴 (blueprints/slack_bot.py 등에 15+ 곳 사용) 그대로.
        import threading as _th

        def _bg_notifications():
            logger.info(f"[BG/START] {final_project_code} 알림 백그라운드 시작")
            try:
                _update_calendar_if_needed(updated_project, project_code, final_project_code)
            except Exception as exc:
                logger.warning(f"[BG/CALENDAR] {final_project_code} 처리 오류: {exc}")
            try:
                from ..services.project_slack_notifier import notify_project_field_changes
                notify_project_field_changes(
                    final_project_code, field_changes, latest_data=updated_project,
                )
            except Exception as exc:
                logger.warning(f"[BG/SLACK] {final_project_code} 알림 처리 오류: {exc}")
            # 2026-07-16: 총액 1 / 부가세 / 사업자명 변경 시 계산서 카드 스레드에도 알림
            try:
                from ..services.project_slack_notifier import notify_invoice_card_amount_change
                notify_invoice_card_amount_change(final_project_code, field_changes)
            except Exception as exc:
                logger.warning(f"[BG/INVOICE] {final_project_code} 알림 처리 오류: {exc}")
            logger.info(f"[BG/DONE] {final_project_code} 알림 처리 완료")

        _th.Thread(target=_bg_notifications, daemon=True).start()

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
        error_id = generate_error_id()
        logger.error(f"[{error_id}] [PUT] 프로젝트 업데이트 오류: {project_code}, {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '프로젝트 업데이트 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.',
            'error_id': error_id
        }), 500


# 서비스 워커 라우트 — 2026-07-08 무력화.
# 이전 동적 sw.js가 Vite 해시 파일 변경 시 stale HTML을 서빙해 매니저 브라우저
# ERR_FAILED 유발 (배포 blocker). 프로덕션 config.SERVICE_WORKER_ENABLED=False로
# 신규 등록은 이미 차단됨. 기존 브라우저에 남은 옛 SW는 이 unregister 스텁을
# 받으면 자기 자신을 해제 + 캐시 삭제.
@projects_bp.route('/sw.js')
def service_worker():
    """Service Worker unregister 스텁 (기존 등록분 자동 정리용)."""
    from flask import Response

    sw_content = """/**
 * Service Worker unregister stub (2026-07-08).
 * 기존에 등록된 SW를 자동 해제하고 캐시를 삭제한다.
 * fetch 훅 없음 — 모든 요청은 브라우저가 직접 처리.
 */
self.addEventListener('install', () => {
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil((async () => {
        try {
            const cacheNames = await caches.keys();
            await Promise.all(cacheNames.map((name) => caches.delete(name)));
        } catch (err) { /* ignore */ }
        try {
            await self.registration.unregister();
        } catch (err) { /* ignore */ }
        try {
            const clients = await self.clients.matchAll({ type: 'window' });
            clients.forEach((client) => {
                if (client.url && 'navigate' in client) {
                    client.navigate(client.url);
                }
            });
        } catch (err) { /* ignore */ }
    })());
});
"""

    resp = Response(sw_content, mimetype='application/javascript')
    # 브라우저가 옛 sw.js를 캐시해 이 스텁을 못 받는 사고 방지
    resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    return resp

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
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캐시 상태 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'error': '캐시 상태 조회 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500

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
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 프로젝트 통계 로드 실패: {e}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '통계 정보를 불러올 수 없습니다.',
            'error_id': error_id
        }), 500


# ===== 헬퍼 함수 1: 기본 검증 및 컬럼 매핑 =====
def _validate_single_memo_request(data):
    """
    단일 메모 저장 요청 기본 검증
    - 필수 필드 검증
    - 필드명 유효성 검증
    - 컬럼 매핑 조회

    Returns:
        tuple: (project_code, field_name, memo, column) or (None, error_response_tuple)
    """
    project_code = data.get('project_code')
    field_name = data.get('field_name')
    memo = data.get('memo')

    # 1. 기본 필수 필드 검증
    if not project_code or not field_name:
        return None, (jsonify({
            'success': False,
            'message': '프로젝트 코드와 필드명이 필요합니다.'
        }), 400)

    # 2. 필드명 유효성 검증
    if field_name not in MEMOABLE_FIELDS:
        return None, (jsonify({
            'success': False,
            'message': ERROR_MESSAGES['memo']['invalid_field']
        }), 400)

    # 3. 컬럼 매핑
    column = PAYMENT_FIELD_TO_COLUMN.get(field_name)
    if not column:
        return None, (jsonify({
            'success': False,
            'message': f'컬럼 매핑을 찾을 수 없습니다: {field_name}'
        }), 500)

    return project_code, field_name, memo, column


# ===== 헬퍼 함수 2: Manager 초기화 및 행 조회 =====
def _find_memo_project_row(project_code):
    """
    Google Sheets Manager 초기화 및 프로젝트 행 조회

    Returns:
        tuple: (manager, sheet_id, sheet_name, row_number) or (None, error_response_tuple)
    """
    # Google Sheets Manager 및 설정 초기화
    manager = get_sheets_manager()
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황')

    # 프로젝트 행 번호 조회
    row_number = manager.find_row_by_project_code(sheet_id, project_code, f'{sheet_name}!A:A')

    if not row_number:
        return None, (jsonify({
            'success': False,
            'message': ERROR_MESSAGES['memo']['project_not_found']
        }), 404)

    return manager, sheet_id, sheet_name, row_number


# ===== 헬퍼 함수 3: Marshmallow 스키마 검증 =====
def _validate_memo_schema(memo, row_number, column, project_code, field_name):
    """
    Marshmallow 스키마로 메모 상세 검증 (메모가 있는 경우에만)

    Returns:
        tuple: (validated_data, None) or (None, error_response_tuple)
    """
    # 메모가 없으면 검증 건너뛰기 (삭제 요청)
    if not memo:
        return None, None

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
        return None, (jsonify({
            'success': False,
            'message': ERROR_MESSAGES['memo']['validation_failed'],
            'validation_errors': error_messages
        }), 400)

    return validated_data, None


# ===== 헬퍼 함수 4: 메모 저장, 캐시 무효화, 감사 로그 =====
def _save_memo_and_audit(manager, sheet_id, sheet_name, cell_address, memo, project_code, field_name):
    """
    메모 저장 + 캐시 무효화 + 감사 로그 기록

    Returns:
        tuple: (success, response_data or error_response)
    """
    logger.info(f"[FIELD_MEMO] 셀 메모 저장 시작: {project_code} / {cell_address}")

    # 기존 메모 값 조회 (감사 로그용)
    old_memo = manager.get_cell_note(sheet_id, sheet_name, cell_address) or ''

    # 구글 시트 셀 메모에 저장
    success = manager.update_cell_note(
        sheet_id,
        sheet_name,
        cell_address,
        memo if memo else None  # 빈 메모는 삭제
    )

    if not success:
        return False, (jsonify({
            'success': False,
            'message': '메모 저장에 실패했습니다.'
        }), 500)

    # 셀 메모 캐시만 무효화 (선택적 캐시 무효화 - MED-4 피드백)
    notes_cache_key = f"cell_notes_{sheet_id}"
    invalidated_notes = smart_invalidate(notes_cache_key)

    logger.info(f"[FIELD_MEMO][PID:{os.getpid()}] 셀 메모 캐시 무효화 완료:")
    logger.info(f"  - {notes_cache_key}: {invalidated_notes}개 항목")

    # 감사 로그 기록
    try:
        from dashboard.utils.user_database import get_audit_repository
        audit_repo = get_audit_repository()
        user_email = session.get('user', {}).get('email', 'unknown')

        action_desc = "삭제" if not memo else "저장"
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

    return True, {
        'success': True,
        'message': '메모가 저장되었습니다.',
        'data': {
            'project_code': project_code,
            'field_name': field_name,
            'memo': memo,
            'cell_address': cell_address
        }
    }


# ===== 메인 함수: save_field_memo (리팩토링됨) =====
@projects_bp.route('/api/projects/field-memo', methods=['POST'])
@editor_required
@track_business_operation("api_field_memo_save")
def save_field_memo():
    """
    필드 메모 저장 API (구글 시트 셀 메모로 저장)

    Request Body:
        {
            "project_code": "G0001",
            "field_name": "계약금" | "중도금" | "잔금",
            "memo": "입금일: 2025-01-10\\n입금자: 홍길동" | null
        }
    """
    try:
        data = request.get_json()

        # 1. 기본 검증 및 컬럼 매핑
        result = _validate_single_memo_request(data)
        if result[0] is None:
            return result[1]
        project_code, field_name, memo, column = result

        # 2. Manager 초기화 및 행 조회
        result = _find_memo_project_row(project_code)
        if result[0] is None:
            return result[1]
        manager, sheet_id, sheet_name, row_number = result

        # 3. Marshmallow 스키마 검증 (메모가 있는 경우)
        validated_data, error = _validate_memo_schema(memo, row_number, column, project_code, field_name)
        if error:
            return error

        # 4. 셀 주소 구성
        cell_address = f"{column}{row_number}"

        # 5. 메모 저장, 캐시 무효화, 감사 로그
        success, response = _save_memo_and_audit(
            manager, sheet_id, sheet_name, cell_address, memo, project_code, field_name
        )

        if success:
            return jsonify(response)
        else:
            return response

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 필드 메모 저장 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'message': '메모 저장 중 오류가 발생했습니다.',
            'error_id': error_id
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
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황')

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

            # 빈값 저장 방어 (2026-07-22 G3823-SJ 사고):
            # 프론트가 편집 안 한 필드까지 빈값으로 보내거나 매니저 실수로 memo=''
            # 저장 요청 → 기존 note 삭제 사고. 명시적 삭제 API 필요 시 별도 endpoint.
            # 여기선 빈값이면 skip (기존 note 유지) + 성공 처리.
            if not memo or not memo.strip():
                results.append({
                    'field_name': field_name,
                    'success': True,
                    'message': '빈값 — 기존 메모 유지 (skip)',
                    'cell_address': cell_address,
                    'memo_index': len(results),
                })
                continue

            # Batch 업데이트 대상에 추가
            cell_notes_to_update[cell_address] = memo

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
        batch_result = manager._execute_with_retry(
            lambda: manager.service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body=body
            ),
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
@editor_required
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

        # 6. 내부용 필드 제거 (프론트엔드 계약 준수)
        for result_item in results:
            result_item.pop('memo_index', None)  # 내부 인덱스 제거
            result_item.pop('cell_address', None)  # 내부 셀 주소 제거

        # 7. 결과 반환
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
        error_id = generate_error_id()
        logger.error(f"[{error_id}] [BATCH_MEMO] 일괄 저장 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'message': '일괄 메모 저장 중 오류가 발생했습니다.',
            'error_id': error_id
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
def _prepare_project_defaults(data, row_number):
    """
    프로젝트 데이터에 기본값 설정 (금액, 수식, 날짜, VAT 등)

    Args:
        data: 원본 데이터 딕셔너리 (in-place 수정됨)
        row_number: Google Sheets에 삽입될 행 번호 (수식에 사용)
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

    # 수식 필드 자동 삽입 (행 번호 동적 삽입 - 성능 최적화 및 가독성 개선)
    # 총액2는 끝자리 1/9원만 보정 (깔끔한 금액 유지)
    #
    # 컬럼 매핑 (2026-07 AO=Lead No 신설로 컬럼 한 칸씩 시프트 반영):
    #   R=총액1, S=부가세, T=총액2, U=계약금, V=중도금, W=잔금, X=미수금
    #   AB=제품대, AC=도급비, AD=자재비, AE=기타비, AF=순익, AG=마진율
    data['총액 2'] = f'=IF(S{row_number}=TRUE, R{row_number}+FLOOR(R{row_number}*0.1,1) + IF(MOD(R{row_number}+FLOOR(R{row_number}*0.1,1), 10)=1, -1, IF(MOD(R{row_number}+FLOOR(R{row_number}*0.1,1), 10)=9, 1, 0)), R{row_number})'
    data['미수금'] = f'=IF(ABS($T{row_number}-$U{row_number}-$V{row_number}-$W{row_number})<2, 0, $T{row_number}-$U{row_number}-$V{row_number}-$W{row_number})'
    data['순익'] = f'=R{row_number}-(AB{row_number}+AC{row_number}+AD{row_number}+AE{row_number})'
    data['마진율'] = f'=IF(OR(R{row_number}=0, AF{row_number}=0), 0, AF{row_number}/R{row_number})'

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

    # 텍스트 필드 빈값 → '-' 채우기 (시트 가독성 + 일관성)
    # 수식/금액/날짜/Boolean/시스템 필드는 제외
    _excluded_fields = {
        '프로젝트 코드', '사업자', '담당자',  # 필수 필드 (검증으로 보장)
        '공사 시작', '공사 종료', '공사 확정',  # 날짜
        '총액 1', '총액 2', '계약금', '중도금', '잔금', '미수금',  # 금액
        '제품대', '도급비', '자재비', '기타비', '순익', '마진율',  # 금액/수식
        '부가세', '수금 확인',  # Boolean
        '계산서', '_version', 'lead_no',  # 시스템 필드
        '수금 날짜',  # 빈값 의도적
    }
    text_fields_to_dash = [
        '사업자명', '발주처 담당자', '발주처 연락처', '발주처 이메일',
        '현장 주소', '공사 구분', '기계 분류', '브랜드',
        '공사 내용', '도급 구분', '시공자', '유입 구분',
        '수금 관련 특이사항', '계약금 입금자명', '중도금 입금자명', '잔금 입금자명',
        '견적서 및 계약서 폴더 경로',
    ]
    for field in text_fields_to_dash:
        if field in _excluded_fields:
            continue
        value = data.get(field)
        if value is None or (isinstance(value, str) and value.strip() == ''):
            data[field] = '-'


# ===== 헬퍼 함수 5: 행 값 배열 구성 =====
def _build_row_values(data, manager, row_number):
    """
    데이터를 Google Sheets 행 배열로 변환 (41 컬럼 A~AO)

    Args:
        row_number: Google Sheets에 삽입될 행 번호 (기본값 수식 생성용)

    Returns:
        list: 42개 요소의 값 배열 (A~AP)
    """
    column_mapping = manager.get_column_mapping()
    values = [''] * 42  # AP 컬럼까지 (Lead No + _version, 2026-07 shift: 41→42)

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
    # 부가세 계산: 절사(FLOOR) 방식 + 끝자리 1/9원 보정 (실무 표준, 행 번호 동적 삽입)
    # 2026-05-22 시프트: 모든 letter가 한 칸씩 뒤로 (E열 사업자명 추가)
    #   총액1: Q→R, 부가세: R→S, 총액2: S→T, 계약금: T→U, 중도금: U→V, 잔금: V→W, 미수금: W→X
    #   제품대: AA→AB, 도급비: AB→AC, 자재비: AC→AD, 기타비: AD→AE, 순익: AE→AF, 마진율: AF→AG, _version: AN→AO
    formula_fields = {
        'T': data.get('총액 2', f'=IF(S{row_number}=TRUE, R{row_number}+FLOOR(R{row_number}*0.1,1) + IF(MOD(R{row_number}+FLOOR(R{row_number}*0.1,1), 10)=1, -1, IF(MOD(R{row_number}+FLOOR(R{row_number}*0.1,1), 10)=9, 1, 0)), R{row_number})'),
        # 2026-07-16 시트 수식 반영: X = T - (U+V+W), 반올림 오차(2원 미만)는 0 처리
        #   X > 0 = 미납, X < 0 = 초과입금. 이전 (U+V+W)-T 는 부호 반대 오류.
        'X': data.get('미수금', f'=IF(ABS($T{row_number}-$U{row_number}-$V{row_number}-$W{row_number})<2, 0, $T{row_number}-$U{row_number}-$V{row_number}-$W{row_number})'),
        'AF': data.get('순익', f'=R{row_number}-(AB{row_number}+AC{row_number}+AD{row_number}+AE{row_number})'),
        'AG': data.get('마진율', f'=IF(OR(R{row_number}=0, AF{row_number}=0), 0, AF{row_number}/R{row_number})'),
        'AO': data.get('Lead No', ''),  # 리드 연결 (2026-07 신규)
        'AP': '0'  # _version 초기값 (낙관적 잠금용) — 옛 AO에서 이동
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

    # 4. 리드 연동 (Lead No가 모달에서 선택됐거나 lead_no로 전달된 경우)
    lead_no = data.get('Lead No') or data.get('lead_no')
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

    # 5. 공사 확정 슬랙 알림 (#공사_확정, 별도 봇)
    try:
        from ..services.project_slack_notifier import send_project_created_notification
        send_project_created_notification(data, code)
    except Exception as slack_error:
        logger.warning(f"[PROJECT/SLACK] 알림 발송 중 오류 (프로젝트 생성은 성공): {slack_error}")

    return {'lead_linked': bool(lead_no)}


# ===== 헬퍼 함수 7: 응답 데이터 구성 (금액 계산 포함) =====
def _build_project_response_data(code, data):
    """
    프로젝트 생성 응답 데이터 구성 (금액 계산 포함)

    Returns:
        dict: 프로젝트 데이터
    """
    from datetime import datetime
    from dashboard.services.project_service import _calculate_total2

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

    # 총액 2 계산: _calculate_total2() 재사용 (FLOOR + 끝자리 1/9원 보정)
    vat_included = data.get('부가세', 'FALSE')
    total_2 = int(_calculate_total2(float(total_1), vat_included == 'TRUE'))

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
        '유입 구분': data.get('유입 구분', ''),
        '담당자': data.get('담당자', ''),
        '발주처 연락처': data.get('발주처 연락처', ''),
        '발주처 이메일': data.get('발주처 이메일', ''),
        '현장 주소': data.get('현장 주소', ''),
        '시공자': data.get('시공자', ''),
        '발주처 담당자': data.get('발주처 담당자', ''),
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



def _dedup_hash(biz: str, owner: str, addr: str, start: str, amount: str) -> str:
    """dedup 키용 해시. Redis 저장·조회 시 사용."""
    import hashlib as _hl
    raw = f'{biz}|{owner}|{addr}|{start}|{amount}'
    return _hl.md5(raw.encode('utf-8')).hexdigest()[:16]


def _find_recent_duplicate(data: dict, minutes: int = 5) -> Optional[str]:
    """최근 오늘자 동일 데이터로 이미 등록된 프로젝트 코드가 있으면 반환.

    (사업자·담당자·주소·공사 시작·총액 1) 모두 같은 프로젝트가 오늘 확정 등록됐으면
    duplicate 로 판단. Idempotency key 유실 시 안전망.

    2026-07-10 강화: Redis 즉시 캐시 우선 조회 → load_data 캐시 지연 문제 회피
      (G3852/G3853 관측: 첫 요청 응답 후 8초 만에 재요청 왔지만 load_data 캐시
      아직 갱신 안 됐음 → dedup 실패).

    성능: 3800+ 행 dict 순회 대신 pandas mask 로 오늘자만 먼저 필터 (수 ms).
    """
    biz = str(data.get('사업자', '') or '').strip()
    owner = str(data.get('담당자', '') or '').strip()
    addr = str(data.get('현장 주소', '') or '').strip()
    start = str(data.get('공사 시작', '') or '').strip()[:10]
    amount_in = str(data.get('총액 1', '') or '').replace(',', '').replace('₩', '').strip()
    if not (biz and owner and addr and start):
        return None

    # Layer A: Redis 즉시 조회 (load_data 캐시 지연 우회)
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        dedup_key = f'project_create_dedup:{_dedup_hash(biz, owner, addr, start, amount_in)}'
        cached_code = rc.get(dedup_key)
        if cached_code:
            code_str = cached_code if isinstance(cached_code, str) else cached_code.decode()
            logger.info(f'[CREATE_PROJECT/dedup:redis] Redis 히트 → 기존 코드 {code_str}')
            return code_str
        else:
            # 2026-07-10 관측성 강화 — 조회 miss 도 남긴다. 첫 요청이면 정상, 재요청인데 miss 면
            # 저장이 안 됐거나 hash 필드가 달라진 것 → 다음 사고 시 근본 원인 짚기 위함.
            logger.info(
                f'[CREATE_PROJECT/dedup:redis] Redis miss — key={dedup_key[-16:]} '
                f'(biz={biz[:20]}, owner={owner}, addr_len={len(addr)}, start={start}, amount={amount_in})'
            )
    except Exception as exc:
        logger.warning(f'[CREATE_PROJECT/dedup:redis] Redis 조회 실패 (계속 진행): {exc}', exc_info=True)

    try:
        from dashboard.services.project_service import load_data
        df = load_data()
        if df is None or df.empty:
            return None

        today = datetime.now().strftime('%Y-%m-%d')
        # 1단계: 오늘 확정된 것만 필터 (수천 → 수십 행)
        confirmed_str = df['공사 확정'].astype(str).str[:10]
        today_mask = confirmed_str == today
        candidates = df[today_mask]
        if candidates.empty:
            return None

        # 2단계: 오늘자 소수 행만 정확 매칭
        for _, r in candidates.iterrows():
            r_biz = str(r.get('사업자', '') or '').strip()
            r_owner = str(r.get('담당자', '') or '').strip()
            r_addr = str(r.get('현장 주소', '') or '').strip()
            r_start = str(r.get('공사 시작', '') or '').strip()[:10]
            r_amount = str(r.get('총액 1', '') or '').replace(',', '').replace('₩', '').strip()
            if (r_biz == biz and r_owner == owner and r_addr == addr
                    and r_start == start and r_amount == amount_in):
                return str(r.get('프로젝트 코드', '') or '').strip()
    except Exception as exc:
        logger.warning(f'[CREATE_PROJECT/dedup] 스캔 실패: {exc}')
    return None


@projects_bp.route('/api/projects/auto', methods=['POST'])
@editor_required
@track_business_operation("api_project_create_auto")
def add_project_auto():
    """신규 프로젝트 자동 코드 생성 및 추가.

    2026-07-09 중복 등록 방지 2단계:
      1. X-Idempotency-Key 헤더 → Redis 캐시 hit 시 이전 응답 그대로 반환
      2. 데이터 기반 안전망 (최근 5분 내 동일 사업자·담당자·주소·시작일·금액 존재 시)
    """
    idem_key = request.headers.get('X-Idempotency-Key', '').strip()
    idem_redis_key = f'project_create_idem:{idem_key}' if idem_key else ''

    # === Layer 1: Idempotency key 캐시 조회 ===
    if idem_key:
        try:
            from dashboard.utils.redis_client import get_redis_client
            rc = get_redis_client().redis
            cached = rc.get(idem_redis_key)
            if cached:
                logger.info(
                    f'[CREATE_PROJECT/idem] 중복 요청 감지 (idem={idem_key[:8]}…) — 이전 응답 반환'
                )
                import json as _json
                return jsonify(_json.loads(cached))
            else:
                # 2026-07-10 fingerprint 기반 idem-key 도입 후 정상적으로 매 요청마다 나올 로그.
                # 재요청 시엔 히트해야 정상 → 재발 시 이 로그가 실마리.
                logger.info(f'[CREATE_PROJECT/idem] miss (idem={idem_key[:8]}…)')
        except Exception as exc:
            logger.warning(
                f'[CREATE_PROJECT/idem] Redis 조회 실패 (계속 진행): {exc}', exc_info=True,
            )

    try:
        data = request.get_json()
        logger.info(f"[POST /api/projects/auto] 받은 데이터: {data}")

        # === Layer 2: 데이터 기반 dedup 안전망 ===
        dup_code = _find_recent_duplicate(data)
        if dup_code:
            logger.warning(
                f'[CREATE_PROJECT/dedup] 최근 5분 내 동일 데이터 감지 → 기존 코드 반환: {dup_code}'
            )
            # dedup 이라도 project_data 찾아서 반환 → 프론트가 전체 시트 재로드(15초) 안 하도록
            dup_project = None
            try:
                from ..services.project_service import get_project_records
                records = get_project_records() or []
                for r in records:
                    if r.get('프로젝트 코드') == dup_code:
                        dup_project = r
                        break
            except Exception as exc:
                logger.debug(f'[CREATE_PROJECT/dedup] project_data 조회 실패 (무시): {exc}')
            resp_body = {
                'success': True,
                'project_code': dup_code,
                'project_data': dup_project,
                'lead_linked': False,
                'deduped': True,
            }
            # idempotency 캐시에도 저장 (같은 key 로 다시 오면 즉시 재응답)
            if idem_key:
                try:
                    import json as _json
                    from dashboard.utils.redis_client import get_redis_client
                    get_redis_client().redis.set(
                        idem_redis_key, _json.dumps(resp_body), ex=600,
                    )
                except Exception:
                    pass
            return jsonify(resp_body)

        # 1. 데이터 검증
        validated_data, error = _validate_project_auto_data(data)
        if error:
            return error

        # 2. 프로젝트 코드 생성 (DB 시퀀스 — 시트와 무관하게 원자적)
        code, df = _generate_project_code(data)
        if not code:
            return df  # 에러 응답
        data["프로젝트 코드"] = code

        # 3. Sheets Manager 초기화
        manager, sheet_id = _initialize_sheets_manager()
        if not manager:
            return sheet_id  # 에러 응답

        # 4. 기본값·행 값 구성 (시트 append 없이도 가능)
        next_row = len(df) + 2
        _prepare_project_defaults(data, next_row)
        values = _build_row_values(data, manager, next_row)

        # 5. 응답 데이터 구성 (계산 필드 포함) — 프론트가 즉시 렌더할 데이터
        project_data = _build_project_response_data(code, data)

        # 6. 응답 본문 준비 (시트 append 전에도 완결)
        resp_body = {
            "success": True,
            "project_code": code,
            "project_data": project_data,
            "lead_linked": bool(data.get('Lead No') or data.get('lead_no')),
        }
        # Idempotency 캐시 저장 (10분 TTL) — 같은 idem_key 재요청 시 즉시 응답
        if idem_key:
            try:
                import json as _json
                from dashboard.utils.redis_client import get_redis_client
                _ok = get_redis_client().redis.set(
                    idem_redis_key, _json.dumps(resp_body), ex=600,
                )
                if _ok:
                    logger.info(
                        f'[CREATE_PROJECT/idem:save] 저장 성공 code={code} idem={idem_key[:8]}…'
                    )
                else:
                    logger.warning(
                        f'[CREATE_PROJECT/idem:save] Redis set 반환값 falsy — idem={idem_key[:8]}…'
                    )
            except Exception as exc:
                logger.warning(
                    f'[CREATE_PROJECT/idem:save] 캐시 저장 실패 code={code}: {exc}',
                    exc_info=True,
                )

        # 7. 백그라운드: 시트 append + 후처리 (감사·캘린더·리드·슬랙)
        # 응답을 먼저 반환하고 Google Sheets 지연이 UX에 영향 없도록 분리 (2026-07-09).
        # 사용자 세션·IP 등 request context 는 여기서 캡처해 background 로 전달.
        _bg_user_email = session.get('user', {}).get('email', 'unknown')
        _bg_ip = request.remote_addr

        def _write_behind():
            # 스레드별 GoogleSheetsManager 필수 — request thread 의 manager 를
            # 그대로 쓰면 heap corruption (2026-07-09 크래시 사고 참조).
            from ..services.project_service import get_sheets_manager as _gsm
            _mgr = _gsm()
            import time as _time
            last_exc = None
            for attempt in range(3):
                try:
                    result = _mgr.append_row(sheet_id, values)
                    if result:
                        _finalize_project_creation_bg(
                            code, data, project_data,
                            user_email=_bg_user_email, ip_address=_bg_ip,
                        )
                        logger.info(f"[CREATE_PROJECT/BG] 시트 append + 후처리 완료: {code}")
                        return
                    else:
                        raise RuntimeError('append_row returned falsy')
                except Exception as exc:
                    last_exc = exc
                    wait = 2.0 * (attempt + 1)
                    logger.warning(
                        f"[CREATE_PROJECT/BG] 시트 append 실패 재시도 {attempt+1}/3 "
                        f"({code}, {wait}s 후): {exc}"
                    )
                    _time.sleep(wait)
            # 3회 재시도 모두 실패 → 관리자 슬랙 DM (sheet_write_queue 의 데드레터 알림 재사용)
            logger.error(
                f"[CREATE_PROJECT/BG] 3회 재시도 모두 실패 — {code} 시트 반영 안 됨. "
                f"관리자 수동 확인 필요: {last_exc}",
                exc_info=True,
            )
            # 관측성: sheet_write_queue 의 FAILED_KEY 에 등록 → /admin/queue-status 에 노출
            # (전체 큐 이관은 리팩토링 리스크 커서 실패 op 만 큐에 등록)
            try:
                from ..services.sheet_write_queue import (
                    _notify_admin_deadletter, FAILED_KEY,
                )
                from dashboard.utils.redis_client import get_redis_client
                import json as _json_q, uuid as _uuid_q, time as _time_q
                failed_op = {
                    'op_id': str(_uuid_q.uuid4()),
                    'op_type': 'project_create_append',
                    'payload': {'tag': code, 'project_code': code},
                    'meta': {'source': 'add_project_auto'},
                    'created_at': _time_q.time(),
                    'attempts': 3,
                    'last_error': str(last_exc)[:500] if last_exc else 'unknown',
                }
                get_redis_client().redis.rpush(
                    FAILED_KEY,
                    _json_q.dumps(failed_op, ensure_ascii=False),
                )
                _notify_admin_deadletter(
                    failed_op,
                    last_exc or Exception('알 수 없는 시트 append 실패'),
                )
            except Exception as notify_exc:
                logger.warning(f'[CREATE_PROJECT/BG] 관리자 알림 실패 (무시): {notify_exc}')
            # Zombie idempotency 캐시 방어 (2026-07-09)
            # bg thread 실패했는데 idem 캐시엔 '성공' 응답이 남아있으면 같은 idem_key 재요청 시
            # '이미 등록됨' 으로 오답 반환하지만 시트엔 실제로 없음. 캐시 삭제해 재시도 가능하게.
            if idem_key:
                try:
                    from dashboard.utils.redis_client import get_redis_client
                    get_redis_client().redis.delete(idem_redis_key)
                    logger.info(f'[CREATE_PROJECT/BG] Zombie 방어: idem 캐시 삭제 ({code})')
                except Exception as del_exc:
                    logger.warning(f'[CREATE_PROJECT/BG] idem 캐시 삭제 실패: {del_exc}')

        threading.Thread(target=_write_behind, daemon=True).start()

        # dedup Redis 캐시 즉시 저장 (2026-07-10 G3852/G3853/R3854/R3855 대응)
        # load_data 캐시 갱신 지연으로 dedup 실패하는 케이스 방지 — 8초 이내 재요청 즉시 감지.
        # 관측성: 저장 성공/필드 부족/예외 모두 명시적으로 로그 → 재발 시 원인 즉시 짚음.
        try:
            from dashboard.utils.redis_client import get_redis_client
            _biz = str(data.get('사업자', '') or '').strip()
            _owner = str(data.get('담당자', '') or '').strip()
            _addr = str(data.get('현장 주소', '') or '').strip()
            _start = str(data.get('공사 시작', '') or '').strip()[:10]
            _amount = str(data.get('총액 1', '') or '').replace(',', '').replace('₩', '').strip()
            if _biz and _owner and _addr and _start:
                _dkey = f'project_create_dedup:{_dedup_hash(_biz, _owner, _addr, _start, _amount)}'
                _ok = get_redis_client().redis.set(_dkey, code, ex=300)  # 5분 TTL
                if _ok:
                    logger.info(
                        f'[CREATE_PROJECT/dedup:save] 저장 성공 code={code} key={_dkey[-16:]}'
                    )
                else:
                    logger.warning(
                        f'[CREATE_PROJECT/dedup:save] Redis set 반환값 falsy — key={_dkey[-16:]}'
                    )
            else:
                logger.warning(
                    f'[CREATE_PROJECT/dedup:save] 필드 부족으로 dedup 저장 스킵 — code={code} '
                    f'biz="{_biz}" owner="{_owner}" addr_len={len(_addr)} start="{_start}"'
                )
        except Exception as _exc:
            logger.warning(
                f'[CREATE_PROJECT/dedup:save] Redis 저장 실패 code={code}: {_exc}',
                exc_info=True,
            )

        logger.info(f"[CREATE_PROJECT] 응답 반환 (BG write 진행 중): {code}")
        return jsonify(resp_body)

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 프로젝트 자동 생성 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            "success": False,
            "error": "프로젝트 생성 중 오류가 발생했습니다.",
            "error_id": error_id,
            "code": "INTERNAL_ERROR"
        }), 500


def _finalize_project_creation_bg(code, data, project_data, user_email='unknown', ip_address=None):
    """`_finalize_project_creation` 의 백그라운드용 변형.

    request context 없이 호출 가능하도록 user_email·ip_address 를 인자로 받음.
    """
    # 1. 캐시 무효화
    invalidate_project_cache(code)

    # 2. 감사 로그 기록
    try:
        from ..utils.user_database import get_audit_repository
        audit_repo = get_audit_repository()
        audit_repo.log_action(
            user_email=user_email,
            action='CREATE_PROJECT',
            details=f'새 프로젝트 등록: {code}',
            project_code=code,
            field_name='전체',
            old_value='-',
            new_value='새 프로젝트 생성',
            ip_address=ip_address,
        )
    except Exception as log_error:
        logger.warning(f"감사 로그 기록 실패: {log_error}")

    # 3. 캘린더 이벤트 생성
    try:
        create_project_calendar_event(project_data)
    except Exception as calendar_error:
        logger.debug(f"[CALENDAR] 이벤트 생성 실패 ({code}): {calendar_error}")

    # 4. 리드 연동
    lead_no = data.get('Lead No') or data.get('lead_no')
    if lead_no:
        try:
            from ..services.lead_service import update_lead_status
            update_lead_status(lead_no, '공사 확정')
        except Exception as lead_error:
            logger.warning(f"[LEAD_LINKED] 리드 상태 업데이트 실패: {lead_error}")

    # 5. 슬랙 알림
    try:
        from ..services.project_slack_notifier import send_project_created_notification
        send_project_created_notification(data, code)
    except Exception as slack_error:
        logger.warning(f"[PROJECT/SLACK] 알림 발송 실패: {slack_error}")



# 임시 디버깅 라우트 (추후 제거 예정)
@projects_bp.route('/debug-frontend')
@login_required
def debug_frontend():
    """Template Context 디버깅용 임시 라우트"""
    return render_template('debug_template.html')


# ============================================
# 공통 헬퍼 함수 (Shared Helpers)
# ============================================

def _validate_status_change_request(data):
    """공사 상태 변경 요청 기본 검증 및 사용자 정보 추출

    Returns:
        tuple: (project_code, user_email, user_name) 또는 (None, response, status_code)
    """
    project_code = data.get('projectCode')

    if not project_code:
        return None, jsonify({'success': False, 'error': '프로젝트 코드가 필요합니다.'}), 400

    user_email = session.get('user', {}).get('email', '')
    user_name = session.get('user', {}).get('name', '')

    return project_code, user_email, user_name


def _find_project_and_row(project_code):
    """프로젝트 조회 및 Google Sheets 행 번호 찾기

    Returns:
        tuple: (project, manager, sheet_id, sheet_name, row_number) 또는 (None, response, status_code)
    """
    # 프로젝트 존재 여부 확인
    projects_data = get_project_records()
    project = next((p for p in projects_data if p.get('프로젝트 코드') == project_code), None)

    if not project:
        return None, jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404

    # Google Sheets Manager 초기화
    manager = get_sheets_manager()
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황')

    # 프로젝트 코드로 행 찾기
    row_number = manager.find_row_by_project_code(sheet_id, project_code, f'{sheet_name}!A:A')

    if not row_number:
        return None, jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404

    return project, manager, sheet_id, sheet_name, row_number


def _update_project_background_color(manager, sheet_id, sheet_name, row_number, color_type, action_name):
    """프로젝트 행 배경색 업데이트 (에러 처리 포함)

    Args:
        color_type: 'dark_grey' (취소) 또는 'normal' (재개)
        action_name: '취소' 또는 '재개' (로그용)
    """
    try:
        success = manager.update_row_background_color(
            spreadsheet_id=sheet_id,
            sheet_name=sheet_name,
            row_number=row_number,
            color_type=color_type
        )
        if success:
            color_desc = '진한 회색' if color_type == 'dark_grey' else '흰색'
            logger.info(f"행 {row_number} 배경색을 {color_desc}으로 변경 완료")
        else:
            logger.warning(f"행 {row_number} 배경색 변경 실패")
    except Exception as color_error:
        logger.error(f"배경색 변경 오류: {str(color_error)}")
        # 색상 변경 실패해도 공사 상태 변경은 진행


def _log_project_status_change(project_code, project, user_email, action, field_name, old_value, new_value):
    """프로젝트 상태 변경 감사 로그 기록"""
    try:
        from dashboard.utils.user_database import get_audit_repository

        # action에 따른 한국어 메시지 생성
        if action == 'CANCEL_PROJECT':
            details = f'프로젝트 공사 취소: {project_code} (수금확인=FALSE, 공사확정일 초기화)'
        elif action == 'RESUME_PROJECT':
            details = f'프로젝트 공사 재개: {project_code}'
        else:
            details = f'프로젝트 상태 변경: {project_code} (action={action})'

        audit_repo = get_audit_repository()
        audit_repo.log_action(
            user_email=user_email,
            action=action,
            details=details,
            project_code=project_code,
            field_name=field_name,
            old_value=old_value,
            new_value=new_value,
            ip_address=request.remote_addr
        )
    except Exception as log_error:
        logger.warning(f"감사 로그 기록 실패: {log_error}")


def _get_and_sanitize_updated_project(project_code):
    """업데이트된 프로젝트 데이터 가져오기 및 JSON 직렬화 가능하게 변환

    Returns:
        dict or None: 직렬화된 프로젝트 데이터
    """
    updated_projects = get_project_records()
    updated_project = next((p for p in updated_projects if p.get('프로젝트 코드') == project_code), None)
    return sanitize_project_for_json(updated_project) if updated_project else None


def _emit_project_status_change(event_name, message, project_code, user_name, sanitized_project, sender_email=''):
    """프로젝트 상태 변경 실시간 알림 (SocketIO)

    Args:
        event_name: 'project_cancelled' 또는 'project_resumed'
        sender_email: 이 액션을 트리거한 사용자 이메일 — 클라이언트가 자기 자신 이벤트를
                      무시하는 용도 (자기 PC에선 이미 로컬 반영돼 있어 재렌더 필요 없음)
    """
    try:
        socketio = current_app.extensions.get('socketio')
        if socketio:
            # 프론트엔드 계약: event_name → action 매핑
            action_map = {
                'project_cancelled': 'cancel_project',
                'project_resumed': 'resume_project'
            }
            socketio.emit(event_name, {
                'message': message,
                'timestamp': datetime.now().isoformat(),
                'action': action_map.get(event_name, event_name),
                'project_code': project_code,
                'user': user_name,
                'sender_email': sender_email,
                'updated_project': sanitized_project
            })
        else:
            logger.warning("SocketIO 인스턴스를 찾을 수 없습니다.")
    except Exception as socket_error:
        logger.warning(f"SocketIO 알림 실패: {socket_error}")


# ============================================
# cancel_project_api() 전용 헬퍼
# ============================================

def _check_already_cancelled(project):
    """프로젝트가 이미 취소되었는지 확인

    Returns:
        tuple: (is_cancelled: bool, response or None, status_code or None)
    """
    if re.search(r'공사\s*취소', project.get('수금 관련 특이사항', '')):
        return True, jsonify({'success': False, 'error': '이미 취소된 프로젝트입니다.'}), 400
    return False, None, None


def _prepare_cancel_updates(sheet_name, row_number):
    """공사 취소 배치 업데이트 준비

    Returns:
        list: Batch update requests (AH: 수금 관련 특이사항, AA: 수금 확인, AM: 공사 확정)

    Note: 2026-07 컬럼 시프트 반영 (AG→AH, Z→AA, AL→AM).
          옛 매핑은 마진율·수금 날짜·폴더 경로에 잘못 write하는 데이터 파괴 버그였음.
    """
    return [
        {
            'range': f'{sheet_name}!AH{row_number}',  # AH: 수금 관련 특이사항
            'values': [['공사 취소']]
        },
        {
            'range': f'{sheet_name}!AA{row_number}',  # AA: 수금 확인
            'values': [['FALSE']]
        },
        {
            'range': f'{sheet_name}!AM{row_number}',  # AM: 공사 확정
            'values': [['']]
        }
    ]


def _delete_calendar_event(project_code):
    """프로젝트 캘린더 이벤트 삭제 (에러 처리 포함)"""
    try:
        delete_project_calendar_event(project_code)
    except Exception as calendar_error:
        logger.warning(f"[CALENDAR] 이벤트 삭제 실패 ({project_code}): {calendar_error}")


# ============================================
# resume_project_api() 전용 헬퍼
# ============================================

def _check_already_active(project, project_code):
    """프로젝트가 이미 활성 상태(취소되지 않음)인지 확인

    Returns:
        tuple: (is_active: bool, response or None, status_code or None)
    """
    if not re.search(r'공사\s*취소', project.get('수금 관련 특이사항', '')):
        logger.info(f"프로젝트 재개 시도 - 이미 정상 상태: {project_code}")
        return True, jsonify({
            'success': True,
            'message': '이미 정상 상태입니다. (취소되지 않은 프로젝트)',
            'project_code': project_code,
            'already_active': True
        }), 200
    return False, None, None


def _prepare_resume_updates(sheet_name, row_number):
    """공사 재개 배치 업데이트 준비

    Returns:
        list: Batch update requests (AH: '', AM: 현재 날짜)

    Note: 2026-07 컬럼 시프트 반영 (AG→AH, AL→AM).
          옛 매핑은 마진율에 빈 값을, 폴더 경로에 날짜를 잘못 write하는 데이터 파괴 버그였음.
    """
    return [
        {
            'range': f'{sheet_name}!AH{row_number}',  # AH: 수금 관련 특이사항
            'values': [['']]
        },
        {
            'range': f'{sheet_name}!AM{row_number}',  # AM: 공사 확정
            'values': [[datetime.now().strftime('%Y-%m-%d')]]
        }
    ]


# ============================================
# Sheet write-behind 큐 관리자 endpoint (2026-07-09)
# ============================================
@projects_bp.route('/admin/queue-status')
@admin_required
def admin_queue_status_page():
    """관리자용 시트 쓰기 큐 상태 HTML 페이지."""
    return render_template('admin_queue_status.html')


@projects_bp.route('/api/admin/sheet-write-queue', methods=['GET'])
@admin_required
def sheet_write_queue_status():
    """큐 상태 조회 (관리자용) — pending/processing/failed 개수."""
    from ..services.sheet_write_queue import get_stats, peek_failed
    stats = get_stats()
    failed_peek = peek_failed(limit=10)
    return jsonify({
        'success': True,
        'stats': stats,
        'failed_recent': failed_peek,
    })


@projects_bp.route('/api/admin/sheet-write-queue/retry/<op_id>', methods=['POST'])
@admin_required
def sheet_write_queue_retry(op_id):
    """실패 데드레터 큐 op 재시도 (관리자용)."""
    from ..services.sheet_write_queue import retry_failed
    ok = retry_failed(op_id)
    return jsonify({'success': ok})


@projects_bp.route('/api/admin/slack-health', methods=['GET'])
@admin_required
def slack_health_stats():
    """슬랙 API 최근 5분 실패 통계 (관리자용)."""
    from ..utils.slack_health import get_stats as _slack_stats
    return jsonify({'success': True, 'stats': _slack_stats()})


@projects_bp.route('/admin/data-integrity')
@admin_required
def admin_data_integrity_page():
    """캐시 ↔ 시트 정합성 감사 페이지."""
    return render_template('admin_data_integrity.html')


@projects_bp.route('/api/admin/cache-metrics', methods=['GET'])
@admin_required
def admin_cache_metrics():
    """캐시 hit/miss 통계 + key별 top miss (관리자용)."""
    try:
        from ..utils.smart_cache_manager import get_smart_cache
        sc = get_smart_cache()
        return jsonify({'success': True, 'stats': sc.get_metrics()})
    except Exception as exc:
        logger.error(f'[CACHE_METRICS] 조회 실패: {exc}', exc_info=True)
        return jsonify({'success': False, 'error': str(exc)}), 500


@projects_bp.route('/api/admin/data-integrity/check', methods=['POST'])
@admin_required
def admin_data_integrity_check():
    """캐시된 프로젝트 목록 vs Google Sheets 실제 read 결과 diff.

    write-behind 큐 도입 후 워커 실패로 캐시-시트 불일치가 조용히 누적될 수 있어
    관리자가 원할 때 강제로 재검증.

    비교 대상: 프로젝트 코드 A열 존재 유무 + 매출·수금·상태 3개 필드값.
    """
    from ..services.project_service import get_project_records
    try:
        cached = get_project_records(force_refresh=False) or []
        fresh = get_project_records(force_refresh=True) or []
    except Exception as exc:
        logger.error(f'[DATA_INTEGRITY] 데이터 fetch 실패: {exc}', exc_info=True)
        return jsonify({'success': False, 'error': str(exc)}), 500

    def _by_code(records):
        return {r.get('프로젝트 코드', ''): r for r in records if r.get('프로젝트 코드')}
    cache_map = _by_code(cached)
    sheet_map = _by_code(fresh)

    only_in_cache = sorted(set(cache_map) - set(sheet_map))
    only_in_sheet = sorted(set(sheet_map) - set(cache_map))
    common = set(cache_map) & set(sheet_map)

    mismatched = []
    compare_fields = ['매출', '수금 확인', '상태']
    for code in sorted(common):
        c = cache_map[code]
        s = sheet_map[code]
        diffs = {}
        for f in compare_fields:
            cv = c.get(f, '')
            sv = s.get(f, '')
            if cv != sv:
                diffs[f] = {'cache': str(cv)[:100], 'sheet': str(sv)[:100]}
        if diffs:
            mismatched.append({'code': code, 'diffs': diffs})

    return jsonify({
        'success': True,
        'summary': {
            'cache_total': len(cache_map),
            'sheet_total': len(sheet_map),
            'only_in_cache': len(only_in_cache),
            'only_in_sheet': len(only_in_sheet),
            'mismatched': len(mismatched),
        },
        'only_in_cache': only_in_cache[:50],
        'only_in_sheet': only_in_sheet[:50],
        'mismatched': mismatched[:50],
    })


# ============================================
# Sheet write-behind 큐 핸들러 (2026-07-09)
# ============================================
from ..services.sheet_write_queue import register as _q_register


@_q_register('project_cancel_sheet')
def _handle_cancel_sheet(payload: dict) -> None:
    """공사 취소 시트 write — batch_update_cells + 행 배경색 dark_grey."""
    from ..services.project_service import get_sheets_manager
    manager = get_sheets_manager()
    sheet_id = payload['sheet_id']
    sheet_name = payload['sheet_name']
    row_number = payload['row_number']
    project_code = payload.get('project_code', '')
    updates = _prepare_cancel_updates(sheet_name, row_number)
    manager.batch_update_cells(sheet_id, updates)
    logger.info(f'[QUEUE/project_cancel] 시트 write 완료: {project_code}')
    # 배경색 (실패해도 큐 재시도 안 하도록 내부 try)
    try:
        _update_project_background_color(
            manager, sheet_id, sheet_name, row_number, 'dark_grey', '취소'
        )
    except Exception as exc:
        logger.warning(f'[QUEUE/project_cancel/color] {project_code}: {exc}')


@_q_register('sheet_batch_write')
def _handle_sheet_batch_write(payload: dict) -> None:
    """범용 batch_update_cells. payload: {sheet_id, updates, tag}.

    슬랙 액션(취소·재개·편집), A/S 시트 갱신 등 다양한 위치에서 재사용.
    """
    from ..services.project_service import get_sheets_manager
    manager = get_sheets_manager()
    manager.batch_update_cells(payload['sheet_id'], payload['updates'])
    tag = payload.get('tag', '')
    logger.info(f'[QUEUE/sheet_batch_write] 완료 (tag={tag})')


@_q_register('sheet_bg_color')
def _handle_sheet_bg_color(payload: dict) -> None:
    """범용 행 배경색 갱신. payload: {sheet_id, sheet_name, row_number, color_type, tag}."""
    from ..services.project_service import get_sheets_manager
    manager = get_sheets_manager()
    manager.update_row_background_color(
        spreadsheet_id=payload['sheet_id'],
        sheet_name=payload['sheet_name'],
        row_number=payload['row_number'],
        color_type=payload['color_type'],
    )
    tag = payload.get('tag', '')
    logger.info(f'[QUEUE/sheet_bg_color] 완료 (tag={tag})')


@_q_register('project_update_sheet')
def _handle_project_update_sheet(payload: dict) -> None:
    """편집 시트 write — update_row + payment comments.

    payload:
      sheet_id, sheet_name, row_number, range_name, current_values, field_changes,
      project_code (로깅용)
    """
    from ..services.project_service import get_sheets_manager
    manager = get_sheets_manager()
    sheet_id = payload['sheet_id']
    sheet_name = payload['sheet_name']
    row_number = payload['row_number']
    current_values = payload['current_values']
    range_name = payload['range_name']
    field_changes = payload.get('field_changes', [])
    code = payload.get('project_code', '')
    manager.update_row(sheet_id, row_number, current_values, range_name)
    logger.info(f'[QUEUE/project_update] 시트 write 완료: {code}')
    try:
        _process_payment_field_comments(manager, sheet_id, sheet_name, row_number, field_changes)
    except Exception as exc:
        logger.warning(f'[QUEUE/project_update/comments] {code}: {exc}')


def _build_updated_project_from_values(current_values: list, field_to_index: dict) -> dict:
    """`_fetch_and_calculate_updated_project` 의 로컬 변형 — 시트 재조회 없이 계산.

    이미 update 를 적용한 current_values 를 그대로 사용해 계산 필드까지 재계산.
    write-behind 시 응답 지연 없이 프론트가 사용할 데이터 반환.
    """
    import pandas as pd
    from dashboard.services.project_service import (
        _safe_parse_amount, _parse_vat_flag,
        _calculate_total2, _calculate_outstanding_amount,
        _calculate_net_profit, _calculate_margin_rate,
    )

    updated_project = {}
    for field_name, index in field_to_index.items():
        if index < len(current_values):
            value = current_values[index]
            if value is None or (isinstance(value, float) and pd.isna(value)):
                updated_project[field_name] = ''
            else:
                updated_project[field_name] = str(value)

    # 날짜 필드 정규화
    #   2026-07-10 회귀 fix — 편집 저장 후 공사 시작/종료가 사라지던 버그.
    #     current_values 가 FORMULA render 로 읽혀서 날짜 셀이 시리얼 넘버 (예: 46236) 로 반환됨.
    #     pd.to_datetime('46236', errors='coerce') 는 NaT → 빈 문자열 → 프론트 카드 재렌더 시 사라짐.
    #     해결: convert_excel_serial_to_date() 로 시리얼·문자열 모두 정확히 파싱.
    #     원본 값을 유지하는 게 사라지는 것보다 안전 (파싱 실패 시 빈 문자열 대신 원본 유지).
    date_fields = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
    for field_name in date_fields:
        if field_name in updated_project:
            date_value = updated_project[field_name]
            if date_value:
                converted = convert_excel_serial_to_date(date_value)
                # 성공: 'YYYY-MM-DD' — 그대로 저장.
                # 실패: 원본 문자열 반환 → 프론트가 새로고침 없이도 렌더 가능한 값 유지.
                updated_project[field_name] = converted

    # 금액 필드 정규화 (통화 기호 제거)
    currency_fields = ['총액 1', '총액 2', '계약금', '중도금', '잔금', '미수금',
                       '제품대', '도급비', '자재비', '기타비']
    for field_name in currency_fields:
        if field_name in updated_project:
            cv = updated_project[field_name]
            if cv is not None and cv != '':
                pa = safe_parse_currency(cv)
                if pa == int(pa):
                    updated_project[field_name] = str(int(pa))
                else:
                    updated_project[field_name] = str(pa)
            else:
                updated_project[field_name] = ''

    # 계산 필드 재계산
    row_series = pd.Series(updated_project)
    total1 = _safe_parse_amount(updated_project.get('총액 1', 0))
    vat_flag = _parse_vat_flag(updated_project.get('부가세'))
    total2 = _calculate_total2(total1, vat_flag)
    outstanding = _calculate_outstanding_amount(row_series)
    net_profit = _calculate_net_profit(row_series)
    margin_rate = _calculate_margin_rate(row_series, net_profit)
    updated_project['총액 2'] = str(int(total2)) if total2 != 0 else ''
    updated_project['미수금'] = str(int(outstanding)) if outstanding != 0 else ''
    updated_project['순익'] = str(int(net_profit)) if net_profit != 0 else ''
    updated_project['마진율'] = str(round(margin_rate, 1)) if margin_rate != 0 else '0'
    return updated_project


@_q_register('project_resume_sheet')
def _handle_resume_sheet(payload: dict) -> None:
    """공사 재개 시트 write — batch_update_cells + 행 배경색 normal."""
    from ..services.project_service import get_sheets_manager
    manager = get_sheets_manager()
    sheet_id = payload['sheet_id']
    sheet_name = payload['sheet_name']
    row_number = payload['row_number']
    project_code = payload.get('project_code', '')
    updates = _prepare_resume_updates(sheet_name, row_number)
    manager.batch_update_cells(sheet_id, updates)
    logger.info(f'[QUEUE/project_resume] 시트 write 완료: {project_code}')
    try:
        _update_project_background_color(
            manager, sheet_id, sheet_name, row_number, 'normal', '재개'
        )
    except Exception as exc:
        logger.warning(f'[QUEUE/project_resume/color] {project_code}: {exc}')


# ============================================
# 메인 함수 (Refactored Main Functions)
# ============================================

@projects_bp.route('/api/project/cancel', methods=['POST'])
@editor_required
@track_business_operation("api_project_cancel")
def cancel_project_api():
    """공사 취소 API - JSON 응답"""
    try:
        data = request.get_json()

        # 1. 기본 검증 및 사용자 정보
        result = _validate_status_change_request(data)
        if result[0] is None:
            return result[1], result[2]
        project_code, user_email, user_name = result

        # 2. 프로젝트 조회 및 행 번호 찾기
        result = _find_project_and_row(project_code)
        if result[0] is None:
            return result[1], result[2]
        project, manager, sheet_id, sheet_name, row_number = result

        # 3. 이미 취소된 프로젝트인지 확인
        is_cancelled, response, status_code = _check_already_cancelled(project)
        if is_cancelled:
            return response, status_code

        # 4. 시트 write 를 큐로 위임 + 캐시 즉시 갱신 (2026-07-09 write-behind)
        try:
            # 캐시 즉시 반영 — 사용자 응답 및 다음 read 에 즉시 노출
            cache_updated = update_project_in_cache(project_code, {
                '수금 관련 특이사항': '공사 취소',
                '수금 확인': False,
                '공사 확정': '',
            })
            if not cache_updated:
                invalidate_project_cache(project_code)

            # 시트 write 는 큐로 (Google 지연과 무관하게 응답 반환)
            from ..services.sheet_write_queue import enqueue as _q_enqueue
            _q_enqueue('project_cancel_sheet', {
                'sheet_id': sheet_id,
                'sheet_name': sheet_name,
                'row_number': row_number,
                'project_code': project_code,
            }, meta={'user_email': user_email})
            updated_cells = 3  # AH/AA/AM 3셀 예상값 (응답 호환용)

            # 6. 감사 로그 기록 (동기 유지 — 감사 이력 즉시 확정)
            _log_project_status_change(
                project_code=project_code,
                project=project,
                user_email=user_email,
                action='CANCEL_PROJECT',
                field_name='수금 관련 특이사항',
                old_value=project.get('수금 관련 특이사항', '-'),
                new_value='공사 취소'
            )

            # 7. 업데이트된 프로젝트 데이터 구성 — 시트 재조회 대신 방금 write한 값을 로컬 반영
            # (2026-07-07): 옛 로직은 _get_and_sanitize_updated_project → get_project_records() 로
            # 시트 재조회했으나 Google Sheets eventual consistency 지연으로 옛 값 반환 확률 있음.
            # 그 옛 값이 socket.io project_cancelled payload로 나가서 클라이언트가 취소 상태를
            # 옛 상태로 롤백하는 UX 버그. batch_update_cells 성공했다는 건 write가 확정됐다는
            # 뜻이므로 그 값을 그대로 쓰는 게 안전.
            updated_project = dict(project)
            updated_project['수금 관련 특이사항'] = '공사 취소'
            updated_project['수금 확인'] = False
            updated_project['공사 확정'] = ''
            sanitized_project = sanitize_project_for_json(updated_project)

            # 8. SocketIO emit — 메인 스레드에서 응답 반환 전 (SIGSEGV 방지)
            # 2026-07-07: 이전엔 백그라운드 스레드에서 emit했으나 반복적 취소·재개에서
            # Flask 프로세스가 SIGSEGV로 크래시 (stderr에 traceback 없음, NSSM 자동 재시작).
            # 원인 추정: async_mode='threading' 하드코딩된 SocketIO에서 백그라운드 스레드
            # emit 시 race condition. 메인 스레드에서 실행 시 크래시 확률 크게 감소.
            # 응답 시간 몇 ms 늘어남 (허용 범위).
            try:
                _emit_project_status_change(
                    event_name='project_cancelled',
                    message=f'프로젝트 공사가 취소되었습니다: {project_code}',
                    project_code=project_code,
                    user_name=user_name,
                    sanitized_project=sanitized_project,
                    sender_email=user_email
                )
            except Exception as exc:
                logger.warning(f"[SOCKETIO] {project_code} 알림 오류: {exc}")

            # 9. 배경색은 큐 핸들러에서 함께 처리. 캘린더 삭제만 백그라운드.
            import threading as _th

            def _bg_side_effects():
                try:
                    _delete_calendar_event(project_code)
                except Exception as exc:
                    logger.debug(f"[BG/CALENDAR] {project_code} 캘린더 삭제 오류: {exc}")

            _th.Thread(target=_bg_side_effects, daemon=True).start()

            logger.info(f"프로젝트 취소 완료: {project_code} by {user_name}")

            # 11. 성공 응답 반환
            return jsonify({
                'success': True,
                'message': '공사가 취소되었습니다.',
                'project_code': project_code,
                'updated_cells': updated_cells,
                'updated_project': sanitized_project
            })

        except Exception as e:
            error_id = generate_error_id()
            logger.error(f"[{error_id}] Google Sheets 업데이트 실패: {str(e)}", exc_info=True)
            return jsonify({
                'success': False,
                'error': '프로젝트 취소에 실패했습니다.',
                'error_id': error_id
            }), 500

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 프로젝트 취소 처리 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '프로젝트 취소 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.',
            'error_id': error_id
        }), 500


@projects_bp.route('/api/project/resume', methods=['POST'])
@editor_required
@track_business_operation("api_project_resume")
def resume_project_api():
    """공사 재개 API - JSON 응답"""
    try:
        data = request.get_json()

        # 1. 기본 검증 및 사용자 정보
        result = _validate_status_change_request(data)
        if result[0] is None:
            return result[1], result[2]
        project_code, user_email, user_name = result

        # 2. 프로젝트 조회 및 행 번호 찾기
        result = _find_project_and_row(project_code)
        if result[0] is None:
            return result[1], result[2]
        project, manager, sheet_id, sheet_name, row_number = result

        # 3. 이미 활성 상태인지 확인 (취소되지 않음)
        is_active, response, status_code = _check_already_active(project, project_code)
        if is_active:
            return response, status_code

        # 4. 시트 write 를 큐로 위임 + 캐시 즉시 갱신 (2026-07-09 write-behind)
        try:
            today_str = datetime.now().strftime('%Y-%m-%d')
            cache_updated = update_project_in_cache(project_code, {
                '수금 관련 특이사항': '',
                '공사 확정': today_str,
            })
            if not cache_updated:
                invalidate_project_cache(project_code)

            from ..services.sheet_write_queue import enqueue as _q_enqueue
            _q_enqueue('project_resume_sheet', {
                'sheet_id': sheet_id,
                'sheet_name': sheet_name,
                'row_number': row_number,
                'project_code': project_code,
            }, meta={'user_email': user_email})
            updated_cells = 2  # AH/AM 2셀 예상값 (응답 호환용)

            # 6. 감사 로그 기록 (동기 유지)
            _log_project_status_change(
                project_code=project_code,
                project=project,
                user_email=user_email,
                action='RESUME_PROJECT',
                field_name='수금 관련 특이사항',
                old_value=project.get('수금 관련 특이사항', '-'),
                new_value=''
            )

            # 7. 업데이트된 프로젝트 데이터 구성 — 시트 재조회 대신 방금 write한 값을 로컬 반영
            # (2026-07-07): 취소 API와 동일한 이유. Google Sheets read consistency 지연 우회.
            updated_project = dict(project)
            updated_project['수금 관련 특이사항'] = ''
            updated_project['공사 확정'] = datetime.now().strftime('%Y-%m-%d')
            sanitized_project = sanitize_project_for_json(updated_project)

            # 8. SocketIO emit — 메인 스레드에서 (SIGSEGV 방지, 취소 API와 대칭)
            try:
                _emit_project_status_change(
                    event_name='project_resumed',
                    message=f'프로젝트 공사가 재개되었습니다: {project_code}',
                    project_code=project_code,
                    user_name=user_name,
                    sanitized_project=sanitized_project,
                    sender_email=user_email
                )
            except Exception as exc:
                logger.warning(f"[SOCKETIO] {project_code} 알림 오류: {exc}")

            # 9. 배경색은 큐 핸들러에서 함께 처리됨.

            logger.info(f"프로젝트 재개 완료: {project_code} by {user_name}")

            # 10. 성공 응답 반환
            return jsonify({
                'success': True,
                'message': '공사가 재개되었습니다.',
                'project_code': project_code,
                'updated_cells': updated_cells,
                'updated_project': sanitized_project
            })

        except Exception as e:
            error_id = generate_error_id()
            logger.error(f"[{error_id}] Google Sheets 업데이트 실패: {str(e)}", exc_info=True)
            return jsonify({
                'success': False,
                'error': '프로젝트 재개에 실패했습니다.',
                'error_id': error_id
            }), 500

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 프로젝트 재개 처리 오류: {e}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '프로젝트 재개 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.',
            'error_id': error_id
        }), 500




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
