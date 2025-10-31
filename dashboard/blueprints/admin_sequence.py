"""
프로젝트 시퀀스 관리 Admin API
DB 시퀀스와 Google Sheets 최댓값 동기화 기능 제공
"""

from flask import Blueprint, jsonify, request
from ..auth import admin_required
from ..utils.user_database import get_user_database
from ..services.project_service import load_data, _extract_number
from ..utils.logging_config import get_logger
from ..utils.error_helpers import generate_error_id

logger = get_logger(__name__)

sequence_bp = Blueprint('admin_sequence', __name__, url_prefix='/admin')

# 별도 페이지는 사용하지 않음 - 모달 방식으로 대체됨
# @sequence_bp.route('/sequence')
# @admin_required
# def sequence_management_page():
#     """
#     프로젝트 시퀀스 관리 페이지 (관리자 전용)
#     """
#     return render_template('sequence_management.html')


@sequence_bp.route('/api/sequence/status', methods=['GET'])
@admin_required
def get_sequence_status():
    """
    프로젝트 시퀀스 현재 상태 조회

    Returns:
        {
            "success": true,
            "db_sequence": 2833,
            "sheet_max": 2833,
            "gap": 0,
            "last_updated": "2025-10-31T15:00:00"
        }
    """
    try:
        db = get_user_database()

        # DB 시퀀스 조회
        sequence_info = db.get_sequence_status("GLOBAL")

        # Google Sheets 최댓값 조회
        df = load_data()
        sheet_max = 0
        if '프로젝트 코드' in df.columns:
            numbers = []
            for code in df['프로젝트 코드'].astype(str):
                number = _extract_number(code)
                if number is not None:
                    numbers.append(number)
            sheet_max = max(numbers) if numbers else 0

        db_current = sequence_info.get('current_number', 0)
        gap = db_current - sheet_max

        return jsonify({
            'success': True,
            'db_sequence': db_current,
            'sheet_max': sheet_max,
            'gap': gap,
            'last_updated': sequence_info.get('last_updated'),
            'message': f'DB 시퀀스: {db_current}, Sheet 최댓값: {sheet_max}, 차이: {gap}'
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 시퀀스 상태 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '시퀀스 상태 조회 중 오류가 발생했습니다',
            'error_id': error_id
        }), 500


@sequence_bp.route('/api/calendar/mapping/status', methods=['GET'])
@admin_required
def get_calendar_mapping_status():
    """
    캘린더 매핑 상태 조회

    Returns:
        {
            "success": true,
            "total_mappings": 150,
            "orphaned_mappings": 3,
            "orphaned_codes": ["G2833-SH", "G2800-XX", ...]
        }
    """
    try:
        from dashboard.utils.user_database import get_calendar_event_repository

        db = get_user_database()
        repo = get_calendar_event_repository()

        # Google Sheets의 모든 프로젝트 코드 조회
        df = load_data()
        sheet_codes = set()
        if '프로젝트 코드' in df.columns:
            sheet_codes = set(df['프로젝트 코드'].astype(str).tolist())

        # DB의 모든 캘린더 매핑 조회
        import sqlite3
        conn = sqlite3.connect(db.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT project_code FROM project_calendar_events")
        all_mappings = cursor.fetchall()
        conn.close()

        # 고아 매핑 찾기 (Sheets에 없는 매핑)
        orphaned = []
        for (code,) in all_mappings:
            if code not in sheet_codes:
                orphaned.append(code)

        return jsonify({
            'success': True,
            'total_mappings': len(all_mappings),
            'orphaned_mappings': len(orphaned),
            'orphaned_codes': orphaned[:10],  # 최대 10개만 표시
            'message': f'총 {len(all_mappings)}개 매핑 중 {len(orphaned)}개 고아 매핑 발견'
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캘린더 매핑 상태 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캘린더 매핑 상태 조회 중 오류가 발생했습니다',
            'error_id': error_id
        }), 500


@sequence_bp.route('/api/calendar/mapping/cleanup', methods=['POST'])
@admin_required
def cleanup_calendar_mappings():
    """
    고아 캘린더 매핑 정리 (Sheets에 없는 매핑 삭제)

    Request Body:
        {
            "confirm": true
        }

    Returns:
        {
            "success": true,
            "deleted_count": 3,
            "deleted_codes": ["G2833-SH", ...]
        }
    """
    try:
        data = request.get_json()

        if not data or not data.get('confirm'):
            return jsonify({
                'success': False,
                'error': '확인(confirm)이 필요합니다'
            }), 400

        from dashboard.utils.user_database import get_calendar_event_repository

        db = get_user_database()
        repo = get_calendar_event_repository()

        # Google Sheets의 모든 프로젝트 코드 조회
        df = load_data()
        sheet_codes = set()
        if '프로젝트 코드' in df.columns:
            sheet_codes = set(df['프로젝트 코드'].astype(str).tolist())

        # DB의 모든 캘린더 매핑 조회
        import sqlite3
        conn = sqlite3.connect(db.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT project_code FROM project_calendar_events")
        all_mappings = cursor.fetchall()

        # 고아 매핑 삭제
        deleted = []
        for (code,) in all_mappings:
            if code not in sheet_codes:
                cursor.execute("DELETE FROM project_calendar_events WHERE project_code = ?", (code,))
                deleted.append(code)

        conn.commit()
        conn.close()

        logger.info(f"[ADMIN_CALENDAR_CLEANUP] 관리자가 {len(deleted)}개의 고아 매핑을 정리했습니다: {deleted}")

        return jsonify({
            'success': True,
            'deleted_count': len(deleted),
            'deleted_codes': deleted,
            'message': f'{len(deleted)}개의 고아 매핑을 정리했습니다'
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캘린더 매핑 정리 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캘린더 매핑 정리 중 오류가 발생했습니다',
            'error_id': error_id
        }), 500


@sequence_bp.route('/api/sequence/reset', methods=['POST'])
@admin_required
def reset_sequence():
    """
    프로젝트 시퀀스를 Google Sheets 최댓값으로 리셋

    Request Body:
        {
            "confirm": true
        }

    Returns:
        {
            "success": true,
            "message": "시퀀스가 2833에서 2830으로 리셋되었습니다.",
            "old_value": 2833,
            "new_value": 2830
        }
    """
    try:
        data = request.get_json()

        if not data or not data.get('confirm'):
            return jsonify({
                'success': False,
                'error': '확인(confirm)이 필요합니다'
            }), 400

        db = get_user_database()

        # Google Sheets 최댓값 조회
        df = load_data()
        sheet_max = 0
        if '프로젝트 코드' in df.columns:
            numbers = []
            for code in df['프로젝트 코드'].astype(str):
                number = _extract_number(code)
                if number is not None:
                    numbers.append(number)
            sheet_max = max(numbers) if numbers else 0

        # 시퀀스 리셋
        result = db.reset_sequence_to_sheet_max("GLOBAL", sheet_max)

        if result['success']:
            logger.info(f"[ADMIN_SEQUENCE_RESET] 관리자가 시퀀스를 리셋했습니다: {result['old_value']} → {result['new_value']}")

        return jsonify(result), 200 if result['success'] else 400

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 시퀀스 리셋 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '시퀀스 리셋 중 오류가 발생했습니다',
            'error_id': error_id
        }), 500
