"""
관리자 블루프린트
캐시 관리, 시스템 모니터링, 관리자 전용 기능들
"""

import logging
from datetime import datetime, timedelta
from flask import Blueprint, render_template, jsonify, request
from dashboard.auth import admin_required, login_required
from dashboard.utils.logging_config import get_logger
from dashboard.utils.error_helpers import generate_error_id

logger = get_logger(__name__)

admin_bp = Blueprint('admin', __name__, url_prefix='/admin')

@admin_bp.route('/cache-monitor')
@admin_required
def cache_monitor_page():
    """캐시 모니터링 페이지 (관리자 전용)"""
    return render_template('cache_monitor.html')

@admin_bp.route('/docs')
@admin_required
def docs_management_page():
    """API 문서 관리 페이지 (관리자 전용)"""
    return render_template('docs_management.html')

# API 엔드포인트들
@admin_bp.route('/api/cache-stats')
@login_required
def get_cache_stats():
    """캐시 상태 모니터링 API"""
    try:
        from dashboard.utils.smart_cache_manager import get_smart_cache, cache_stats

        # 기본 캐시 통계
        basic_stats = cache_stats()

        # 스마트 캐시 통계
        smart_cache = get_smart_cache()
        smart_stats = smart_cache.get_cache_info()

        return jsonify({
            'basic_cache': basic_stats,
            'smart_cache': smart_stats,
            'timestamp': datetime.now().isoformat(),
            'status': 'healthy'
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캐시 통계 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캐시 통계를 조회할 수 없습니다',
            'error_id': error_id
        }), 500

@admin_bp.route('/api/cache-clear', methods=['POST'])
@admin_required
def clear_cache():
    """캐시 전체 삭제 API (관리자 전용)"""
    try:
        from dashboard.utils.smart_cache_manager import get_smart_cache, cache_clear
        from dashboard.utils.user_database import get_audit_repository
        from flask import session

        # 기본 캐시 삭제
        cache_clear()

        # 스마트 캐시 삭제
        smart_cache = get_smart_cache()
        smart_cache.clear()

        # 감사 로그 기록
        try:
            audit_repo = get_audit_repository()
            user_email = session.get('user', {}).get('email', 'unknown')
            audit_repo.log_action(
                user_email=user_email,
                action='CACHE_CLEAR',
                details='전체 캐시 삭제',
                field_name='cache',
                old_value='전체 캐시',
                new_value='삭제됨',
                ip_address=request.remote_addr
            )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")

        return jsonify({
            'success': True,
            'message': '캐시가 성공적으로 삭제되었습니다.',
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캐시 삭제 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캐시를 삭제할 수 없습니다',
            'error_id': error_id
        }), 500

@admin_bp.route('/api/cache/status')
@login_required
def get_cache_status():
    """캐시 상태 조회 (모든 사용자)"""
    try:
        from dashboard.utils.smart_cache_manager import get_smart_cache
        from dashboard.utils.background_prefetch import get_background_prefetch

        cache = get_smart_cache()
        prefetch = get_background_prefetch()

        # 캐시 기본 정보
        cache_info = cache.get_cache_info()

        # 프리패치 상태 (간소화된 정보만)
        prefetch_stats = prefetch.get_stats()

        # 마지막 성공적인 프리패치 시간
        last_success = None
        has_recent_errors = False

        if 'current_sheet_data' in prefetch_stats.get('jobs', {}):
            job_info = prefetch_stats['jobs']['current_sheet_data']
            last_success = job_info.get('last_prefetch')
            has_recent_errors = job_info.get('error_count', 0) > 0

        # 캐시 건강 상태 판단
        cache_health = 'healthy'
        if not prefetch_stats.get('running', False):
            cache_health = 'warning'  # 프리패치가 실행되지 않음
        elif has_recent_errors:
            cache_health = 'error'    # 최근 오류 발생
        elif last_success:
            # 마지막 성공이 10분 이상 전이면 경고
            try:
                # ISO 형식 날짜 파싱 (YYYY-MM-DDTHH:MM:SS.ffffff)
                last_time = datetime.fromisoformat(last_success.replace('Z', '+00:00'))
                if datetime.now() - last_time > timedelta(minutes=10):
                    cache_health = 'warning'
            except Exception as e:
                logger.warning(f"처리 중 오류 무시: {e}")
                pass

        return jsonify({
            'cache_info': cache_info,
            'prefetch_running': prefetch_stats.get('running', False),
            'last_successful_update': last_success,
            'cache_health': cache_health,
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캐시 상태 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캐시 상태를 조회할 수 없습니다',
            'error_id': error_id
        }), 500

@admin_bp.route('/api/cache/refresh', methods=['POST'])
@login_required
def manual_cache_refresh():
    """수동 캐시 새로고침 (폴백 전략)"""
    try:
        from dashboard.services.project_service import invalidate_project_cache, load_data
        from flask import session

        # 기존 캐시 무효화
        invalidate_project_cache()

        # 새 데이터 강제 로드
        fresh_data = load_data()

        logger.info(f"[MANUAL_REFRESH] 사용자 요청으로 캐시 수동 새로고침 완료")

        data_count = 0
        if fresh_data is not None:
            if hasattr(fresh_data, '__len__'):
                data_count = len(fresh_data)
            elif hasattr(fresh_data, 'shape'):  # pandas DataFrame
                data_count = fresh_data.shape[0]

        # 감사 로그 기록
        try:
            from dashboard.utils.user_database import get_audit_repository
            audit_repo = get_audit_repository()
            user_email = session.get('user', {}).get('email', 'unknown')
            audit_repo.log_action(
                user_email=user_email,
                action='CACHE_REFRESH',
                details=f'수동 캐시 새로고침 (데이터 {data_count}건)',
                field_name='cache',
                old_value='이전 캐시',
                new_value='새로고침됨',
                ip_address=request.remote_addr
            )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")

        return jsonify({
            'success': True,
            'message': f'캐시가 성공적으로 새로고침되었습니다. ({data_count}건의 데이터 로드)',
            'data_count': data_count,
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캐시 새로고침 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캐시를 새로고침할 수 없습니다',
            'error_id': error_id
        }), 500

@admin_bp.route('/api/calendar/sync-all', methods=['POST'])
@admin_required
def sync_all_projects_to_calendar():
    """
    기존 프로젝트를 캘린더에 일괄 등록 (관리자 전용)

    - 공사 시작일이 있는 모든 프로젝트 조회
    - 아직 캘린더에 등록되지 않은 프로젝트만 필터링
    - 백그라운드에서 순차 등록
    - 진행 상황을 로그로 출력

    Returns:
        JSON: 등록 시작 확인 메시지 (비동기 실행)
    """
    try:
        from dashboard.services.project_service import get_project_records
        from dashboard.services.calendar_service import create_project_calendar_event, is_calendar_enabled
        from dashboard.utils.calendar_event_repository import get_calendar_event_repository
        from flask import session
        import threading

        # 캘린더 활성화 확인
        if not is_calendar_enabled():
            return jsonify({
                'success': False,
                'error': 'Google Calendar 연동이 설정되지 않았습니다.'
            }), 400

        def background_sync():
            """백그라운드 동기화 작업"""
            try:
                logger.info("[CALENDAR_SYNC_ALL] 일괄 동기화 시작")

                # 모든 프로젝트 조회
                all_projects = get_project_records()
                logger.info(f"[CALENDAR_SYNC_ALL] 전체 프로젝트: {len(all_projects)}개")

                # 공사 시작일이 있는 프로젝트만 필터링
                projects_with_date = [
                    p for p in all_projects
                    if p.get('공사 시작') and str(p.get('공사 시작')).strip()
                ]
                logger.info(f"[CALENDAR_SYNC_ALL] 공사 시작일 있음: {len(projects_with_date)}개")

                # 이미 등록된 프로젝트 확인
                repo = get_calendar_event_repository()
                registered_codes = set()
                for project in projects_with_date:
                    code = project.get('프로젝트 코드')
                    if code and repo.get_event(code):
                        registered_codes.add(code)

                logger.info(f"[CALENDAR_SYNC_ALL] 이미 등록됨: {len(registered_codes)}개")

                # 등록되지 않은 프로젝트만 추출
                projects_to_sync = [
                    p for p in projects_with_date
                    if p.get('프로젝트 코드') not in registered_codes
                ]

                logger.info(f"[CALENDAR_SYNC_ALL] 등록 대상: {len(projects_to_sync)}개")

                # 순차 등록
                success_count = 0
                error_count = 0

                for i, project in enumerate(projects_to_sync, 1):
                    project_code = project.get('프로젝트 코드')
                    try:
                        event_id = create_project_calendar_event(project)
                        if event_id:
                            success_count += 1
                            logger.info(f"[CALENDAR_SYNC_ALL] [{i}/{len(projects_to_sync)}] 등록 성공: {project_code}")
                        else:
                            error_count += 1
                            logger.warning(f"[CALENDAR_SYNC_ALL] [{i}/{len(projects_to_sync)}] 등록 실패: {project_code} (이벤트 ID 없음)")
                    except Exception as e:
                        error_count += 1
                        logger.error(f"[CALENDAR_SYNC_ALL] [{i}/{len(projects_to_sync)}] 등록 오류: {project_code} - {e}")

                logger.info(
                    f"[CALENDAR_SYNC_ALL] 일괄 동기화 완료 - "
                    f"성공: {success_count}개, 실패: {error_count}개, "
                    f"이미 등록: {len(registered_codes)}개"
                )

            except Exception as e:
                logger.error(f"[CALENDAR_SYNC_ALL] 백그라운드 동기화 오류: {e}", exc_info=True)

        # 백그라운드 스레드로 실행
        sync_thread = threading.Thread(target=background_sync, daemon=True)
        sync_thread.start()

        # 감사 로그 기록
        try:
            from dashboard.utils.user_database import get_audit_repository
            audit_repo = get_audit_repository()
            user_email = session.get('user', {}).get('email', 'unknown')
            audit_repo.log_action(
                user_email=user_email,
                action='CALENDAR_SYNC_ALL',
                details='기존 프로젝트 캘린더 일괄 등록 시작',
                field_name='calendar',
                old_value='미등록',
                new_value='일괄 등록 시작',
                ip_address=request.remote_addr
            )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")

        return jsonify({
            'success': True,
            'message': '캘린더 일괄 동기화를 시작했습니다. 서버 로그에서 진행 상황을 확인하세요.',
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캘린더 일괄 동기화 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캘린더 일괄 동기화를 시작할 수 없습니다',
            'error_id': error_id
        }), 500