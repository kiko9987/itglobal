"""
시스템 모니터링 블루프린트
헬스체크, 감사 로그, API 사용량, 잠금 상태 등 모니터링 기능들
"""

import os
from datetime import datetime
from flask import Blueprint, jsonify, request, session
from dashboard.auth import admin_required, login_required

from ..utils.logging_config import get_logger

logger = get_logger(__name__)
monitoring_bp = Blueprint('monitoring', __name__)


@monitoring_bp.route('/api/health')
def health_check():
    """
    상세한 Readiness Probe
    서비스가 트래픽을 받을 준비가 되었는지 확인
    Redis, 파일시스템 등 중요 서비스 체크
    """
    try:
        health_status = {
            'status': 'healthy',
            'timestamp': datetime.now().isoformat(),
            'version': '1.0',
            'service': 'ITG Dashboard',
            'services': {}
        }

        # Redis 연결 체크
        try:
            from dashboard.utils.redis_client import get_redis_client
            redis = get_redis_client()
            redis.ping()
            health_status['services']['redis'] = 'up'
        except Exception as e:
            logger.warning(f"Redis 헬스체크 실패: {e}")
            health_status['services']['redis'] = 'down'
            health_status['status'] = 'degraded'

        # 파일시스템 체크
        try:
            log_dir = os.path.join(os.path.dirname(__file__), '..', 'logs')
            if os.path.exists(log_dir):
                health_status['services']['filesystem'] = 'accessible'
            else:
                health_status['services']['filesystem'] = 'limited'
                health_status['status'] = 'degraded'
        except Exception as e:
            logger.warning(f"파일시스템 헬스체크 경고: {e}")
            health_status['services']['filesystem'] = 'error'
            health_status['status'] = 'degraded'

        # 데이터베이스 체크
        try:
            from dashboard.utils.user_database import get_user_database
            db = get_user_database()
            # 간단한 쿼리로 DB 연결 확인
            users = db.get_all_users()
            if isinstance(users, list):
                health_status['services']['database'] = 'up'
            else:
                health_status['services']['database'] = 'down'
                health_status['status'] = 'degraded'
        except Exception as e:
            logger.warning(f"Database 헬스체크 실패: {e}")
            health_status['services']['database'] = 'down'
            health_status['status'] = 'degraded'

        # HTTP 상태 코드 결정
        status_code = 200 if health_status['status'] == 'healthy' else 503

        return jsonify(health_status), status_code

    except Exception as e:
        logger.error(f"헬스체크 실패: {e}")
        return jsonify({
            'status': 'unhealthy',
            'error': str(e),
            'timestamp': datetime.now().isoformat()
        }), 500

@monitoring_bp.route('/api/audit-logs', methods=['GET'])
@login_required
def get_audit_logs_api():
    """감사 로그 조회 API (페이지네이션 지원)"""
    try:
        from dashboard.utils.user_database import get_audit_repository
        from dashboard.utils.user_database import get_user_database

        # 쿼리 파라미터 처리 및 검증
        try:
            days = int(request.args.get('days', 7))
            page = int(request.args.get('page', 1))
            per_page = int(request.args.get('per_page', 50))
        except (ValueError, TypeError):
            return jsonify({
                'success': False,
                'message': '유효하지 않은 파라미터입니다.'
            }), 400

        # days: 1~90일로 제한 (3개월 이내)
        if days < 1 or days > 90:
            return jsonify({
                'success': False,
                'message': '조회 기간은 1~90일 사이여야 합니다.'
            }), 400

        # per_page: 1~200으로 제한
        if per_page < 1 or per_page > 200:
            return jsonify({
                'success': False,
                'message': '페이지당 항목 수는 1~200 사이여야 합니다.'
            }), 400

        # page: 최소 1
        if page < 1:
            return jsonify({
                'success': False,
                'message': '페이지 번호는 1 이상이어야 합니다.'
            }), 400

        # 로그 조회 (SQL 레벨 필터링)
        audit_repo = get_audit_repository()
        user_db = get_user_database()
        user = session.get('user', {})
        user_email = user.get('email', '')
        user_role = user.get('permission_level', 'viewer')

        limit = days * 100  # 대략적인 제한

        # 관리자가 아닌 경우 SQL WHERE 절로 필터링 (성능 향상)
        if user_role != 'admin':
            all_logs = audit_repo.get_recent_logs(limit=limit, user_email=user_email)
        else:
            all_logs = audit_repo.get_recent_logs(limit=limit)

        # 사용자 정보 추가 (user_name, user_role)
        user_cache = {}  # 캐시로 중복 조회 방지
        for log in all_logs:
            log_email = log.get('user_email')
            if log_email:
                if log_email not in user_cache:
                    user = user_db.get_user_by_email(log_email)
                    if user:
                        user_cache[log_email] = {
                            'name': user.get('name', log_email),
                            'role': user.get('permission_level', 'viewer')
                        }
                    else:
                        user_cache[log_email] = {
                            'name': log_email,
                            'role': 'viewer'
                        }

                log['user_name'] = user_cache[log_email]['name']
                log['user_role'] = user_cache[log_email]['role']
            else:
                log['user_name'] = 'Unknown'
                log['user_role'] = 'viewer'

            # SQLite의 timestamp는 UTC이므로 명시적으로 'Z'를 붙여서 ISO 8601 형식으로 변환
            # 예: '2025-01-13 12:34:56' -> '2025-01-13T12:34:56Z'
            if log.get('timestamp'):
                timestamp_str = log['timestamp']
                # 공백을 'T'로 변경하고 끝에 'Z' 추가
                if ' ' in timestamp_str and not timestamp_str.endswith('Z'):
                    log['timestamp'] = timestamp_str.replace(' ', 'T') + 'Z'

        # 최신순 정렬 (timestamp 기준 내림차순)
        all_logs.sort(key=lambda x: x.get('timestamp', ''), reverse=True)

        # 페이지네이션 계산
        total_count = len(all_logs)
        total_pages = (total_count + per_page - 1) // per_page  # 올림 계산
        start_index = (page - 1) * per_page
        end_index = start_index + per_page

        # 페이지 데이터 추출
        page_logs = all_logs[start_index:end_index]

        return jsonify({
            'success': True,
            'logs': page_logs,
            'pagination': {
                'current_page': page,
                'total_pages': total_pages,
                'total_count': total_count,
                'per_page': per_page,
                'has_next': page < total_pages,
                'has_prev': page > 1
            }
        })

    except Exception as e:
        logger.error(f"감사 로그 조회 오류: {str(e)}")
        return jsonify({'success': False, 'message': '감사 로그를 불러올 수 없습니다.'}), 500

@monitoring_bp.route('/api/monitoring/usage')
@admin_required
def get_api_usage():
    """API 사용량 통계 조회 (관리자 전용)"""
    try:
        from dashboard.utils.api_usage_monitor import get_api_monitor, check_google_sheets_rate_limit

        monitor = get_api_monitor()
        stats = monitor.get_usage_stats()
        rate_limit = check_google_sheets_rate_limit()

        return jsonify({
            'usage_stats': stats,
            'rate_limit': rate_limit,
            'timestamp': datetime.now().isoformat()
        })
    except Exception as e:
        logger.error(f"API 사용량 통계 조회 오류: {e}")
        return jsonify({'error': str(e)}), 500

@monitoring_bp.route('/api/monitoring/prefetch')
@admin_required
def get_prefetch_stats():
    """백그라운드 프리패치 통계 조회 (관리자 전용)"""
    try:
        from dashboard.utils.background_prefetch import get_background_prefetch
        prefetch = get_background_prefetch()
        stats = prefetch.get_stats()
        return jsonify(stats)
    except Exception as e:
        logger.error(f"프리패치 통계 조회 오류: {e}")
        return jsonify({'error': str(e)}), 500