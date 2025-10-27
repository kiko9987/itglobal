"""
프로젝트 잠금 관리 블루프린트
통합 편집 모드에서 프로젝트 단위 잠금 획득/해제 및 실시간 협업 기능
"""

import logging
from datetime import datetime
from flask import Blueprint, jsonify, request, session
from dashboard.auth import login_required
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

locks_bp = Blueprint('locks', __name__, url_prefix='/api')

@locks_bp.route('/project-lock/acquire', methods=['POST'])
@login_required
def acquire_project_lock():
    """프로젝트 잠금 획득 (통합 편집 모드 진입)"""
    try:
        from dashboard.utils.project_lock_manager import get_project_lock_manager

        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'message': '요청 데이터가 필요합니다.'}), 400

        project_code = data.get('project_code')
        if not project_code:
            return jsonify({'success': False, 'message': '프로젝트 코드가 필요합니다.'}), 400

        # 사용자 정보 추출
        user_data = session.get('user', {})
        user_email = user_data.get('email', 'unknown@test.com')
        user_name = user_data.get('name', 'Unknown User')

        # 탭 ID 추출 (다중 탭 구분용 - 프론트에서 전달)
        tab_id = data.get('tab_id')
        if not tab_id:
            return jsonify({'success': False, 'message': '탭 ID가 필요합니다.'}), 400

        # 프로젝트 잠금 획득
        lock_manager = get_project_lock_manager()
        result = lock_manager.acquire_lock(project_code, user_email, user_name, tab_id)

        # 잠금 상태 변경을 실시간으로 브로드캐스트
        if result['success']:
            try:
                from dashboard import socketio
                socketio.emit('project_lock_changed', {
                    'project_code': project_code,
                    'action': 'acquire',
                    'user_email': user_email,
                    'user_name': user_name,
                    'timestamp': datetime.now().isoformat()
                }, broadcast=True)
                logger.info(f"[LOCK] 프로젝트 잠금 획득: {project_code} by {user_name} ({user_email})")
            except Exception as e:
                logger.debug(f"SocketIO 브로드캐스트 실패: {e}")

        return jsonify(result)

    except Exception as e:
        logger.error(f"프로젝트 잠금 획득 오류: {str(e)}")
        return jsonify({'success': False, 'message': '잠금 획득 중 오류가 발생했습니다.'}), 500

@locks_bp.route('/project-lock/release', methods=['POST'])
@login_required
def release_project_lock():
    """프로젝트 잠금 해제 (통합 편집 모드 종료)"""
    try:
        from dashboard.utils.project_lock_manager import get_project_lock_manager

        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'message': '요청 데이터가 필요합니다.'}), 400

        project_code = data.get('project_code')
        if not project_code:
            return jsonify({'success': False, 'message': '프로젝트 코드가 필요합니다.'}), 400

        # 사용자 정보 추출
        user_data = session.get('user', {})
        user_email = user_data.get('email', 'unknown@test.com')

        # 탭 ID 추출 (다중 탭 구분용 - 프론트에서 전달)
        tab_id = data.get('tab_id')
        if not tab_id:
            return jsonify({'success': False, 'message': '탭 ID가 필요합니다.'}), 400

        # 프로젝트 잠금 해제 (사용자 + 탭 ID 일치 확인)
        lock_manager = get_project_lock_manager()
        result = lock_manager.release_lock(project_code, user_email, tab_id)

        # 잠금 해제를 실시간으로 브로드캐스트
        if result['success']:
            try:
                from dashboard import socketio
                socketio.emit('project_lock_changed', {
                    'project_code': project_code,
                    'action': 'release',
                    'user_email': user_email,
                    'timestamp': datetime.now().isoformat()
                }, broadcast=True)
                logger.info(f"[LOCK] 프로젝트 잠금 해제: {project_code} by {user_email}")
            except Exception as e:
                logger.debug(f"SocketIO 브로드캐스트 실패: {e}")

        return jsonify(result)

    except Exception as e:
        logger.error(f"프로젝트 잠금 해제 오류: {str(e)}")
        return jsonify({'success': False, 'message': '잠금 해제 중 오류가 발생했습니다.'}), 500

@locks_bp.route('/project-lock/status/<project_code>')
@login_required
def get_project_lock_status(project_code):
    """프로젝트 잠금 상태 조회"""
    try:
        from dashboard.utils.project_lock_manager import get_project_lock_manager

        lock_manager = get_project_lock_manager()
        lock_info = lock_manager.get_lock_status(project_code)

        return jsonify({
            'success': True,
            'lock_info': lock_info,
            'is_locked': lock_info is not None
        })

    except Exception as e:
        logger.error(f"잠금 상태 조회 오류: {str(e)}")
        return jsonify({
            'success': False,
            'message': '잠금 상태 조회 중 오류가 발생했습니다.'
        }), 500

@locks_bp.route('/project-lock/force-release', methods=['POST'])
@login_required
def force_release_project_lock():
    """관리자용 강제 프로젝트 잠금 해제"""
    try:
        from dashboard.utils.project_lock_manager import get_project_lock_manager

        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'message': '요청 데이터가 필요합니다.'}), 400

        project_code = data.get('project_code')
        if not project_code:
            return jsonify({'success': False, 'message': '프로젝트 코드가 필요합니다.'}), 400

        # 관리자 정보
        user_data = session.get('user', {})
        admin_email = user_data.get('email', 'unknown')
        admin_permission = user_data.get('permission_level', 'viewer')

        # 강제 잠금 해제
        lock_manager = get_project_lock_manager()
        result = lock_manager.force_release_lock(project_code, admin_email, admin_permission)

        # 권한 부족 시 403 반환
        if not result['success'] and '권한이 필요합니다' in result.get('message', ''):
            return jsonify(result), 403

        # 감사 로그 기록
        if result['success']:
            try:
                from dashboard.utils.user_database import get_audit_repository
                audit_repo = get_audit_repository()
                audit_repo.log_action(
                    user_email=admin_email,
                    action='FORCE_PROJECT_LOCK_RELEASE',
                    details=f'프로젝트 {project_code} 잠금 강제 해제 (원 소유자: {result.get("previous_owner", "unknown")})',
                    field_name='project_lock',
                    old_value='잠금됨',
                    new_value='해제됨',
                    ip_address=request.remote_addr
                )
            except Exception as log_error:
                logger.warning(f"감사 로그 기록 실패: {log_error}")

            # 실시간 브로드캐스트
            try:
                from dashboard import socketio
                socketio.emit('project_lock_changed', {
                    'project_code': project_code,
                    'action': 'admin_force_release',
                    'admin_email': admin_email,
                    'timestamp': datetime.now().isoformat()
                }, broadcast=True)
            except Exception as e:
                logger.debug(f"SocketIO 브로드캐스트 실패: {e}")

        return jsonify(result)

    except Exception as e:
        logger.error(f"강제 잠금 해제 오류: {str(e)}")
        return jsonify({
            'success': False,
            'message': '강제 잠금 해제 중 오류가 발생했습니다.'
        }), 500

@locks_bp.route('/project-lock/heartbeat', methods=['POST'])
@login_required
def project_lock_heartbeat():
    """프로젝트 잠금 연장 (heartbeat)"""
    try:
        from dashboard.utils.project_lock_manager import get_project_lock_manager

        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'message': '요청 데이터가 필요합니다.'}), 400

        project_code = data.get('project_code')
        if not project_code:
            return jsonify({'success': False, 'message': '프로젝트 코드가 필요합니다.'}), 400

        # 사용자 정보 추출
        user_data = session.get('user', {})
        user_email = user_data.get('email', 'unknown@test.com')
        user_name = user_data.get('name', 'Unknown User')

        # 탭 ID 추출 (다중 탭 구분용 - 프론트에서 전달)
        tab_id = data.get('tab_id')
        if not tab_id:
            return jsonify({'success': False, 'message': '탭 ID가 필요합니다.'}), 400

        # 잠금 연장 (acquire_lock은 같은 사용자 + 같은 탭이면 자동으로 연장)
        lock_manager = get_project_lock_manager()
        result = lock_manager.acquire_lock(project_code, user_email, user_name, tab_id)

        # 다른 사용자가 잠금 중이면 실패
        if not result['success']:
            logger.warning(f"[LOCK] Heartbeat 실패: {project_code} - {result.get('message')}")
            return jsonify(result), 409  # Conflict

        logger.debug(f"[LOCK] Heartbeat: {project_code} by {user_email}")
        return jsonify(result)

    except Exception as e:
        logger.error(f"프로젝트 잠금 연장 오류: {str(e)}")
        return jsonify({'success': False, 'message': '잠금 연장 중 오류가 발생했습니다.'}), 500

@locks_bp.route('/project-lock/release-all-user-locks', methods=['POST'])
@login_required
def release_all_user_project_locks():
    """사용자의 모든 프로젝트 잠금 해제 (페이지 종료 시 등)"""
    try:
        from dashboard.utils.project_lock_manager import get_project_lock_manager

        data = request.get_json() if request.get_json() else {}

        # 현재 로그인한 사용자 정보
        current_user = session.get('user', {})
        current_user_email = current_user.get('email', '')
        current_user_permission = current_user.get('permission_level', 'viewer')

        # 요청에서 받은 이메일
        requested_user_email = data.get('user_email')

        # 해제할 대상 사용자 이메일 결정
        if not requested_user_email:
            # 요청에 이메일이 없으면 본인 것만 해제
            user_email = current_user_email
        else:
            # 요청에 이메일이 있는 경우
            if requested_user_email != current_user_email:
                # 다른 사용자의 락을 해제하려는 경우 - 관리자 권한 필요
                if current_user_permission not in ['admin', 'super_admin']:
                    logger.warning(f"[SECURITY] 권한 없는 잠금 해제 시도: {current_user_email}이(가) {requested_user_email}의 락 해제 시도")
                    return jsonify({
                        'success': False,
                        'message': '다른 사용자의 잠금을 해제할 권한이 없습니다.'
                    }), 403

                # 관리자는 다른 사용자 락 해제 가능
                logger.info(f"[ADMIN] 관리자 {current_user_email}이(가) {requested_user_email}의 락 해제 요청")

            user_email = requested_user_email

        reason = data.get('reason', 'manual_cleanup')

        # 잠금 해제
        lock_manager = get_project_lock_manager()
        user_locks = lock_manager.get_user_locks(user_email)
        released_count = lock_manager.release_all_user_locks(user_email, reason)

        logger.info(f"[LOCK_CLEANUP] 사용자 {user_email}의 {released_count}개 잠금 해제됨. 사유: {reason}")

        # 실시간 브로드캐스트
        try:
            from dashboard import socketio
            for lock in user_locks:
                socketio.emit('project_lock_changed', {
                    'project_code': lock.get('project_code'),
                    'action': 'user_cleanup',
                    'user_email': user_email,
                    'reason': reason,
                    'timestamp': datetime.now().isoformat()
                }, broadcast=True)
        except Exception as e:
            logger.debug(f"SocketIO 브로드캐스트 실패: {e}")

        return jsonify({
            'success': True,
            'message': f'{released_count}개의 잠금이 해제되었습니다.',
            'released_count': released_count,
            'released_locks': [lock.get('project_code') for lock in user_locks]
        })

    except Exception as e:
        logger.error(f"사용자 잠금 해제 오류: {str(e)}")
        return jsonify({
            'success': False,
            'message': '잠금 해제 중 오류가 발생했습니다.',
            'error': str(e)
        }), 500

@locks_bp.route('/project-lock/all-locks')
@login_required
def get_all_project_locks():
    """모든 활성 프로젝트 잠금 조회 (모니터링용)"""
    try:
        from dashboard.utils.project_lock_manager import get_project_lock_manager

        lock_manager = get_project_lock_manager()
        all_locks = lock_manager.get_all_locks()

        return jsonify({
            'success': True,
            'locks': all_locks,
            'count': len(all_locks)
        })

    except Exception as e:
        logger.error(f"모든 잠금 조회 오류: {str(e)}")
        return jsonify({
            'success': False,
            'message': '잠금 목록 조회 중 오류가 발생했습니다.'
        }), 500