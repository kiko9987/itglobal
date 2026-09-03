"""
사용자 관리 블루프린트
사용자 CRUD, 권한 관리, 사용자 상태 관리 기능들
"""

import logging
from flask import Blueprint, jsonify, request, session
from dashboard.auth import admin_required, login_required, get_user_role
from dashboard.utils.logging_config import get_logger
from dashboard.utils.user_database import get_user_database
from dashboard.utils.error_helpers import generate_error_id

logger = get_logger(__name__)

users_bp = Blueprint('users', __name__, url_prefix='/api')

def _get_user_suffix(user_email, user_name):
    """
    사용자의 프로젝트 코드 접미사 조회 (내부 헬퍼 함수)

    Args:
        user_email: 사용자 이메일
        user_name: 사용자 이름

    Returns:
        str or None: 프로젝트 코드 접미사 (없으면 None)
    """
    try:
        from dashboard.utils.redis_client import get_redis_client
        from dashboard.services import project_service

        # Redis에서 조회
        redis_client = get_redis_client()
        redis_key = f"user:project_code_suffix:{user_email}"
        suffix = redis_client.get(redis_key)

        if suffix:
            return suffix

        # Redis에 없으면 project_config.json에서 조회
        config_suffix = project_service.PROJECT_CONFIG.get('owner_suffix_map', {}).get(user_name, '')

        if config_suffix:
            return config_suffix

        return None
    except Exception as e:
        logger.error(f"프로젝트 코드 접미사 조회 오류 (user: {user_email}): {str(e)}")
        return None

@users_bp.route('/user/role')
@login_required
def get_user_role_api():
    """사용자 역할 정보 반환"""
    try:
        user_role = get_user_role()
        return jsonify({'success': True, 'role': user_role})
    except Exception as e:
        logger.error(f"사용자 역할 조회 오류: {str(e)}")
        return jsonify({'success': False, 'role': 'user'}), 500

@users_bp.route('/users', methods=['GET'])
@admin_required
def get_users():
    """사용자 목록 조회 (관리자만, 페이지네이션 지원)"""
    try:
        # 쿼리 파라미터 처리 및 검증
        try:
            page = int(request.args.get('page', 1))
            per_page = int(request.args.get('per_page', 100))  # 기본 100개 (사용자는 많지 않음)
        except (ValueError, TypeError):
            return jsonify({
                'success': False,
                'message': '유효하지 않은 파라미터입니다.'
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

        user_db = get_user_database()
        all_users = user_db.get_all_users()

        # 정렬: 권한 순서(admin -> editor -> viewer) -> 퇴사자 -> 이름 가나다순
        def sort_key(user):
            # 권한 우선순위 (숫자가 작을수록 먼저)
            permission_priority = {
                'admin': 1,
                'super_admin': 1,  # admin과 동일 취급
                'editor': 2,
                'viewer': 3
            }

            # 퇴사자는 가장 마지막
            is_resigned = user.get('is_resigned', False)
            resigned_priority = 1 if is_resigned else 0

            # 정렬 키: (퇴사여부, 권한우선순위, 이름)
            return (
                resigned_priority,
                permission_priority.get(user.get('permission_level', 'viewer'), 99),
                user.get('name', '')
            )

        all_users = sorted(all_users, key=sort_key)

        # 페이지네이션 계산
        total_count = len(all_users)
        total_pages = (total_count + per_page - 1) // per_page  # 올림 계산
        start_index = (page - 1) * per_page
        end_index = start_index + per_page

        # 페이지 데이터 추출
        page_users = all_users[start_index:end_index]

        # 각 사용자에 프로젝트 코드 접미사 추가
        for user in page_users:
            user_email = user.get('email', '')
            user_name = user.get('name', '')
            user['project_code_suffix'] = _get_user_suffix(user_email, user_name)

        return jsonify({
            'success': True,
            'users': page_users,
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
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 사용자 목록 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'message': '사용자 목록을 불러올 수 없습니다.',
            'error_id': error_id
        }), 500

@users_bp.route('/resigned-managers', methods=['GET'])
@login_required
def get_resigned_managers():
    """퇴사 처리된 담당자 정보 목록 조회 (모든 로그인 사용자)"""
    try:
        from dashboard.utils.user_database import get_user_database
        user_db = get_user_database()
        resigned_data = user_db.get_resigned_managers()
        logger.info(f"[API] 퇴사자 {len(resigned_data)}명 조회됨: {[user['name'] for user in resigned_data]}")
        return jsonify({'success': True, 'resigned_managers': resigned_data})
    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 퇴사자 목록 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'message': '퇴사자 목록을 불러올 수 없습니다.',
            'resigned_managers': [],
            'error_id': error_id
        }), 500

@users_bp.route('/inactive-managers', methods=['GET'])
@login_required
def get_inactive_managers():
    """담당자 후보 제외 대상(비활성+퇴사) 목록 조회 (모든 로그인 사용자).

    새 프로젝트 등록 모달의 담당자 드롭다운에서 이 이름들을 제외한다.
    """
    try:
        from dashboard.utils.user_database import get_user_database
        user_db = get_user_database()
        data = user_db.get_inactive_managers()
        return jsonify({'success': True, 'inactive_managers': data})
    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 비활성/퇴사 담당자 목록 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'message': '비활성/퇴사 담당자 목록을 불러올 수 없습니다.',
            'inactive_managers': [],
            'error_id': error_id
        }), 500


@users_bp.route('/users/permission', methods=['POST'])
@admin_required
def update_user_permission():
    """사용자 권한 업데이트 (관리자만)"""
    try:
        data = request.get_json()
        email = data.get('email')
        permission = data.get('permission')

        if not email or not permission:
            return jsonify({'success': False, 'message': '이메일과 권한이 필요합니다.'}), 400

        user_db = get_user_database()

        # 현재 권한 조회 (감사 로그용)
        user_info = user_db.get_user_by_email(email)
        old_permission = user_info.get('permission_level', '-') if user_info else '-'

        success, message = user_db.update_user_permission(email, permission)
        if success:
            # 감사 로그 기록
            try:
                from dashboard.utils.user_database import get_audit_repository
                audit_repo = get_audit_repository()
                user_email = session.get('user', {}).get('email', 'unknown')
                audit_repo.log_action(
                    user_email=user_email,
                    action='USER_PERMISSION_UPDATE',
                    details=f'사용자 {email}의 권한을 {permission}으로 변경',
                    field_name='permission_level',
                    old_value=old_permission,
                    new_value=permission,
                    ip_address=request.remote_addr
                )
            except Exception as log_error:
                logger.warning(f"감사 로그 기록 실패: {log_error}")

            return jsonify({'success': True, 'message': '권한이 성공적으로 업데이트되었습니다.'})
        else:
            return jsonify({'success': False, 'message': '권한 업데이트에 실패했습니다.'})

    except Exception as e:
        logger.error(f"권한 업데이트 오류: {str(e)}")
        return jsonify({'success': False, 'message': '권한 업데이트에 실패했습니다.'}), 500

@users_bp.route('/users/status', methods=['POST'])
@admin_required
def toggle_user_status():
    """사용자 상태 변경 (관리자만) - 활성/비활성/퇴사"""
    try:
        data = request.get_json()
        email = data.get('email')

        # 새로운 형식: status 문자열 ('활성', '비활성', '퇴사')
        if 'status' in data:
            status = data.get('status')
            from dashboard.utils.user_database import get_user_database
            user_db = get_user_database()

            # 현재 상태 조회 (감사 로그용)
            user_info = user_db.get_user_by_email(email)
            if user_info:
                is_active = user_info.get('is_active', False)
                is_resigned = user_info.get('is_resigned', False)
                if is_resigned:
                    old_status = '퇴사'
                elif is_active:
                    old_status = '활성'
                else:
                    old_status = '비활성'
            else:
                old_status = '-'

            # 새 상태를 한글로 변환 (영문이 들어올 수도 있음)
            status_to_korean = {
                '활성': '활성',
                'active': '활성',
                '비활성': '비활성',
                'inactive': '비활성',
                '퇴사': '퇴사',
                'resigned': '퇴사'
            }
            new_status = status_to_korean.get(status, status)

            success, message = user_db.update_user_status(email, status)

            if success:
                # 감사 로그 기록
                try:
                    from dashboard.utils.user_database import get_audit_repository
                    audit_repo = get_audit_repository()
                    user_email = session.get('user', {}).get('email', 'unknown')
                    audit_repo.log_action(
                        user_email=user_email,
                        action='USER_STATUS_CHANGE',
                        details=f'사용자 {email}을 {new_status} 상태로 변경',
                        field_name='status',
                        old_value=old_status,
                        new_value=new_status,
                        ip_address=request.remote_addr
                    )
                except Exception as log_error:
                    logger.warning(f"감사 로그 기록 실패: {log_error}")

                return jsonify({'success': True, 'message': message})
            else:
                return jsonify({'success': False, 'message': message})

        # 기존 형식: is_active 불리언 (하위 호환성)
        is_active = data.get('is_active')
        if email is None or is_active is None:
            return jsonify({'success': False, 'message': '이메일과 상태가 필요합니다.'}), 400

        user_db = get_user_database()
        status = '활성' if is_active else '비활성'
        success, message = user_db.update_user_status(email, status)
        if success:
            # 감사 로그 기록
            try:
                from dashboard.utils.user_database import get_audit_repository
                audit_repo = get_audit_repository()
                user_email = session.get('user', {}).get('email', 'unknown')
                status_text = '활성화' if is_active else '비활성화'
                audit_repo.log_action(
                    user_email=user_email,
                    action='USER_STATUS_CHANGE',
                    details=f'사용자 {email}을 {status_text}',
                    field_name='is_active',
                    old_value=str(not is_active),
                    new_value=str(is_active),
                    ip_address=request.remote_addr
                )
            except Exception as log_error:
                logger.warning(f"감사 로그 기록 실패: {log_error}")

            return jsonify({'success': True, 'message': f'사용자 상태가 성공적으로 변경되었습니다.'})
        else:
            return jsonify({'success': False, 'message': '사용자 상태 변경에 실패했습니다.'})

    except Exception as e:
        logger.error(f"사용자 상태 변경 오류: {str(e)}")
        return jsonify({'success': False, 'message': '사용자 상태 변경에 실패했습니다.'}), 500

# Google OAuth 전용 시스템: 비밀번호 기반 사용자 생성 엔드포인트 제거됨
# 사용자는 Google OAuth를 통해 자동 등록됩니다.

@users_bp.route('/users/<email>', methods=['DELETE'])
@admin_required
def delete_user(email):
    """사용자 삭제 (관리자만)"""
    try:
        # 본인 계정 삭제 방지
        if session['user']['email'] == email:
            return jsonify({'success': False, 'message': '본인 계정은 삭제할 수 없습니다.'}), 400

        # 삭제하려는 사용자 정보 조회 (로그를 위해)
        user_db = get_user_database()
        user_info = user_db.get_user_by_email(email)
        user_name = user_info.get('name', '알 수 없음') if user_info else '알 수 없음'

        success, message = user_db.delete_user(email)
        if success:
            # 감사 로그 기록
            try:
                from dashboard.utils.user_database import get_audit_repository
                audit_repo = get_audit_repository()
                admin_email = session.get('user', {}).get('email', 'unknown')
                audit_repo.log_action(
                    user_email=admin_email,
                    action='USER_DELETE',
                    details=f'사용자 삭제: {user_name} ({email})',
                    field_name='user',
                    old_value=f'{user_name} ({email})',
                    new_value='삭제됨',
                    ip_address=request.remote_addr
                )
            except Exception as log_error:
                logger.warning(f"감사 로그 기록 실패: {log_error}")

            return jsonify({'success': True, 'message': '사용자가 성공적으로 삭제되었습니다.'})
        else:
            return jsonify({'success': False, 'message': '사용자 삭제에 실패했습니다.'})

    except Exception as e:
        logger.error(f"사용자 삭제 오류: {str(e)}")
        return jsonify({'success': False, 'message': '사용자 삭제에 실패했습니다.'}), 500

@users_bp.route('/users/<email>/project-code-suffix', methods=['GET'])
@login_required
def get_project_code_suffix(email):
    """사용자 프로젝트 코드 접미사 조회"""
    try:
        from dashboard.utils.redis_client import get_redis_client

        redis_client = get_redis_client()
        redis_key = f"user:project_code_suffix:{email}"

        # Redis에서 조회
        suffix = redis_client.get(redis_key)

        if suffix:
            logger.debug(f"프로젝트 코드 접미사 조회: {email} → {suffix}")
            return jsonify({
                'success': True,
                'suffix': suffix,
                'source': 'custom'
            })

        # Redis에 없으면 project_config.json에서 조회
        from dashboard.services import project_service
        user_db = get_user_database()
        user_info = user_db.get_user_by_email(email)
        user_name = user_info.get('name', '') if user_info else ''

        config_suffix = project_service.PROJECT_CONFIG.get('owner_suffix_map', {}).get(user_name, '')

        if config_suffix:
            return jsonify({
                'success': True,
                'suffix': config_suffix,
                'source': 'config'
            })

        # 둘 다 없으면 null
        return jsonify({
            'success': True,
            'suffix': None,
            'source': 'none'
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 프로젝트 코드 접미사 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'message': '프로젝트 코드 접미사 조회 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500

@users_bp.route('/users/<email>/project-code-suffix', methods=['PUT'])
@admin_required
def update_project_code_suffix(email):
    """사용자 프로젝트 코드 접미사 업데이트 (관리자만)"""
    try:
        data = request.get_json()
        suffix = data.get('suffix', '').strip().upper()

        # Validation
        if suffix and (len(suffix) < 1 or len(suffix) > 3):
            return jsonify({
                'success': False,
                'message': '접미사는 1~3자 사이여야 합니다.'
            }), 400

        if suffix and not suffix.isalpha():
            return jsonify({
                'success': False,
                'message': '접미사는 영문자만 입력 가능합니다.'
            }), 400

        from dashboard.utils.redis_client import get_redis_client
        redis_client = get_redis_client()
        redis_key = f"user:project_code_suffix:{email}"

        # 이전 값 조회 (감사 로그용)
        old_suffix = redis_client.get(redis_key) or '-'

        if suffix:
            # 값 설정
            redis_client.set(redis_key, suffix)
            logger.info(f"프로젝트 코드 접미사 설정: {email} → {suffix}")
        else:
            # 값 삭제 (빈 문자열이면 삭제)
            redis_client.delete(redis_key)
            logger.info(f"프로젝트 코드 접미사 삭제: {email}")

        # 감사 로그 기록
        try:
            from dashboard.utils.user_database import get_audit_repository
            audit_repo = get_audit_repository()
            admin_email = session.get('user', {}).get('email', 'unknown')
            audit_repo.log_action(
                user_email=admin_email,
                action='PROJECT_CODE_SUFFIX_UPDATE',
                details=f'사용자 {email}의 프로젝트 코드 접미사를 {suffix or "(삭제)"}로 변경',
                field_name='project_code_suffix',
                old_value=old_suffix,
                new_value=suffix or '(삭제)',
                ip_address=request.remote_addr
            )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")

        return jsonify({
            'success': True,
            'message': '프로젝트 코드 접미사가 성공적으로 업데이트되었습니다.',
            'suffix': suffix or None
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 프로젝트 코드 접미사 업데이트 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'message': '프로젝트 코드 접미사 업데이트 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500

@users_bp.route('/release-all-user-locks', methods=['POST'])
@login_required
def release_all_user_locks():
    """특정 사용자의 모든 잠금 강제 해제 (페이지 새로고침/종료 시)"""
    try:
        data = request.get_json()
        if not data:
            data = {}

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

        # 해제할 잠금들을 먼저 조회 (WebSocket 이벤트를 위해)
        user_locks = []

        try:
            # 사용자의 활성 프로젝트 잠금 조회 (새 시스템)
            from dashboard.utils.project_lock_manager import get_project_lock_manager
            lock_manager = get_project_lock_manager()
            user_locks = lock_manager.get_user_locks(user_email)

            # 모든 사용자 프로젝트 잠금 해제
            released_count = lock_manager.release_all_user_locks(user_email, reason)

            logger.info(f"[LOCK_CLEANUP] 사용자 {user_email}의 {released_count}개 프로젝트 잠금 해제됨. 사유: {reason}")

            # 실시간 브로드캐스트 (WebSocket)
            try:
                from dashboard import socketio
                from datetime import datetime
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

            # 성공 응답
            return jsonify({
                'success': True,
                'message': f'{released_count}개의 잠금이 해제되었습니다.',
                'released_count': released_count,
                'released_locks': [lock.get('project_code') for lock in user_locks]
            })

        except ImportError:
            # project_lock_manager가 없는 경우 기본 응답
            logger.warning("project_lock_manager를 찾을 수 없습니다. 기본 응답을 반환합니다.")
            return jsonify({
                'success': True,
                'message': '잠금 해제 요청이 처리되었습니다.',
                'released_count': 0
            })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 사용자 잠금 해제 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'message': '잠금 해제 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500