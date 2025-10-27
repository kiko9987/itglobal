"""
프로젝트 잠금 관리 시스템
통합 편집 모드에서 프로젝트 전체 단위 잠금 제공
"""

import os
import time
import threading
from datetime import datetime, timedelta
from typing import Dict, Optional, List
from dataclasses import dataclass, asdict
import logging

logger = logging.getLogger(__name__)

# Grace Period: 잠금 만료 후 원래 사용자가 재획득할 수 있는 유예 시간 (초)
# 네트워크 지연으로 인한 Lock 연장 실패를 방지
GRACE_PERIOD_SECONDS = 30

@dataclass
class ProjectLock:
    """프로젝트 잠금 정보"""
    project_code: str
    user_email: str
    user_name: str
    tab_id: str
    locked_at: datetime
    expires_at: datetime

    def to_dict(self) -> dict:
        """딕셔너리로 변환"""
        return {
            'project_code': self.project_code,
            'user_email': self.user_email,
            'user_name': self.user_name,
            'tab_id': self.tab_id,
            'locked_at': self.locked_at.isoformat(),
            'expires_at': self.expires_at.isoformat(),
            'remaining_minutes': max(0, (self.expires_at - datetime.now()).total_seconds() / 60)
        }

class ProjectLockManager:
    """프로젝트 단위 잠금 관리자"""

    def __init__(self, lock_timeout_minutes=None):
        """
        Args:
            lock_timeout_minutes: 잠금 자동 만료 시간 (기본 5분, 환경 변수로 설정 가능)
        """
        # 환경 변수에서 타임아웃 값 가져오기 (기본 5분)
        if lock_timeout_minutes is None:
            lock_timeout_minutes = int(os.getenv('LOCK_TIMEOUT_MINUTES', 5))

        self.locks: Dict[str, ProjectLock] = {}  # key: project_code
        self.lock_timeout_minutes = lock_timeout_minutes
        self._lock = threading.RLock()
        self._shutdown_event = threading.Event()  # 종료 이벤트

        # 만료된 잠금 정리를 위한 백그라운드 스레드
        self._cleanup_thread = threading.Thread(target=self._cleanup_expired_locks, daemon=True)
        self._cleanup_thread.start()

        logger.info(f"ProjectLockManager 초기화 완료 (타임아웃: {lock_timeout_minutes}분)")

    def acquire_lock(self, project_code: str, user_email: str, user_name: str, tab_id: str) -> Dict:
        """
        프로젝트 잠금 획득 시도

        Args:
            project_code: 프로젝트 코드
            user_email: 사용자 이메일
            user_name: 사용자 이름
            tab_id: 탭 ID (다중 탭 구분용)

        Returns:
            {
                'success': bool,
                'message': str,
                'lock_info': dict (optional)
            }
        """
        with self._lock:
            existing_lock = self.locks.get(project_code)

            if existing_lock:
                current_time = datetime.now()

                # 만료된 잠금인지 확인
                if current_time > existing_lock.expires_at:
                    # Grace Period 확인: 만료 후 일정 시간 내에 원래 사용자가 재획득 가능
                    grace_deadline = existing_lock.expires_at + timedelta(seconds=GRACE_PERIOD_SECONDS)

                    if current_time <= grace_deadline and existing_lock.user_email == user_email and existing_lock.tab_id == tab_id:
                        # 원래 사용자가 Grace Period 내에 재획득 시도 - 잠금 복구
                        existing_lock.expires_at = current_time + timedelta(minutes=self.lock_timeout_minutes)
                        logger.info(f"Grace Period 내 잠금 복구: {project_code} by {user_email} (탭 ID: {tab_id[:8]}...)")
                        return {
                            'success': True,
                            'message': '잠금이 복구되었습니다.',
                            'recovered': True,
                            'lock_info': existing_lock.to_dict()
                        }
                    else:
                        # Grace Period 만료 또는 다른 사용자 - 잠금 제거
                        del self.locks[project_code]
                        logger.info(f"만료된 잠금 제거: {project_code}")
                else:
                    # 같은 사용자 + 같은 탭인지 확인
                    if existing_lock.user_email == user_email and existing_lock.tab_id == tab_id:
                        # 같은 사용자 + 같은 탭이면 잠금 연장
                        existing_lock.expires_at = datetime.now() + timedelta(minutes=self.lock_timeout_minutes)
                        logger.info(f"잠금 연장: {project_code} by {user_email} (탭 ID: {tab_id[:8]}...)")
                        return {
                            'success': True,
                            'message': '잠금이 연장되었습니다.',
                            'lock_info': existing_lock.to_dict()
                        }
                    elif existing_lock.user_email == user_email:
                        # 같은 사용자지만 다른 탭
                        return {
                            'success': False,
                            'message': f'다른 탭에서 이미 편집 중입니다. 해당 탭으로 이동하세요.',
                            'locked_by': existing_lock.user_name,
                            'locked_by_email': existing_lock.user_email,
                            'same_user': True,
                            'lock_info': existing_lock.to_dict()
                        }
                    else:
                        # 다른 사용자가 잠금 중
                        return {
                            'success': False,
                            'message': f'{existing_lock.user_name}님이 편집 중입니다.',
                            'locked_by': existing_lock.user_name,
                            'locked_by_email': existing_lock.user_email,
                            'same_user': False,
                            'lock_info': existing_lock.to_dict()
                        }

            # 새 잠금 생성
            new_lock = ProjectLock(
                project_code=project_code,
                user_email=user_email,
                user_name=user_name,
                tab_id=tab_id,
                locked_at=datetime.now(),
                expires_at=datetime.now() + timedelta(minutes=self.lock_timeout_minutes)
            )

            self.locks[project_code] = new_lock
            logger.info(f"새 잠금 생성: {project_code} by {user_email} ({user_name}) - 탭 ID: {tab_id[:8]}...")

            return {
                'success': True,
                'message': '편집 모드로 전환되었습니다.',
                'lock_info': new_lock.to_dict()
            }

    def release_lock(self, project_code: str, user_email: str, tab_id: str) -> Dict:
        """
        프로젝트 잠금 해제

        Args:
            project_code: 프로젝트 코드
            user_email: 사용자 이메일
            tab_id: 탭 ID (잠금 획득 시 사용한 탭 ID와 일치해야 함)

        Returns:
            {'success': bool, 'message': str}
        """
        with self._lock:
            existing_lock = self.locks.get(project_code)

            if not existing_lock:
                return {
                    'success': True,
                    'message': '잠금이 이미 해제되었습니다.'
                }

            # 잠금 소유자 확인 (사용자 + 탭 ID 모두 일치해야 함)
            if existing_lock.user_email != user_email:
                return {
                    'success': False,
                    'message': '다른 사용자의 잠금은 해제할 수 없습니다.',
                    'locked_by': existing_lock.user_name
                }

            # 같은 사용자지만 다른 탭에서 해제 시도
            if existing_lock.tab_id != tab_id:
                return {
                    'success': False,
                    'message': '다른 탭에서 획득한 잠금은 해제할 수 없습니다. 해당 탭으로 이동하세요.',
                    'locked_by': existing_lock.user_name,
                    'same_user': True
                }

            # 잠금 해제 (사용자 + 탭 ID 모두 일치)
            del self.locks[project_code]
            logger.info(f"잠금 해제: {project_code} by {user_email} (탭 ID: {tab_id[:8]}...)")

            return {
                'success': True,
                'message': '편집 모드가 종료되었습니다.'
            }

    def force_release_lock(self, project_code: str, admin_email: str, admin_permission: str) -> Dict:
        """
        관리자 권한으로 강제 잠금 해제

        Args:
            project_code: 프로젝트 코드
            admin_email: 관리자 이메일
            admin_permission: 관리자 권한 레벨

        Returns:
            {'success': bool, 'message': str}
        """
        with self._lock:
            # 관리자 권한 확인
            if admin_permission not in ['admin', 'super_admin']:
                return {
                    'success': False,
                    'message': '관리자 권한이 필요합니다.'
                }

            existing_lock = self.locks.get(project_code)

            if not existing_lock:
                return {
                    'success': True,
                    'message': '잠금이 이미 해제되어 있습니다.'
                }

            locked_by_name = existing_lock.user_name
            del self.locks[project_code]

            logger.warning(f"[ADMIN] 강제 잠금 해제: {project_code} (원 소유자: {locked_by_name}) by 관리자 {admin_email}")

            return {
                'success': True,
                'message': f'{locked_by_name}님의 잠금이 해제되었습니다.',
                'forced': True,
                'previous_owner': locked_by_name
            }

    def get_lock_status(self, project_code: str) -> Optional[Dict]:
        """
        프로젝트 잠금 상태 조회

        Returns:
            잠금 정보 딕셔너리 또는 None (잠금 없음)
        """
        with self._lock:
            existing_lock = self.locks.get(project_code)

            if not existing_lock:
                return None

            # 만료 확인
            if datetime.now() > existing_lock.expires_at:
                del self.locks[project_code]
                logger.info(f"만료된 잠금 제거: {project_code}")
                return None

            return existing_lock.to_dict()

    def get_all_locks(self) -> List[Dict]:
        """
        모든 활성 잠금 조회

        Returns:
            활성 잠금 리스트
        """
        with self._lock:
            active_locks = []

            for project_code, lock in list(self.locks.items()):
                # 만료 확인
                if datetime.now() <= lock.expires_at:
                    active_locks.append(lock.to_dict())
                else:
                    # 만료된 잠금 제거
                    del self.locks[project_code]

            return active_locks

    def get_user_locks(self, user_email: str) -> List[Dict]:
        """
        특정 사용자의 모든 잠금 조회

        Returns:
            사용자의 활성 잠금 리스트
        """
        with self._lock:
            user_locks = []

            for lock in self.locks.values():
                if lock.user_email == user_email and datetime.now() <= lock.expires_at:
                    user_locks.append(lock.to_dict())

            return user_locks

    def release_all_user_locks(self, user_email: str, reason: str = 'manual_cleanup') -> int:
        """
        특정 사용자의 모든 잠금 해제

        Args:
            user_email: 사용자 이메일
            reason: 해제 사유

        Returns:
            해제된 잠금 개수
        """
        with self._lock:
            project_codes_to_remove = []

            for project_code, lock in self.locks.items():
                if lock.user_email == user_email:
                    project_codes_to_remove.append(project_code)

            for project_code in project_codes_to_remove:
                del self.locks[project_code]
                logger.info(f"사용자 잠금 해제: {project_code} for {user_email} (사유: {reason})")

            return len(project_codes_to_remove)

    def shutdown(self):
        """매니저 종료 (백그라운드 스레드 정리)"""
        logger.info("ProjectLockManager 종료 중...")
        self._shutdown_event.set()
        if self._cleanup_thread.is_alive():
            self._cleanup_thread.join(timeout=5)
        logger.info("ProjectLockManager 종료 완료")

    def _cleanup_expired_locks(self):
        """만료된 잠금 정리 (백그라운드 스레드)"""
        while not self._shutdown_event.is_set():
            try:
                with self._lock:
                    current_time = datetime.now()
                    expired_codes = []

                    for project_code, lock in self.locks.items():
                        if current_time > lock.expires_at:
                            expired_codes.append(project_code)

                    for code in expired_codes:
                        del self.locks[code]
                        logger.info(f"만료된 잠금 자동 정리: {code}")

                # 1분마다 정리 (또는 종료 이벤트 대기)
                self._shutdown_event.wait(60)

            except Exception as e:
                logger.error(f"잠금 정리 중 오류: {e}")
                self._shutdown_event.wait(60)

        logger.info("잠금 정리 스레드 종료")

# 전역 싱글톤 인스턴스
_lock_manager_instance = None

def get_project_lock_manager() -> ProjectLockManager:
    """프로젝트 잠금 매니저 싱글톤 인스턴스 반환"""
    global _lock_manager_instance
    if _lock_manager_instance is None:
        _lock_manager_instance = ProjectLockManager()
    return _lock_manager_instance
