"""
프로젝트 잠금 관리 시스템
통합 편집 모드에서 프로젝트 전체 단위 잠금 제공

변경 이력:
- 2025-01: Dict → Redis 전환 (분산 환경 지원)
- Redis TTL 자동 만료, Grace Period 유지
"""

import os
import json
import threading
from datetime import datetime, timedelta
from typing import Dict, Optional, List
from dataclasses import dataclass, asdict
import logging

from dashboard.utils.redis_client import get_redis_client
from dashboard.utils.exceptions import ServiceUnavailable

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
    """프로젝트 단위 잠금 관리자 (Redis 기반)

    변경사항:
    - Dict → Redis로 전환
    - TTL은 Redis가 자동 관리
    - 정리 스레드 제거 (Redis 자동 만료)
    - Grace Period는 별도 키로 저장
    """

    def __init__(self, lock_timeout_minutes=None):
        """
        Args:
            lock_timeout_minutes: 잠금 자동 만료 시간 (기본 5분, 환경 변수로 설정 가능)
        """
        # 환경 변수에서 타임아웃 값 가져오기 (기본 5분)
        if lock_timeout_minutes is None:
            lock_timeout_minutes = int(os.getenv('LOCK_TIMEOUT_MINUTES', 5))

        self.redis = get_redis_client()
        self.lock_timeout_minutes = lock_timeout_minutes
        self._lock = threading.RLock()  # 로컬 동기화용

        logger.info(f"ProjectLockManager 초기화 완료 (타임아웃: {lock_timeout_minutes}분, Redis 기반)")

    def _get_lock_from_redis(self, project_code: str) -> Optional[ProjectLock]:
        """Redis에서 잠금 정보 조회 (내부 메서드)"""
        try:
            lock_key = f"lock:{project_code}"
            lock_data = self.redis.get(lock_key)

            if lock_data is None:
                return None

            # JSON 역직렬화하여 ProjectLock 객체 생성
            return ProjectLock(
                project_code=lock_data['project_code'],
                user_email=lock_data['user_email'],
                user_name=lock_data['user_name'],
                tab_id=lock_data['tab_id'],
                locked_at=datetime.fromisoformat(lock_data['locked_at']),
                expires_at=datetime.fromisoformat(lock_data['expires_at'])
            )
        except (ServiceUnavailable, KeyError, ValueError) as e:
            logger.error(f"Redis에서 잠금 조회 실패: {project_code}, error={e}")
            return None

    def _set_lock_in_redis(self, lock: ProjectLock) -> bool:
        """Redis에 잠금 정보 저장 (내부 메서드)"""
        try:
            lock_key = f"lock:{lock.project_code}"

            # TTL 계산 (초 단위)
            ttl_seconds = int((lock.expires_at - datetime.now()).total_seconds())
            if ttl_seconds <= 0:
                logger.warning(f"잠금 TTL이 0 이하: {lock.project_code}")
                return False

            # ProjectLock를 JSON 직렬화 가능한 dict로 변환
            lock_data = {
                'project_code': lock.project_code,
                'user_email': lock.user_email,
                'user_name': lock.user_name,
                'tab_id': lock.tab_id,
                'locked_at': lock.locked_at.isoformat(),
                'expires_at': lock.expires_at.isoformat()
            }

            # Redis에 저장 (TTL 설정)
            self.redis.set(lock_key, lock_data, ex=ttl_seconds)

            # Grace Period 마커 저장 (잠금 만료 후에도 GRACE_PERIOD_SECONDS 동안 유지)
            grace_key = f"grace:{lock.project_code}"
            grace_data = {
                'user_email': lock.user_email,
                'tab_id': lock.tab_id,
                'lock_expired_at': lock.expires_at.isoformat()
            }
            grace_ttl = ttl_seconds + GRACE_PERIOD_SECONDS
            self.redis.set(grace_key, grace_data, ex=grace_ttl)

            return True

        except ServiceUnavailable as e:
            logger.error(f"Redis에 잠금 저장 실패: {lock.project_code}, error={e}")
            raise

    def _get_grace_info(self, project_code: str) -> Optional[Dict]:
        """Grace Period 정보 조회 (내부 메서드)"""
        try:
            grace_key = f"grace:{project_code}"
            grace_data = self.redis.get(grace_key)
            return grace_data
        except ServiceUnavailable as e:
            logger.error(f"Grace Period 조회 실패: {project_code}, error={e}")
            return None

    def acquire_lock(self, project_code: str, user_email: str, user_name: str, tab_id: str) -> Dict:
        """
        프로젝트 잠금 획득 시도 (Redis 기반)

        Note (2026-07-07): Grace period 마커 조회와 main 잠금 설정이 원자적이지 않지만
        `with self._lock:` (threading.Lock) 으로 프로세스 내 순차 처리 보장. Waitress
        단일 프로세스 아키텍처에서는 실질 race 없음. 멀티프로세스 확장 시점에 Lua
        스크립트로 원자화 검토.

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
            try:
                # Redis에서 현재 잠금 조회
                existing_lock = self._get_lock_from_redis(project_code)

                if existing_lock:
                    # 잠금이 존재하는 경우
                    # 같은 사용자 + 같은 탭인지 확인
                    if existing_lock.user_email == user_email and existing_lock.tab_id == tab_id:
                        # 같은 사용자 + 같은 탭이면 잠금 연장
                        existing_lock.expires_at = datetime.now() + timedelta(minutes=self.lock_timeout_minutes)
                        self._set_lock_in_redis(existing_lock)
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

                # 잠금이 없는 경우 - Grace Period 확인
                grace_info = self._get_grace_info(project_code)
                if grace_info:
                    # Grace Period 마커가 존재
                    if grace_info['user_email'] == user_email and grace_info['tab_id'] == tab_id:
                        # 원래 사용자가 Grace Period 내에 재획득 시도 - 잠금 복구
                        recovered_lock = ProjectLock(
                            project_code=project_code,
                            user_email=user_email,
                            user_name=user_name,
                            tab_id=tab_id,
                            locked_at=datetime.now(),
                            expires_at=datetime.now() + timedelta(minutes=self.lock_timeout_minutes)
                        )
                        self._set_lock_in_redis(recovered_lock)
                        logger.info(f"Grace Period 내 잠금 복구: {project_code} by {user_email} (탭 ID: {tab_id[:8]}...)")
                        return {
                            'success': True,
                            'message': '잠금이 복구되었습니다.',
                            'recovered': True,
                            'lock_info': recovered_lock.to_dict()
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

                self._set_lock_in_redis(new_lock)
                logger.info(f"새 잠금 생성: {project_code} by {user_email} ({user_name}) - 탭 ID: {tab_id[:8]}...")

                return {
                    'success': True,
                    'message': '편집 모드로 전환되었습니다.',
                    'lock_info': new_lock.to_dict()
                }

            except ServiceUnavailable as e:
                logger.error(f"Redis 장애로 잠금 획득 실패: {project_code}, error={e}")
                return {
                    'success': False,
                    'message': '일시적인 서버 오류입니다. 잠시 후 다시 시도해주세요.',
                    'error': str(e)
                }

    def release_lock(self, project_code: str, user_email: str, tab_id: str) -> Dict:
        """
        프로젝트 잠금 해제 (Redis 기반)

        Args:
            project_code: 프로젝트 코드
            user_email: 사용자 이메일
            tab_id: 탭 ID (잠금 획득 시 사용한 탭 ID와 일치해야 함)

        Returns:
            {'success': bool, 'message': str}
        """
        with self._lock:
            try:
                existing_lock = self._get_lock_from_redis(project_code)

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
                lock_key = f"lock:{project_code}"
                self.redis.delete(lock_key)
                logger.info(f"잠금 해제: {project_code} by {user_email} (탭 ID: {tab_id[:8]}...)")

                return {
                    'success': True,
                    'message': '편집 모드가 종료되었습니다.'
                }

            except ServiceUnavailable as e:
                logger.error(f"Redis 장애로 잠금 해제 실패: {project_code}, error={e}")
                return {
                    'success': False,
                    'message': '일시적인 서버 오류입니다. 잠시 후 다시 시도해주세요.',
                    'error': str(e)
                }

    def force_release_lock(self, project_code: str, admin_email: str, admin_permission: str) -> Dict:
        """
        관리자 권한으로 강제 잠금 해제 (Redis 기반)

        Args:
            project_code: 프로젝트 코드
            admin_email: 관리자 이메일
            admin_permission: 관리자 권한 레벨

        Returns:
            {'success': bool, 'message': str}
        """
        with self._lock:
            try:
                # 관리자 권한 확인
                if admin_permission not in ['admin', 'super_admin']:
                    return {
                        'success': False,
                        'message': '관리자 권한이 필요합니다.'
                    }

                existing_lock = self._get_lock_from_redis(project_code)

                if not existing_lock:
                    return {
                        'success': True,
                        'message': '잠금이 이미 해제되어 있습니다.'
                    }

                locked_by_name = existing_lock.user_name
                lock_key = f"lock:{project_code}"
                self.redis.delete(lock_key)

                logger.warning(f"[ADMIN] 강제 잠금 해제: {project_code} (원 소유자: {locked_by_name}) by 관리자 {admin_email}")

                return {
                    'success': True,
                    'message': f'{locked_by_name}님의 잠금이 해제되었습니다.',
                    'forced': True,
                    'previous_owner': locked_by_name
                }

            except ServiceUnavailable as e:
                logger.error(f"Redis 장애로 강제 해제 실패: {project_code}, error={e}")
                return {
                    'success': False,
                    'message': '일시적인 서버 오류입니다.',
                    'error': str(e)
                }

    def get_lock_status(self, project_code: str) -> Optional[Dict]:
        """
        프로젝트 잠금 상태 조회 (Redis 기반)

        Returns:
            잠금 정보 딕셔너리 또는 None (잠금 없음)

        Note:
            Redis TTL 자동 만료로 만료 확인 불필요
        """
        try:
            existing_lock = self._get_lock_from_redis(project_code)

            if not existing_lock:
                return None

            return existing_lock.to_dict()

        except ServiceUnavailable as e:
            logger.error(f"Redis 장애로 잠금 상태 조회 실패: {project_code}, error={e}")
            return None

    def get_all_locks(self) -> List[Dict]:
        """
        모든 활성 잠금 조회 (Redis 기반)

        Returns:
            활성 잠금 리스트

        Note:
            Redis KEYS 명령 사용 (O(N)) - 대량 키가 있을 때 주의
        """
        try:
            active_locks = []
            # 2026-07-08: KEYS(O(N) blocking) → SCAN(non-blocking iterator)으로 교체.
            # 20명 동시 사용 대비 Redis 블로킹 회피.
            lock_keys = list(self.redis.scan_iter(match="lock:*", count=100))

            for lock_key in lock_keys:
                # lock:PROJECT_CODE 형식에서 project_code 추출
                project_code = lock_key.replace("lock:", "", 1)
                lock = self._get_lock_from_redis(project_code)

                if lock:
                    active_locks.append(lock.to_dict())

            return active_locks

        except ServiceUnavailable as e:
            logger.error(f"Redis 장애로 전체 잠금 조회 실패: error={e}")
            return []

    def get_user_locks(self, user_email: str) -> List[Dict]:
        """
        특정 사용자의 모든 잠금 조회 (Redis 기반)

        Returns:
            사용자의 활성 잠금 리스트
        """
        try:
            user_locks = []
            # 2026-07-08: KEYS(O(N) blocking) → SCAN(non-blocking iterator)으로 교체.
            # 20명 동시 사용 대비 Redis 블로킹 회피.
            lock_keys = list(self.redis.scan_iter(match="lock:*", count=100))

            for lock_key in lock_keys:
                project_code = lock_key.replace("lock:", "", 1)
                lock = self._get_lock_from_redis(project_code)

                if lock and lock.user_email == user_email:
                    user_locks.append(lock.to_dict())

            return user_locks

        except ServiceUnavailable as e:
            logger.error(f"Redis 장애로 사용자 잠금 조회 실패: {user_email}, error={e}")
            return []

    def release_all_user_locks(self, user_email: str, reason: str = 'manual_cleanup') -> int:
        """
        특정 사용자의 모든 잠금 해제 (Redis 기반)

        Args:
            user_email: 사용자 이메일
            reason: 해제 사유

        Returns:
            해제된 잠금 개수
        """
        try:
            project_codes_to_remove = []
            # 2026-07-08: KEYS(O(N) blocking) → SCAN(non-blocking iterator)으로 교체.
            # 20명 동시 사용 대비 Redis 블로킹 회피.
            lock_keys = list(self.redis.scan_iter(match="lock:*", count=100))

            # 해당 사용자의 잠금 찾기
            for lock_key in lock_keys:
                project_code = lock_key.replace("lock:", "", 1)
                lock = self._get_lock_from_redis(project_code)

                if lock and lock.user_email == user_email:
                    project_codes_to_remove.append(project_code)

            # 잠금 해제
            for project_code in project_codes_to_remove:
                lock_key = f"lock:{project_code}"
                self.redis.delete(lock_key)
                logger.info(f"사용자 잠금 해제: {project_code} for {user_email} (사유: {reason})")

            return len(project_codes_to_remove)

        except ServiceUnavailable as e:
            logger.error(f"Redis 장애로 사용자 잠금 해제 실패: {user_email}, error={e}")
            return 0

    def shutdown(self):
        """매니저 종료 (Redis 기반에서는 정리 불필요)

        Note:
            Redis 전환 후 백그라운드 스레드 없음 (Redis TTL 자동 관리)
        """
        logger.info("ProjectLockManager 종료 완료 (Redis 기반 - 정리 불필요)")

# 전역 싱글톤 인스턴스
_lock_manager_instance = None

def get_project_lock_manager() -> ProjectLockManager:
    """프로젝트 잠금 매니저 싱글톤 인스턴스 반환"""
    global _lock_manager_instance
    if _lock_manager_instance is None:
        _lock_manager_instance = ProjectLockManager()
    return _lock_manager_instance
