"""
필드별 잠금 관리 시스템
동시 편집 방지를 위한 실시간 잠금 기능
"""

import time
import threading
from datetime import datetime, timedelta
from typing import Dict, Optional, Set
from dataclasses import dataclass
import logging

logger = logging.getLogger(__name__)

@dataclass
class FieldLock:
    project_code: str
    field_name: str
    user_email: str
    user_name: str
    locked_at: datetime
    expires_at: datetime

class FieldLockManager:
    def __init__(self, lock_timeout_minutes=5):
        self.locks: Dict[str, FieldLock] = {}  # key: "project_code:field_name"
        self.lock_timeout_minutes = lock_timeout_minutes
        self._lock = threading.RLock()
        
        # 만료된 잠금 정리를 위한 백그라운드 스레드
        self._cleanup_thread = threading.Thread(target=self._cleanup_expired_locks, daemon=True)
        self._cleanup_thread.start()
    
    def _get_lock_key(self, project_code: str, field_name: str) -> str:
        """잠금 키 생성"""
        return f"{project_code}:{field_name}"
    
    def acquire_lock(self, project_code: str, field_name: str, user_email: str, user_name: str) -> Dict:
        """필드 잠금 획득 시도"""
        with self._lock:
            lock_key = self._get_lock_key(project_code, field_name)
            
            # 기존 잠금 확인
            existing_lock = self.locks.get(lock_key)
            
            if existing_lock:
                # 만료된 잠금인지 확인
                if datetime.now() > existing_lock.expires_at:
                    # 만료된 잠금 제거
                    del self.locks[lock_key]
                    logger.info(f"만료된 잠금 제거: {lock_key}")
                else:
                    # 같은 사용자인지 확인
                    if existing_lock.user_email == user_email:
                        # 같은 사용자면 잠금 연장
                        existing_lock.expires_at = datetime.now() + timedelta(minutes=self.lock_timeout_minutes)
                        logger.info(f"잠금 연장: {lock_key} by {user_email}")
                        return {
                            'success': True,
                            'message': '잠금이 연장되었습니다.',
                            'lock_info': self._lock_to_dict(existing_lock)
                        }
                    else:
                        # 다른 사용자가 잠금 중
                        return {
                            'success': False,
                            'message': f'{existing_lock.user_name}님이 수정 중입니다.',
                            'locked_by': existing_lock.user_name,
                            'locked_at': existing_lock.locked_at.isoformat(),
                            'lock_info': self._lock_to_dict(existing_lock)
                        }
            
            # 새 잠금 생성
            new_lock = FieldLock(
                project_code=project_code,
                field_name=field_name,
                user_email=user_email,
                user_name=user_name,
                locked_at=datetime.now(),
                expires_at=datetime.now() + timedelta(minutes=self.lock_timeout_minutes)
            )
            
            self.locks[lock_key] = new_lock
            
            logger.info(f"새 잠금 생성: {lock_key} by {user_email}")
            
            return {
                'success': True,
                'message': '잠금을 획득했습니다.',
                'lock_info': self._lock_to_dict(new_lock)
            }
    
    def release_lock(self, project_code: str, field_name: str, user_email: str) -> Dict:
        """필드 잠금 해제"""
        with self._lock:
            lock_key = self._get_lock_key(project_code, field_name)
            existing_lock = self.locks.get(lock_key)
            
            if not existing_lock:
                return {
                    'success': True,
                    'message': '잠금이 이미 해제되었습니다.'
                }
            
            # 잠금 소유자 확인
            if existing_lock.user_email != user_email:
                return {
                    'success': False,
                    'message': '다른 사용자의 잠금은 해제할 수 없습니다.',
                    'locked_by': existing_lock.user_name
                }
            
            # 잠금 해제
            del self.locks[lock_key]
            logger.info(f"잠금 해제: {lock_key} by {user_email}")
            
            return {
                'success': True,
                'message': '잠금이 해제되었습니다.'
            }
    
    def get_project_locks(self, project_code: str) -> Dict[str, Dict]:
        """특정 프로젝트의 모든 잠금 조회"""
        with self._lock:
            project_locks = {}
            
            for lock_key, lock in self.locks.items():
                if lock.project_code == project_code:
                    # 만료 확인
                    if datetime.now() <= lock.expires_at:
                        project_locks[lock.field_name] = self._lock_to_dict(lock)
            
            return project_locks
    
    def get_all_locks(self) -> Dict[str, Dict]:
        """모든 활성 잠금 조회"""
        with self._lock:
            active_locks = {}
            
            for lock_key, lock in self.locks.items():
                # 만료 확인
                if datetime.now() <= lock.expires_at:
                    active_locks[lock_key] = self._lock_to_dict(lock)
            
            return active_locks
    
    def force_release_user_locks(self, user_email: str) -> int:
        """특정 사용자의 모든 잠금 강제 해제 (로그아웃 시 등)"""
        with self._lock:
            keys_to_remove = []
            
            for lock_key, lock in self.locks.items():
                if lock.user_email == user_email:
                    keys_to_remove.append(lock_key)
            
            for key in keys_to_remove:
                del self.locks[key]
                logger.info(f"사용자 잠금 강제 해제: {key} for {user_email}")
            
            return len(keys_to_remove)
    
    def _lock_to_dict(self, lock: FieldLock) -> Dict:
        """FieldLock을 딕셔너리로 변환"""
        return {
            'project_code': lock.project_code,
            'field_name': lock.field_name,
            'user_email': lock.user_email,
            'user_name': lock.user_name,
            'locked_at': lock.locked_at.isoformat(),
            'expires_at': lock.expires_at.isoformat(),
            'remaining_minutes': max(0, (lock.expires_at - datetime.now()).total_seconds() / 60)
        }
    
    def _cleanup_expired_locks(self):
        """만료된 잠금 정리 (백그라운드 스레드)"""
        while True:
            try:
                with self._lock:
                    current_time = datetime.now()
                    expired_keys = []
                    
                    for lock_key, lock in self.locks.items():
                        if current_time > lock.expires_at:
                            expired_keys.append(lock_key)
                    
                    for key in expired_keys:
                        del self.locks[key]
                        logger.info(f"만료된 잠금 자동 정리: {key}")
                
                # 30초마다 정리
                time.sleep(30)
                
            except Exception as e:
                logger.error(f"잠금 정리 중 오류: {e}")
                time.sleep(60)  # 오류 시 1분 대기

# 전역 인스턴스
field_lock_manager = FieldLockManager()