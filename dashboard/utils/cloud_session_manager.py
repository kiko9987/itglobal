"""
클라우드 세션 관리자
- 멀티 인스턴스 세션 공유
- Redis 기반 세션 스토어
- 세션 보안 강화
"""

import os
import json
import hashlib
import time
import uuid
from datetime import datetime, timedelta
from typing import Dict, Any, Optional
import logging
from flask import session, request, g
import redis
from .cloud_database_manager import get_cloud_db_manager

logger = logging.getLogger(__name__)

class CloudSessionManager:
    """클라우드 환경 세션 관리자"""

    def __init__(self):
        self.redis_client = None
        self.db_manager = None
        self.session_timeout = 28800  # 8시간
        self.init_storage()

    def init_storage(self):
        """세션 저장소 초기화"""
        try:
            # Redis 연결 (1순위)
            redis_host = os.getenv('CLOUD_REDIS_HOST', 'localhost')
            redis_port = int(os.getenv('CLOUD_REDIS_PORT', 6379))
            redis_auth = os.getenv('CLOUD_REDIS_AUTH')

            self.redis_client = redis.Redis(
                host=redis_host,
                port=redis_port,
                password=redis_auth,
                db=1,  # 세션 전용 DB
                decode_responses=True,
                socket_connect_timeout=3,
                socket_timeout=3
            )

            # 연결 테스트
            self.redis_client.ping()
            logger.info("Redis 세션 스토어 연결 성공")

        except Exception as e:
            logger.warning(f"Redis 연결 실패, 데이터베이스 폴백: {e}")
            self.redis_client = None

        # 데이터베이스 매니저 (2순위 폴백)
        try:
            self.db_manager = get_cloud_db_manager()
        except Exception as e:
            logger.error(f"데이터베이스 연결 실패: {e}")

    def generate_session_id(self, user_id: str = None) -> str:
        """안전한 세션 ID 생성"""
        timestamp = str(int(time.time()))
        random_part = str(uuid.uuid4())
        client_info = request.remote_addr if request else 'unknown'

        data = f"{user_id}:{timestamp}:{random_part}:{client_info}"
        session_id = hashlib.sha256(data.encode()).hexdigest()

        return f"sess_{session_id[:32]}"

    def save_session(self, session_id: str, user_id: str, session_data: Dict[str, Any]) -> bool:
        """세션 데이터 저장"""
        try:
            # 세션 메타데이터 추가
            enhanced_data = {
                **session_data,
                '_session_id': session_id,
                '_user_id': user_id,
                '_created_at': datetime.now().isoformat(),
                '_last_accessed': datetime.now().isoformat(),
                '_ip_address': request.remote_addr if request else 'unknown',
                '_user_agent': request.headers.get('User-Agent', '') if request else ''
            }

            expires_at = datetime.now() + timedelta(seconds=self.session_timeout)

            # Redis 저장 (1순위)
            if self.redis_client:
                try:
                    session_key = f"session:{session_id}"
                    self.redis_client.hset(session_key, mapping={
                        'data': json.dumps(enhanced_data, default=str),
                        'user_id': user_id,
                        'expires_at': expires_at.isoformat()
                    })
                    self.redis_client.expire(session_key, self.session_timeout)

                    logger.debug(f"Redis 세션 저장: {session_id}")
                    return True

                except Exception as redis_error:
                    logger.warning(f"Redis 세션 저장 실패: {redis_error}")

            # 데이터베이스 저장 (폴백)
            if self.db_manager:
                return self.db_manager.save_session(
                    session_id, user_id, enhanced_data, expires_at
                )

            return False

        except Exception as e:
            logger.error(f"세션 저장 실패: {e}")
            return False

    def load_session(self, session_id: str) -> Optional[Dict[str, Any]]:
        """세션 데이터 로드"""
        if not session_id:
            return None

        try:
            # Redis 조회 (1순위)
            if self.redis_client:
                try:
                    session_key = f"session:{session_id}"
                    session_hash = self.redis_client.hgetall(session_key)

                    if session_hash and 'data' in session_hash:
                        # 만료 시간 확인
                        expires_at = datetime.fromisoformat(session_hash['expires_at'])
                        if datetime.now() > expires_at:
                            self.delete_session(session_id)
                            return None

                        # 마지막 접근 시간 업데이트
                        self._update_last_accessed(session_id)

                        session_data = json.loads(session_hash['data'])
                        logger.debug(f"Redis 세션 로드: {session_id}")
                        return session_data

                except Exception as redis_error:
                    logger.warning(f"Redis 세션 조회 실패: {redis_error}")

            # 데이터베이스 조회 (폴백)
            if self.db_manager:
                session_data = self.db_manager.get_session(session_id)
                if session_data:
                    self._update_last_accessed(session_id)
                    return session_data

            return None

        except Exception as e:
            logger.error(f"세션 로드 실패: {e}")
            return None

    def delete_session(self, session_id: str) -> bool:
        """세션 삭제"""
        try:
            success = False

            # Redis 삭제
            if self.redis_client:
                try:
                    session_key = f"session:{session_id}"
                    result = self.redis_client.delete(session_key)
                    success = bool(result)
                except Exception as redis_error:
                    logger.warning(f"Redis 세션 삭제 실패: {redis_error}")

            # 데이터베이스 삭제
            if self.db_manager:
                try:
                    with self.db_manager.Session() as db_session:
                        db_session.execute(
                            "DELETE FROM user_sessions WHERE session_id = :session_id",
                            {"session_id": session_id}
                        )
                        db_session.commit()
                        success = True
                except Exception as db_error:
                    logger.warning(f"DB 세션 삭제 실패: {db_error}")

            if success:
                logger.debug(f"세션 삭제: {session_id}")

            return success

        except Exception as e:
            logger.error(f"세션 삭제 실패: {e}")
            return False

    def _update_last_accessed(self, session_id: str):
        """마지막 접근 시간 업데이트"""
        try:
            current_time = datetime.now().isoformat()

            if self.redis_client:
                session_key = f"session:{session_id}"
                # 세션 데이터 업데이트
                session_data_str = self.redis_client.hget(session_key, 'data')
                if session_data_str:
                    session_data = json.loads(session_data_str)
                    session_data['_last_accessed'] = current_time
                    self.redis_client.hset(session_key, 'data', json.dumps(session_data, default=str))

        except Exception as e:
            logger.debug(f"마지막 접근 시간 업데이트 실패: {e}")

    def cleanup_expired_sessions(self) -> int:
        """만료된 세션 정리"""
        cleaned_count = 0

        try:
            # Redis 정리 (TTL 기반이므로 자동 정리됨)
            if self.redis_client:
                pattern = "session:*"
                expired_keys = []

                for key in self.redis_client.scan_iter(match=pattern):
                    if self.redis_client.ttl(key) == -1:  # TTL이 설정되지 않은 경우
                        expired_keys.append(key)

                if expired_keys:
                    self.redis_client.delete(*expired_keys)
                    cleaned_count += len(expired_keys)

            # 데이터베이스 정리
            if self.db_manager:
                db_cleaned = self.db_manager.cleanup_expired_data()
                cleaned_count += db_cleaned

            if cleaned_count > 0:
                logger.info(f"만료된 세션 정리: {cleaned_count}개")

        except Exception as e:
            logger.error(f"세션 정리 실패: {e}")

        return cleaned_count

    def get_active_sessions(self, user_id: str = None) -> list:
        """활성 세션 목록 조회"""
        sessions = []

        try:
            if self.redis_client:
                pattern = "session:*"
                for key in self.redis_client.scan_iter(match=pattern):
                    session_hash = self.redis_client.hgetall(key)
                    if session_hash and 'data' in session_hash:
                        session_data = json.loads(session_hash['data'])

                        # 특정 사용자 필터링
                        if user_id and session_data.get('_user_id') != user_id:
                            continue

                        sessions.append({
                            'session_id': session_data.get('_session_id'),
                            'user_id': session_data.get('_user_id'),
                            'created_at': session_data.get('_created_at'),
                            'last_accessed': session_data.get('_last_accessed'),
                            'ip_address': session_data.get('_ip_address'),
                            'user_agent': session_data.get('_user_agent', '')[:100]
                        })

        except Exception as e:
            logger.error(f"활성 세션 조회 실패: {e}")

        return sessions

    def revoke_user_sessions(self, user_id: str, except_session_id: str = None) -> int:
        """사용자의 모든 세션 삭제 (현재 세션 제외)"""
        revoked_count = 0

        try:
            active_sessions = self.get_active_sessions(user_id)

            for session_info in active_sessions:
                session_id = session_info['session_id']
                if session_id != except_session_id:
                    if self.delete_session(session_id):
                        revoked_count += 1

            logger.info(f"사용자 세션 삭제: {user_id} - {revoked_count}개")

        except Exception as e:
            logger.error(f"사용자 세션 삭제 실패: {e}")

        return revoked_count

    def extend_session(self, session_id: str, additional_seconds: int = None) -> bool:
        """세션 유효기간 연장"""
        if additional_seconds is None:
            additional_seconds = self.session_timeout

        try:
            if self.redis_client:
                session_key = f"session:{session_id}"
                return bool(self.redis_client.expire(session_key, additional_seconds))

        except Exception as e:
            logger.error(f"세션 연장 실패: {e}")

        return False

    def get_session_stats(self) -> Dict[str, Any]:
        """세션 통계"""
        stats = {
            'total_sessions': 0,
            'unique_users': set(),
            'storage_type': 'unknown'
        }

        try:
            if self.redis_client:
                stats['storage_type'] = 'redis'
                pattern = "session:*"
                session_keys = list(self.redis_client.scan_iter(match=pattern))
                stats['total_sessions'] = len(session_keys)

                # 고유 사용자 수 계산
                for key in session_keys:
                    session_hash = self.redis_client.hgetall(key)
                    if 'user_id' in session_hash:
                        stats['unique_users'].add(session_hash['user_id'])

            elif self.db_manager:
                stats['storage_type'] = 'database'
                # 데이터베이스에서 통계 조회
                # (구현 필요)

            stats['unique_users'] = len(stats['unique_users'])

        except Exception as e:
            logger.error(f"세션 통계 조회 실패: {e}")

        return stats


# 전역 인스턴스
cloud_session_manager = None

def get_cloud_session_manager() -> CloudSessionManager:
    """클라우드 세션 매니저 인스턴스 반환"""
    global cloud_session_manager
    if cloud_session_manager is None:
        cloud_session_manager = CloudSessionManager()
    return cloud_session_manager

def init_cloud_sessions():
    """클라우드 세션 시스템 초기화"""
    global cloud_session_manager
    cloud_session_manager = CloudSessionManager()
    return cloud_session_manager

# Flask 세션 인터페이스 확장
class CloudSessionInterface:
    """Flask와 클라우드 세션 통합"""

    def __init__(self, session_manager: CloudSessionManager):
        self.session_manager = session_manager

    def load_session_data(self, session_id: str) -> Dict[str, Any]:
        """Flask 세션 데이터 로드"""
        return self.session_manager.load_session(session_id) or {}

    def save_session_data(self, session_id: str, session_data: Dict[str, Any]):
        """Flask 세션 데이터 저장"""
        user_id = session_data.get('user', {}).get('id', 'anonymous')
        return self.session_manager.save_session(session_id, user_id, session_data)

    def delete_session_data(self, session_id: str):
        """Flask 세션 데이터 삭제"""
        return self.session_manager.delete_session(session_id)