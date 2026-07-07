"""
Redis Pub/Sub 기반 실시간 캐시 무효화 브로드캐스트

멀티 워커 환경에서 캐시 일관성 보장:
- 워커 A가 데이터 수정 → Redis Pub/Sub로 무효화 메시지 발행
- 워커 B, C, D가 메시지 수신 → 로컬 캐시 무효화
- 모든 워커가 최신 데이터 반영
"""

import json
import threading
import time
from typing import Optional, List, Callable
import logging

from dashboard.utils.redis_client import get_redis_client

logger = logging.getLogger(__name__)


class CacheInvalidationService:
    """
    Redis Pub/Sub 기반 캐시 무효화 서비스

    특징:
    - 멀티 워커 간 캐시 일관성 보장
    - 실시간 무효화 메시지 브로드캐스트
    - 패턴 기반 무효화 지원
    - SocketIO 통합 (프론트엔드 실시간 알림)
    """

    CHANNEL_NAME = "cache:invalidation"

    def __init__(self):
        self.redis = get_redis_client()
        self._pubsub = None
        self._subscriber_thread = None
        self._running = False
        self._handlers: List[Callable] = []
        logger.info("CacheInvalidationService 초기화")

    def publish_invalidation(self, keys: List[str], reason: str = "data_updated") -> bool:
        """
        캐시 무효화 메시지 발행 (모든 워커에게 브로드캐스트)

        Args:
            keys: 무효화할 캐시 키 목록
            reason: 무효화 사유 (data_updated, user_action, admin_force 등)

        Returns:
            bool: 발행 성공 여부
        """
        try:
            message = {
                "action": "invalidate",
                "keys": keys,
                "reason": reason,
                "timestamp": time.time()
            }

            # Redis Pub/Sub로 메시지 발행
            subscribers = self.redis.redis.publish(
                self.CHANNEL_NAME,
                json.dumps(message)
            )

            logger.info(f"캐시 무효화 메시지 발행: {len(keys)}개 키, {subscribers}개 구독자")
            logger.debug(f"무효화 키: {keys}")

            return True

        except Exception as e:
            logger.error(f"캐시 무효화 메시지 발행 실패: {e}")

            # 🔄 FALLBACK: Pub/Sub 실패 시 로컬 캐시라도 무효화
            try:
                from dashboard.utils.smart_cache_manager import get_smart_cache
                cache = get_smart_cache()

                for key in keys:
                    # broadcast=False로 무한 루프 방지
                    cache.delete(key, set_marker=True, broadcast=False)

                logger.warning(
                    f"⚠️ Pub/Sub 실패 → 로컬 캐시만 무효화 완료: {len(keys)}개 키\n"
                    f"다중 워커 환경에서는 다른 워커의 캐시가 무효화되지 않을 수 있습니다."
                )
                return True  # 로컬 캐시는 무효화 성공

            except Exception as fallback_error:
                logger.error(f"폴백 캐시 무효화도 실패: {fallback_error}")
                return False

    def publish_pattern_invalidation(self, pattern: str, reason: str = "pattern_invalidation") -> bool:
        """
        패턴 기반 캐시 무효화 (예: "project:*", "user:admin:*")

        Args:
            pattern: Redis 키 패턴 (와일드카드 지원)
            reason: 무효화 사유

        Returns:
            bool: 발행 성공 여부
        """
        try:
            message = {
                "action": "invalidate_pattern",
                "pattern": pattern,
                "reason": reason,
                "timestamp": time.time()
            }

            subscribers = self.redis.redis.publish(
                self.CHANNEL_NAME,
                json.dumps(message)
            )

            logger.info(f"패턴 무효화 메시지 발행: {pattern}, {subscribers}개 구독자")

            return True

        except Exception as e:
            logger.error(f"패턴 무효화 메시지 발행 실패: {e}")

            # 🔄 FALLBACK: Pub/Sub 실패 시 로컬 캐시라도 무효화
            try:
                from dashboard.utils.smart_cache_manager import get_smart_cache
                cache = get_smart_cache()

                # broadcast=False로 무한 루프 방지
                invalidated_count = cache.invalidate_by_pattern(pattern, set_marker=True, broadcast=False)

                logger.warning(
                    f"⚠️ Pub/Sub 실패 → 로컬 캐시만 패턴 무효화 완료: {pattern} ({invalidated_count}개)\n"
                    f"다중 워커 환경에서는 다른 워커의 캐시가 무효화되지 않을 수 있습니다."
                )
                return True  # 로컬 캐시는 무효화 성공

            except Exception as fallback_error:
                logger.error(f"폴백 패턴 무효화도 실패: {fallback_error}")
                return False

    def publish_cache_clear(self, reason: str = "admin_clear") -> bool:
        """
        전체 캐시 무효화 (관리자 강제 초기화)

        Args:
            reason: 무효화 사유

        Returns:
            bool: 발행 성공 여부
        """
        try:
            message = {
                "action": "clear_all",
                "reason": reason,
                "timestamp": time.time()
            }

            subscribers = self.redis.redis.publish(
                self.CHANNEL_NAME,
                json.dumps(message)
            )

            logger.warning(f"전체 캐시 무효화 메시지 발행: {subscribers}개 구독자")

            return True

        except Exception as e:
            logger.error(f"전체 캐시 무효화 메시지 발행 실패: {e}")
            return False

    def register_handler(self, handler: Callable[[dict], None]):
        """
        무효화 메시지 핸들러 등록

        Args:
            handler: 메시지 처리 함수 (dict 파라미터)
        """
        self._handlers.append(handler)
        logger.info(f"무효화 핸들러 등록: {handler.__name__}")

    def start_subscriber(self):
        """
        Pub/Sub 구독자 시작 (백그라운드 스레드)

        Note (2026-07-07): 현재 아키텍처는 Waitress 단일 프로세스이므로 Pub/Sub는
        향후 멀티프로세스/다중 워커 확장 대비. 단일 프로세스에서는 로컬 캐시 무효화
        직접 호출로 즉시 반영되고, 여기 pub/sub는 안전망 역할.
        """
        if self._running:
            logger.warning("Pub/Sub 구독자가 이미 실행 중입니다.")
            return

        self._running = True
        self._subscriber_thread = threading.Thread(
            target=self._subscriber_loop,
            daemon=True,
            name="CacheInvalidationSubscriber"
        )
        self._subscriber_thread.start()
        logger.info(f"Pub/Sub 구독자 시작: 채널={self.CHANNEL_NAME}")

    def stop_subscriber(self):
        """
        Pub/Sub 구독자 중지
        """
        if not self._running:
            return

        self._running = False

        if self._pubsub:
            try:
                self._pubsub.unsubscribe()
                self._pubsub.close()
            except Exception as e:
                logger.debug(f"Pub/Sub 구독 해제 중 오류: {e}")

        if self._subscriber_thread:
            self._subscriber_thread.join(timeout=5)

        logger.info("Pub/Sub 구독자 중지")

    def _subscriber_loop(self):
        """
        Pub/Sub 메시지 수신 루프 (백그라운드 스레드)
        """
        try:
            # Pub/Sub 연결 생성 (스레드 안전)
            self._pubsub = self.redis.redis.pubsub()
            self._pubsub.subscribe(self.CHANNEL_NAME)

            logger.info(f"Pub/Sub 구독 시작: {self.CHANNEL_NAME}")

            # 메시지 수신 루프
            for message in self._pubsub.listen():
                if not self._running:
                    break

                # 메시지 타입이 'message'인 경우만 처리
                if message['type'] == 'message':
                    try:
                        data = json.loads(message['data'])
                        self._handle_message(data)
                    except json.JSONDecodeError as e:
                        logger.error(f"Pub/Sub 메시지 파싱 오류: {e}")
                    except Exception as e:
                        logger.error(f"Pub/Sub 메시지 처리 오류: {e}")

        except Exception as e:
            logger.error(f"Pub/Sub 구독자 루프 오류: {e}")

        finally:
            logger.info("Pub/Sub 구독자 루프 종료")

    def _handle_message(self, message: dict):
        """
        무효화 메시지 처리

        Args:
            message: 무효화 메시지 (dict)
        """
        action = message.get("action")

        if action == "invalidate":
            keys = message.get("keys", [])
            reason = message.get("reason", "unknown")
            logger.debug(f"캐시 무효화 메시지 수신: {len(keys)}개 키, 사유={reason}")

        elif action == "invalidate_pattern":
            pattern = message.get("pattern", "")
            reason = message.get("reason", "unknown")
            logger.debug(f"패턴 무효화 메시지 수신: {pattern}, 사유={reason}")

        elif action == "clear_all":
            reason = message.get("reason", "unknown")
            logger.warning(f"전체 캐시 무효화 메시지 수신: 사유={reason}")

        # 등록된 핸들러 실행
        for handler in self._handlers:
            try:
                handler(message)
            except Exception as e:
                logger.error(f"핸들러 실행 오류 ({handler.__name__}): {e}")


# 싱글톤 인스턴스
_cache_invalidation_service: Optional[CacheInvalidationService] = None


def get_cache_invalidation_service() -> CacheInvalidationService:
    """
    CacheInvalidationService 싱글톤 인스턴스 반환

    Returns:
        CacheInvalidationService: 캐시 무효화 서비스 인스턴스
    """
    global _cache_invalidation_service
    if _cache_invalidation_service is None:
        _cache_invalidation_service = CacheInvalidationService()
    return _cache_invalidation_service
