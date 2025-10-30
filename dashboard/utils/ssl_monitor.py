"""
SSL 에러 모니터링 시스템

전문가 권장사항:
- SSL 에러 빈도 추적
- 재시도 성공/실패 통계
- 임계값 초과 시 알림
"""
import time
from datetime import datetime, timedelta
from typing import Optional, Dict
from .logging_config import get_logger
from .redis_client import get_redis_client

logger = get_logger(__name__)


class SSLErrorMonitor:
    """SSL 에러 발생 빈도 모니터링"""

    # Redis 키 prefix
    KEY_PREFIX = "ssl_monitor"

    # 통계 추적 기간 (초)
    WINDOW_SIZE = 3600  # 1시간

    # 알림 임계값
    ALERT_THRESHOLD = {
        'error_count_per_hour': 10,  # 시간당 10회 이상
        'failure_rate': 0.3,  # 실패율 30% 이상
    }

    def __init__(self):
        self.redis = get_redis_client().redis  # 실제 redis.Redis 객체 사용

    def record_ssl_error(self, operation_name: str, error_type: str = "SSLError"):
        """
        SSL 에러 발생 기록

        Args:
            operation_name: 작업 이름
            error_type: 에러 유형 (SSLError, requests.SSLError)
        """
        try:
            timestamp = int(time.time())

            # 전역 카운터 증가
            key = f"{self.KEY_PREFIX}:errors:total"
            self.redis.incr(key)
            self.redis.expire(key, self.WINDOW_SIZE)

            # 시간대별 카운터 (분 단위)
            minute_key = f"{self.KEY_PREFIX}:errors:minute:{timestamp // 60}"
            self.redis.incr(minute_key)
            self.redis.expire(minute_key, 3600)

            # 에러 타입별 카운터
            type_key = f"{self.KEY_PREFIX}:errors:type:{error_type}"
            self.redis.incr(type_key)
            self.redis.expire(type_key, self.WINDOW_SIZE)

            # 작업별 카운터
            op_key = f"{self.KEY_PREFIX}:errors:operation:{operation_name}"
            self.redis.incr(op_key)
            self.redis.expire(op_key, self.WINDOW_SIZE)

            # 최근 에러 리스트 (최대 100개)
            recent_key = f"{self.KEY_PREFIX}:errors:recent"
            error_info = {
                'timestamp': timestamp,
                'operation': operation_name,
                'type': error_type,
                'datetime': datetime.now().isoformat()
            }
            self.redis.lpush(recent_key, str(error_info))
            self.redis.ltrim(recent_key, 0, 99)
            self.redis.expire(recent_key, 86400)  # 24시간

            logger.warning(
                f"[SSL_MONITOR] SSL 에러 기록됨: {operation_name} ({error_type})"
            )

        except Exception as e:
            logger.error(f"[SSL_MONITOR] 에러 기록 실패: {e}")

    def record_ssl_recovery(self, operation_name: str, attempt_count: int, recovery_time_ms: float):
        """
        SSL 에러 복구 성공 기록

        Args:
            operation_name: 작업 이름
            attempt_count: 재시도 횟수
            recovery_time_ms: 복구 소요 시간 (밀리초)
        """
        try:
            timestamp = int(time.time())

            # 복구 성공 카운터
            key = f"{self.KEY_PREFIX}:recoveries:total"
            self.redis.incr(key)
            self.redis.expire(key, self.WINDOW_SIZE)

            # 재시도 횟수 통계
            retry_key = f"{self.KEY_PREFIX}:recoveries:retry_count:{attempt_count}"
            self.redis.incr(retry_key)
            self.redis.expire(retry_key, self.WINDOW_SIZE)

            # 복구 시간 통계 (평균 계산용)
            time_key = f"{self.KEY_PREFIX}:recoveries:total_time"
            self.redis.incrbyfloat(time_key, recovery_time_ms)
            self.redis.expire(time_key, self.WINDOW_SIZE)

            logger.info(
                f"[SSL_MONITOR] SSL 복구 성공: {operation_name} "
                f"(시도: {attempt_count}, 소요: {recovery_time_ms:.0f}ms)"
            )

        except Exception as e:
            logger.error(f"[SSL_MONITOR] 복구 기록 실패: {e}")

    def record_ssl_failure(self, operation_name: str):
        """
        SSL 에러 최대 재시도 실패 기록 (SSLRetryFailedError)

        Args:
            operation_name: 작업 이름
        """
        try:
            # 실패 카운터
            key = f"{self.KEY_PREFIX}:failures:total"
            self.redis.incr(key)
            self.redis.expire(key, self.WINDOW_SIZE)

            # 작업별 실패 카운터
            op_key = f"{self.KEY_PREFIX}:failures:operation:{operation_name}"
            self.redis.incr(op_key)
            self.redis.expire(op_key, self.WINDOW_SIZE)

            logger.error(
                f"[SSL_MONITOR] SSL 복구 실패 (최대 재시도 초과): {operation_name}"
            )

            # 실패율 체크
            self._check_failure_rate_alert()

        except Exception as e:
            logger.error(f"[SSL_MONITOR] 실패 기록 실패: {e}")

    def get_statistics(self) -> Dict:
        """
        SSL 에러 통계 조회

        Returns:
            Dict: 통계 정보
        """
        try:
            total_errors = int(self.redis.get(f"{self.KEY_PREFIX}:errors:total") or 0)
            total_recoveries = int(self.redis.get(f"{self.KEY_PREFIX}:recoveries:total") or 0)
            total_failures = int(self.redis.get(f"{self.KEY_PREFIX}:failures:total") or 0)

            total_time = float(self.redis.get(f"{self.KEY_PREFIX}:recoveries:total_time") or 0)
            avg_recovery_time = total_time / total_recoveries if total_recoveries > 0 else 0

            failure_rate = total_failures / total_errors if total_errors > 0 else 0

            return {
                'window_hours': self.WINDOW_SIZE / 3600,
                'total_errors': total_errors,
                'total_recoveries': total_recoveries,
                'total_failures': total_failures,
                'failure_rate': failure_rate,
                'avg_recovery_time_ms': avg_recovery_time,
                'errors_per_hour': total_errors / (self.WINDOW_SIZE / 3600),
            }

        except Exception as e:
            logger.error(f"[SSL_MONITOR] 통계 조회 실패: {e}")
            return {}

    def _check_failure_rate_alert(self):
        """실패율 임계값 초과 시 알림"""
        try:
            stats = self.get_statistics()

            if not stats:
                return

            # 임계값 체크
            if stats['total_errors'] >= 5:  # 최소 5회 이상 발생 시에만 체크
                if stats['failure_rate'] >= self.ALERT_THRESHOLD['failure_rate']:
                    logger.critical(
                        f"[SSL_MONITOR] ⚠️ SSL 실패율 임계값 초과! "
                        f"실패율: {stats['failure_rate']:.1%} "
                        f"(임계값: {self.ALERT_THRESHOLD['failure_rate']:.1%}) | "
                        f"총 에러: {stats['total_errors']}, "
                        f"복구: {stats['total_recoveries']}, "
                        f"실패: {stats['total_failures']}"
                    )

            if stats['errors_per_hour'] >= self.ALERT_THRESHOLD['error_count_per_hour']:
                logger.critical(
                    f"[SSL_MONITOR] ⚠️ SSL 에러 빈도 임계값 초과! "
                    f"시간당 {stats['errors_per_hour']:.1f}회 "
                    f"(임계값: {self.ALERT_THRESHOLD['error_count_per_hour']}회)"
                )

        except Exception as e:
            logger.error(f"[SSL_MONITOR] 알림 체크 실패: {e}")

    def print_report(self):
        """통계 리포트 출력 (로그)"""
        try:
            stats = self.get_statistics()

            if not stats or stats['total_errors'] == 0:
                logger.info("[SSL_MONITOR] 📊 SSL 에러 통계: 에러 없음")
                return

            logger.info(
                f"[SSL_MONITOR] 📊 SSL 에러 통계 (최근 {stats['window_hours']:.1f}시간):\n"
                f"  - 총 에러: {stats['total_errors']}회\n"
                f"  - 복구 성공: {stats['total_recoveries']}회\n"
                f"  - 복구 실패: {stats['total_failures']}회\n"
                f"  - 실패율: {stats['failure_rate']:.1%}\n"
                f"  - 시간당 에러: {stats['errors_per_hour']:.1f}회\n"
                f"  - 평균 복구 시간: {stats['avg_recovery_time_ms']:.0f}ms"
            )

        except Exception as e:
            logger.error(f"[SSL_MONITOR] 리포트 출력 실패: {e}")


# 전역 싱글턴 인스턴스
_ssl_monitor: Optional[SSLErrorMonitor] = None


def get_ssl_monitor() -> SSLErrorMonitor:
    """SSL 모니터 싱글턴 인스턴스 반환"""
    global _ssl_monitor
    if _ssl_monitor is None:
        _ssl_monitor = SSLErrorMonitor()
    return _ssl_monitor
