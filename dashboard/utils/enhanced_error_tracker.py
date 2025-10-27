"""
강화된 에러 추적 및 알림 시스템
실시간 에러 모니터링, 상관관계 분석, 자동 복구 전략 포함
"""

import time
import threading
import hashlib
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Callable, Tuple
from collections import defaultdict, deque
from dataclasses import dataclass, asdict
import json
import os
import traceback
from enum import Enum

from .metrics_collector import get_metrics_collector, MetricType
from .logging_config import get_logger
from .error_handler import ErrorCategory, get_error_handler
from .notification_system import NotificationSystem


class AlertSeverity(Enum):
    """알림 심각도"""
    INFORMATIONAL = "info"
    WARNING = "warning"
    CRITICAL = "critical"
    EMERGENCY = "emergency"


class ErrorPattern(Enum):
    """에러 패턴 타입"""
    SPIKE = "spike"           # 급격한 증가
    BURST = "burst"           # 집중 발생
    TREND = "trend"           # 지속적 증가
    ANOMALY = "anomaly"       # 이상 패턴


@dataclass
class ErrorSignature:
    """에러 시그니처"""
    error_type: str
    module: str
    function: str
    signature_hash: str
    first_seen: datetime
    last_seen: datetime
    occurrences: int
    pattern: ErrorPattern = None


@dataclass
class ErrorAlert:
    """에러 알림 정보"""
    alert_id: str
    severity: AlertSeverity
    title: str
    message: str
    error_signature: ErrorSignature
    threshold_exceeded: Dict[str, Any]
    suggested_actions: List[str]
    created_at: datetime
    auto_recovery_attempted: bool = False


class EnhancedErrorTracker:
    """강화된 에러 추적 시스템"""

    def __init__(self):
        self.logger = get_logger(__name__)
        self.metrics_collector = get_metrics_collector()
        self.base_error_handler = get_error_handler()
        self.notification_system = NotificationSystem()

        # 에러 시그니처 및 패턴 저장소
        self.error_signatures: Dict[str, ErrorSignature] = {}
        self.error_timeline = deque(maxlen=10000)  # 최근 에러 타임라인
        self.alert_history = deque(maxlen=1000)

        # 임계값 설정
        self.thresholds = {
            'error_rate_per_minute': 10,      # 분당 에러 수
            'error_rate_per_hour': 100,       # 시간당 에러 수
            'unique_errors_per_hour': 20,     # 시간당 고유 에러 수
            'critical_error_threshold': 1,     # 크리티컬 에러 즉시 알림
            'error_burst_threshold': 5,        # 10초 내 동일 에러 발생 시
            'error_trend_window': 3600         # 트렌드 분석 윈도우 (초)
        }

        # 알림 제한 (스팸 방지)
        self.alert_throttle = defaultdict(lambda: {'last_sent': None, 'count': 0})
        self.throttle_window = 300  # 5분

        # 자동 복구 전략
        self.recovery_strategies = {}
        self.register_default_recovery_strategies()

        # 백그라운드 분석 스레드
        self.analysis_running = False
        self.analysis_thread = None
        self.start_background_analysis()

    def register_default_recovery_strategies(self):
        """기본 자동 복구 전략 등록"""
        self.recovery_strategies.update({
            'DatabaseConnectionError': self._recover_database_connection,
            'RedisConnectionError': self._recover_redis_connection,
            'FileNotFoundError': self._recover_missing_file,
            'PermissionError': self._recover_permission_error,
            'TimeoutError': self._recover_timeout_error
        })

    def create_error_signature(self, error: Exception, context: Dict[str, Any]) -> ErrorSignature:
        """에러 시그니처 생성"""
        error_type = type(error).__name__
        module = context.get('module', 'unknown')
        function = context.get('function', 'unknown')

        # 시그니처 해시 생성 (에러 타입, 모듈, 함수 기반)
        signature_data = f"{error_type}:{module}:{function}"
        signature_hash = hashlib.md5(signature_data.encode()).hexdigest()[:12]

        now = datetime.now()

        if signature_hash in self.error_signatures:
            # 기존 시그니처 업데이트
            signature = self.error_signatures[signature_hash]
            signature.last_seen = now
            signature.occurrences += 1
        else:
            # 새로운 시그니처 생성
            signature = ErrorSignature(
                error_type=error_type,
                module=module,
                function=function,
                signature_hash=signature_hash,
                first_seen=now,
                last_seen=now,
                occurrences=1
            )
            self.error_signatures[signature_hash] = signature

        return signature

    def track_error(self, error: Exception, context: Dict[str, Any] = None,
                   category: str = ErrorCategory.MEDIUM, user_id: str = None):
        """강화된 에러 추적"""
        context = context or {}

        # 기본 에러 로깅
        error_info = self.base_error_handler.log_error(error, context, category, user_id)

        # 에러 시그니처 생성
        signature = self.create_error_signature(error, context)

        # 타임라인에 추가
        timeline_entry = {
            'timestamp': datetime.now(),
            'signature_hash': signature.signature_hash,
            'category': category,
            'user_id': user_id,
            'context': context
        }
        self.error_timeline.append(timeline_entry)

        # 메트릭 기록
        self._record_error_metrics(signature, category, context)

        # 실시간 분석 및 알림
        self._analyze_and_alert(signature, category, context)

        # 자동 복구 시도
        if category in [ErrorCategory.CRITICAL, ErrorCategory.HIGH]:
            self._attempt_auto_recovery(error, signature, context)

        self.logger.debug(
            f"Enhanced error tracking: {signature.error_type}",
            signature_hash=signature.signature_hash,
            occurrences=signature.occurrences,
            category=category
        )

        return signature

    def _record_error_metrics(self, signature: ErrorSignature, category: str, context: Dict[str, Any]):
        """에러 메트릭 기록"""
        tags = {
            'error_type': signature.error_type,
            'module': signature.module,
            'function': signature.function,
            'category': category,
            'signature_hash': signature.signature_hash
        }

        # 사용자 정보 추가
        if context.get('user_id'):
            tags['user_id'] = str(context['user_id'])

        # 요청 정보 추가
        if context.get('endpoint'):
            tags['endpoint'] = context['endpoint']
        if context.get('method'):
            tags['method'] = context['method']

        # 에러 카운터 증가
        self.metrics_collector.increment_counter("errors.total", tags=tags)
        self.metrics_collector.increment_counter(f"errors.{category.lower()}", tags=tags)
        self.metrics_collector.increment_counter(f"errors.by_type", tags={'error_type': signature.error_type})

        # 에러 발생 빈도 기록
        self.metrics_collector.set_gauge(
            "errors.signature_occurrences",
            signature.occurrences,
            tags={'signature_hash': signature.signature_hash}
        )

    def _analyze_and_alert(self, signature: ErrorSignature, category: str, context: Dict[str, Any]):
        """실시간 에러 분석 및 알림"""
        now = datetime.now()

        # 크리티컬 에러 즉시 알림
        if category == ErrorCategory.CRITICAL:
            self._send_critical_alert(signature, context)
            return

        # 에러 패턴 분석
        pattern = self._detect_error_pattern(signature)
        if pattern:
            signature.pattern = pattern
            self._send_pattern_alert(signature, pattern, context)

        # 에러율 분석
        error_rate = self._calculate_error_rate()
        if error_rate['per_minute'] > self.thresholds['error_rate_per_minute']:
            self._send_rate_alert(error_rate, 'minute')
        elif error_rate['per_hour'] > self.thresholds['error_rate_per_hour']:
            self._send_rate_alert(error_rate, 'hour')

    def _detect_error_pattern(self, signature: ErrorSignature) -> Optional[ErrorPattern]:
        """에러 패턴 감지"""
        now = datetime.now()
        recent_window = now - timedelta(seconds=300)  # 5분 윈도우

        # 최근 동일 에러 발생 횟수
        recent_occurrences = sum(
            1 for entry in self.error_timeline
            if (entry['signature_hash'] == signature.signature_hash and
                entry['timestamp'] > recent_window)
        )

        # Burst 패턴 감지 (짧은 시간 내 집중 발생)
        if recent_occurrences >= self.thresholds['error_burst_threshold']:
            return ErrorPattern.BURST

        # Spike 패턴 감지 (급격한 증가)
        hour_ago = now - timedelta(hours=1)
        hour_occurrences = sum(
            1 for entry in self.error_timeline
            if (entry['signature_hash'] == signature.signature_hash and
                entry['timestamp'] > hour_ago)
        )

        if hour_occurrences > signature.occurrences * 0.5:  # 최근 1시간에 절반 이상 발생
            return ErrorPattern.SPIKE

        # Trend 패턴 감지 (지속적 증가)
        if signature.occurrences > 10:  # 충분한 데이터가 있을 때만
            recent_rate = recent_occurrences / 5  # 분당 발생률
            overall_rate = signature.occurrences / max(1,
                (now - signature.first_seen).total_seconds() / 60)

            if recent_rate > overall_rate * 2:  # 최근 발생률이 전체 평균의 2배 이상
                return ErrorPattern.TREND

        return None

    def _calculate_error_rate(self) -> Dict[str, int]:
        """에러 발생률 계산"""
        now = datetime.now()
        minute_ago = now - timedelta(minutes=1)
        hour_ago = now - timedelta(hours=1)

        per_minute = sum(1 for entry in self.error_timeline if entry['timestamp'] > minute_ago)
        per_hour = sum(1 for entry in self.error_timeline if entry['timestamp'] > hour_ago)

        return {'per_minute': per_minute, 'per_hour': per_hour}

    def _send_critical_alert(self, signature: ErrorSignature, context: Dict[str, Any]):
        """크리티컬 에러 즉시 알림"""
        alert = ErrorAlert(
            alert_id=f"critical_{signature.signature_hash}_{int(time.time())}",
            severity=AlertSeverity.CRITICAL,
            title=f"🚨 크리티컬 에러 발생: {signature.error_type}",
            message=f"""
크리티컬 에러가 발생했습니다.

에러 타입: {signature.error_type}
모듈: {signature.module}
함수: {signature.function}
발생 시간: {signature.last_seen.strftime('%Y-%m-%d %H:%M:%S')}

컨텍스트:
{json.dumps(context, indent=2, ensure_ascii=False)}

즉시 확인이 필요합니다.
""",
            error_signature=signature,
            threshold_exceeded={'critical_error': True},
            suggested_actions=[
                "시스템 로그 확인",
                "서비스 상태 점검",
                "필요시 긴급 대응팀 연락",
                "자동 복구 시도 여부 확인"
            ],
            created_at=datetime.now()
        )

        self._send_alert(alert, bypass_throttle=True)

    def _send_pattern_alert(self, signature: ErrorSignature, pattern: ErrorPattern, context: Dict[str, Any]):
        """패턴 기반 알림"""
        pattern_messages = {
            ErrorPattern.BURST: f"에러 집중 발생 (5분 내 {self.thresholds['error_burst_threshold']}회 이상)",
            ErrorPattern.SPIKE: f"에러 급증 감지 (최근 1시간 집중 발생)",
            ErrorPattern.TREND: f"에러 지속 증가 추세 감지"
        }

        alert = ErrorAlert(
            alert_id=f"pattern_{pattern.value}_{signature.signature_hash}_{int(time.time())}",
            severity=AlertSeverity.WARNING,
            title=f"[WARN] 에러 패턴 감지: {signature.error_type}",
            message=f"""
에러 패턴이 감지되었습니다.

패턴: {pattern_messages.get(pattern, pattern.value)}
에러 타입: {signature.error_type}
총 발생 횟수: {signature.occurrences}
모듈: {signature.module}
함수: {signature.function}

모니터링이 필요합니다.
""",
            error_signature=signature,
            threshold_exceeded={'pattern': pattern.value},
            suggested_actions=[
                "에러 패턴 분석",
                "관련 시스템 점검",
                "로그 상세 분석",
                "필요시 예방 조치 수행"
            ],
            created_at=datetime.now()
        )

        self._send_alert(alert)

    def _send_rate_alert(self, error_rate: Dict[str, int], timeframe: str):
        """에러율 기반 알림"""
        threshold = self.thresholds[f'error_rate_per_{timeframe}']
        current_rate = error_rate[f'per_{timeframe}']

        alert = ErrorAlert(
            alert_id=f"rate_{timeframe}_{int(time.time())}",
            severity=AlertSeverity.WARNING,
            title=f"[WARN] 높은 에러 발생률 감지",
            message=f"""
시스템에서 높은 에러 발생률이 감지되었습니다.

현재 에러율: {current_rate}건/{timeframe}
임계값: {threshold}건/{timeframe}
초과율: {((current_rate / threshold) - 1) * 100:.1f}%

시스템 전반적인 점검이 필요합니다.
""",
            error_signature=None,
            threshold_exceeded={'error_rate': current_rate, 'threshold': threshold, 'timeframe': timeframe},
            suggested_actions=[
                "시스템 리소스 확인",
                "최근 배포/변경사항 점검",
                "데이터베이스 상태 확인",
                "네트워크 연결 상태 점검"
            ],
            created_at=datetime.now()
        )

        self._send_alert(alert)

    def _send_alert(self, alert: ErrorAlert, bypass_throttle: bool = False):
        """알림 발송 (스팸 방지 포함)"""
        alert_key = f"{alert.severity.value}_{alert.title}"

        # 스팸 방지 체크
        if not bypass_throttle:
            throttle_info = self.alert_throttle[alert_key]
            now = datetime.now()

            if throttle_info['last_sent']:
                time_since_last = (now - throttle_info['last_sent']).total_seconds()
                if time_since_last < self.throttle_window:
                    throttle_info['count'] += 1
                    self.logger.debug(f"Alert throttled: {alert_key}")
                    return

            throttle_info['last_sent'] = now
            throttle_info['count'] = 1

        # 알림 히스토리에 추가
        self.alert_history.append(alert)

        # 슬랙 알림 발송
        slack_message = self._format_slack_message(alert)
        success = self.notification_system.send_slack_notification(slack_message)

        # 메트릭 기록
        self.metrics_collector.increment_counter(
            "alerts.sent",
            tags={
                'severity': alert.severity.value,
                'alert_type': alert.alert_id.split('_')[0],
                'success': str(success)
            }
        )

        self.logger.info(
            f"Alert sent: {alert.title}",
            alert_id=alert.alert_id,
            severity=alert.severity.value,
            success=success
        )

    def _format_slack_message(self, alert: ErrorAlert) -> str:
        """슬랙 메시지 포맷"""
        severity_icons = {
            AlertSeverity.INFORMATIONAL: "[INFO]",
            AlertSeverity.WARNING: "[WARN]",
            AlertSeverity.CRITICAL: "[CRIT]",
            AlertSeverity.EMERGENCY: "[EMRG]"
        }

        icon = severity_icons.get(alert.severity, "[UNKN]")

        message = f"{icon} *{alert.title}*\n\n{alert.message}"

        if alert.suggested_actions:
            message += "\n\n*권장 조치:*\n"
            for action in alert.suggested_actions:
                message += f"• {action}\n"

        message += f"\n*알림 ID:* `{alert.alert_id}`"
        message += f"\n*발생 시간:* {alert.created_at.strftime('%Y-%m-%d %H:%M:%S')}"

        return message

    def _attempt_auto_recovery(self, error: Exception, signature: ErrorSignature, context: Dict[str, Any]):
        """자동 복구 시도"""
        recovery_func = self.recovery_strategies.get(signature.error_type)

        if recovery_func:
            try:
                self.logger.info(f"Attempting auto recovery for {signature.error_type}")
                success = recovery_func(error, signature, context)

                # 복구 시도 메트릭 기록
                self.metrics_collector.increment_counter(
                    "auto_recovery.attempts",
                    tags={
                        'error_type': signature.error_type,
                        'success': str(success)
                    }
                )

                if success:
                    self.logger.info(f"Auto recovery successful for {signature.error_type}")
                    # 복구 성공 알림
                    recovery_message = f"""
[SUCCESS] 자동 복구 성공

에러 타입: {signature.error_type}
복구 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
시그니처: {signature.signature_hash}

시스템이 자동으로 복구되었습니다.
"""
                    self.notification_system.send_slack_notification(recovery_message)

            except Exception as recovery_error:
                self.logger.error(f"Auto recovery failed: {recovery_error}")
                self.metrics_collector.increment_counter(
                    "auto_recovery.failures",
                    tags={'error_type': signature.error_type}
                )

    def _recover_database_connection(self, error: Exception, signature: ErrorSignature, context: Dict[str, Any]) -> bool:
        """데이터베이스 연결 복구"""
        try:
            # 데이터베이스 연결 재시도
            from ..utils.user_database import get_user_database
            user_db = get_user_database()
            # 간단한 쿼리로 연결 테스트
            user_db.get_all_users()
            return True
        except Exception as e:
            return False

    def _recover_redis_connection(self, error: Exception, signature: ErrorSignature, context: Dict[str, Any]) -> bool:
        """Redis 연결 복구"""
        # Redis 복구 로직 구현
        return False

    def _recover_missing_file(self, error: Exception, signature: ErrorSignature, context: Dict[str, Any]) -> bool:
        """누락 파일 복구"""
        # 파일 복구 로직 구현 (백업에서 복원 등)
        return False

    def _recover_permission_error(self, error: Exception, signature: ErrorSignature, context: Dict[str, Any]) -> bool:
        """권한 오류 복구"""
        # 권한 복구 로직 구현
        return False

    def _recover_timeout_error(self, error: Exception, signature: ErrorSignature, context: Dict[str, Any]) -> bool:
        """타임아웃 오류 복구"""
        # 타임아웃 복구 로직 구현 (재시도 등)
        return False

    def start_background_analysis(self):
        """백그라운드 분석 시작"""
        if self.analysis_running:
            return

        self.analysis_running = True
        self.analysis_thread = threading.Thread(target=self._background_analysis, daemon=True)
        self.analysis_thread.start()
        self.logger.info("Background error analysis started")

    def stop_background_analysis(self):
        """백그라운드 분석 중지"""
        self.analysis_running = False
        if self.analysis_thread:
            self.analysis_thread.join(timeout=5)
        self.logger.info("Background error analysis stopped")

    def _background_analysis(self):
        """백그라운드 에러 분석"""
        while self.analysis_running:
            try:
                # 주기적 에러 분석
                self._analyze_error_trends()
                self._cleanup_old_data()

                # 30초마다 분석
                time.sleep(30)

            except Exception as e:
                self.logger.error(f"Background analysis error: {e}")
                time.sleep(60)  # 오류 시 1분 대기

    def _analyze_error_trends(self):
        """에러 트렌드 분석"""
        now = datetime.now()
        hour_ago = now - timedelta(hours=1)

        # 시간대별 에러 분포 분석
        hourly_errors = defaultdict(int)
        for entry in self.error_timeline:
            if entry['timestamp'] > hour_ago:
                hour_key = entry['timestamp'].strftime('%H:%M')[:4] + '0'  # 10분 단위
                hourly_errors[hour_key] += 1

        # 급격한 증가 감지
        if len(hourly_errors) >= 3:
            values = list(hourly_errors.values())
            if values[-1] > sum(values[:-1]) / len(values[:-1]) * 3:  # 마지막 구간이 평균의 3배 이상
                self.logger.warning("Error trend spike detected")

    def _cleanup_old_data(self):
        """오래된 데이터 정리"""
        cutoff_time = datetime.now() - timedelta(hours=24)

        # 24시간 이상 된 에러 시그니처 정리
        to_remove = []
        for hash_key, signature in self.error_signatures.items():
            if signature.last_seen < cutoff_time:
                to_remove.append(hash_key)

        for hash_key in to_remove:
            del self.error_signatures[hash_key]

        if to_remove:
            self.logger.debug(f"Cleaned up {len(to_remove)} old error signatures")

    def get_error_dashboard_data(self) -> Dict[str, Any]:
        """에러 대시보드 데이터 반환"""
        now = datetime.now()
        hour_ago = now - timedelta(hours=1)
        day_ago = now - timedelta(days=1)

        # 최근 에러 통계
        recent_errors = [e for e in self.error_timeline if e['timestamp'] > hour_ago]
        daily_errors = [e for e in self.error_timeline if e['timestamp'] > day_ago]

        # 에러 타입별 분포
        error_types = defaultdict(int)
        for signature in self.error_signatures.values():
            error_types[signature.error_type] += signature.occurrences

        # 최근 알림
        recent_alerts = [alert for alert in self.alert_history if alert.created_at > day_ago]

        return {
            'summary': {
                'total_signatures': len(self.error_signatures),
                'errors_last_hour': len(recent_errors),
                'errors_last_day': len(daily_errors),
                'alerts_last_day': len(recent_alerts)
            },
            'error_types': dict(error_types),
            'recent_signatures': [
                {
                    'signature_hash': sig.signature_hash,
                    'error_type': sig.error_type,
                    'module': sig.module,
                    'function': sig.function,
                    'occurrences': sig.occurrences,
                    'last_seen': sig.last_seen.isoformat(),
                    'pattern': sig.pattern.value if sig.pattern else None
                }
                for sig in sorted(self.error_signatures.values(),
                                key=lambda x: x.last_seen, reverse=True)[:10]
            ],
            'recent_alerts': [
                {
                    'alert_id': alert.alert_id,
                    'severity': alert.severity.value,
                    'title': alert.title,
                    'created_at': alert.created_at.isoformat()
                }
                for alert in sorted(recent_alerts,
                                  key=lambda x: x.created_at, reverse=True)[:5]
            ],
            'thresholds': self.thresholds,
            'analysis_running': self.analysis_running
        }


# 전역 인스턴스
_enhanced_error_tracker = None


def get_enhanced_error_tracker() -> EnhancedErrorTracker:
    """전역 강화 에러 트래커 반환"""
    global _enhanced_error_tracker
    if _enhanced_error_tracker is None:
        _enhanced_error_tracker = EnhancedErrorTracker()
    return _enhanced_error_tracker


def enhanced_error_handler(category: str = ErrorCategory.MEDIUM,
                          context: Dict[str, Any] = None,
                          enable_auto_recovery: bool = True):
    """강화된 에러 처리 데코레이터"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                tracker = get_enhanced_error_tracker()

                # 컨텍스트 구성
                error_context = context or {}
                error_context.update({
                    'function': func.__name__,
                    'module': func.__module__,
                    'args_count': len(args),
                    'kwargs_keys': list(kwargs.keys())
                })

                # 강화된 에러 추적
                tracker.track_error(e, error_context, category)

                # 재발생
                raise

        return wrapper
    return decorator