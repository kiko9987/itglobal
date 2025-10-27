"""
API 사용량 모니터링 시스템
Google Sheets API 호출 횟수와 응답 시간을 추적하며 임계치 알림 기능을 제공합니다.
"""

import time
import threading
from typing import Dict, Any, Optional, Callable, Tuple
from datetime import datetime, timedelta
import logging
from collections import defaultdict, deque
import atexit

logger = logging.getLogger(__name__)

class APIUsageMonitor:
    """API 사용량을 모니터링하는 클래스"""

    def __init__(self, window_minutes: int = 60):
        self.window_minutes = window_minutes
        self.lock = threading.RLock()

        # 시간 윈도우별 API 호출 기록
        self._api_calls: Dict[str, deque] = defaultdict(lambda: deque())

        # 총 통계
        self._total_calls = defaultdict(int)
        self._total_errors = defaultdict(int)
        self._total_response_time = defaultdict(float)

        # 현재 요청 추적 (응답 시간 측정용)
        # 값: (시작 시간, API 이름) 튜플
        self._active_requests: Dict[str, Tuple[float, str]] = {}

        # 알림 시스템
        self._alert_callbacks: Dict[str, Callable] = {}
        self._last_alerts: Dict[str, float] = {}  # 마지막 알림 시간
        self._alert_cooldown = 300  # 5분 쿨다운

    def start_request(self, api_name: str) -> str:
        """API 요청 시작을 기록"""
        request_id = f"{api_name}_{time.time()}_{threading.current_thread().ident}"
        current_time = time.time()

        with self.lock:
            # API 이름도 함께 저장하여 end_request에서 정확히 복원
            self._active_requests[request_id] = (current_time, api_name)

            # 시간 윈도우 기록
            self._api_calls[api_name].append({
                'timestamp': current_time,
                'type': 'start',
                'request_id': request_id
            })

            # 윈도우 정리
            self._cleanup_old_records(api_name)

        logger.debug(f"[API_MONITOR] {api_name} 요청 시작: {request_id}")
        return request_id

    def end_request(self, request_id: str, success: bool = True, error_msg: Optional[str] = None):
        """API 요청 완료를 기록"""
        current_time = time.time()

        with self.lock:
            if request_id in self._active_requests:
                # start_request에서 저장한 (start_time, api_name) 튜플 복원
                start_time, api_name = self._active_requests.pop(request_id)
                response_time = current_time - start_time

                # 통계 업데이트
                self._total_calls[api_name] += 1
                self._total_response_time[api_name] += response_time

                if not success:
                    self._total_errors[api_name] += 1

                # 시간 윈도우 기록
                self._api_calls[api_name].append({
                    'timestamp': current_time,
                    'type': 'end',
                    'request_id': request_id,
                    'response_time': response_time,
                    'success': success,
                    'error_msg': error_msg
                })

                # 윈도우 정리
                self._cleanup_old_records(api_name)

                status = "성공" if success else f"실패 ({error_msg})"
                logger.debug(f"[API_MONITOR] {api_name} 요청 완료: {status}, 응답시간: {response_time:.3f}초")

    def _cleanup_old_records(self, api_name: str):
        """윈도우 시간을 벗어난 오래된 기록 정리"""
        cutoff_time = time.time() - (self.window_minutes * 60)

        while self._api_calls[api_name] and self._api_calls[api_name][0]['timestamp'] < cutoff_time:
            self._api_calls[api_name].popleft()

    def get_usage_stats(self, api_name: Optional[str] = None) -> Dict[str, Any]:
        """API 사용량 통계 반환"""
        with self.lock:
            if api_name:
                apis = [api_name] if api_name in self._total_calls else []
            else:
                apis = list(self._total_calls.keys())

            stats = {
                'window_minutes': self.window_minutes,
                'timestamp': datetime.now().isoformat(),
                'apis': {}
            }

            for api in apis:
                # 현재 윈도우 내 호출 수 계산
                current_time = time.time()
                cutoff_time = current_time - (self.window_minutes * 60)

                window_calls = [
                    call for call in self._api_calls[api]
                    if call['timestamp'] >= cutoff_time and call['type'] == 'end'
                ]

                window_errors = len([call for call in window_calls if not call['success']])
                window_response_times = [call['response_time'] for call in window_calls]

                avg_response_time = (
                    sum(window_response_times) / len(window_response_times)
                    if window_response_times else 0
                )

                total_avg_response_time = (
                    self._total_response_time[api] / self._total_calls[api]
                    if self._total_calls[api] > 0 else 0
                )

                stats['apis'][api] = {
                    'window_calls': len(window_calls),
                    'window_errors': window_errors,
                    'window_error_rate': window_errors / len(window_calls) if window_calls else 0,
                    'window_avg_response_time': avg_response_time,
                    'total_calls': self._total_calls[api],
                    'total_errors': self._total_errors[api],
                    'total_error_rate': self._total_errors[api] / self._total_calls[api] if self._total_calls[api] > 0 else 0,
                    'total_avg_response_time': total_avg_response_time,
                    'active_requests': len([req for req in self._active_requests.keys() if req.startswith(api)])
                }

            return stats

    def get_rate_limit_status(self, api_name: str, limit_per_minute: int) -> Dict[str, Any]:
        """API 속도 제한 상태 확인"""
        with self.lock:
            current_time = time.time()
            cutoff_time = current_time - 60  # 1분 윈도우

            recent_calls = [
                call for call in self._api_calls[api_name]
                if call['timestamp'] >= cutoff_time and call['type'] == 'end'
            ]

            calls_in_last_minute = len(recent_calls)
            remaining_calls = max(0, limit_per_minute - calls_in_last_minute)

            return {
                'api_name': api_name,
                'limit_per_minute': limit_per_minute,
                'calls_in_last_minute': calls_in_last_minute,
                'remaining_calls': remaining_calls,
                'usage_percentage': (calls_in_last_minute / limit_per_minute) * 100 if limit_per_minute > 0 else 0,
                'is_near_limit': calls_in_last_minute >= (limit_per_minute * 0.8),  # 80% 이상 사용 시 경고
                'timestamp': datetime.now().isoformat()
            }

    def register_alert_callback(self, alert_type: str, callback: Callable):
        """알림 콜백 등록"""
        with self.lock:
            self._alert_callbacks[alert_type] = callback
            logger.info(f"[API_MONITOR] 알림 콜백 등록: {alert_type}")

    def check_thresholds_and_alert(self):
        """임계치 체크 및 알림 발송"""
        with self.lock:
            for api_name in self._total_calls.keys():
                # 속도 제한 체크
                rate_limit_status = self.get_rate_limit_status(api_name, GOOGLE_SHEETS_READ_LIMIT)
                if rate_limit_status['is_near_limit']:
                    self._send_alert('rate_limit_warning', {
                        'api_name': api_name,
                        'usage_percentage': rate_limit_status['usage_percentage'],
                        'calls_in_last_minute': rate_limit_status['calls_in_last_minute'],
                        'limit': rate_limit_status['limit_per_minute']
                    })

                # 오류율 체크
                stats = self.get_usage_stats(api_name)
                if api_name in stats['apis']:
                    api_stats = stats['apis'][api_name]
                    if api_stats['window_error_rate'] > 0.1:  # 10% 이상 오류율
                        self._send_alert('high_error_rate', {
                            'api_name': api_name,
                            'error_rate': api_stats['window_error_rate'] * 100,
                            'window_errors': api_stats['window_errors'],
                            'window_calls': api_stats['window_calls']
                        })

                    # 응답 시간 체크
                    if api_stats['window_avg_response_time'] > 10.0:  # 10초 이상
                        self._send_alert('slow_response', {
                            'api_name': api_name,
                            'avg_response_time': api_stats['window_avg_response_time'],
                            'window_calls': api_stats['window_calls']
                        })

    def _send_alert(self, alert_type: str, alert_data: Dict[str, Any]):
        """알림 발송 (쿨다운 적용)"""
        current_time = time.time()
        alert_key = f"{alert_type}_{alert_data.get('api_name', 'unknown')}"

        # 쿨다운 체크
        if alert_key in self._last_alerts:
            if current_time - self._last_alerts[alert_key] < self._alert_cooldown:
                return  # 쿨다운 중이므로 알림 스킵

        self._last_alerts[alert_key] = current_time

        # 기본 로그 알림
        self._log_alert(alert_type, alert_data)

        # 등록된 콜백 실행
        if alert_type in self._alert_callbacks:
            try:
                self._alert_callbacks[alert_type](alert_data)
            except Exception as e:
                logger.error(f"[API_MONITOR] 알림 콜백 실행 오류: {e}")

    def _log_alert(self, alert_type: str, alert_data: Dict[str, Any]):
        """로그로 알림 발송"""
        api_name = alert_data.get('api_name', 'unknown')

        if alert_type == 'rate_limit_warning':
            logger.warning(
                f"[API_ALERT] 속도 제한 경고: {api_name} API 사용량 {alert_data['usage_percentage']:.1f}% "
                f"({alert_data['calls_in_last_minute']}/{alert_data['limit']} calls/min)"
            )
        elif alert_type == 'high_error_rate':
            logger.warning(
                f"[API_ALERT] 높은 오류율: {api_name} API 오류율 {alert_data['error_rate']:.1f}% "
                f"({alert_data['window_errors']}/{alert_data['window_calls']} 오류/호출)"
            )
        elif alert_type == 'slow_response':
            logger.warning(
                f"[API_ALERT] 느린 응답: {api_name} API 평균 응답시간 {alert_data['avg_response_time']:.2f}초 "
                f"({alert_data['window_calls']}회 호출)"
            )


class APIMonitorScheduler:
    """API 모니터링 스케줄러 - 주기적으로 임계치 체크 및 사용량 리포트 수행"""

    def __init__(self, monitor: APIUsageMonitor, check_interval: int = 300, report_interval: int = 3600):
        self.monitor = monitor
        self.check_interval = check_interval    # 임계치 체크 간격 (기본 5분)
        self.report_interval = report_interval  # 사용량 리포트 간격 (기본 1시간)

        self._running = False
        self._scheduler_thread: Optional[threading.Thread] = None
        self._shutdown_event = threading.Event()
        self._last_report_time = 0

        # 로거
        self.logger = logging.getLogger(f"{__name__}.scheduler")

    def start(self):
        """스케줄러 시작"""
        if self._running:
            self.logger.warning("[API_SCHEDULER] 스케줄러가 이미 실행 중입니다")
            return

        self._running = True
        self._shutdown_event.clear()

        self._scheduler_thread = threading.Thread(
            target=self._scheduler_loop,
            name="APIMonitorScheduler",
            daemon=True
        )
        self._scheduler_thread.start()

        self.logger.info(f"[API_SCHEDULER] 스케줄러 시작 (체크: {self.check_interval}초, 리포트: {self.report_interval}초)")

    def stop(self):
        """스케줄러 중지"""
        if not self._running:
            return

        self._running = False
        self._shutdown_event.set()

        if self._scheduler_thread and self._scheduler_thread.is_alive():
            self._scheduler_thread.join(timeout=10)

            if self._scheduler_thread.is_alive():
                self.logger.warning("[API_SCHEDULER] 정상 종료 실패, 강제 종료 시도")

        self.logger.info("[API_SCHEDULER] 스케줄러 중지")

    def _scheduler_loop(self):
        """스케줄러 메인 루프"""
        self.logger.info("[API_SCHEDULER] 스케줄러 루프 시작")

        while self._running and not self._shutdown_event.is_set():
            try:
                current_time = time.time()

                # 임계치 체크 수행
                self._check_thresholds()

                # 사용량 리포트 수행 (시간 간격 확인)
                if current_time - self._last_report_time >= self.report_interval:
                    self._generate_usage_report()
                    self._last_report_time = current_time

                # 인터럽트 가능한 대기
                if self._shutdown_event.wait(timeout=self.check_interval):
                    break

            except KeyboardInterrupt:
                self.logger.info("[API_SCHEDULER] 키보드 인터럽트로 중지")
                break
            except Exception as e:
                self.logger.error(f"[API_SCHEDULER] 스케줄러 루프 오류: {e}")
                # 오류 발생 시 짧은 대기 후 재시도
                if self._shutdown_event.wait(timeout=30):
                    break

        self.logger.info("[API_SCHEDULER] 스케줄러 루프 종료")

    def _check_thresholds(self):
        """API 임계치 체크 수행"""
        try:
            self.monitor.check_thresholds_and_alert()
            self.logger.debug("[API_SCHEDULER] 임계치 체크 완료")
        except Exception as e:
            self.logger.error(f"[API_SCHEDULER] 임계치 체크 오류: {e}")

    def _generate_usage_report(self):
        """사용량 리포트 생성"""
        try:
            log_api_usage_summary()
            self.logger.debug("[API_SCHEDULER] 사용량 리포트 생성 완료")
        except Exception as e:
            self.logger.error(f"[API_SCHEDULER] 사용량 리포트 생성 오류: {e}")

    def get_scheduler_stats(self) -> Dict[str, Any]:
        """스케줄러 상태 정보 반환"""
        return {
            'running': self._running,
            'check_interval': self.check_interval,
            'report_interval': self.report_interval,
            'last_report_time': datetime.fromtimestamp(self._last_report_time).isoformat() if self._last_report_time > 0 else None,
            'next_report_time': datetime.fromtimestamp(self._last_report_time + self.report_interval).isoformat() if self._last_report_time > 0 else None,
            'thread_alive': self._scheduler_thread.is_alive() if self._scheduler_thread else False
        }


# 전역 모니터 인스턴스
_global_monitor = APIUsageMonitor()
_global_scheduler = None

def get_api_monitor() -> APIUsageMonitor:
    """글로벌 API 모니터 인스턴스 반환"""
    return _global_monitor

def get_api_scheduler() -> Optional[APIMonitorScheduler]:
    """글로벌 API 스케줄러 인스턴스 반환"""
    return _global_scheduler

def start_api_monitoring_scheduler(check_interval: int = 300, report_interval: int = 3600) -> bool:
    """API 모니터링 스케줄러 시작"""
    global _global_scheduler

    if _global_scheduler and _global_scheduler._running:
        logger.warning("[API_SCHEDULER] 스케줄러가 이미 실행 중입니다")
        return False

    try:
        _global_scheduler = APIMonitorScheduler(
            monitor=_global_monitor,
            check_interval=check_interval,
            report_interval=report_interval
        )
        _global_scheduler.start()

        # 프로세스 종료 시 정리를 위한 atexit 핸들러 등록
        atexit.register(stop_api_monitoring_scheduler)

        logger.info("[API_SCHEDULER] API 모니터링 스케줄러 시작 완료")
        return True

    except Exception as e:
        logger.error(f"[API_SCHEDULER] 스케줄러 시작 실패: {e}")
        return False

def stop_api_monitoring_scheduler():
    """API 모니터링 스케줄러 중지"""
    global _global_scheduler

    if _global_scheduler:
        _global_scheduler.stop()
        _global_scheduler = None
        logger.info("[API_SCHEDULER] API 모니터링 스케줄러 중지 완료")
    else:
        logger.debug("[API_SCHEDULER] 중지할 스케줄러가 없습니다")

def monitor_api_call(api_name: str):
    """API 호출을 모니터링하는 데코레이터"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            monitor = get_api_monitor()
            request_id = monitor.start_request(api_name)

            try:
                result = func(*args, **kwargs)
                monitor.end_request(request_id, success=True)
                return result
            except Exception as e:
                monitor.end_request(request_id, success=False, error_msg=str(e))
                raise

        return wrapper
    return decorator

# Google Sheets API 제한 상수 (분당 100회 읽기 요청)
# Google Sheets API 제한 상수 (실제 제한에 맞게 조정)
GOOGLE_SHEETS_READ_LIMIT = 100   # 분당 100회 읽기 요청 (실제는 더 높을 수 있음)
GOOGLE_SHEETS_WRITE_LIMIT = 100  # 분당 100회 쓰기 요청 (실제는 더 낮을 수 있음)

# 기본 스케줄러 설정
DEFAULT_CHECK_INTERVAL = 300     # 5분마다 임계치 체크
DEFAULT_REPORT_INTERVAL = 3600   # 1시간마다 사용량 리포트

def check_google_sheets_rate_limit() -> Dict[str, Any]:
    """Google Sheets API 속도 제한 상태 확인"""
    monitor = get_api_monitor()
    return monitor.get_rate_limit_status('google_sheets', GOOGLE_SHEETS_READ_LIMIT)

def check_api_thresholds():
    """API 임계치 체크 및 알림"""
    monitor = get_api_monitor()
    monitor.check_thresholds_and_alert()

def setup_api_alerts():
    """API 알림 시스템 설정"""
    monitor = get_api_monitor()

    # 기본 알림 콜백 등록 (로그만 출력)
    def default_alert_callback(alert_data):
        logger.info(f"[API_ALERT] 기본 알림 처리: {alert_data}")

    monitor.register_alert_callback('rate_limit_warning', default_alert_callback)
    monitor.register_alert_callback('high_error_rate', default_alert_callback)
    monitor.register_alert_callback('slow_response', default_alert_callback)

    logger.info("[API_MONITOR] 알림 시스템 설정 완료")

def setup_api_monitoring(auto_start_scheduler: bool = True, check_interval: int = 300, report_interval: int = 3600):
    """API 모니터링 시스템 전체 설정"""
    # 알림 시스템 설정
    setup_api_alerts()

    # 스케줄러 자동 시작
    if auto_start_scheduler:
        success = start_api_monitoring_scheduler(
            check_interval=check_interval,
            report_interval=report_interval
        )
        if success:
            logger.info("[API_MONITOR] API 모니터링 시스템 전체 설정 완료")
        else:
            logger.warning("[API_MONITOR] 스케줄러 시작 실패, 수동 모니터링만 가능")
    else:
        logger.info("[API_MONITOR] API 모니터링 시스템 설정 완료 (스케줄러 수동 시작 필요)")

def get_api_monitoring_status() -> Dict[str, Any]:
    """API 모니터링 시스템 전체 상태 반환"""
    monitor = get_api_monitor()
    scheduler = get_api_scheduler()

    status = {
        'timestamp': datetime.now().isoformat(),
        'monitor': {
            'active': True,  # 모니터는 항상 활성화
            'stats': monitor.get_usage_stats()
        },
        'scheduler': {
            'active': scheduler._running if scheduler else False,
            'stats': scheduler.get_scheduler_stats() if scheduler else None
        }
    }

    return status

def log_api_usage_summary():
    """API 사용량 요약을 로그에 출력"""
    monitor = get_api_monitor()
    stats = monitor.get_usage_stats()

    if not stats['apis']:
        logger.info("[API_MONITOR] 사용량 요약: API 호출 기록이 없습니다")
        return

    logger.info("[API_MONITOR] 사용량 요약:")
    for api_name, api_stats in stats['apis'].items():
        # 상태 판단
        status_indicators = []
        if api_stats['window_error_rate'] > 0.1:
            status_indicators.append("⚠️ 높은 오류율")
        if api_stats['window_avg_response_time'] > 10.0:
            status_indicators.append("🐌 느린 응답")
        if api_stats['active_requests'] > 0:
            status_indicators.append(f"🔄 {api_stats['active_requests']}개 진행중")

        status_suffix = f" [{', '.join(status_indicators)}]" if status_indicators else ""

        logger.info(
            f"  {api_name}: {api_stats['window_calls']}회 호출 "
            f"(오류: {api_stats['window_errors']}회, "
            f"평균응답시간: {api_stats['window_avg_response_time']:.3f}초)"
            f"{status_suffix}"
        )

    # Google Sheets API 특별 체크
    if 'google_sheets' in stats['apis']:
        rate_limit_status = check_google_sheets_rate_limit()
        if rate_limit_status['is_near_limit']:
            logger.warning(
                f"[API_MONITOR] Google Sheets API 사용량 주의: "
                f"{rate_limit_status['usage_percentage']:.1f}% 사용중"
            )