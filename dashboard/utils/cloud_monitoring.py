"""
GCP 클라우드 모니터링 시스템
- Cloud Monitoring 통합
- 사용자 정의 메트릭
- 실시간 알림
- 헬스체크 시스템
"""

import os
import time
import logging
import threading
from datetime import datetime, timedelta
from typing import Dict, Any, List, Optional
from google.cloud import monitoring_v3
from google.cloud import logging as cloud_logging
import psutil
import requests
from dataclasses import dataclass
from enum import Enum

logger = logging.getLogger(__name__)

class AlertSeverity(Enum):
    """알림 심각도"""
    CRITICAL = "critical"
    WARNING = "warning"
    INFO = "info"

@dataclass
class Metric:
    """메트릭 데이터"""
    name: str
    value: float
    timestamp: datetime
    labels: Dict[str, str] = None

class CloudMonitoring:
    """GCP 클라우드 모니터링"""

    def __init__(self):
        self.project_id = os.getenv('GOOGLE_CLOUD_PROJECT')
        self.client = None
        self.logging_client = None
        self.metrics_queue = []
        self.monitoring_thread = None
        self.running = False
        self.init_clients()

    def init_clients(self):
        """모니터링 클라이언트 초기화"""
        try:
            if self.project_id:
                self.client = monitoring_v3.MetricServiceClient()
                self.logging_client = cloud_logging.Client()
                self.project_name = f"projects/{self.project_id}"
                logger.info("Cloud Monitoring 초기화 완료")
            else:
                logger.warning("GOOGLE_CLOUD_PROJECT 미설정 - 로컬 모니터링 모드")
        except Exception as e:
            logger.error(f"Cloud Monitoring 초기화 실패: {e}")

    def start_monitoring(self):
        """모니터링 시작"""
        if self.running:
            return

        self.running = True
        self.monitoring_thread = threading.Thread(target=self._monitoring_loop, daemon=True)
        self.monitoring_thread.start()
        logger.info("모니터링 시스템 시작")

    def stop_monitoring(self):
        """모니터링 중지"""
        self.running = False
        if self.monitoring_thread:
            self.monitoring_thread.join(timeout=5)
        logger.info("모니터링 시스템 중지")

    def _monitoring_loop(self):
        """모니터링 메인 루프"""
        while self.running:
            try:
                # 시스템 메트릭 수집
                self._collect_system_metrics()

                # 애플리케이션 메트릭 수집
                self._collect_app_metrics()

                # 데이터베이스 메트릭 수집
                self._collect_database_metrics()

                # 헬스체크 수행
                self._perform_health_checks()

                # 메트릭 전송
                self._send_metrics()

                time.sleep(60)  # 1분마다 수집

            except Exception as e:
                logger.error(f"모니터링 루프 오류: {e}")
                time.sleep(10)

    def _collect_system_metrics(self):
        """시스템 메트릭 수집"""
        try:
            # CPU 사용률
            cpu_percent = psutil.cpu_percent(interval=1)
            self.add_metric("system/cpu_usage", cpu_percent, {"unit": "percent"})

            # 메모리 사용률
            memory = psutil.virtual_memory()
            self.add_metric("system/memory_usage", memory.percent, {"unit": "percent"})
            self.add_metric("system/memory_available", memory.available / (1024**3), {"unit": "GB"})

            # 디스크 사용률
            disk = psutil.disk_usage('/')
            disk_percent = (disk.used / disk.total) * 100
            self.add_metric("system/disk_usage", disk_percent, {"unit": "percent"})

            # 네트워크 트래픽
            network = psutil.net_io_counters()
            self.add_metric("system/network_bytes_sent", network.bytes_sent, {"unit": "bytes"})
            self.add_metric("system/network_bytes_recv", network.bytes_recv, {"unit": "bytes"})

        except Exception as e:
            logger.error(f"시스템 메트릭 수집 실패: {e}")

    def _collect_app_metrics(self):
        """애플리케이션 메트릭 수집"""
        try:
            # 현재 프로세스 정보
            process = psutil.Process()

            # 프로세스 메모리 사용량
            memory_info = process.memory_info()
            self.add_metric("app/memory_rss", memory_info.rss / (1024**2), {"unit": "MB"})

            # 프로세스 CPU 사용률
            cpu_percent = process.cpu_percent()
            self.add_metric("app/cpu_usage", cpu_percent, {"unit": "percent"})

            # 열린 파일 디스크립터 수
            try:
                num_fds = process.num_fds()
                self.add_metric("app/open_files", num_fds, {"unit": "count"})
            except (AttributeError, OSError):
                pass  # Windows에서는 지원하지 않음

            # 스레드 수
            num_threads = process.num_threads()
            self.add_metric("app/thread_count", num_threads, {"unit": "count"})

        except Exception as e:
            logger.error(f"애플리케이션 메트릭 수집 실패: {e}")

    def _collect_database_metrics(self):
        """데이터베이스 메트릭 수집"""
        try:
            from .cloud_database_manager import get_cloud_db_manager

            db_manager = get_cloud_db_manager()
            health = db_manager.get_system_health()

            # 데이터베이스 연결 상태
            self.add_metric("database/connection_status", 1 if health['database'] else 0, {"type": "postgres"})
            self.add_metric("database/redis_status", 1 if health['redis'] else 0, {"type": "redis"})

            # Redis 메트릭 (가능한 경우)
            if db_manager.redis_client:
                try:
                    redis_info = db_manager.redis_client.info()
                    self.add_metric("redis/used_memory", redis_info.get('used_memory', 0), {"unit": "bytes"})
                    self.add_metric("redis/connected_clients", redis_info.get('connected_clients', 0), {"unit": "count"})
                    self.add_metric("redis/keyspace_hits", redis_info.get('keyspace_hits', 0), {"unit": "count"})
                    self.add_metric("redis/keyspace_misses", redis_info.get('keyspace_misses', 0), {"unit": "count"})
                except Exception as redis_error:
                    logger.debug(f"Redis 메트릭 수집 실패: {redis_error}")

        except Exception as e:
            logger.error(f"데이터베이스 메트릭 수집 실패: {e}")

    def _perform_health_checks(self):
        """헬스체크 수행"""
        try:
            # 로컬 헬스체크
            health_status = self._check_local_health()
            self.add_metric("health/overall_status", 1 if health_status['healthy'] else 0)

            # 외부 서비스 헬스체크
            self._check_external_services()

        except Exception as e:
            logger.error(f"헬스체크 실패: {e}")

    def _check_local_health(self) -> Dict[str, Any]:
        """로컬 서비스 상태 확인"""
        health = {
            'healthy': True,
            'checks': {},
            'timestamp': datetime.now().isoformat()
        }

        # 메모리 확인 (90% 이상이면 경고)
        memory = psutil.virtual_memory()
        memory_healthy = memory.percent < 90
        health['checks']['memory'] = {
            'healthy': memory_healthy,
            'usage_percent': memory.percent
        }
        if not memory_healthy:
            health['healthy'] = False

        # 디스크 확인 (95% 이상이면 경고)
        disk = psutil.disk_usage('/')
        disk_percent = (disk.used / disk.total) * 100
        disk_healthy = disk_percent < 95
        health['checks']['disk'] = {
            'healthy': disk_healthy,
            'usage_percent': disk_percent
        }
        if not disk_healthy:
            health['healthy'] = False

        # 프로세스 상태 확인
        try:
            process = psutil.Process()
            process_healthy = process.is_running()
            health['checks']['process'] = {
                'healthy': process_healthy,
                'status': process.status()
            }
        except Exception as e:
            health['checks']['process'] = {
                'healthy': False,
                'error': str(e)
            }
            health['healthy'] = False

        return health

    def _check_external_services(self):
        """외부 서비스 상태 확인"""
        # Google Sheets API 확인
        try:
            # 간단한 인증 확인 (실제 API 호출 없이)
            import os
            sheets_healthy = bool(os.getenv('GOOGLE_APPLICATION_CREDENTIALS') or
                                os.getenv('GOOGLE_OAUTH_CLIENT_ID'))
            self.add_metric("external/google_sheets_status", 1 if sheets_healthy else 0)
        except Exception as e:
            logger.debug(f"Google Sheets 상태 확인 실패: {e}")
            self.add_metric("external/google_sheets_status", 0)

        # 인터넷 연결 확인
        try:
            response = requests.get('https://www.google.com', timeout=5)
            internet_healthy = response.status_code == 200
            self.add_metric("external/internet_status", 1 if internet_healthy else 0)
        except Exception as e:
            logger.debug(f"인터넷 연결 확인 실패: {e}")
            self.add_metric("external/internet_status", 0)

    def add_metric(self, metric_name: str, value: float, labels: Dict[str, str] = None):
        """메트릭 추가"""
        metric = Metric(
            name=metric_name,
            value=value,
            timestamp=datetime.now(),
            labels=labels or {}
        )
        self.metrics_queue.append(metric)

    def _send_metrics(self):
        """메트릭을 Cloud Monitoring으로 전송"""
        if not self.client or not self.metrics_queue:
            return

        try:
            series = []
            for metric in self.metrics_queue:
                # 메트릭 디스크립터 생성
                descriptor = monitoring_v3.MetricDescriptor()
                descriptor.type = f"custom.googleapis.com/{metric.name}"
                descriptor.metric_kind = monitoring_v3.MetricDescriptor.MetricKind.GAUGE
                descriptor.value_type = monitoring_v3.MetricDescriptor.ValueType.DOUBLE

                # 시계열 데이터 생성
                series_data = monitoring_v3.TimeSeries()
                series_data.metric.type = descriptor.type

                # 라벨 추가
                for key, value in metric.labels.items():
                    series_data.metric.labels[key] = value

                # 리소스 정보
                series_data.resource.type = "gce_instance"
                series_data.resource.labels["instance_id"] = os.getenv('INSTANCE_ID', 'unknown')
                series_data.resource.labels["zone"] = os.getenv('ZONE', 'unknown')

                # 데이터 포인트
                point = monitoring_v3.Point()
                point.value.double_value = metric.value
                point.interval.end_time.seconds = int(metric.timestamp.timestamp())
                series_data.points = [point]

                series.append(series_data)

            # 배치 전송
            if series:
                self.client.create_time_series(
                    name=self.project_name,
                    time_series=series
                )

            # 큐 비우기
            self.metrics_queue.clear()

        except Exception as e:
            logger.error(f"메트릭 전송 실패: {e}")

    def send_alert(self, message: str, severity: AlertSeverity = AlertSeverity.WARNING,
                   tags: Dict[str, str] = None):
        """알림 전송"""
        try:
            alert_data = {
                'message': message,
                'severity': severity.value,
                'timestamp': datetime.now().isoformat(),
                'tags': tags or {},
                'source': 'itglobal-dashboard'
            }

            # Cloud Logging으로 알림 로그 전송
            if self.logging_client:
                self.logging_client.logger('alerts').log_struct(
                    alert_data,
                    severity=self._get_log_severity(severity)
                )

            # 추가 알림 채널 (이메일, Slack 등)
            self._send_notification_to_channels(alert_data)

        except Exception as e:
            logger.error(f"알림 전송 실패: {e}")

    def _get_log_severity(self, severity: AlertSeverity) -> str:
        """로그 심각도 변환"""
        mapping = {
            AlertSeverity.CRITICAL: 'CRITICAL',
            AlertSeverity.WARNING: 'WARNING',
            AlertSeverity.INFO: 'INFO'
        }
        return mapping.get(severity, 'WARNING')

    def _send_notification_to_channels(self, alert_data: Dict[str, Any]):
        """다양한 채널로 알림 전송"""
        # Slack 웹훅 (환경변수에서 URL 가져오기)
        slack_webhook = os.getenv('SLACK_WEBHOOK_URL')
        if slack_webhook:
            try:
                slack_message = {
                    'text': f"[{alert_data['severity'].upper()}] {alert_data['message']}",
                    'username': 'IT Global Monitor',
                    'icon_emoji': ':warning:' if alert_data['severity'] == 'warning' else ':fire:'
                }
                requests.post(slack_webhook, json=slack_message, timeout=5)
            except Exception as e:
                logger.debug(f"Slack 알림 전송 실패: {e}")

        # 이메일 알림 (SendGrid, SES 등)
        email_api_key = os.getenv('EMAIL_API_KEY')
        if email_api_key and alert_data['severity'] == 'critical':
            try:
                self._send_email_alert(alert_data)
            except Exception as e:
                logger.debug(f"이메일 알림 전송 실패: {e}")

    def _send_email_alert(self, alert_data: Dict[str, Any]):
        """이메일 알림 전송"""
        try:
            # SendGrid API 사용 (패키지가 설치된 경우에만)
            import sendgrid
            from sendgrid.helpers.mail import Mail

            sg = sendgrid.SendGridAPIClient(api_key=os.getenv('EMAIL_API_KEY'))

            # 관리자 이메일 목록
            admin_emails = os.getenv('ADMIN_EMAILS', '').split(',')

            for email in admin_emails:
                if email.strip():
                    message = Mail(
                        from_email='alerts@itglobal.com',
                        to_emails=email.strip(),
                        subject=f"[CRITICAL] IT Global Dashboard Alert",
                        html_content=f"""
                        <h2>시스템 알림</h2>
                        <p><strong>메시지:</strong> {alert_data['message']}</p>
                        <p><strong>심각도:</strong> {alert_data['severity']}</p>
                        <p><strong>시간:</strong> {alert_data['timestamp']}</p>
                        <p><strong>태그:</strong> {alert_data['tags']}</p>
                        """
                    )
                    sg.send(message)
        except ImportError:
            logger.debug("SendGrid 패키지가 설치되지 않음 - 이메일 알림 비활성화")
        except Exception as e:
            logger.error(f"이메일 알림 전송 실패: {e}")

    def create_custom_dashboard(self):
        """커스텀 대시보드 생성"""
        if not self.client:
            return

        try:
            # 대시보드 설정
            dashboard_config = {
                'displayName': 'IT Global Dashboard Monitoring',
                'mosaicLayout': {
                    'tiles': [
                        {
                            'width': 6,
                            'height': 4,
                            'widget': {
                                'title': 'System CPU Usage',
                                'xyChart': {
                                    'dataSets': [{
                                        'timeSeriesQuery': {
                                            'timeSeriesFilter': {
                                                'filter': 'metric.type="custom.googleapis.com/system/cpu_usage"'
                                            }
                                        }
                                    }]
                                }
                            }
                        },
                        {
                            'width': 6,
                            'height': 4,
                            'xPos': 6,
                            'widget': {
                                'title': 'Memory Usage',
                                'xyChart': {
                                    'dataSets': [{
                                        'timeSeriesQuery': {
                                            'timeSeriesFilter': {
                                                'filter': 'metric.type="custom.googleapis.com/system/memory_usage"'
                                            }
                                        }
                                    }]
                                }
                            }
                        }
                    ]
                }
            }

            # 대시보드 생성 API 호출
            # (실제 구현에서는 Dashboard Service 클라이언트 사용)

        except Exception as e:
            logger.error(f"대시보드 생성 실패: {e}")

    def get_monitoring_status(self) -> Dict[str, Any]:
        """모니터링 시스템 상태"""
        return {
            'running': self.running,
            'metrics_queue_size': len(self.metrics_queue),
            'cloud_monitoring_enabled': self.client is not None,
            'project_id': self.project_id,
            'last_health_check': self._check_local_health()
        }


# 전역 모니터링 인스턴스
cloud_monitoring = None

def get_cloud_monitoring() -> CloudMonitoring:
    """클라우드 모니터링 인스턴스 반환"""
    global cloud_monitoring
    if cloud_monitoring is None:
        cloud_monitoring = CloudMonitoring()
    return cloud_monitoring

def init_monitoring():
    """모니터링 시스템 초기화"""
    global cloud_monitoring
    cloud_monitoring = CloudMonitoring()
    cloud_monitoring.start_monitoring()
    return cloud_monitoring

def send_alert(message: str, severity: AlertSeverity = AlertSeverity.WARNING):
    """편의 함수: 알림 전송"""
    monitoring = get_cloud_monitoring()
    monitoring.send_alert(message, severity)