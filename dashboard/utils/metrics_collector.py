"""
메트릭 수집 및 관리 시스템
시스템 성능, 애플리케이션 메트릭, 비즈니스 지표 수집
"""

import time
import threading
import psutil
import os
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Union
from collections import defaultdict, deque
import json
from dataclasses import dataclass, asdict
from enum import Enum
import sqlite3

from .logging_config import get_logger


class MetricType(Enum):
    """메트릭 타입 정의"""
    COUNTER = "counter"          # 누적 카운터 (요청 수, 에러 수)
    GAUGE = "gauge"              # 현재 값 (메모리 사용량, 활성 사용자)
    HISTOGRAM = "histogram"      # 분포 (응답 시간, 파일 크기)
    TIMER = "timer"             # 시간 측정 (작업 소요 시간)


@dataclass
class MetricData:
    """메트릭 데이터 구조"""
    name: str
    value: Union[int, float]
    metric_type: MetricType
    timestamp: datetime
    tags: Dict[str, str] = None
    unit: str = None

    def __post_init__(self):
        if self.tags is None:
            self.tags = {}
        if isinstance(self.metric_type, str):
            self.metric_type = MetricType(self.metric_type)


class MetricsCollector:
    """메트릭 수집 및 관리 클래스"""

    def __init__(self, storage_path: str = None):
        self.logger = get_logger(__name__)

        # 메트릭 저장소
        self.metrics_buffer = defaultdict(deque)  # 최근 메트릭 버퍼
        self.aggregated_metrics = defaultdict(dict)  # 집계된 메트릭

        # 설정
        self.buffer_size = 1000  # 메트릭 버퍼 크기
        self.collection_interval = 60  # 시스템 메트릭 수집 간격 (초)

        # 스레드 안전성
        self.lock = threading.RLock()
        self.running = False
        self.collection_thread = None

        # 데이터베이스 설정
        self.storage_path = storage_path or os.path.join(
            os.path.dirname(__file__), '..', 'logs', 'metrics.db'
        )
        self._init_database()

        # 시스템 메트릭 수집 시작
        self.start_collection()

    def _init_database(self):
        """메트릭 데이터베이스 초기화"""
        os.makedirs(os.path.dirname(self.storage_path), exist_ok=True)

        conn = sqlite3.connect(self.storage_path)
        try:
            conn.execute('''
                CREATE TABLE IF NOT EXISTS metrics (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    value REAL NOT NULL,
                    metric_type TEXT NOT NULL,
                    timestamp DATETIME NOT NULL,
                    tags TEXT,
                    unit TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            conn.execute('''
                CREATE INDEX IF NOT EXISTS idx_metrics_name_time
                ON metrics(name, timestamp)
            ''')

            conn.execute('''
                CREATE INDEX IF NOT EXISTS idx_metrics_timestamp
                ON metrics(timestamp)
            ''')

            conn.commit()
            self.logger.info("메트릭 데이터베이스 초기화 완료")
        finally:
            conn.close()

    def record_metric(self, name: str, value: Union[int, float],
                     metric_type: MetricType, tags: Dict[str, str] = None,
                     unit: str = None, persist: bool = True):
        """메트릭 기록"""

        metric = MetricData(
            name=name,
            value=value,
            metric_type=metric_type,
            timestamp=datetime.now(),
            tags=tags or {},
            unit=unit
        )

        with self.lock:
            # 버퍼에 추가
            self.metrics_buffer[name].append(metric)

            # 버퍼 크기 제한
            if len(self.metrics_buffer[name]) > self.buffer_size:
                self.metrics_buffer[name].popleft()

        # 데이터베이스에 저장
        if persist:
            self._persist_metric(metric)

        self.logger.debug(
            f"Metric recorded: {name}={value}",
            metric_name=name,
            metric_value=value,
            metric_type=metric_type.value
        )

    def _persist_metric(self, metric: MetricData):
        """메트릭을 데이터베이스에 저장"""
        try:
            conn = sqlite3.connect(self.storage_path)
            try:
                conn.execute('''
                    INSERT INTO metrics (name, value, metric_type, timestamp, tags, unit)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (
                    metric.name,
                    metric.value,
                    metric.metric_type.value,
                    metric.timestamp.isoformat(),
                    json.dumps(metric.tags) if metric.tags else None,
                    metric.unit
                ))
                conn.commit()
            finally:
                conn.close()
        except Exception as e:
            self.logger.error(f"메트릭 저장 실패: {e}")

    def increment_counter(self, name: str, value: int = 1, tags: Dict[str, str] = None):
        """카운터 메트릭 증가"""
        self.record_metric(name, value, MetricType.COUNTER, tags)

    def set_gauge(self, name: str, value: Union[int, float],
                  tags: Dict[str, str] = None, unit: str = None):
        """게이지 메트릭 설정"""
        self.record_metric(name, value, MetricType.GAUGE, tags, unit)

    def record_histogram(self, name: str, value: Union[int, float],
                        tags: Dict[str, str] = None, unit: str = None):
        """히스토그램 메트릭 기록"""
        self.record_metric(name, value, MetricType.HISTOGRAM, tags, unit)

    def record_timer(self, name: str, duration: float,
                    tags: Dict[str, str] = None):
        """타이머 메트릭 기록 (밀리초 단위)"""
        self.record_metric(name, duration, MetricType.TIMER, tags, "ms")

    def start_collection(self):
        """시스템 메트릭 수집 시작"""
        if self.running:
            return

        self.running = True
        self.collection_thread = threading.Thread(
            target=self._collect_system_metrics,
            daemon=True
        )
        self.collection_thread.start()
        self.logger.info("시스템 메트릭 수집 시작")

    def stop_collection(self):
        """시스템 메트릭 수집 중지"""
        self.running = False
        if self.collection_thread:
            self.collection_thread.join(timeout=5)
        self.logger.info("시스템 메트릭 수집 중지")

    def _collect_system_metrics(self):
        """시스템 메트릭 수집 (백그라운드 스레드)"""
        while self.running:
            try:
                # CPU 사용률
                cpu_percent = psutil.cpu_percent(interval=1)
                self.set_gauge("system.cpu.usage", cpu_percent, unit="percent")

                # 메모리 사용률
                memory = psutil.virtual_memory()
                self.set_gauge("system.memory.usage", memory.percent, unit="percent")
                self.set_gauge("system.memory.available", memory.available, unit="bytes")
                self.set_gauge("system.memory.used", memory.used, unit="bytes")

                # 디스크 사용률
                disk = psutil.disk_usage('/')
                self.set_gauge("system.disk.usage",
                              (disk.used / disk.total) * 100, unit="percent")
                self.set_gauge("system.disk.free", disk.free, unit="bytes")

                # 프로세스 정보
                process = psutil.Process()
                self.set_gauge("process.memory.usage",
                              process.memory_info().rss, unit="bytes")
                self.set_gauge("process.cpu.usage",
                              process.cpu_percent(), unit="percent")
                self.set_gauge("process.threads.count",
                              process.num_threads(), unit="count")

                # 네트워크 I/O
                net_io = psutil.net_io_counters()
                if net_io:
                    self.set_gauge("system.network.bytes_sent",
                                  net_io.bytes_sent, unit="bytes")
                    self.set_gauge("system.network.bytes_recv",
                                  net_io.bytes_recv, unit="bytes")

                time.sleep(self.collection_interval)

            except Exception as e:
                self.logger.error(f"시스템 메트릭 수집 오류: {e}")
                time.sleep(10)  # 오류 시 10초 대기

    def get_metrics(self, name: str = None,
                   start_time: datetime = None,
                   end_time: datetime = None,
                   limit: int = 100) -> List[Dict]:
        """메트릭 조회"""

        query = "SELECT * FROM metrics WHERE 1=1"
        params = []

        if name:
            query += " AND name = ?"
            params.append(name)

        if start_time:
            query += " AND timestamp >= ?"
            params.append(start_time.isoformat())

        if end_time:
            query += " AND timestamp <= ?"
            params.append(end_time.isoformat())

        query += " ORDER BY timestamp DESC LIMIT ?"
        params.append(limit)

        try:
            conn = sqlite3.connect(self.storage_path)
            conn.row_factory = sqlite3.Row
            try:
                cursor = conn.execute(query, params)
                results = []
                for row in cursor:
                    metric_dict = dict(row)
                    if metric_dict['tags']:
                        metric_dict['tags'] = json.loads(metric_dict['tags'])
                    results.append(metric_dict)
                return results
            finally:
                conn.close()
        except Exception as e:
            self.logger.error(f"메트릭 조회 실패: {e}")
            return []

    def get_aggregated_metrics(self, name: str,
                              start_time: datetime,
                              end_time: datetime,
                              aggregation: str = "avg",
                              interval: str = "1h") -> List[Dict]:
        """집계된 메트릭 조회"""

        # 간격 설정
        interval_seconds = {
            "1m": 60,
            "5m": 300,
            "1h": 3600,
            "1d": 86400
        }.get(interval, 3600)

        # SQL 집계 함수
        agg_func = {
            "avg": "AVG",
            "sum": "SUM",
            "max": "MAX",
            "min": "MIN",
            "count": "COUNT"
        }.get(aggregation, "AVG")

        query = f'''
            SELECT
                name,
                {agg_func}(value) as value,
                datetime((strftime('%s', timestamp) / {interval_seconds}) * {interval_seconds}, 'unixepoch') as time_bucket
            FROM metrics
            WHERE name = ? AND timestamp BETWEEN ? AND ?
            GROUP BY name, time_bucket
            ORDER BY time_bucket
        '''

        try:
            conn = sqlite3.connect(self.storage_path)
            conn.row_factory = sqlite3.Row
            try:
                cursor = conn.execute(query, (name, start_time.isoformat(), end_time.isoformat()))
                return [dict(row) for row in cursor]
            finally:
                conn.close()
        except Exception as e:
            self.logger.error(f"집계 메트릭 조회 실패: {e}")
            return []

    def get_metrics_summary(self) -> Dict[str, Any]:
        """메트릭 요약 정보"""
        try:
            conn = sqlite3.connect(self.storage_path)
            conn.row_factory = sqlite3.Row
            try:
                # 전체 메트릭 수
                total_count = conn.execute("SELECT COUNT(*) as count FROM metrics").fetchone()['count']

                # 메트릭 이름별 수
                name_counts = conn.execute('''
                    SELECT name, COUNT(*) as count
                    FROM metrics
                    GROUP BY name
                    ORDER BY count DESC
                    LIMIT 10
                ''').fetchall()

                # 최근 24시간 메트릭 수
                recent_count = conn.execute('''
                    SELECT COUNT(*) as count
                    FROM metrics
                    WHERE timestamp > datetime('now', '-24 hours')
                ''').fetchone()['count']

                # 메트릭 타입별 분포
                type_distribution = conn.execute('''
                    SELECT metric_type, COUNT(*) as count
                    FROM metrics
                    GROUP BY metric_type
                ''').fetchall()

                return {
                    'total_metrics': total_count,
                    'recent_24h_metrics': recent_count,
                    'top_metrics': [dict(row) for row in name_counts],
                    'type_distribution': [dict(row) for row in type_distribution],
                    'buffer_size': {name: len(buffer) for name, buffer in self.metrics_buffer.items()},
                    'collection_running': self.running
                }
            finally:
                conn.close()
        except Exception as e:
            self.logger.error(f"메트릭 요약 정보 조회 실패: {e}")
            return {}

    def cleanup_old_metrics(self, days: int = 30):
        """오래된 메트릭 데이터 정리"""
        cutoff_date = datetime.now() - timedelta(days=days)

        try:
            conn = sqlite3.connect(self.storage_path)
            try:
                result = conn.execute(
                    "DELETE FROM metrics WHERE timestamp < ?",
                    (cutoff_date.isoformat(),)
                )
                deleted_count = result.rowcount
                conn.commit()

                self.logger.info(f"오래된 메트릭 {deleted_count}개 삭제 완료")
                return deleted_count
            finally:
                conn.close()
        except Exception as e:
            self.logger.error(f"메트릭 정리 실패: {e}")
            return 0


class TimerContext:
    """타이머 컨텍스트 매니저"""

    def __init__(self, collector: MetricsCollector, name: str, tags: Dict[str, str] = None):
        self.collector = collector
        self.name = name
        self.tags = tags or {}
        self.start_time = None

    def __enter__(self):
        self.start_time = time.time()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.start_time:
            duration = (time.time() - self.start_time) * 1000  # ms
            self.collector.record_timer(self.name, duration, self.tags)


def timer_decorator(collector: MetricsCollector, name: str = None, tags: Dict[str, str] = None):
    """타이머 데코레이터"""
    def decorator(func):
        metric_name = name or f"function.{func.__module__}.{func.__name__}.duration"

        def wrapper(*args, **kwargs):
            with TimerContext(collector, metric_name, tags):
                return func(*args, **kwargs)
        return wrapper
    return decorator


# 전역 메트릭 컬렉터
_global_collector = None


def get_metrics_collector() -> MetricsCollector:
    """전역 메트릭 컬렉터 반환"""
    global _global_collector
    if _global_collector is None:
        _global_collector = MetricsCollector()
    return _global_collector


# 편의 함수들
def increment(name: str, value: int = 1, tags: Dict[str, str] = None):
    """카운터 증가"""
    get_metrics_collector().increment_counter(name, value, tags)


def gauge(name: str, value: Union[int, float], tags: Dict[str, str] = None, unit: str = None):
    """게이지 설정"""
    get_metrics_collector().set_gauge(name, value, tags, unit)


def histogram(name: str, value: Union[int, float], tags: Dict[str, str] = None, unit: str = None):
    """히스토그램 기록"""
    get_metrics_collector().record_histogram(name, value, tags, unit)


def timer(name: str, duration: float, tags: Dict[str, str] = None):
    """타이머 기록"""
    get_metrics_collector().record_timer(name, duration, tags)


def time_function(name: str = None, tags: Dict[str, str] = None):
    """함수 실행 시간 측정 데코레이터"""
    return timer_decorator(get_metrics_collector(), name, tags)


def time_block(name: str, tags: Dict[str, str] = None):
    """코드 블록 실행 시간 측정 컨텍스트"""
    return TimerContext(get_metrics_collector(), name, tags)