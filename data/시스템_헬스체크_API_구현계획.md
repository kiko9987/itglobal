# 시스템 헬스 체크 API 구현 계획

**작성일:** 2025-10-31
**목적:** 24/7 운영 환경에서 메모리, 스레드, Redis 상태 등을 실시간 모니터링
**우선순위:** P3 (선택적 개선)

---

## 📋 목차

1. [개요](#개요)
2. [요구사항 분석](#요구사항-분석)
3. [API 설계](#api-설계)
4. [백엔드 구현](#백엔드-구현)
5. [프론트엔드 통합](#프론트엔드-통합)
6. [테스트 계획](#테스트-계획)
7. [배포 및 모니터링](#배포-및-모니터링)

---

## 개요

### 목적
- 시스템 리소스(메모리, 스레드, GC) 실시간 모니터링
- Redis 연결 상태 및 Fallback 모드 감지
- 관리자에게 임계값 초과 시 조기 경고
- 장기 운영 시 메모리 누수 패턴 조기 발견

### 기대 효과
- 🟢 **사전 예방:** 메모리 500MB 초과 시 경고 → OOM 방지
- 🟢 **빠른 대응:** Redis 장애 즉시 감지 → 수동 복구 가능
- 🟢 **운영 가시성:** 시스템 상태를 실시간 대시보드로 확인

---

## 요구사항 분석

### 기능 요구사항

#### 1. 메모리 모니터링
- **지표:**
  - RSS (Resident Set Size): 프로세스가 실제 사용 중인 물리 메모리
  - Percent: 시스템 전체 대비 사용률
- **임계값:**
  - ⚠️ 경고: 500MB
  - 🔴 위험: 1GB

#### 2. 스레드 모니터링
- **지표:**
  - Active thread count
  - Thread list (선택적, 디버그 모드)
- **임계값:**
  - ⚠️ 경고: 30개
  - 🔴 위험: 50개

#### 3. GC (Garbage Collector) 모니터링
- **지표:**
  - Generation counts (gen0, gen1, gen2)
  - GC threshold 설정값
- **의미:**
  - gen0 높음: 짧은 생명주기 객체 많음 (정상)
  - gen2 지속 증가: 메모리 누수 의심

#### 4. Redis 상태 모니터링
- **지표:**
  - Connected: Redis 정상 연결 여부
  - Fallback active: 메모리 캐시 사용 중 여부
  - Fallback 캐시 사용률
- **알림:**
  - Fallback 모드 진입 시 즉시 경고

#### 5. 캐시 통계
- **지표:**
  - Hit rate (캐시 히트율)
  - Miss rate
  - Total items
  - Eviction count (만료/제거된 항목 수)

### 비기능 요구사항

1. **성능:**
   - API 응답 시간 < 100ms
   - 프로세스에 오버헤드 최소화 (< 1%)

2. **보안:**
   - 관리자 권한만 접근 가능 (`@admin_required`)
   - 민감한 시스템 정보 노출 방지

3. **확장성:**
   - 추가 지표 쉽게 추가 가능한 구조
   - 모니터링 플랫폼(Prometheus, Grafana) 연동 가능

---

## API 설계

### 엔드포인트 명세

#### 1. 기본 헬스 체크 API

**Endpoint:** `GET /api/system/health`

**권한:** `@login_required` (모든 로그인 사용자)

**응답 예시:**
```json
{
  "status": "healthy",
  "timestamp": "2025-10-31T10:30:45.123456",
  "uptime_seconds": 86400,
  "memory": {
    "rss_mb": 245.6,
    "percent": 12.3,
    "threshold_warning": 500,
    "threshold_critical": 1000,
    "status": "normal"
  },
  "threads": {
    "active": 18,
    "threshold_warning": 30,
    "threshold_critical": 50,
    "status": "normal"
  },
  "gc": {
    "counts": [450, 12, 3],
    "threshold": [700, 10, 10],
    "status": "normal"
  },
  "redis": {
    "connected": true,
    "fallback_active": false,
    "status": "normal"
  },
  "overall_status": "healthy"
}
```

**Status 값:**
- `"healthy"`: 모든 지표 정상
- `"warning"`: 하나 이상의 지표가 경고 수준
- `"critical"`: 하나 이상의 지표가 위험 수준

---

#### 2. 상세 시스템 정보 API (관리자 전용)

**Endpoint:** `GET /api/system/health/detailed`

**권한:** `@admin_required`

**응답 예시:**
```json
{
  "status": "healthy",
  "timestamp": "2025-10-31T10:30:45.123456",
  "process": {
    "pid": 12345,
    "ppid": 1000,
    "create_time": "2025-10-30T10:00:00.000000",
    "uptime_seconds": 86400,
    "cpu_percent": 2.5,
    "num_threads": 18,
    "open_files": 45,
    "connections": 12
  },
  "memory": {
    "rss_mb": 245.6,
    "vms_mb": 1024.0,
    "percent": 12.3,
    "available_mb": 7800.0,
    "total_mb": 16384.0,
    "status": "normal"
  },
  "threads": {
    "active": 18,
    "list": [
      "MainThread",
      "Thread-1 (CachePreloaderRefresh)",
      "Thread-2 (BackgroundPrefetch)",
      "..."
    ],
    "daemon_count": 15,
    "non_daemon_count": 3,
    "status": "normal"
  },
  "gc": {
    "counts": [450, 12, 3],
    "threshold": [700, 10, 10],
    "collected_objects": {
      "gen0": 12450,
      "gen1": 345,
      "gen2": 12
    },
    "status": "normal"
  },
  "redis": {
    "connected": true,
    "fallback_active": false,
    "ping_latency_ms": 0.5,
    "info": {
      "version": "7.0.12",
      "uptime_seconds": 259200,
      "used_memory_mb": 45.2
    },
    "status": "normal"
  },
  "cache": {
    "type": "redis",
    "hit_rate": 95.6,
    "miss_rate": 4.4,
    "total_requests": 12450,
    "hits": 11901,
    "misses": 549,
    "total_items": 156,
    "evictions": 23
  },
  "fallback_cache": {
    "active": false,
    "total_items": 0,
    "max_size": 2000,
    "usage_percent": 0.0
  },
  "background_workers": {
    "cache_preloader": {
      "running": true,
      "last_refresh": "2025-10-31T10:28:30.000000"
    },
    "background_prefetch": {
      "running": true,
      "jobs": 1,
      "last_prefetch": "2025-10-31T10:29:50.000000"
    },
    "calendar_sync": {
      "running": false,
      "reason": "credentials_missing"
    }
  }
}
```

---

#### 3. 헬스 체크 히스토리 API (선택적)

**Endpoint:** `GET /api/system/health/history`

**권한:** `@admin_required`

**쿼리 파라미터:**
- `hours`: 최근 N시간 데이터 (기본값: 24)
- `metric`: 특정 지표만 조회 (memory, threads, gc, redis)

**응답 예시:**
```json
{
  "metric": "memory",
  "hours": 24,
  "data_points": [
    {
      "timestamp": "2025-10-31T10:00:00.000000",
      "value": 245.6
    },
    {
      "timestamp": "2025-10-31T11:00:00.000000",
      "value": 247.2
    },
    "..."
  ],
  "statistics": {
    "min": 230.5,
    "max": 260.8,
    "avg": 246.3,
    "trend": "stable"
  }
}
```

**Note:** 히스토리 데이터는 Redis에 저장 (TTL: 7일)

---

## 백엔드 구현

### 파일 구조

```
dashboard/
├── blueprints/
│   └── monitoring.py  # 기존 파일에 추가
├── utils/
│   ├── system_monitor.py  # 신규 파일 (핵심 로직)
│   └── health_check.py    # 신규 파일 (헬스 체크 통합)
└── services/
    └── health_history.py  # 신규 파일 (히스토리 저장/조회)
```

---

### 1. SystemMonitor 클래스 (utils/system_monitor.py)

```python
"""
시스템 리소스 모니터링 유틸리티
"""

import psutil
import gc
import threading
import time
from typing import Dict, Any, List
from datetime import datetime
from dashboard.utils.logging_config import get_logger
from dashboard.utils.redis_client import get_redis_client
from dashboard.utils.smart_cache_manager import get_smart_cache
from dashboard.utils.fallback_cache import get_fallback_cache

logger = get_logger(__name__)


class SystemMonitor:
    """시스템 리소스 모니터링 클래스"""

    # 임계값 설정
    MEMORY_WARNING_MB = 500
    MEMORY_CRITICAL_MB = 1000
    THREADS_WARNING = 30
    THREADS_CRITICAL = 50

    def __init__(self):
        self.process = psutil.Process()
        self.start_time = time.time()

    def get_memory_info(self) -> Dict[str, Any]:
        """메모리 정보 조회"""
        try:
            mem_info = self.process.memory_info()
            rss_mb = mem_info.rss / 1024 / 1024
            vms_mb = mem_info.vms / 1024 / 1024
            percent = self.process.memory_percent()

            # 시스템 전체 메모리
            system_mem = psutil.virtual_memory()

            # 상태 판단
            if rss_mb >= self.MEMORY_CRITICAL_MB:
                status = "critical"
            elif rss_mb >= self.MEMORY_WARNING_MB:
                status = "warning"
            else:
                status = "normal"

            return {
                "rss_mb": round(rss_mb, 2),
                "vms_mb": round(vms_mb, 2),
                "percent": round(percent, 2),
                "available_mb": round(system_mem.available / 1024 / 1024, 2),
                "total_mb": round(system_mem.total / 1024 / 1024, 2),
                "threshold_warning": self.MEMORY_WARNING_MB,
                "threshold_critical": self.MEMORY_CRITICAL_MB,
                "status": status
            }
        except Exception as e:
            logger.error(f"메모리 정보 조회 실패: {e}")
            return {"status": "error", "error": str(e)}

    def get_thread_info(self, include_list: bool = False) -> Dict[str, Any]:
        """스레드 정보 조회"""
        try:
            active_count = threading.active_count()

            result = {
                "active": active_count,
                "threshold_warning": self.THREADS_WARNING,
                "threshold_critical": self.THREADS_CRITICAL,
                "status": "normal"
            }

            # 상태 판단
            if active_count >= self.THREADS_CRITICAL:
                result["status"] = "critical"
            elif active_count >= self.THREADS_WARNING:
                result["status"] = "warning"

            # 스레드 목록 포함 (선택적, 관리자 전용)
            if include_list:
                threads = threading.enumerate()
                result["list"] = [t.name for t in threads]
                result["daemon_count"] = sum(1 for t in threads if t.daemon)
                result["non_daemon_count"] = sum(1 for t in threads if not t.daemon)

            return result
        except Exception as e:
            logger.error(f"스레드 정보 조회 실패: {e}")
            return {"status": "error", "error": str(e)}

    def get_gc_info(self) -> Dict[str, Any]:
        """GC 정보 조회"""
        try:
            counts = gc.get_count()
            threshold = gc.get_threshold()

            # GC 통계
            stats = gc.get_stats()

            return {
                "counts": list(counts),
                "threshold": list(threshold),
                "collected_objects": {
                    "gen0": stats[0].get("collected", 0) if stats else 0,
                    "gen1": stats[1].get("collected", 0) if len(stats) > 1 else 0,
                    "gen2": stats[2].get("collected", 0) if len(stats) > 2 else 0
                },
                "status": "normal"
            }
        except Exception as e:
            logger.error(f"GC 정보 조회 실패: {e}")
            return {"status": "error", "error": str(e)}

    def get_redis_info(self) -> Dict[str, Any]:
        """Redis 상태 정보 조회"""
        try:
            redis_client = get_redis_client()
            fallback_active = getattr(redis_client, '_use_fallback', False)

            result = {
                "connected": not fallback_active,
                "fallback_active": fallback_active,
                "status": "warning" if fallback_active else "normal"
            }

            # Redis 연결 시 추가 정보
            if not fallback_active:
                try:
                    # Ping latency 측정
                    start = time.time()
                    redis_client.redis.ping()
                    latency = (time.time() - start) * 1000
                    result["ping_latency_ms"] = round(latency, 2)

                    # Redis 서버 정보
                    info = redis_client.redis.info()
                    result["info"] = {
                        "version": info.get("redis_version", "unknown"),
                        "uptime_seconds": info.get("uptime_in_seconds", 0),
                        "used_memory_mb": round(info.get("used_memory", 0) / 1024 / 1024, 2)
                    }
                except Exception as e:
                    logger.warning(f"Redis 상세 정보 조회 실패: {e}")

            return result
        except Exception as e:
            logger.error(f"Redis 정보 조회 실패: {e}")
            return {"status": "error", "error": str(e)}

    def get_cache_info(self) -> Dict[str, Any]:
        """캐시 통계 정보 조회"""
        try:
            cache = get_smart_cache()
            cache_info = cache.get_cache_info()

            # 히트율 계산
            total_requests = cache_info.get("hit_count", 0) + cache_info.get("miss_count", 0)
            hit_rate = (cache_info.get("hit_count", 0) / total_requests * 100) if total_requests > 0 else 0
            miss_rate = 100 - hit_rate

            return {
                "type": "redis" if not getattr(get_redis_client(), '_use_fallback', False) else "fallback",
                "hit_rate": round(hit_rate, 2),
                "miss_rate": round(miss_rate, 2),
                "total_requests": total_requests,
                "hits": cache_info.get("hit_count", 0),
                "misses": cache_info.get("miss_count", 0),
                "total_items": cache_info.get("total_items", 0),
                "evictions": cache_info.get("eviction_count", 0)
            }
        except Exception as e:
            logger.error(f"캐시 정보 조회 실패: {e}")
            return {"status": "error", "error": str(e)}

    def get_fallback_cache_info(self) -> Dict[str, Any]:
        """Fallback 캐시 정보 조회"""
        try:
            redis_client = get_redis_client()
            fallback_active = getattr(redis_client, '_use_fallback', False)

            result = {
                "active": fallback_active
            }

            if fallback_active:
                fallback_cache = get_fallback_cache()
                stats = fallback_cache.get_stats()
                result.update({
                    "total_items": stats.get("total_items", 0),
                    "max_size": stats.get("max_size", 2000),
                    "usage_percent": round((stats.get("total_items", 0) / stats.get("max_size", 2000)) * 100, 2)
                })
            else:
                result.update({
                    "total_items": 0,
                    "max_size": 2000,
                    "usage_percent": 0.0
                })

            return result
        except Exception as e:
            logger.error(f"Fallback 캐시 정보 조회 실패: {e}")
            return {"status": "error", "error": str(e)}

    def get_background_workers_info(self) -> Dict[str, Any]:
        """백그라운드 워커 상태 조회"""
        result = {}

        try:
            # CachePreloader 상태
            from dashboard.utils.cache_preloader import get_cache_preloader
            preloader = get_cache_preloader()
            result["cache_preloader"] = {
                "running": preloader._running if hasattr(preloader, '_running') else False
            }
        except Exception as e:
            result["cache_preloader"] = {"running": False, "error": str(e)}

        try:
            # BackgroundPrefetch 상태
            from dashboard.utils.background_prefetch import get_background_prefetch
            prefetch = get_background_prefetch()
            stats = prefetch.get_stats()
            result["background_prefetch"] = {
                "running": stats.get("running", False),
                "jobs": len(stats.get("jobs", {})),
                "last_prefetch": stats.get("jobs", {}).get("current_sheet_data", {}).get("last_prefetch")
            }
        except Exception as e:
            result["background_prefetch"] = {"running": False, "error": str(e)}

        try:
            # Calendar Sync 상태
            from dashboard.services.calendar_sync_scheduler import get_calendar_sync_scheduler
            sync_scheduler = get_calendar_sync_scheduler()
            result["calendar_sync"] = {
                "running": sync_scheduler._running if hasattr(sync_scheduler, '_running') else False
            }
        except Exception as e:
            result["calendar_sync"] = {"running": False, "reason": "credentials_missing"}

        return result

    def get_process_info(self) -> Dict[str, Any]:
        """프로세스 정보 조회"""
        try:
            return {
                "pid": self.process.pid,
                "ppid": self.process.ppid(),
                "create_time": datetime.fromtimestamp(self.process.create_time()).isoformat(),
                "uptime_seconds": int(time.time() - self.start_time),
                "cpu_percent": round(self.process.cpu_percent(interval=0.1), 2),
                "num_threads": self.process.num_threads(),
                "open_files": len(self.process.open_files()),
                "connections": len(self.process.connections())
            }
        except Exception as e:
            logger.error(f"프로세스 정보 조회 실패: {e}")
            return {"status": "error", "error": str(e)}

    def determine_overall_status(self, components: Dict[str, Any]) -> str:
        """전체 상태 판단"""
        statuses = []
        for key, value in components.items():
            if isinstance(value, dict) and "status" in value:
                statuses.append(value["status"])

        if "critical" in statuses:
            return "critical"
        elif "warning" in statuses:
            return "warning"
        elif "error" in statuses:
            return "degraded"
        else:
            return "healthy"


# 싱글톤 인스턴스
_system_monitor: SystemMonitor = None


def get_system_monitor() -> SystemMonitor:
    """SystemMonitor 싱글톤 인스턴스 반환"""
    global _system_monitor
    if _system_monitor is None:
        _system_monitor = SystemMonitor()
    return _system_monitor
```

---

### 2. HealthCheck 클래스 (utils/health_check.py)

```python
"""
헬스 체크 통합 클래스
"""

from datetime import datetime
from typing import Dict, Any
from dashboard.utils.system_monitor import get_system_monitor
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


class HealthCheck:
    """헬스 체크 통합"""

    def __init__(self):
        self.monitor = get_system_monitor()

    def get_basic_health(self) -> Dict[str, Any]:
        """기본 헬스 체크 정보"""
        memory = self.monitor.get_memory_info()
        threads = self.monitor.get_thread_info(include_list=False)
        gc_info = self.monitor.get_gc_info()
        redis = self.monitor.get_redis_info()

        components = {
            "memory": memory,
            "threads": threads,
            "gc": gc_info,
            "redis": redis
        }

        overall_status = self.monitor.determine_overall_status(components)

        return {
            "status": overall_status,
            "timestamp": datetime.now().isoformat(),
            "uptime_seconds": int(self.monitor.process.create_time()),
            "memory": memory,
            "threads": threads,
            "gc": gc_info,
            "redis": redis,
            "overall_status": overall_status
        }

    def get_detailed_health(self) -> Dict[str, Any]:
        """상세 헬스 체크 정보 (관리자 전용)"""
        basic = self.get_basic_health()

        # 추가 정보
        basic["process"] = self.monitor.get_process_info()
        basic["threads"] = self.monitor.get_thread_info(include_list=True)
        basic["cache"] = self.monitor.get_cache_info()
        basic["fallback_cache"] = self.monitor.get_fallback_cache_info()
        basic["background_workers"] = self.monitor.get_background_workers_info()

        return basic


# 싱글톤 인스턴스
_health_check: HealthCheck = None


def get_health_check() -> HealthCheck:
    """HealthCheck 싱글톤 인스턴스 반환"""
    global _health_check
    if _health_check is None:
        _health_check = HealthCheck()
    return _health_check
```

---

### 3. Blueprint 라우트 추가 (blueprints/monitoring.py)

```python
# 기존 monitoring.py 파일에 추가

from dashboard.utils.health_check import get_health_check

@monitoring_bp.route('/api/system/health')
@login_required
def system_health():
    """시스템 헬스 체크 (기본)"""
    try:
        health_check = get_health_check()
        result = health_check.get_basic_health()
        return jsonify(result)
    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 헬스 체크 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '시스템 상태를 조회할 수 없습니다',
            'error_id': error_id
        }), 500


@monitoring_bp.route('/api/system/health/detailed')
@admin_required
def system_health_detailed():
    """시스템 헬스 체크 (상세, 관리자 전용)"""
    try:
        health_check = get_health_check()
        result = health_check.get_detailed_health()
        return jsonify(result)
    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 상세 헬스 체크 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '시스템 상태를 조회할 수 없습니다',
            'error_id': error_id
        }), 500
```

---

## 프론트엔드 통합

### 1. 시스템 상태 위젯 (간단한 버전)

```html
<!-- templates/base.html 또는 dashboard.html에 추가 -->

<div id="system-status-widget" class="status-widget">
    <div class="status-indicator" id="status-indicator">
        <span class="status-dot"></span>
        <span class="status-text">시스템 확인 중...</span>
    </div>
</div>
```

```css
/* static/css/system-status.css */

.status-widget {
    position: fixed;
    top: 60px;
    right: 20px;
    background: white;
    border-radius: 8px;
    padding: 12px 16px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    z-index: 1000;
}

.status-indicator {
    display: flex;
    align-items: center;
    gap: 8px;
}

.status-dot {
    width: 12px;
    height: 12px;
    border-radius: 50%;
    animation: pulse 2s infinite;
}

.status-dot.healthy {
    background-color: #10b981;
}

.status-dot.warning {
    background-color: #f59e0b;
}

.status-dot.critical {
    background-color: #ef4444;
}

@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.5; }
}
```

```javascript
// static/js/system-status.js

class SystemStatusMonitor {
    constructor() {
        this.checkInterval = 5 * 60 * 1000; // 5분마다
        this.lastStatus = 'unknown';
        this.init();
    }

    async init() {
        await this.checkHealth();
        setInterval(() => this.checkHealth(), this.checkInterval);
    }

    async checkHealth() {
        try {
            const response = await fetch('/api/system/health');
            const data = await response.json();

            this.updateUI(data);
            this.checkAlerts(data);
        } catch (error) {
            console.error('시스템 상태 확인 실패:', error);
            this.updateUI({ overall_status: 'error' });
        }
    }

    updateUI(data) {
        const indicator = document.getElementById('status-indicator');
        const dot = indicator.querySelector('.status-dot');
        const text = indicator.querySelector('.status-text');

        dot.className = `status-dot ${data.overall_status}`;

        const statusText = {
            'healthy': '시스템 정상',
            'warning': '주의 필요',
            'critical': '긴급 조치 필요',
            'error': '상태 확인 불가'
        };

        text.textContent = statusText[data.overall_status] || '상태 확인 중';

        // 클릭 시 상세 정보 표시
        indicator.onclick = () => this.showDetails(data);
    }

    checkAlerts(data) {
        // 메모리 경고
        if (data.memory && data.memory.status === 'warning') {
            this.showWarning(`메모리 사용량 높음: ${data.memory.rss_mb} MB`);
        }

        if (data.memory && data.memory.status === 'critical') {
            this.showCritical(`메모리 위험 수준: ${data.memory.rss_mb} MB`);
        }

        // Redis fallback 경고
        if (data.redis && data.redis.fallback_active) {
            this.showWarning('Redis 연결 실패 - Fallback 모드 실행 중');
        }

        // 스레드 경고
        if (data.threads && data.threads.status === 'warning') {
            this.showWarning(`스레드 수 높음: ${data.threads.active}개`);
        }
    }

    showWarning(message) {
        if (this.lastStatus === 'critical') return; // 위험 알림 우선

        console.warn('[시스템 경고]', message);
        // Toast 알림 또는 알림 센터에 표시
        this.showNotification('warning', message);
    }

    showCritical(message) {
        this.lastStatus = 'critical';
        console.error('[시스템 위험]', message);
        this.showNotification('critical', message);
    }

    showNotification(level, message) {
        // 기존 알림 시스템 활용 또는 브라우저 알림
        if (window.Notification && Notification.permission === 'granted') {
            new Notification(`시스템 ${level === 'critical' ? '위험' : '경고'}`, {
                body: message,
                icon: '/static/img/alert-icon.png'
            });
        }
    }

    showDetails(data) {
        // 모달 또는 사이드 패널로 상세 정보 표시
        const details = `
            메모리: ${data.memory.rss_mb} MB (${data.memory.percent}%)
            스레드: ${data.threads.active}개
            Redis: ${data.redis.connected ? '정상' : 'Fallback 모드'}
        `;
        alert(details); // 실제로는 모달 사용
    }
}

// 초기화
document.addEventListener('DOMContentLoaded', () => {
    new SystemStatusMonitor();
});
```

---

### 2. 관리자 대시보드 페이지 (선택적)

```html
<!-- templates/system_health_dashboard.html -->

<!DOCTYPE html>
<html>
<head>
    <title>시스템 헬스 모니터링</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head>
<body>
    <div class="container">
        <h1>시스템 헬스 모니터링</h1>

        <!-- 전체 상태 카드 -->
        <div class="status-card">
            <h2>전체 상태: <span id="overall-status">확인 중...</span></h2>
        </div>

        <!-- 메모리 차트 -->
        <div class="chart-container">
            <canvas id="memory-chart"></canvas>
        </div>

        <!-- 스레드 차트 -->
        <div class="chart-container">
            <canvas id="threads-chart"></canvas>
        </div>

        <!-- Redis 상태 -->
        <div class="info-card">
            <h3>Redis 상태</h3>
            <div id="redis-status"></div>
        </div>

        <!-- 캐시 통계 -->
        <div class="info-card">
            <h3>캐시 통계</h3>
            <div id="cache-stats"></div>
        </div>
    </div>

    <script>
        // Chart.js로 실시간 그래프 표시
        // 5분마다 /api/system/health/detailed 호출하여 업데이트
    </script>
</body>
</html>
```

---

## 테스트 계획

### 1. 단위 테스트

```python
# tests/unit/test_system_monitor.py

import pytest
from dashboard.utils.system_monitor import SystemMonitor

def test_memory_info():
    """메모리 정보 조회 테스트"""
    monitor = SystemMonitor()
    result = monitor.get_memory_info()

    assert "rss_mb" in result
    assert "status" in result
    assert result["rss_mb"] > 0

def test_thread_info():
    """스레드 정보 조회 테스트"""
    monitor = SystemMonitor()
    result = monitor.get_thread_info()

    assert "active" in result
    assert result["active"] > 0

def test_redis_fallback_detection():
    """Redis fallback 감지 테스트"""
    monitor = SystemMonitor()
    result = monitor.get_redis_info()

    assert "fallback_active" in result
    assert isinstance(result["fallback_active"], bool)
```

### 2. 통합 테스트

```python
# tests/integration/test_health_api.py

def test_health_api_basic(client, login_as_user):
    """기본 헬스 체크 API 테스트"""
    response = client.get('/api/system/health')
    assert response.status_code == 200

    data = response.get_json()
    assert data["overall_status"] in ["healthy", "warning", "critical"]

def test_health_api_detailed(client, login_as_admin):
    """상세 헬스 체크 API 테스트 (관리자)"""
    response = client.get('/api/system/health/detailed')
    assert response.status_code == 200

    data = response.get_json()
    assert "process" in data
    assert "cache" in data

def test_health_api_permission(client, login_as_viewer):
    """권한 없는 사용자 접근 테스트"""
    response = client.get('/api/system/health/detailed')
    assert response.status_code == 403
```

### 3. 부하 테스트

```python
# tests/performance/test_health_api_performance.py

import time
import statistics

def test_health_api_performance(client, login_as_user):
    """헬스 체크 API 성능 테스트"""
    response_times = []

    for _ in range(100):
        start = time.time()
        response = client.get('/api/system/health')
        elapsed = (time.time() - start) * 1000

        assert response.status_code == 200
        response_times.append(elapsed)

    avg = statistics.mean(response_times)
    p95 = statistics.quantiles(response_times, n=20)[18]

    print(f"평균: {avg:.2f}ms, P95: {p95:.2f}ms")

    # 요구사항: P95 < 100ms
    assert p95 < 100
```

---

## 배포 및 모니터링

### 1. 배포 체크리스트

- [ ] `dashboard/utils/system_monitor.py` 파일 추가
- [ ] `dashboard/utils/health_check.py` 파일 추가
- [ ] `dashboard/blueprints/monitoring.py`에 라우트 추가
- [ ] 프론트엔드 파일 추가 (HTML/CSS/JS)
- [ ] 단위 테스트 실행 및 통과
- [ ] 통합 테스트 실행 및 통과
- [ ] 성능 테스트 실행 (P95 < 100ms)
- [ ] 관리자 권한 확인
- [ ] 프로덕션 배포

### 2. 운영 가이드

#### 정상 상태 모니터링
```bash
# 수동 확인
curl http://localhost:5000/api/system/health

# 예상 응답
{
  "overall_status": "healthy",
  "memory": { "status": "normal", "rss_mb": 245 },
  "threads": { "status": "normal", "active": 18 },
  "redis": { "status": "normal", "connected": true }
}
```

#### 경고 발생 시 조치
```bash
# 메모리 경고
- 현상: memory.rss_mb > 500MB
- 조치:
  1. 프로세스 재시작 고려
  2. 메모리 프로파일링 실행 (tracemalloc)
  3. 불필요한 캐시 삭제 (/admin/api/cache-clear)

# Redis fallback 경고
- 현상: redis.fallback_active == true
- 조치:
  1. Redis 서버 상태 확인 (redis-cli ping)
  2. Redis 재시작
  3. Fallback 캐시 사용률 모니터링
```

### 3. Prometheus 연동 (선택적)

```python
# dashboard/utils/prometheus_exporter.py

from prometheus_client import Gauge, generate_latest, CONTENT_TYPE_LATEST
from flask import Response

# 메트릭 정의
memory_usage = Gauge('dashboard_memory_mb', 'Memory usage in MB')
thread_count = Gauge('dashboard_threads', 'Active thread count')
redis_connected = Gauge('dashboard_redis_connected', 'Redis connection status (1=connected, 0=fallback)')

def update_metrics():
    """메트릭 업데이트"""
    monitor = get_system_monitor()

    mem = monitor.get_memory_info()
    threads = monitor.get_thread_info()
    redis = monitor.get_redis_info()

    memory_usage.set(mem.get("rss_mb", 0))
    thread_count.set(threads.get("active", 0))
    redis_connected.set(1 if redis.get("connected") else 0)

@monitoring_bp.route('/metrics')
def metrics():
    """Prometheus 메트릭 엔드포인트"""
    update_metrics()
    return Response(generate_latest(), mimetype=CONTENT_TYPE_LATEST)
```

---

## 마무리

### 구현 순서 권장

1. **Phase 1: 백엔드 기본 구현** (2-3시간)
   - `system_monitor.py` 작성
   - `health_check.py` 작성
   - 기본 API 라우트 추가

2. **Phase 2: 테스트** (1-2시간)
   - 단위 테스트 작성 및 실행
   - 통합 테스트 작성 및 실행

3. **Phase 3: 프론트엔드 기본 위젯** (1-2시간)
   - 시스템 상태 표시 위젯 추가
   - 5분마다 자동 체크 구현

4. **Phase 4: 관리자 대시보드** (선택적, 3-4시간)
   - 상세 페이지 구현
   - Chart.js 그래프 추가

5. **Phase 5: 고급 기능** (선택적)
   - 히스토리 저장/조회 구현
   - Prometheus 연동
   - 알림 시스템 통합

### 예상 작업 시간
- **최소 구현** (Phase 1-2): 4-5시간
- **권장 구현** (Phase 1-3): 6-7시간
- **전체 구현** (Phase 1-5): 10-15시간

### 다음 단계

이 문서를 바탕으로 구현을 시작하시겠습니까?
1. Phase 1부터 단계적으로 구현
2. 특정 Phase만 선택적으로 구현
3. 추가 요구사항 확인 후 구현
