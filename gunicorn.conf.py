"""
Gunicorn configuration for Redis-based distributed deployment
Optimized for 20 internal users on office network

변경 이력:
- 2025-01: Redis 도입 후 다중 worker 지원 가능
- 단계별 workers 증가 전략 수립 (1→2→4)
"""

import os
import multiprocessing

# ====================================================================
# Workers 설정: Redis 기반 분산 환경 대응
# ====================================================================
#
# 단계별 workers 증가 전략:
#
# 【1단계】 workers=1 (현재) ← 초기 Redis 안정화
#   - Redis 전환 직후 안정성 검증 기간
#   - 모니터링 지표: CPU 사용률, 응답 시간, 에러율
#   - 기간: 1-2주
#
# 【2단계】 workers=2 (CPU 사용률 60% 초과 시)
#   - CPU 병목 확인 후 증가
#   - 모니터링 지표: 락 경합, 캐시 히트율
#   - 조건:
#     * CPU 사용률 평균 60% 이상
#     * 응답 시간 증가 추세
#     * 에러율 1% 미만 유지
#
# 【3단계】 workers=4 (동시 사용자 15명 이상)
#   - 사용자 증가 시 최종 증가
#   - 로컬 PC 환경 고려 (4 core 기준)
#   - 조건:
#     * 피크 타임 동시 접속 15명+
#     * CPU 여유 확보 (70% 미만)
#
# 환경 변수로 재정의 가능:
#   GUNICORN_WORKERS=2 gunicorn ...
#
# ====================================================================

workers = int(os.getenv('GUNICORN_WORKERS', 1))  # 기본값: 1

# Thread pool for concurrent request handling
threads = 4

# Use threaded worker class
worker_class = "gthread"

# Timeout for requests (2 minutes)
timeout = 120

# Binding address
bind = "0.0.0.0:8000"

# Logging
loglevel = "info"
accesslog = "-"
errorlog = "-"

# Graceful timeout for worker shutdown
graceful_timeout = 30

# Keep-alive
keepalive = 5

# Worker restart configuration (메모리 누수 방지)
max_requests = 1000
max_requests_jitter = 50

# ====================================================================
# Redis 환경 모니터링 가이드
# ====================================================================
#
# 1. CPU 사용률 확인:
#    - Windows: 작업 관리자 → 성능 탭
#    - Linux: top, htop
#
# 2. Gunicorn 프로세스 모니터링:
#    - ps aux | grep gunicorn  (Linux)
#    - Get-Process python      (PowerShell)
#
# 3. Redis 모니터링:
#    - redis-cli INFO stats
#    - redis-cli INFO keyspace
#
# 4. 애플리케이션 로그 확인:
#    - 락 경합: "잠금 획득 실패"
#    - 캐시 성능: "캐시 히트/미스"
#    - Redis 장애: "ServiceUnavailable"
#
# 5. 성능 지표:
#    - 응답 시간: /api/health (< 100ms)
#    - 캐시 히트율: > 80%
#    - 락 경합률: < 5%
#
# ====================================================================
