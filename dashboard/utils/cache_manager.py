"""
메모리 기반 캐싱 시스템
Thread-Safe하고 TTL(Time To Live) 지원
"""

import time
import threading
from typing import Any, Optional, Dict
from datetime import datetime, timedelta
import logging

logger = logging.getLogger(__name__)

class MemoryCache:
    """Thread-Safe 메모리 캐시 클래스"""
    
    def __init__(self, default_ttl=300):  # 기본 5분
        self._cache: Dict[str, Dict] = {}
        self._lock = threading.RLock()
        self._default_ttl = default_ttl
        self._cleanup_interval = 60  # 1분마다 정리
        self._last_cleanup = time.time()
    
    def get(self, key: str) -> Optional[Any]:
        """캐시에서 값 가져오기"""
        with self._lock:
            self._cleanup_if_needed()
            
            if key not in self._cache:
                return None
            
            item = self._cache[key]
            
            # TTL 체크
            if time.time() > item['expires']:
                del self._cache[key]
                logger.debug(f"캐시 만료로 삭제: {key}")
                return None
            
            # 접근 시간 업데이트
            item['last_accessed'] = time.time()
            logger.debug(f"캐시 히트: {key}")
            return item['value']
    
    def set(self, key: str, value: Any, ttl: Optional[int] = None) -> None:
        """캐시에 값 저장"""
        with self._lock:
            expires = time.time() + (ttl or self._default_ttl)
            
            self._cache[key] = {
                'value': value,
                'expires': expires,
                'created': time.time(),
                'last_accessed': time.time()
            }
            
            logger.debug(f"캐시 저장: {key}, TTL: {ttl or self._default_ttl}초")
    
    def delete(self, key: str) -> bool:
        """캐시에서 키 삭제"""
        with self._lock:
            if key in self._cache:
                del self._cache[key]
                logger.debug(f"캐시 삭제: {key}")
                return True
            return False
    
    def clear(self) -> None:
        """모든 캐시 삭제"""
        with self._lock:
            count = len(self._cache)
            self._cache.clear()
            logger.info(f"전체 캐시 삭제: {count}개 항목")
    
    def exists(self, key: str) -> bool:
        """키 존재 여부 확인"""
        with self._lock:
            if key not in self._cache:
                return False
            
            # TTL 체크
            if time.time() > self._cache[key]['expires']:
                del self._cache[key]
                return False
            
            return True
    
    def get_stats(self) -> Dict[str, Any]:
        """캐시 통계 정보"""
        with self._lock:
            self._cleanup_if_needed()
            
            total_items = len(self._cache)
            expired_items = 0
            current_time = time.time()
            
            for item in self._cache.values():
                if current_time > item['expires']:
                    expired_items += 1
            
            return {
                'total_items': total_items,
                'expired_items': expired_items,
                'active_items': total_items - expired_items,
                'memory_usage_estimate': total_items * 1024  # 대략적 추정
            }
    
    def _cleanup_if_needed(self) -> None:
        """필요시 만료된 항목 정리"""
        current_time = time.time()
        
        if current_time - self._last_cleanup < self._cleanup_interval:
            return
        
        expired_keys = []
        for key, item in self._cache.items():
            if current_time > item['expires']:
                expired_keys.append(key)
        
        for key in expired_keys:
            del self._cache[key]
        
        if expired_keys:
            logger.debug(f"만료된 캐시 정리: {len(expired_keys)}개 항목")
        
        self._last_cleanup = current_time

# 전역 캐시 인스턴스
_global_cache = MemoryCache()

def get_cache() -> MemoryCache:
    """전역 캐시 인스턴스 반환"""
    return _global_cache

# 편의 함수들
def cache_get(key: str) -> Optional[Any]:
    """캐시에서 값 가져오기"""
    return _global_cache.get(key)

def cache_set(key: str, value: Any, ttl: Optional[int] = None) -> None:
    """캐시에 값 저장"""
    _global_cache.set(key, value, ttl)

def cache_delete(key: str) -> bool:
    """캐시에서 키 삭제"""
    return _global_cache.delete(key)

def cache_clear() -> None:
    """모든 캐시 삭제"""
    _global_cache.clear()

def cache_exists(key: str) -> bool:
    """키 존재 여부 확인"""
    return _global_cache.exists(key)

def cache_stats() -> Dict[str, Any]:
    """캐시 통계 정보"""
    return _global_cache.get_stats()

# 캐시 데코레이터
def cached(ttl: int = 300, key_prefix: str = ""):
    """함수 결과를 캐싱하는 데코레이터"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            # 캐시 키 생성
            import hashlib
            key_parts = [key_prefix, func.__name__]
            if args:
                key_parts.append(str(args))
            if kwargs:
                key_parts.append(str(sorted(kwargs.items())))
            
            cache_key = hashlib.md5(":".join(key_parts).encode()).hexdigest()
            
            # 캐시에서 확인
            cached_result = cache_get(cache_key)
            if cached_result is not None:
                return cached_result
            
            # 함수 실행 및 결과 캐싱
            result = func(*args, **kwargs)
            cache_set(cache_key, result, ttl)
            return result
        
        return wrapper
    return decorator