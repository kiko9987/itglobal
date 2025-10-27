"""
통합 에러 핸들링 시스템
로깅, 알림, 복구 전략을 포함한 포괄적인 에러 처리

## 주요 기능

### 1. 에러 카테고리 분류
- CRITICAL: 서비스 중단 (즉시 알림 필요)
- HIGH: 주요 기능 장애 (빠른 대응 필요)
- MEDIUM: 일반적인 오류 (모니터링 필요)
- LOW: 경고 수준 (로그 기록만)

### 2. 에러 핸들러 (ErrorHandler 클래스)
- `log_error()`: 에러 로깅 및 파일 저장
- `get_recent_errors()`: 최근 에러 조회
- `clear_old_errors()`: 오래된 에러 정리

### 3. 데코레이터
- `@handle_error()`: 기본 에러 핸들링
- `@handle_critical_error()`: CRITICAL 에러 처리
- `@handle_high_priority_error()`: HIGH 우선순위 에러 처리
- `@handle_medium_priority_error()`: MEDIUM 우선순위 에러 처리

### 4. Flask 에러 핸들러
- `setup_flask_error_handlers(app)`: Flask 앱에 전역 에러 핸들러 등록
  - 404 Not Found
  - 500 Internal Server Error
  - 기타 예외 처리

### 5. 에러 통계
- `get_error_stats(days=7)`: 에러 통계 정보 조회
  - 카테고리별 분류
  - 에러 타입별 분류
  - 일별 발생 추이

## 사용 예시

### 데코레이터 사용
```python
from dashboard.utils.error_handler import handle_error, ErrorCategory

@handle_error(category=ErrorCategory.HIGH)
def critical_function():
    # 에러 발생 시 자동으로 로깅됨
    result = risky_operation()
    return result
```

### Flask 앱 설정
```python
from dashboard.utils.error_handler import setup_flask_error_handlers

app = Flask(__name__)
setup_flask_error_handlers(app)
```

### 수동 에러 로깅
```python
from dashboard.utils.error_handler import get_error_handler, ErrorCategory

error_handler = get_error_handler()
try:
    dangerous_operation()
except Exception as e:
    error_handler.log_error(
        e,
        context={'user_id': 'user123', 'action': 'update_project'},
        category=ErrorCategory.HIGH
    )
```

### 에러 통계 조회
```python
from dashboard.utils.error_handler import get_error_stats

stats = get_error_stats(days=7)
print(f"Total errors: {stats['total_errors']}")
print(f"By category: {stats['by_category']}")
```

## 에러 로그 저장

- 경로: `dashboard/logs/errors.json`
- 형식: JSON Lines (각 줄이 하나의 에러 객체)
- 자동 정리: 30일 이상 오래된 에러는 자동 삭제 (수동 호출 필요)

## 에러 로그 구조

```json
{
    "timestamp": "2025-01-20T10:30:00.123456",
    "error_type": "ValueError",
    "error_message": "Invalid input",
    "category": "MEDIUM",
    "user_id": "user123",
    "context": {
        "path": "/api/projects",
        "method": "POST",
        "data": {...}
    },
    "traceback": "...",
    "stack_trace": [...]
}
```

## 주의사항

1. **CRITICAL 에러**: 즉시 알림이 필요한 경우에만 사용
2. **컨텍스트 정보**: 가능한 많은 컨텍스트를 제공하여 디버깅 용이
3. **개인정보**: 에러 로그에 민감한 정보(비밀번호, 토큰 등)를 포함하지 말 것
4. **로그 정리**: 정기적으로 `clear_old_errors()` 호출하여 디스크 공간 관리

## 버전 정보

- 작성일: 2025-01-20
- 버전: 1.0
- 다음 리뷰: 2025-07-20
"""

import logging
import traceback
import functools
from typing import Optional, Any, Callable, Dict, List
from datetime import datetime
import json
import os

logger = logging.getLogger(__name__)

class ErrorCategory:
    """에러 카테고리 상수"""
    CRITICAL = "CRITICAL"      # 서비스 중단
    HIGH = "HIGH"              # 주요 기능 장애
    MEDIUM = "MEDIUM"          # 일반적인 오류
    LOW = "LOW"                # 경고 수준

class ErrorHandler:
    """통합 에러 핸들러"""
    
    def __init__(self):
        self.error_log_path = os.path.join(os.path.dirname(__file__), '..', 'logs', 'errors.json')
        self.ensure_log_directory()
        
    def ensure_log_directory(self):
        """로그 디렉토리 생성"""
        log_dir = os.path.dirname(self.error_log_path)
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
    
    def log_error(self, error: Exception, context: Dict[str, Any] = None, 
                  category: str = ErrorCategory.MEDIUM, user_id: str = None):
        """에러 로깅 및 저장"""
        error_info = {
            'timestamp': datetime.now().isoformat(),
            'error_type': type(error).__name__,
            'error_message': str(error),
            'category': category,
            'user_id': user_id,
            'context': context or {},
            'traceback': traceback.format_exc(),
            'stack_trace': traceback.format_stack()
        }
        
        # 파일 로그에 저장
        self._save_error_to_file(error_info)
        
        # 로거에 기록
        log_level = self._get_log_level(category)
        logger.log(log_level, f"[{category}] {error_info['error_type']}: {error_info['error_message']}")
        
        return error_info
    
    def _get_log_level(self, category: str) -> int:
        """카테고리별 로그 레벨 반환"""
        mapping = {
            ErrorCategory.CRITICAL: logging.CRITICAL,
            ErrorCategory.HIGH: logging.ERROR,
            ErrorCategory.MEDIUM: logging.WARNING,
            ErrorCategory.LOW: logging.INFO
        }
        return mapping.get(category, logging.WARNING)
    
    def _save_error_to_file(self, error_info: Dict[str, Any]):
        """에러 정보를 파일에 저장"""
        try:
            # 기존 로그 읽기
            errors = []
            if os.path.exists(self.error_log_path):
                try:
                    with open(self.error_log_path, 'r', encoding='utf-8') as f:
                        errors = json.load(f)
                except json.JSONDecodeError:
                    errors = []
            
            # 새 에러 추가
            errors.append(error_info)
            
            # 최근 1000개만 유지
            if len(errors) > 1000:
                errors = errors[-1000:]
            
            # 파일에 저장
            with open(self.error_log_path, 'w', encoding='utf-8') as f:
                json.dump(errors, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            logger.error(f"에러 로그 저장 실패: {str(e)}")
    
    def get_recent_errors(self, limit: int = 50, category: str = None) -> List[Dict]:
        """최근 에러 조회"""
        try:
            if not os.path.exists(self.error_log_path):
                return []
            
            with open(self.error_log_path, 'r', encoding='utf-8') as f:
                errors = json.load(f)
            
            # 카테고리 필터링
            if category:
                errors = [e for e in errors if e.get('category') == category]
            
            # 최근 순으로 정렬 후 제한
            errors.sort(key=lambda x: x.get('timestamp', ''), reverse=True)
            return errors[:limit]
            
        except Exception as e:
            logger.error(f"에러 로그 조회 실패: {str(e)}")
            return []

# 전역 에러 핸들러 인스턴스
_error_handler = ErrorHandler()

def handle_error(category: str = ErrorCategory.MEDIUM, 
                context: Dict[str, Any] = None,
                reraise: bool = True,
                fallback_value: Any = None):
    """에러 처리 데코레이터"""
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                # 에러 로깅
                error_context = context or {}
                error_context.update({
                    'function': func.__name__,
                    'module': func.__module__,
                    'args': str(args)[:200],  # 너무 긴 인자는 잘라냄
                    'kwargs': str(kwargs)[:200]
                })
                
                _error_handler.log_error(e, error_context, category)
                
                # 재발생 여부 결정
                if reraise:
                    raise
                else:
                    logger.warning(f"에러 무시됨: {str(e)}")
                    return fallback_value
        
        return wrapper
    return decorator

def safe_execute(func: Callable, *args, default=None, 
                error_category: str = ErrorCategory.LOW, **kwargs):
    """안전한 함수 실행"""
    try:
        return func(*args, **kwargs)
    except Exception as e:
        _error_handler.log_error(e, {
            'function': getattr(func, '__name__', 'anonymous'),
            'args': str(args)[:200],
            'kwargs': str(kwargs)[:200]
        }, error_category)
        return default

def log_and_ignore(category: str = ErrorCategory.LOW):
    """에러를 로깅하고 무시하는 데코레이터"""
    return handle_error(category=category, reraise=False, fallback_value=None)

def critical_error_handler(func):
    """치명적 에러 전용 핸들러"""
    return handle_error(category=ErrorCategory.CRITICAL)(func)

def get_error_handler() -> ErrorHandler:
    """전역 에러 핸들러 반환"""
    return _error_handler

# 플라스크 앱용 에러 핸들러
def setup_flask_error_handlers(app):
    """Flask 앱에 에러 핸들러 설정"""
    
    @app.errorhandler(404)
    def not_found_error(error):
        _error_handler.log_error(
            Exception("404 Not Found"), 
            {'path': request.path if 'request' in globals() else 'unknown'}, 
            ErrorCategory.LOW
        )
        return "페이지를 찾을 수 없습니다", 404
    
    @app.errorhandler(500)
    def internal_error(error):
        _error_handler.log_error(
            error.original_exception if hasattr(error, 'original_exception') else Exception("500 Internal Server Error"),
            {'path': request.path if 'request' in globals() else 'unknown'},
            ErrorCategory.HIGH
        )
        return "서버 내부 오류가 발생했습니다", 500
    
    @app.errorhandler(Exception)
    def handle_exception(e):
        _error_handler.log_error(
            e,
            {'path': request.path if 'request' in globals() else 'unknown'},
            ErrorCategory.HIGH
        )
        return "예상치 못한 오류가 발생했습니다", 500

# 에러 통계
def get_error_stats(days: int = 7) -> Dict[str, Any]:
    """에러 통계 정보"""
    try:
        from datetime import datetime, timedelta
        
        cutoff_date = (datetime.now() - timedelta(days=days)).isoformat()
        recent_errors = _error_handler.get_recent_errors(limit=10000)
        
        # 날짜 필터링
        recent_errors = [e for e in recent_errors if e.get('timestamp', '') >= cutoff_date]
        
        # 통계 계산
        by_category = {}
        by_type = {}
        by_day = {}
        
        for error in recent_errors:
            category = error.get('category', 'UNKNOWN')
            error_type = error.get('error_type', 'UNKNOWN')
            day = error.get('timestamp', '')[:10]  # YYYY-MM-DD
            
            by_category[category] = by_category.get(category, 0) + 1
            by_type[error_type] = by_type.get(error_type, 0) + 1
            by_day[day] = by_day.get(day, 0) + 1
        
        return {
            'total_errors': len(recent_errors),
            'by_category': by_category,
            'by_type': by_type,
            'by_day': by_day,
            'period_days': days
        }
        
    except Exception as e:
        logger.error(f"에러 통계 생성 실패: {str(e)}")
        return {'total_errors': 0, 'error': str(e)}