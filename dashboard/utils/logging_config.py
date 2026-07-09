"""
통합 로깅 설정 시스템
중앙집중식 로깅 설정, 구조화된 로그, 회전 파일 시스템
"""

import logging
import logging.handlers
import json
import os
import sys
from datetime import datetime
from typing import Dict, Any, Optional
from pathlib import Path
import traceback
import threading
import atexit
import signal
from multiprocessing import current_process


class JSONFormatter(logging.Formatter):
    """JSON 형태로 로그를 포맷하는 커스텀 포매터"""

    def format(self, record: logging.LogRecord) -> str:
        """로그 레코드를 JSON 형태로 포맷"""
        # 기본 로그 정보
        log_entry = {
            'timestamp': datetime.fromtimestamp(record.created).isoformat(),
            'level': record.levelname,
            'logger': record.name,
            'message': record.getMessage(),
            'module': record.module,
            'function': record.funcName,
            'line': record.lineno,
            'thread': record.thread,
            'thread_name': record.threadName,
            'process': record.process
        }

        # 예외 정보가 있으면 추가
        if record.exc_info:
            log_entry['exception'] = {
                'type': record.exc_info[0].__name__ if record.exc_info[0] else None,
                'message': str(record.exc_info[1]) if record.exc_info[1] else None,
                'traceback': self.formatException(record.exc_info)
            }

        # 추가 필드들 (extra 파라미터로 전달된 것들)
        for key, value in record.__dict__.items():
            if key not in ['name', 'msg', 'args', 'levelname', 'levelno', 'pathname',
                          'filename', 'module', 'exc_info', 'exc_text', 'stack_info',
                          'lineno', 'funcName', 'created', 'msecs', 'relativeCreated',
                          'thread', 'threadName', 'processName', 'process', 'message']:
                log_entry[key] = value

        return json.dumps(log_entry, ensure_ascii=False, default=str)


class ContextualFormatter(logging.Formatter):
    """컨텍스트 정보를 포함한 인간 친화적 포매터"""

    def __init__(self):
        super().__init__()

    def format(self, record: logging.LogRecord) -> str:
        """컨텍스트 정보를 포함한 로그 포맷"""
        # 기본 포맷
        timestamp = datetime.fromtimestamp(record.created).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]

        # 레벨별 이모지
        level_emoji = {
            'DEBUG': '[DEBUG]',
            'INFO': '[INFO]',
            'WARNING': '[WARN]',
            'ERROR': '[ERROR]',
            'CRITICAL': '[CRIT]'
        }

        emoji = level_emoji.get(record.levelname, '[LOG]')

        # 기본 메시지 구성
        message = f"{timestamp} {emoji} [{record.levelname}] {record.name}:{record.lineno} - {record.getMessage()}"

        # 추가 컨텍스트 정보
        context_info = []

        # 사용자 정보
        if hasattr(record, 'user_id'):
            context_info.append(f"user_id={record.user_id}")
        if hasattr(record, 'user_email'):
            context_info.append(f"user_email={record.user_email}")

        # 요청 정보
        if hasattr(record, 'request_id'):
            context_info.append(f"request_id={record.request_id}")
        if hasattr(record, 'endpoint'):
            context_info.append(f"endpoint={record.endpoint}")
        if hasattr(record, 'method'):
            context_info.append(f"method={record.method}")

        # 성능 정보
        if hasattr(record, 'duration'):
            context_info.append(f"duration={record.duration}ms")
        if hasattr(record, 'memory_usage'):
            context_info.append(f"memory={record.memory_usage}MB")

        # 비즈니스 컨텍스트
        if hasattr(record, 'project_code'):
            context_info.append(f"project={record.project_code}")
        if hasattr(record, 'operation'):
            context_info.append(f"operation={record.operation}")

        if context_info:
            message += f" | {' | '.join(context_info)}"

        # 예외 정보
        if record.exc_info:
            message += f"\n{self.formatException(record.exc_info)}"

        return message


class LoggingConfig:
    """중앙집중식 로깅 설정 관리자"""

    def __init__(self, app_name: str = "itglobal_dashboard"):
        self.app_name = app_name
        self.log_dir = Path(__file__).parent.parent / 'logs'
        self.log_dir.mkdir(exist_ok=True)

        # 로그 레벨 설정
        self.default_level = os.getenv('LOG_LEVEL', 'INFO').upper()
        self.debug_mode = os.getenv('DEBUG', 'False').lower() == 'true'

        # 파일 회전 설정
        self.max_bytes = int(os.getenv('LOG_MAX_BYTES', 10 * 1024 * 1024))  # 10MB
        self.backup_count = int(os.getenv('LOG_BACKUP_COUNT', 5))

        self._setup_root_logger()

    def _setup_root_logger(self):
        """루트 로거 설정"""
        root_logger = logging.getLogger()
        root_logger.setLevel(getattr(logging, self.default_level))

        # 기존 핸들러 제거
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)

        # 콘솔 핸들러
        self._add_console_handler(root_logger)

        # 파일 핸들러들
        self._add_file_handlers(root_logger)

        # Flask 관련 로거 설정
        self._configure_third_party_loggers()

        # 슬랙 API 실패 감지 필터 (P1-6, 2026-07-09~)
        root_logger.addFilter(_SlackErrorSnooper())


class _SlackErrorSnooper(logging.Filter):
    """[SLACK/...] 프리픽스 + ERROR 레벨 로그를 감시.

    발생 시 dashboard.utils.slack_health.record_slack_api_failure 호출.
    항상 True 반환 (로그 흐름 변경 없음).
    """

    def filter(self, record: logging.LogRecord) -> bool:
        try:
            if record.levelno < logging.ERROR:
                return True
            msg = record.getMessage()
            if '[SLACK/' not in msg and 'slack' not in record.name.lower():
                return True
            # 지연 import (부팅 순서 이슈 방지)
            from dashboard.utils.slack_health import record_slack_api_failure
            # API 이름 추출: '[SLACK/AS] 요청 모달 열기 실패' → 'AS'
            api = 'unknown'
            if '[SLACK/' in msg:
                try:
                    api = msg.split('[SLACK/', 1)[1].split(']', 1)[0]
                except Exception:
                    pass
            exc = record.exc_info[1] if record.exc_info else Exception(msg[:200])
            record_slack_api_failure(f'slack.{api}', exc, context={'logger': record.name, 'msg': msg[:150]})
        except Exception:
            pass  # 감시 로직 자체가 로그 흐름을 막지 못하게
        return True

    def _add_console_handler(self, logger: logging.Logger):
        """콘솔 핸들러 추가 (UTF-8 인코딩)"""
        # UTF-8 인코딩된 stdout 스트림 생성 (Windows CP949 문제 해결)
        import io
        utf8_stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        console_handler = logging.StreamHandler(utf8_stdout)

        if self.debug_mode:
            console_handler.setLevel(logging.DEBUG)
            console_handler.setFormatter(ContextualFormatter())
        else:
            console_handler.setLevel(logging.INFO)
            # 프로덕션에서는 간단한 포맷
            simple_formatter = logging.Formatter(
                '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            console_handler.setFormatter(simple_formatter)

        logger.addHandler(console_handler)

    def _add_file_handlers(self, logger: logging.Logger):
        """파일 핸들러들 추가"""

        # 1. 일반 로그 파일 (INFO 이상)
        info_file = self.log_dir / f'{self.app_name}.log'
        info_handler = logging.handlers.RotatingFileHandler(
            info_file, maxBytes=self.max_bytes, backupCount=self.backup_count,
            encoding='utf-8'
        )
        info_handler.setLevel(logging.INFO)
        info_handler.setFormatter(ContextualFormatter())
        logger.addHandler(info_handler)

        # 2. 에러 로그 파일 (ERROR 이상)
        error_file = self.log_dir / f'{self.app_name}_error.log'
        error_handler = logging.handlers.RotatingFileHandler(
            error_file, maxBytes=self.max_bytes, backupCount=self.backup_count,
            encoding='utf-8'
        )
        error_handler.setLevel(logging.ERROR)
        error_handler.setFormatter(ContextualFormatter())
        logger.addHandler(error_handler)

        # 3. JSON 로그 파일 (구조화된 로그)
        json_file = self.log_dir / f'{self.app_name}_json.log'
        json_handler = logging.handlers.RotatingFileHandler(
            json_file, maxBytes=self.max_bytes, backupCount=self.backup_count,
            encoding='utf-8'
        )
        json_handler.setLevel(logging.INFO)
        json_handler.setFormatter(JSONFormatter())
        logger.addHandler(json_handler)

        # 4. 디버그 로그 파일 (DEBUG 모드일 때만)
        if self.debug_mode:
            debug_file = self.log_dir / f'{self.app_name}_debug.log'
            debug_handler = logging.handlers.RotatingFileHandler(
                debug_file, maxBytes=self.max_bytes, backupCount=self.backup_count,
                encoding='utf-8'
            )
            debug_handler.setLevel(logging.DEBUG)
            debug_handler.setFormatter(ContextualFormatter())
            logger.addHandler(debug_handler)

    def _configure_third_party_loggers(self):
        """서드파티 라이브러리 로거 설정"""
        # 프로덕션 환경 감지
        is_production = os.getenv('FLASK_ENV') == 'production' or not self.debug_mode

        # Flask 관련
        logging.getLogger('werkzeug').setLevel(logging.ERROR if is_production else logging.WARNING)
        logging.getLogger('urllib3').setLevel(logging.ERROR if is_production else logging.WARNING)

        # Google API 관련 (프로덕션에서는 CRITICAL만)
        logging.getLogger('googleapiclient.discovery').setLevel(logging.CRITICAL if is_production else logging.ERROR)
        logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.CRITICAL if is_production else logging.ERROR)
        logging.getLogger('google.auth').setLevel(logging.ERROR if is_production else logging.WARNING)

        # 기타
        logging.getLogger('requests').setLevel(logging.ERROR if is_production else logging.WARNING)
        logging.getLogger('schedule').setLevel(logging.ERROR if is_production else logging.WARNING)

    def get_logger(self, name: str) -> logging.Logger:
        """모듈별 로거 반환"""
        return logging.getLogger(name)

    def log_with_context(self, logger: logging.Logger, level: str, message: str, **context):
        """컨텍스트 정보와 함께 로깅"""
        log_method = getattr(logger, level.lower())
        log_method(message, extra=context)


class RequestContextLogger:
    """Flask 요청 컨텍스트를 포함한 로거"""

    def __init__(self, logger: logging.Logger):
        self.logger = logger

    def _get_request_context(self) -> Dict[str, Any]:
        """현재 Flask 요청의 컨텍스트 정보 추출"""
        context = {}

        try:
            from flask import request, session, g

            if request:
                context.update({
                    'method': request.method,
                    'endpoint': request.endpoint,
                    'url': request.url,
                    'remote_addr': request.remote_addr,
                    'user_agent': request.headers.get('User-Agent', '')[:100]
                })

                # 요청 ID (있다면)
                if hasattr(g, 'request_id'):
                    context['request_id'] = g.request_id

                # 사용자 정보 (세션에서)
                if session and 'user' in session:
                    user_info = session['user']
                    context.update({
                        'user_id': user_info.get('id'),
                        'user_email': user_info.get('email')
                    })
        except:
            # Flask 컨텍스트 외부에서 호출된 경우
            pass

        return context

    def debug(self, message: str, **extra):
        context = self._get_request_context()
        # exc_info, stack_info, stacklevel은 별도 파라미터로 전달
        exc_info = extra.pop('exc_info', None)
        stack_info = extra.pop('stack_info', None)
        stacklevel = extra.pop('stacklevel', 1)
        context.update(extra)
        self.logger.debug(message, exc_info=exc_info, stack_info=stack_info,
                         stacklevel=stacklevel, extra=context)

    def info(self, message: str, **extra):
        context = self._get_request_context()
        exc_info = extra.pop('exc_info', None)
        stack_info = extra.pop('stack_info', None)
        stacklevel = extra.pop('stacklevel', 1)
        context.update(extra)
        self.logger.info(message, exc_info=exc_info, stack_info=stack_info,
                        stacklevel=stacklevel, extra=context)

    def warning(self, message: str, **extra):
        context = self._get_request_context()
        exc_info = extra.pop('exc_info', None)
        stack_info = extra.pop('stack_info', None)
        stacklevel = extra.pop('stacklevel', 1)
        context.update(extra)
        self.logger.warning(message, exc_info=exc_info, stack_info=stack_info,
                           stacklevel=stacklevel, extra=context)

    def error(self, message: str, **extra):
        context = self._get_request_context()
        exc_info = extra.pop('exc_info', None)
        stack_info = extra.pop('stack_info', None)
        stacklevel = extra.pop('stacklevel', 1)
        context.update(extra)
        self.logger.error(message, exc_info=exc_info, stack_info=stack_info,
                         stacklevel=stacklevel, extra=context)

    def critical(self, message: str, **extra):
        context = self._get_request_context()
        exc_info = extra.pop('exc_info', None)
        stack_info = extra.pop('stack_info', None)
        stacklevel = extra.pop('stacklevel', 1)
        context.update(extra)
        self.logger.critical(message, exc_info=exc_info, stack_info=stack_info,
                            stacklevel=stacklevel, extra=context)


# 전역 로깅 설정 인스턴스 및 충돌 방지
_logging_config = None
_logging_lock = threading.RLock()  # 재진입 가능한 락
_initialized_processes = set()     # 초기화된 프로세스 ID 추적
_cleanup_registered = False        # atexit 핸들러 등록 여부


def setup_logging(app_name: str = "itglobal_dashboard") -> LoggingConfig:
    """
    로깅 시스템 초기화 (중복 초기화 방지, LOG_LEVEL 보존)

    멀티프로세스 환경에서 로그 설정이 중복 초기화되는 것을 방지하고,
    기존에 설정된 LOG_LEVEL을 보존합니다.
    """
    global _logging_config, _logging_lock, _initialized_processes

    current_pid = current_process().pid

    with _logging_lock:
        # 이미 현재 프로세스에서 초기화되었다면 기존 설정 반환
        if _logging_config is not None and current_pid in _initialized_processes:
            return _logging_config

        # 새로 초기화하는 경우에만 LOG_LEVEL 환경변수 확인
        current_log_level = os.getenv('LOG_LEVEL', 'INFO').upper()

        # 기존 로거가 있다면 현재 레벨과 환경변수 레벨 비교
        if _logging_config is not None:
            existing_level = logging.getLogger().level
            existing_level_name = logging.getLevelName(existing_level)

            # 환경변수가 변경되지 않았다면 기존 설정 유지
            if existing_level_name == current_log_level:
                _initialized_processes.add(current_pid)
                return _logging_config

        # 새로운 설정 또는 LOG_LEVEL이 변경된 경우에만 재초기화
        _logging_config = LoggingConfig(app_name)
        _initialized_processes.add(current_pid)

        return _logging_config


def get_logger(name: str) -> RequestContextLogger:
    """컨텍스트 정보를 포함한 로거 반환"""
    if _logging_config is None:
        setup_logging()

    base_logger = _logging_config.get_logger(name)
    return RequestContextLogger(base_logger)


def log_performance(func):
    """함수 실행 시간을 로깅하는 데코레이터"""
    def wrapper(*args, **kwargs):
        logger = get_logger(func.__module__)
        start_time = datetime.now()

        try:
            result = func(*args, **kwargs)
            duration = (datetime.now() - start_time).total_seconds() * 1000

            logger.info(
                f"Function {func.__name__} completed",
                operation=func.__name__,
                duration=f"{duration:.2f}",
                status="success"
            )

            return result
        except Exception as e:
            duration = (datetime.now() - start_time).total_seconds() * 1000

            logger.error(
                f"Function {func.__name__} failed",
                operation=func.__name__,
                duration=f"{duration:.2f}",
                status="error",
                error_type=type(e).__name__,
                error_message=str(e)
            )
            raise

    return wrapper


def log_business_event(event_type: str, **context):
    """비즈니스 이벤트 로깅"""
    logger = get_logger('business_events')

    logger.info(
        f"Business event: {event_type}",
        event_type=event_type,
        **context
    )


# 편의 함수들
def debug(message: str, **context):
    """디버그 로그"""
    logger = get_logger('app')
    logger.debug(message, **context)


def info(message: str, **context):
    """정보 로그"""
    logger = get_logger('app')
    logger.info(message, **context)


def warning(message: str, **context):
    """경고 로그"""
    logger = get_logger('app')
    logger.warning(message, **context)


def error(message: str, **context):
    """에러 로그"""
    logger = get_logger('app')
    logger.error(message, **context)


def critical(message: str, **context):
    """치명적 오류 로그"""
    logger = get_logger('app')
    logger.critical(message, **context)