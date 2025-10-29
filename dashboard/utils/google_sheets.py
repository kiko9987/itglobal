import os
import logging
import pandas as pd
import time
import ssl
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential_jitter,
    retry_if_exception_type,
    before_sleep_log,
    after_log
)
from .logging_config import get_logger
from .api_usage_monitor import get_api_monitor

# SSL 에러 처리를 위한 import
try:
    import requests.exceptions
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

# 전역 로깅 설정을 사용하는 로거
logger = get_logger(__name__)

# Tenacity 재시도 조건 함수
def _should_retry_http_error(exception):
    """
    HttpError가 재시도 가능한 에러인지 판단

    Args:
        exception: 발생한 예외

    Returns:
        bool: 재시도 가능 여부
    """
    if not isinstance(exception, HttpError):
        return False

    # Rate Limit 및 서버 오류만 재시도
    RATE_LIMIT_ERROR_CODES = [429, 500, 503]
    error_code = exception.resp.status
    return error_code in RATE_LIMIT_ERROR_CODES


class GoogleSheetsManager:
    """구글 시트 연동 관리 클래스 (Thread-Safe 싱글톤)"""

    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.readonly'
    ]
    _instance = None
    _lock = None

    # 프로젝트 코드 → 행번호 캐시 (메모리)
    _row_cache = {}
    _row_cache_timestamp = None

    # Rate Limiting 및 Retry 설정
    MAX_RETRIES = 5  # 최대 재시도 횟수
    INITIAL_BACKOFF = 1  # 초기 대기 시간 (초)
    MAX_BACKOFF = 32  # 최대 대기 시간 (초)
    RATE_LIMIT_ERROR_CODES = [429, 500, 503]  # 재시도할 HTTP 에러 코드

    # Backoff Jitter 설정
    JITTER_MIN = 0  # Exponential Backoff 지터 최소값 (초)
    JITTER_MAX = 1  # Exponential Backoff 지터 최대값 (초)

    # 데이터 유효성 검사
    MIN_REQUIRED_ROWS = 2  # 최소 필요 행 수 (헤더 + 데이터 1행)

    # 디버그 로깅 설정
    DEBUG_SAMPLE_PROJECTS = 10  # 디버그 로깅할 최근 프로젝트 수
    DEBUG_SAMPLE_NOTES = 3  # 디버그 로깅할 노트 샘플 행 수

    # UI 표시 설정
    NOTE_PREVIEW_LENGTH = 30  # 노트 미리보기 문자 수

    # Google Sheets 컬럼 설정
    MAX_COLUMN_INDEX = 50  # 배경색 업데이트 최대 컬럼 수 (A~AX)
    
    def __new__(cls, credentials_file='credentials.json'):
        """Thread-Safe 싱글톤 구현"""
        if cls._lock is None:
            import threading
            cls._lock = threading.RLock()
        
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(GoogleSheetsManager, cls).__new__(cls)
                cls._instance._initialized = False
            return cls._instance
    
    def __init__(self, credentials_file='credentials.json'):
        """
        구글 시트 매니저 초기화 (한 번만 실행)
        
        Args:
            credentials_file: 구글 API 서비스 계정 자격증명 파일 경로
        """
        if self._initialized:
            return
            
        env_credentials = os.getenv('GOOGLE_APPLICATION_CREDENTIALS') or os.getenv('GOOGLE_CREDENTIALS_FILE')
        if env_credentials:
            credentials_file = env_credentials

        self.credentials_file = os.path.abspath(credentials_file)
        logger.debug(f"Google credentials file resolved to: {self.credentials_file}")
        self.service = None
        self._authenticate()
        self._initialized = True
    
    def _authenticate(self):
        """구글 API 인증 처리 (서비스 계정 방식)"""
        if not os.path.exists(self.credentials_file):
            raise FileNotFoundError(
                f"구글 API 자격증명 파일이 없습니다: {self.credentials_file}\n"
                "Google Cloud Console에서 서비스 계정 JSON 키를 다운로드하여 credentials.json으로 저장하거나,\n"
                "환경 변수 GOOGLE_APPLICATION_CREDENTIALS / GOOGLE_CREDENTIALS_FILE 값을 올바른 경로로 설정하세요."
            )

        try:
            # 서비스 계정 자격증명 로드
            creds = ServiceAccountCredentials.from_service_account_file(
                self.credentials_file, scopes=self.SCOPES
            )

            # 서비스 객체 생성
            self.service = build('sheets', 'v4', credentials=creds)
            logger.info("구글 시트 API 인증 완료 (서비스 계정)")

        except Exception as e:
            logger.error(f"구글 API 인증 실패: {str(e)}")
            raise

    def reset_service(self):
        """
        서비스 객체 재초기화 (타임아웃/에러 후 복구용)

        Thread-Safe: _lock을 사용하여 동시 재초기화 방지
        """
        with self._lock:
            try:
                logger.warning("[SERVICE_RESET] Google Sheets 서비스 객체 재초기화 시작")

                # 기존 서비스 객체 정리
                if self.service:
                    try:
                        # HTTP 연결 정리 시도
                        if hasattr(self.service, '_http'):
                            self.service._http.close()
                    except Exception as cleanup_err:
                        logger.debug(f"[SERVICE_RESET] 기존 연결 정리 중 경고: {cleanup_err}")

                # 새로운 서비스 객체 생성
                self._authenticate()

                logger.info("[SERVICE_RESET] Google Sheets 서비스 객체 재초기화 완료")
                return True

            except Exception as e:
                logger.error(f"[SERVICE_RESET] 서비스 재초기화 실패: {e}")
                return False

    def _execute_with_retry(self, request, operation_name="API 호출"):
        """
        Exponential Backoff를 사용한 Google Sheets API 호출 재시도 로직

        Thread-Safe: _lock을 사용하여 동시 API 호출 방지 (OpenSSL race condition 해결)

        Args:
            request: Google API 요청 객체
            operation_name: 로깅용 작업 이름

        Returns:
            API 호출 결과

        Raises:
            Exception: 최대 재시도 후에도 실패 시 (사용자 친화적 메시지 포함)
        """
        for attempt in range(self.MAX_RETRIES):
            try:
                # 🔒 스레드 안전성 확보: Google API service 객체 동시 접근 방지
                # OpenSSL TLS 세션 race condition 방지
                with self._lock:
                    result = request.execute()

                if attempt > 0:
                    logger.info(
                        f"[재시도 성공] {operation_name} - "
                        f"{attempt + 1}번째 시도에서 성공 (소요 시간: 약 {sum([self.INITIAL_BACKOFF * (2 ** i) for i in range(attempt)])}초)"
                    )
                return result

            except ssl.SSLError as ssl_err:
                # SSL 에러는 즉시 로그 남기고 프로세스 강제 종료
                logger.critical(
                    f"[SSL_ERROR] {operation_name} - SSL/TLS 에러 발생: {ssl_err}\n"
                    f"프로세스를 즉시 종료하여 추가 손상을 방지합니다. 프로세스 매니저가 자동으로 재시작합니다."
                )
                # os._exit(1): 모든 스레드를 즉시 강제 종료 (sys.exit과 달리 프로세스 전체 종료)
                os._exit(1)

            except HttpError as e:
                error_code = e.resp.status
                error_reason = e._get_reason()

                # 사용자 친화적 에러 메시지 생성
                user_friendly_message = self._get_user_friendly_error_message(error_code, error_reason)

                # Rate Limit 또는 서버 에러인 경우 재시도
                if error_code in self.RATE_LIMIT_ERROR_CODES:
                    if attempt < self.MAX_RETRIES - 1:
                        # Exponential Backoff 계산 (지터 추가)
                        import random
                        backoff_time = min(
                            self.INITIAL_BACKOFF * (2 ** attempt) + random.uniform(self.JITTER_MIN, self.JITTER_MAX),
                            self.MAX_BACKOFF
                        )

                        logger.warning(
                            f"[Rate Limit/Server Error] {operation_name} - "
                            f"HTTP {error_code}: {error_reason} | "
                            f"시도 {attempt + 1}/{self.MAX_RETRIES} | "
                            f"{backoff_time:.2f}초 후 재시도 | "
                            f"사용자 메시지: {user_friendly_message}"
                        )

                        time.sleep(backoff_time)
                        continue
                    else:
                        # 최대 재시도 횟수 초과
                        error_msg = (
                            f"Google Sheets API 오류: {user_friendly_message}\n"
                            f"최대 재시도 횟수({self.MAX_RETRIES}회)를 초과했습니다. "
                            f"잠시 후 다시 시도해주세요."
                        )
                        logger.error(
                            f"[재시도 실패] {operation_name} - "
                            f"최대 재시도 횟수 {self.MAX_RETRIES}회 초과 | "
                            f"HTTP {error_code}: {error_reason} | "
                            f"총 소요 시간: 약 {sum([self.INITIAL_BACKOFF * (2 ** i) for i in range(self.MAX_RETRIES)])}초"
                        )
                        raise Exception(error_msg) from e

                # Rate Limit이 아닌 다른 에러는 즉시 raise
                else:
                    error_msg = f"Google Sheets API 오류: {user_friendly_message}"
                    logger.error(
                        f"[API 에러] {operation_name} - "
                        f"HTTP {error_code}: {error_reason} | "
                        f"사용자 메시지: {user_friendly_message}"
                    )
                    raise Exception(error_msg) from e

            except Exception as e:
                # requests 라이브러리의 SSL 에러 체크 (HttpError보다 나중에 체크)
                if HAS_REQUESTS and isinstance(e, requests.exceptions.SSLError):
                    logger.critical(
                        f"[REQUESTS_SSL_ERROR] {operation_name} - Requests SSL 에러 발생: {e}\n"
                        f"프로세스를 즉시 종료하여 추가 손상을 방지합니다. 프로세스 매니저가 자동으로 재시작합니다."
                    )
                    # os._exit(1): 모든 스레드를 즉시 강제 종료
                    os._exit(1)

                # HttpError가 아닌 다른 예외 (네트워크 에러 등)
                if attempt < self.MAX_RETRIES - 1:
                    import random
                    backoff_time = min(
                        self.INITIAL_BACKOFF * (2 ** attempt) + random.uniform(self.JITTER_MIN, self.JITTER_MAX),
                        self.MAX_BACKOFF
                    )

                    logger.warning(
                        f"[일반 에러] {operation_name} - "
                        f"{type(e).__name__}: {str(e)} | "
                        f"시도 {attempt + 1}/{self.MAX_RETRIES} | "
                        f"{backoff_time:.2f}초 후 재시도"
                    )

                    time.sleep(backoff_time)
                    continue
                else:
                    # 사용자 친화적 에러 메시지
                    error_msg = (
                        f"Google Sheets 연결 오류: {type(e).__name__} - {str(e)}\n"
                        f"네트워크 연결을 확인하거나 잠시 후 다시 시도해주세요."
                    )
                    logger.error(
                        f"[재시도 실패] {operation_name} - "
                        f"최대 재시도 횟수 {self.MAX_RETRIES}회 초과 | "
                        f"{type(e).__name__}: {str(e)}"
                    )
                    raise Exception(error_msg) from e

        # 이 라인에 도달하면 안 되지만, 안전을 위해
        raise Exception(f"{operation_name} 실패: 알 수 없는 에러")

    def _get_user_friendly_error_message(self, error_code, error_reason):
        """
        HTTP 에러 코드를 사용자 친화적 메시지로 변환

        Args:
            error_code: HTTP 에러 코드
            error_reason: 원본 에러 메시지

        Returns:
            str: 사용자 친화적 에러 메시지
        """
        error_messages = {
            400: "잘못된 요청입니다. 데이터 형식을 확인해주세요.",
            401: "인증 실패. 서비스 계정 권한을 확인해주세요.",
            403: "접근 권한이 없습니다. Google Sheets 공유 설정을 확인해주세요.",
            404: "시트를 찾을 수 없습니다. 시트 ID와 범위를 확인해주세요.",
            429: "API 사용량 한도를 초과했습니다. 잠시 후 다시 시도됩니다.",
            500: "Google 서버 오류입니다. 자동으로 재시도합니다.",
            503: "Google Sheets 서비스가 일시적으로 사용 불가능합니다. 자동으로 재시도합니다.",
        }

        user_message = error_messages.get(
            error_code,
            f"예상치 못한 오류가 발생했습니다 (HTTP {error_code})"
        )

        return f"{user_message} (상세: {error_reason})"
    
    def get_sheet_data(self, sheet_id, range_name='공사 현황의 사본!A:AN'):
        """
        구글 시트에서 데이터 가져오기 (에러 처리 강화)

        Args:
            sheet_id: 구글 시트 ID
            range_name: 데이터 범위

        Returns:
            pandas.DataFrame: 시트 데이터
        """
        # API 모니터링 시작
        monitor = get_api_monitor()
        request_id = monitor.start_request('google_sheets')

        try:
            if not sheet_id:
                raise ValueError("시트 ID가 제공되지 않았습니다.")

            # 시트 데이터 가져오기 (함수 계산 결과 포함) - Rate Limiting 처리 포함
            request = self.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name,
                valueRenderOption='FORMATTED_VALUE',  # 함수 계산 결과를 포맷된 값으로
                dateTimeRenderOption='FORMATTED_STRING'  # 날짜 포맷된 문자열로
            )
            result = self._execute_with_retry(request, f"get_sheet_data({range_name})")
            
            values = result.get('values', [])
            
            if not values:
                logger.warning("시트에 데이터가 없습니다.")
                return pd.DataFrame()

            if len(values) < self.MIN_REQUIRED_ROWS:
                logger.warning("시트에 헤더만 있고 데이터가 없습니다.")
                return pd.DataFrame(columns=values[0] if values else [])

            # DataFrame 생성 (불완전한 행 처리)
            header = values[0]
            data_rows = []
            for row in values[1:]:
                # 부족한 컬럼은 빈 문자열로 채움
                padded_row = row + [''] * (len(header) - len(row))
                # 초과 컬럼은 잘라냄
                padded_row = padded_row[:len(header)]
                data_rows.append(padded_row)

            df = pd.DataFrame(data_rows, columns=header)
            
            # 공사 종료 날짜 원시 데이터 재확인 (디버그 모드에서만)
            # RequestContextLogger와 호환성 문제로 가드 제거 (logger가 레벨 체크 수행)
            if '공사 종료' in df.columns:
                logger.debug("=== 최근 10개 프로젝트의 공사 종료 원시 데이터 재확인 ===")
                recent_projects = df.tail(self.DEBUG_SAMPLE_PROJECTS)
                for i, (idx, row) in enumerate(recent_projects.iterrows()):
                    project_code = row.get('프로젝트 코드', 'N/A')
                    start_date = row.get('공사 시작', 'N/A')
                    end_date = row.get('공사 종료', 'N/A')
                    logger.debug(f"  프로젝트 {project_code}: 시작='{start_date}' 종료='{end_date}' (종료타입:{type(end_date)})")
            
            # 데이터 전처리
            df = self._preprocess_data(df)
            
            # 운영 환경에서는 경고만 표시
            if len(df) == 0:
                logger.warning("구글 시트에서 데이터를 가져오지 못했습니다.")
            else:
                logger.debug(f"구글 시트에서 {len(df)}행의 데이터를 가져왔습니다.")

            # API 모니터링 성공 기록
            monitor.end_request(request_id, success=True)
            return df

        except Exception as e:
            error_msg = f"구글 시트 데이터 가져오기 오류: {str(e)}"
            logger.error(error_msg)

            # API 모니터링 실패 기록
            monitor.end_request(request_id, success=False, error_msg=str(e))
            
            # 구체적인 에러 유형별 처리
            if "404" in str(e):
                logger.error("시트를 찾을 수 없습니다. 시트 ID와 권한을 확인하세요.")
            elif "403" in str(e):
                logger.error("시트 접근 권한이 없습니다. 서비스 계정에 뷰어 권한을 부여하세요.")
            elif "401" in str(e):
                logger.error("인증 실패. 서비스 계정 키를 확인하세요.")
            
            raise
    
    def _preprocess_data(self, df):
        """데이터 전처리"""
        # 빈 행 제거
        df = df.dropna(how='all')
        
        # 프로젝트 코드가 있는 행만 필터링 (실제 데이터만)
        if '프로젝트 코드' in df.columns:
            original_count = len(df)
            df = df[df['프로젝트 코드'].notna() & (df['프로젝트 코드'].astype(str).str.strip() != '')]
            filtered_count = len(df)
            logger.debug(f"프로젝트 코드 필터링: {original_count}행 → {filtered_count}행")
        else:
            logger.warning("프로젝트 코드 컬럼을 찾을 수 없습니다.")
        
        # 날짜 컬럼 처리 (빈 값은 빈 문자열로 유지)
        date_columns = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
        for col in date_columns:
            if col in df.columns:
                # 빈 값이 아닌 경우에만 날짜 변환 시도
                df[col] = df[col].apply(lambda x: 
                    pd.to_datetime(x, errors='coerce') if x and str(x).strip() and str(x).strip().lower() != 'none' 
                    else pd.NaT
                )
        
        # 숫자 컬럼 처리 (함수 계산 결과 포함) - 성능 최적화
        numeric_columns = ['총액 1', '총액 2', '총액2', '계약금', '중도금', '잔금', 
                          '미수금', '미수금W', '제품대', '도급비', '자재비', '기타비', '순익', '마진율']
        
        # 존재하는 숫자 컬럼만 필터링하여 처리 속도 향상
        existing_numeric_columns = [col for col in numeric_columns if col in df.columns]
        
        if existing_numeric_columns:
            # 벡터화된 일괄 처리로 성능 대폭 향상
            for col in existing_numeric_columns:
                df[col] = (df[col].astype(str, copy=False)
                          .str.replace(r'[,￦₩\-]', '', regex=True)
                          .str.strip()
                          .replace(['', 'nan', 'None', 'NaN'], pd.NA))
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # 불린 컬럼 처리 (FutureWarning 해결: fillna 다운캐스팅 경고 방지)
        boolean_columns = ['부가세', '수금 확인']
        for col in boolean_columns:
            if col in df.columns:
                # map 후 boolean dtype으로 변환 (NaN은 자동으로 False로 변환됨)
                df[col] = df[col].map({'TRUE': True, 'FALSE': False}).astype('boolean').fillna(False)
        
        return df
    
    def get_sheet_metadata(self, sheet_id):
        """시트 메타데이터 가져오기 - Rate Limiting 처리 포함"""
        try:
            request = self.service.spreadsheets().get(spreadsheetId=sheet_id)
            result = self._execute_with_retry(request, f"get_sheet_metadata({sheet_id})")
            return {
                'title': result.get('properties', {}).get('title', ''),
                'sheets': [{'title': sheet.get('properties', {}).get('title', ''),
                           'sheetId': sheet.get('properties', {}).get('sheetId', 0)}
                          for sheet in result.get('sheets', [])],
                'last_updated': datetime.now().isoformat(),
                'raw_sheets': result.get('sheets', [])
            }
        except Exception as e:
            logger.error(f"시트 메타데이터 가져오기 오류: {str(e)}")
            return {}

    def get_sheet_id_by_name(self, spreadsheet_id, sheet_name):
        """시트 이름으로 시트 ID(숫자) 가져오기"""
        try:
            metadata = self.get_sheet_metadata(spreadsheet_id)
            for sheet in metadata.get('sheets', []):
                if sheet['title'] == sheet_name:
                    return sheet['sheetId']
            logger.warning(f"시트 '{sheet_name}'을 찾을 수 없습니다.")
            return 0  # 기본값
        except Exception as e:
            logger.error(f"시트 ID 조회 오류: {str(e)}")
            return 0

    def update_row_background_color(self, spreadsheet_id, sheet_name, row_number, color_type='normal'):
        """
        구글 시트 특정 행의 배경색 변경

        Args:
            spreadsheet_id: 스프레드시트 ID
            sheet_name: 시트 이름
            row_number: 행 번호 (1부터 시작)
            color_type: 'cancelled' (연한 빨간색), 'dark_grey' (진한 회색), 'normal' (흰색)
        """
        try:
            # 시트 ID 가져오기
            sheet_id = self.get_sheet_id_by_name(spreadsheet_id, sheet_name)

            # 색상 설정
            if color_type == 'cancelled':
                # 연한 빨간색 (#ffeaea)
                background_color = {
                    "red": 1.0,
                    "green": 0.918,
                    "blue": 0.918
                }
            elif color_type == 'dark_grey':
                # 진한 회색 (#808080)
                background_color = {
                    "red": 0.5,
                    "green": 0.5,
                    "blue": 0.5
                }
            else:
                # 기본 흰색 (#ffffff)
                background_color = {
                    "red": 1.0,
                    "green": 1.0,
                    "blue": 1.0
                }

            # batchUpdate 요청 생성
            request = {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_number - 1,  # 0-based index
                        "endRowIndex": row_number,
                        "startColumnIndex": 0,
                        "endColumnIndex": self.MAX_COLUMN_INDEX  # A부터 AX 컬럼까지 (충분히 넓게)
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": background_color
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            }

            batch_update_request = {
                "requests": [request]
            }

            # API 호출 - Rate Limiting 처리 포함
            request = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=batch_update_request
            )
            result = self._execute_with_retry(request, f"update_row_background_color(row={row_number})")

            logger.debug(f"[SUCCESS] 행 {row_number} 배경색 변경 완료: {color_type}")
            return True

        except Exception as e:
            logger.error(f"[ERROR] 행 배경색 변경 실패: {str(e)}")
            return False
    
    def validate_connection(self, sheet_id):
        """구글 시트 연결 테스트"""
        try:
            metadata = self.get_sheet_metadata(sheet_id)
            if metadata:
                logger.debug(f"구글 시트 연결 성공: {metadata['title']}")
                return True
            else:
                logger.error("구글 시트 연결 실패")
                return False
        except Exception as e:
            logger.error(f"구글 시트 연결 테스트 오류: {str(e)}")
            return False

    def find_next_empty_row(self, sheet_id, range_name='공사 현황의 사본!A:A'):
        """
        다음 빈 행 번호 찾기 (수식이 미리 설정된 행에 데이터 추가용)
        
        Returns:
            int: 다음 빈 행 번호 (1부터 시작)
        """
        try:
            request = self.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name,
                valueRenderOption='FORMATTED_VALUE'
            )
            result = self._execute_with_retry(request, f"find_next_empty_row({range_name})")
            
            values = result.get('values', [])
            
            # 프로젝트 코드(A열)가 비어있는 첫 번째 행 찾기
            for i, row in enumerate(values):
                if i == 0:  # 헤더 행 스킵
                    continue
                    
                if not row or len(row) == 0 or not row[0] or not row[0].strip():
                    return i + 1  # 행 번호는 1부터 시작
            
            # 빈 행을 찾지 못한 경우 마지막 다음 행 반환
            return len(values) + 1
            
        except Exception as e:
            logger.error(f"빈 행 찾기 오류: {str(e)}")
            return None

    def append_row(self, sheet_id, values, range_name='공사 현황의 사본!A:AN'):
        """
        구글 시트의 다음 빈 행에 데이터 추가 (수식이 미리 설정된 행에 덮어쓰기)
        
        Args:
            sheet_id: 구글 시트 ID
            values: 추가할 데이터 리스트
            range_name: 데이터 범위 (사용하지 않음, 호환성용)
            
        Returns:
            dict: 추가 결과
        """
        try:
            # 다음 빈 행 번호 찾기
            next_row = self.find_next_empty_row(sheet_id)
            if not next_row:
                raise Exception("다음 빈 행을 찾을 수 없습니다")
            
            # 특정 행에 데이터 업데이트 (수식이 있는 빈 행에 덮어쓰기)
            actual_range = f'공사 현황의 사본!A{next_row}:AM{next_row}'
            body = {
                'values': [values]
            }
            
            request = self.service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=actual_range,
                valueInputOption='USER_ENTERED',
                body=body
            )
            result = self._execute_with_retry(request, f"append_row(row={next_row})")

            logger.debug(f"빈 행({next_row})에 데이터 추가 성공: {result.get('updatedCells', 0)}셀 업데이트")
            return result
            
        except Exception as e:
            logger.error(f"빈 행 데이터 추가 오류: {str(e)}")
            raise
    
    def update_row(self, sheet_id, row_number, values, range_name='공사 현황의 사본!A{row}:AM{row}'):
        """
        구글 시트의 특정 행 업데이트
        
        Args:
            sheet_id: 구글 시트 ID
            row_number: 행 번호 (1부터 시작)
            values: 업데이트할 데이터 리스트
            range_name: 데이터 범위 템플릿
            
        Returns:
            dict: 업데이트 결과
        """
        try:
            # 범위 설정
            actual_range = range_name.format(row=row_number)
            
            body = {
                'values': [values]
            }
            
            request = self.service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=actual_range,
                valueInputOption='USER_ENTERED',
                body=body
            )
            result = self._execute_with_retry(request, f"update_row(row={row_number})")

            logger.debug(f"행 업데이트 성공: {row_number}행, {result.get('updatedCells', 0)}셀")
            return result
            
        except Exception as e:
            logger.error(f"행 업데이트 오류: {str(e)}")
            raise
    
    def batch_update_cells(self, sheet_id, updates):
        """
        구글 시트의 여러 셀을 일괄 업데이트

        Args:
            sheet_id: 구글 시트 ID
            updates: 업데이트할 데이터 리스트 [{'range': 'A1', 'values': [['data']]}]

        Returns:
            dict: Google Sheets API 응답 (totalUpdatedCells 등 포함)

        Raises:
            Exception: API 호출 실패 시 발생한 예외를 그대로 전달
        """
        try:
            data = []
            for update in updates:
                data.append({
                    'range': update['range'],
                    'values': update['values']
                })

            body = {
                'valueInputOption': 'USER_ENTERED',
                'data': data
            }

            request = self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=sheet_id,
                body=body
            )
            result = self._execute_with_retry(
                request,
                f"batch_update_cells({len(data)} ranges)"
            )

            updated_cells = result.get('totalUpdatedCells', 0)
            logger.debug(f"배치 업데이트 성공: {updated_cells}개 셀 업데이트")
            return result

        except Exception as e:
            logger.error(f"배치 업데이트 오류: {str(e)}")
            raise
    
    def find_row_by_project_code(self, sheet_id, project_code, range_name=None, use_cache=True):
        """
        프로젝트 코드로 행 번호 찾기 (메모리 캐시 사용)

        Args:
            sheet_id: 구글 시트 ID
            project_code: 찾을 프로젝트 코드
            range_name: 검색할 범위 (None이면 환경 변수 GOOGLE_SHEET_NAME 사용)
            use_cache: 캐시 사용 여부 (기본: True)

        Returns:
            int: 행 번호 (없으면 None)
        """
        # 기본 시트명 (환경 변수에서 읽기)
        default_sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황')

        # range_name이 None이면 기본값 설정
        if range_name is None:
            range_name = f"{default_sheet_name}!A:A"

        # range_name 형식 정규화 (방어 로직)
        if '!' not in range_name:
            # 시트명이 없으면 기본 시트명 추가
            range_name = f"{default_sheet_name}!{range_name}"
            logger.debug(f"[RANGE_NAME_NORMALIZE] 시트명 추가: {range_name}")

        # 캐시 확인 및 검증 (use_cache=True이고 캐시가 있으면)
        if use_cache and project_code in self._row_cache:
            cached_row = self._row_cache[project_code]
            logger.debug(f"[ROW_CACHE_HIT] {project_code} → 행 {cached_row} (검증 중...)")

            # 캐시 검증: 실제로 해당 행에 같은 프로젝트 코드가 있는지 확인
            try:
                sheet_name = range_name.split('!')[0] or default_sheet_name  # 빈 문자열 방어
                single_row_range = f"{sheet_name}!A{cached_row}"
                verify_result = self.service.spreadsheets().values().get(
                    spreadsheetId=sheet_id,
                    range=single_row_range
                ).execute()

                verify_values = verify_result.get('values', [])
                if verify_values and len(verify_values[0]) > 0 and verify_values[0][0] == project_code:
                    logger.debug(f"[ROW_CACHE_VALID] {project_code} → 행 {cached_row}")
                    return cached_row
                else:
                    # 캐시가 틀렸음 - 무효화하고 전체 검색으로 진행
                    actual_value = verify_values[0][0] if (verify_values and verify_values[0]) else '(빈 행)'
                    logger.warning(
                        f"[ROW_CACHE_INVALID] {project_code}: 캐시된 행 {cached_row}이 틀림 "
                        f"(실제 값: {actual_value}). 전체 검색으로 재시도"
                    )
                    del self._row_cache[project_code]
            except Exception as e:
                logger.warning(f"[ROW_CACHE_VERIFY_ERROR] {project_code}: {str(e)}. 전체 검색으로 재시도")
                if project_code in self._row_cache:
                    del self._row_cache[project_code]

        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [])

            for i, row in enumerate(values):
                if row and len(row) > 0 and row[0] == project_code:
                    row_number = i + 1  # 1부터 시작하는 행 번호

                    # 캐시 저장
                    if use_cache:
                        self._row_cache[project_code] = row_number
                        logger.debug(f"[ROW_CACHE_SET] {project_code} → 행 {row_number}")

                    return row_number

            return None

        except Exception as e:
            logger.error(f"행 찾기 오류: {str(e)}")
            return None
    
    def get_next_project_code(self, sheet_id, region_code='IT'):
        """
        다음 프로젝트 코드 생성
        
        Args:
            sheet_id: 구글 시트 ID
            region_code: 지역 코드 (예: IT, YG, JW 등)
            
        Returns:
            str: 새 프로젝트 코드
        """
        try:
            # 기존 데이터에서 해당 지역의 최대 번호 찾기
            df = self.get_sheet_data(sheet_id)
            
            if df.empty or '프로젝트 코드' not in df.columns:
                return f"G0001-{region_code}"
            
            # 해당 지역 코드가 포함된 프로젝트 코드 찾기
            region_projects = df[df['프로젝트 코드'].str.contains(f'-{region_code}', na=False)]
            
            if region_projects.empty:
                return f"G0001-{region_code}"
            
            # 번호 추출 및 최대값 찾기
            max_num = 0
            for code in region_projects['프로젝트 코드']:
                try:
                    # G0001-IT 형태에서 숫자 부분 추출
                    num_part = code.split('-')[0][1:]  # G 제거 후 숫자 부분
                    num = int(num_part)
                    max_num = max(max_num, num)
                except (ValueError, IndexError):
                    continue
            
            next_num = max_num + 1
            return f"G{next_num:04d}-{region_code}"
            
        except Exception as e:
            logger.error(f"프로젝트 코드 생성 오류: {str(e)}")
            return f"G0001-{region_code}"
    
    def get_column_mapping(self):
        """컬럼 매핑 정보 반환"""
        return {
            'A': '프로젝트 코드',
            'B': '사업자',
            'C': '담당자',
            'D': '거래처',
            'E': '현장 주소',
            'F': '공사 구분',
            'G': '기계 분류',
            'H': '브랜드',
            'I': '공사 시작',
            'J': '공사 종료',
            'K': '공사 내용',
            'L': '도급 구분',
            'M': '시공자',
            'N': '현장 담당자',
            'O': '담당자 연락처',
            'P': '담당자 이메일',
            'Q': '총액 1',
            'R': '부가세',
            'S': '총액 2',
            'T': '계약금',
            'U': '중도금',
            'V': '잔금',
            'W': '미수금',
            'X': '계산서',
            'Y': '수금 날짜',
            'Z': '수금 확인',
            'AA': '제품대',
            'AB': '도급비',
            'AC': '자재비',
            'AD': '기타비',
            'AE': '순익',
            'AF': '마진율',
            'AG': '수금 관련 특이사항',
            'AH': '계약금 입금자명',
            'AI': '중도금 입금자명',
            'AJ': '잔금 입금자명',
            'AK': '견적서 및 계약서 폴더 경로',
            'AL': '공사 확정',
            'AM': 'Airtable Record ID',
            'AN': '_version'
        }

    def get_field_column_mapping(self):
        """레거시 방식의 필드-컬럼 매핑 (인라인 편집용)"""
        return {
            # 기본정보
            '사업자': 'B',
            '현장 담당자': 'N',
            '도급 구분': 'L',
            '담당자 연락처': 'O',
            '시공자': 'M',
            '담당자 이메일': 'P',
            '현장 주소': 'E',
            # 공사정보
            '공사 구분': 'F',
            '기계 분류': 'G',
            '브랜드': 'H',
            '공사 시작': 'I',
            '공사 종료': 'J',
            '공사 내용': 'K',
            '공사 확정': 'AL',
            # 문서 정보
            '견적서 및 계약서 폴더 경로': 'AK',
            # 거래처 정보
            '거래처': 'D',
            '담당자': 'C'
        }

    def batch_update_fields(self, sheet_id, sheet_name, row_number, field_data, project_code=None):
        """
        레거시 방식의 개별 필드 배치 업데이트 (batchUpdate API 사용)

        Args:
            sheet_id: 구글 시트 ID
            sheet_name: 시트 이름
            row_number: 행 번호 (1부터 시작)
            field_data: 필드 데이터 딕셔너리 {'필드명': '값', ...}
            project_code: 프로젝트 코드 (로깅용)

        Returns:
            dict: {
                'success': bool,
                'updated_cells': int,
                'old_values': dict,
                'changes': list
            }
        """
        try:
            field_column_mapping = self.get_field_column_mapping()
            updates = []
            old_values = {}
            changes = []

            # 1. 업데이트할 셀들의 현재 값을 먼저 조회 (감사 로깅용)
            for field_name, new_value in field_data.items():
                if field_name in field_column_mapping:
                    column = field_column_mapping[field_name]

                    # 현재 값 조회
                    try:
                        current_value_result = self.service.spreadsheets().values().get(
                            spreadsheetId=sheet_id,
                            range=f'{sheet_name}!{column}{row_number}'
                        ).execute()
                        current_values = current_value_result.get('values', [['']])
                        old_value = current_values[0][0] if current_values and current_values[0] else ''
                        old_values[field_name] = old_value

                        # 변경사항 기록
                        if str(old_value) != str(new_value):
                            changes.append({
                                'field_name': field_name,
                                'old_value': old_value,
                                'new_value': new_value,
                                'column': column
                            })

                            # 업데이트할 셀 추가
                            updates.append({
                                'range': f'{sheet_name}!{column}{row_number}',
                                'values': [[new_value]]
                            })

                    except Exception as e:
                        logger.warning(f"현재 값 조회 실패 ({field_name}): {e}")
                        old_values[field_name] = ''

                        # 값 조회가 실패해도 업데이트는 진행
                        updates.append({
                            'range': f'{sheet_name}!{column}{row_number}',
                            'values': [[new_value]]
                        })
                        changes.append({
                            'field_name': field_name,
                            'old_value': '',
                            'new_value': new_value,
                            'column': column
                        })

            # 2. 업데이트할 항목이 없으면 성공으로 반환
            if not updates:
                logger.info(f"업데이트할 변경사항이 없습니다: {project_code}")
                return {
                    'success': True,
                    'updated_cells': 0,
                    'old_values': old_values,
                    'changes': []
                }

            # 3. batchUpdate API로 일괄 업데이트
            batch_update_body = {
                'valueInputOption': 'USER_ENTERED',
                'data': updates
            }

            result = self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=sheet_id,
                body=batch_update_body
            ).execute()

            updated_cells = result.get('totalUpdatedCells', 0)
            logger.info(f"[SUCCESS] 필드 배치 업데이트 완료: {project_code}, {updated_cells}개 셀 업데이트")

            return {
                'success': True,
                'updated_cells': updated_cells,
                'old_values': old_values,
                'changes': changes
            }

        except Exception as e:
            logger.error(f"[ERROR] 필드 배치 업데이트 실패: {project_code}, {str(e)}")
            return {
                'success': False,
                'updated_cells': 0,
                'old_values': {},
                'changes': [],
                'error': str(e)
            }

    def get_single_cell_value(self, sheet_id, sheet_name, column, row_number):
        """
        단일 셀 값 조회

        Args:
            sheet_id: 구글 시트 ID
            sheet_name: 시트 이름
            column: 컬럼 (A, B, C 등)
            row_number: 행 번호 (1부터 시작)

        Returns:
            str: 셀 값
        """
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=f'{sheet_name}!{column}{row_number}'
            ).execute()
            values = result.get('values', [['']])
            return values[0][0] if values and values[0] else ''
        except Exception as e:
            logger.warning(f"셀 값 조회 실패 ({column}{row_number}): {e}")
            return ''

    def get_row_values(self, sheet_id, sheet_name, row_number, start_col='A', end_col='AM', value_render_option='FORMATTED_VALUE'):
        """
        특정 행의 값을 조회 (재시도/로깅/모니터링 포함)

        Args:
            sheet_id: 구글 시트 ID
            sheet_name: 시트 이름
            row_number: 행 번호 (1부터 시작)
            start_col: 시작 컬럼 (기본값: 'A')
            end_col: 끝 컬럼 (기본값: 'AM')
            value_render_option: 값 렌더링 옵션
                - 'FORMATTED_VALUE': 포맷된 값 (기본값)
                - 'UNFORMATTED_VALUE': 포맷되지 않은 원시 값
                - 'FORMULA': 수식 포함

        Returns:
            list: 행의 값들 (빈 리스트 반환 가능)

        Raises:
            Exception: API 호출 실패 시 (재시도 후에도 실패하면)
        """
        row_range = f'{sheet_name}!{start_col}{row_number}:{end_col}{row_number}'

        # _execute_with_retry를 사용하여 재시도/로깅/모니터링 적용
        request = self.service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=row_range,
            valueRenderOption=value_render_option
        )

        result = self._execute_with_retry(
            request,
            operation_name=f"행 값 조회 (Sheet: {sheet_name}, Row: {row_number}, Range: {start_col}-{end_col})"
        )

        # 값 추출 (값이 없으면 빈 리스트)
        values = result.get('values', [[]])
        return values[0] if values else []

    def _column_letter_to_number(self, col):
        """
        컬럼 문자를 숫자 인덱스로 변환 (A->0, B->1, Z->25, AA->26, ...)

        Args:
            col: 컬럼 문자 (A, B, ..., Z, AA, AB, ...)

        Returns:
            int: 0-based 컬럼 인덱스
        """
        num = 0
        for char in col.upper():
            num = num * 26 + (ord(char) - ord('A') + 1)
        return num - 1

    def _number_to_column_letter(self, n):
        """
        숫자 인덱스를 컬럼 문자로 변환 (0->A, 1->B, 25->Z, 26->AA, ...)

        Args:
            n: 0-based 컬럼 인덱스

        Returns:
            str: 컬럼 문자
        """
        result = ""
        while n >= 0:
            result = chr(n % 26 + ord('A')) + result
            n = n // 26 - 1
            if n < 0:
                break
        return result

    def get_cell_notes(self, sheet_id, sheet_name='공사 현황의 사본', columns=['T', 'U', 'V']):
        """
        특정 컬럼들의 셀 노트(댓글)를 모두 가져오기

        Args:
            sheet_id: 구글 시트 ID
            sheet_name: 시트 이름
            columns: 읽을 컬럼 리스트 (기본: T, U, V = 계약금, 중도금, 잔금)

        Returns:
            dict: {
                행번호(1-based): {
                    'T': '계약금 메모',
                    'U': '중도금 메모',
                    'V': '잔금 메모'
                }
            }

        Example:
            {
                5: {'T': '입금일: 2025-01-10\\n입금자: 홍길동', 'U': '입금일: 2025-02-15\\n입금자: 김철수'},
                10: {'V': '입금일: 2025-03-20\\n입금자: 이영희'}
            }
        """
        try:
            # 컬럼 범위 구성 (예: T:V)
            start_col = min(columns)
            end_col = max(columns)
            range_name = f"{sheet_name}!{start_col}:{end_col}"

            logger.info(f"[GET_CELL_NOTES] 셀 노트 읽기 시작: {range_name}")

            # includeGridData=True로 셀 데이터와 노트 함께 가져오기
            request = self.service.spreadsheets().get(
                spreadsheetId=sheet_id,
                ranges=[range_name],
                includeGridData=True
            )
            result = self._execute_with_retry(request, f"get_cell_notes({range_name})")

            notes_by_row = {}
            sheets = result.get('sheets', [])

            if not sheets:
                logger.warning(f"[GET_CELL_NOTES] 시트를 찾을 수 없습니다: {sheet_name}")
                return notes_by_row

            # 첫 번째 시트의 gridData 파싱
            for sheet in sheets:
                grid_data = sheet.get('data', [])

                for data in grid_data:
                    start_row = data.get('startRow', 0)  # 0-based
                    start_col = data.get('startColumn', 0)  # 0-based

                    row_data_list = data.get('rowData', [])

                    for row_idx, row_data in enumerate(row_data_list):
                        actual_row = start_row + row_idx + 1  # 1-based 행번호

                        values = row_data.get('values', [])

                        for col_idx, cell_data in enumerate(values):
                            # 셀에 노트가 있는지 확인
                            note = cell_data.get('note')

                            if note:
                                actual_col_index = start_col + col_idx
                                col_letter = self._number_to_column_letter(actual_col_index)

                                # 해당 컬럼이 요청한 컬럼 리스트에 있는지 확인
                                if col_letter in columns:
                                    if actual_row not in notes_by_row:
                                        notes_by_row[actual_row] = {}
                                    notes_by_row[actual_row][col_letter] = note

            logger.info(f"[GET_CELL_NOTES] 완료: {len(notes_by_row)}개 행에서 노트 발견")

            # 디버그 로깅 (처음 3개 행만)
            # RequestContextLogger와 호환성 문제로 가드 제거 (logger가 레벨 체크 수행)
            if notes_by_row:
                sample_rows = list(notes_by_row.items())[:self.DEBUG_SAMPLE_NOTES]
                for row_num, notes in sample_rows:
                    logger.debug(f"  행 {row_num}: {notes}")

            return notes_by_row

        except Exception as e:
            logger.error(f"[ERROR] 셀 노트 읽기 오류: {str(e)}")
            import traceback
            logger.error(f"스택 트레이스:\n{traceback.format_exc()}")
            return {}

    # ========================================================================
    # 헬퍼 함수: update_cell_note() 리팩토링
    # ========================================================================

    def _parse_cell_address_for_note(self, cell_address):
        """
        셀 주소 파싱 (예: 'T5' → col_letter='T', row_number=5)

        Args:
            cell_address: 셀 주소 (예: 'T5', 'U10', 'V20')

        Returns:
            tuple: (col_letter, row_number, col_index, row_index)
        """
        logger.debug(f"[MEMO_SAVE_STEP_2] 셀 주소 파싱 시작: {cell_address}")
        col_letter = ''.join(filter(str.isalpha, cell_address))
        row_number = int(''.join(filter(str.isdigit, cell_address)))
        col_index = self._column_letter_to_number(col_letter)
        row_index = row_number - 1  # 0-based
        logger.debug(f"[MEMO_SAVE_STEP_2] 파싱 완료: 행={row_index}, 열={col_index}")
        return col_letter, row_number, col_index, row_index

    def _create_note_batch_request(self, numeric_sheet_id, row_index, col_index, note_text):
        """
        셀 노트 업데이트를 위한 batchUpdate 요청 생성

        Args:
            numeric_sheet_id: 시트 ID (숫자)
            row_index: 행 인덱스 (0-based)
            col_index: 열 인덱스 (0-based)
            note_text: 메모 내용 (None이면 삭제)

        Returns:
            list: batchUpdate requests
        """
        logger.debug(f"[MEMO_SAVE_STEP_3] batchUpdate 요청 생성 시작")
        requests = [{
            'updateCells': {
                'range': {
                    'sheetId': numeric_sheet_id,
                    'startRowIndex': row_index,
                    'endRowIndex': row_index + 1,
                    'startColumnIndex': col_index,
                    'endColumnIndex': col_index + 1
                },
                'rows': [{
                    'values': [{
                        'note': note_text if note_text else None
                    }]
                }],
                'fields': 'note'
            }
        }]
        logger.debug(f"[MEMO_SAVE_STEP_3] 요청 생성 완료")
        return requests

    def _execute_note_api_with_monitoring(self, sheet_id, requests, cell_address, process):
        """
        Google Sheets API 호출 (메모리 모니터링 포함)

        Args:
            sheet_id: 구글 시트 ID
            requests: batchUpdate 요청 리스트
            cell_address: 셀 주소 (로깅용)
            process: psutil.Process 객체

        Returns:
            result: API 응답 객체
        """
        logger.debug(f"[MEMO_SAVE_STEP_4] Google Sheets API 호출 시작 (CRITICAL SECTION)")
        request = self.service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={'requests': requests}
        )
        logger.debug(f"[MEMO_SAVE_STEP_4] Request 객체 생성 완료, execute() 호출 직전")

        # API 호출 전 메모리 체크
        pre_api_memory = process.memory_info().rss / 1024 / 1024
        logger.debug(f"[MEMO_SAVE_STEP_4] API 호출 전 메모리: {pre_api_memory:.1f}MB")

        result = self._execute_with_retry(request, f"update_cell_note({cell_address})")

        # API 호출 후 메모리 체크
        post_api_memory = process.memory_info().rss / 1024 / 1024
        memory_increase = post_api_memory - pre_api_memory
        logger.debug(f"[MEMO_SAVE_STEP_4] API 호출 후 메모리: {post_api_memory:.1f}MB (증가: {memory_increase:+.1f}MB)")

        logger.debug(f"[MEMO_SAVE_STEP_4] Google Sheets API 호출 성공")
        return result


    def _log_note_operation_error(self, e, cell_address, elapsed_time, memory_change, process):
        """
        셀 노트 저장 실패 시 상세 에러 로깅

        Args:
            e: 발생한 예외
            cell_address: 셀 주소
            elapsed_time: 소요 시간
            memory_change: 메모리 변화량
            process: psutil.Process 객체
        """
        import sys
        import traceback

        logger.error(f"[MEMO_SAVE_ERROR] 셀: {cell_address}")
        logger.error(f"[MEMO_SAVE_ERROR] 에러 타입: {type(e).__name__}")
        logger.error(f"[MEMO_SAVE_ERROR] 에러 메시지: {str(e)}")
        logger.error(f"[MEMO_SAVE_ERROR] 소요시간: {elapsed_time:.2f}초")
        logger.error(f"[MEMO_SAVE_ERROR] 메모리 변화: {memory_change:+.1f}MB")

        import threading
        logger.error(f"[MEMO_SAVE_ERROR] 스레드: {threading.current_thread().name}")
        logger.error(f"[MEMO_SAVE_ERROR] Python 버전: {sys.version}")

        stack_trace = traceback.format_exc()
        logger.error(f"[MEMO_SAVE_ERROR] 스택 트레이스:\n{stack_trace}")

        # 시스템 리소스 정보 덤프
        try:
            cpu_percent = process.cpu_percent(interval=0.1)
            open_files = len(process.open_files())
            num_threads = process.num_threads()
            logger.error(f"[MEMO_SAVE_ERROR] 시스템 상태: CPU={cpu_percent}%, 열린파일={open_files}, 스레드수={num_threads}")
        except Exception as resource_error:
            logger.warning(f"처리 중 오류 무시: {resource_error}")

    # ========================================================================
    # 메인 함수 (리팩토링됨)
    # ========================================================================

    def update_cell_note(self, sheet_id, sheet_name, cell_address, note_text):
        """
        특정 셀에 노트(댓글) 추가/수정/삭제 (스레드 안전, 상세 로깅) - 리팩토링됨

        Args:
            sheet_id: 구글 시트 ID
            sheet_name: 시트 이름
            cell_address: 셀 주소 (예: 'T5', 'U10', 'V20')
            note_text: 메모 내용 (None 또는 빈 문자열이면 댓글 삭제)

        Returns:
            bool: 성공 여부

        Example:
            manager.update_cell_note(sheet_id, '공사 현황의 사본', 'T5', '입금일: 2025-01-10\\n입금자: 홍길동')
            manager.update_cell_note(sheet_id, '공사 현황의 사본', 'U10', None)  # 댓글 삭제
        """
        import psutil
        import threading

        # 스레드 안전성 확보: 동시에 여러 메모 저장 방지
        with self._lock:
            start_time = time.time()
            process = psutil.Process()
            initial_memory = process.memory_info().rss / 1024 / 1024  # MB

            try:
                logger.info(f"[MEMO_SAVE_START] 셀: {cell_address}, 스레드: {threading.current_thread().name}, 메모리: {initial_memory:.1f}MB")

                # 1단계: 시트 ID 가져오기
                logger.debug(f"[MEMO_SAVE_STEP_1] 시트 ID 조회 시작: {sheet_name}")
                numeric_sheet_id = self.get_sheet_id_by_name(sheet_id, sheet_name)
                logger.debug(f"[MEMO_SAVE_STEP_1] 시트 ID 조회 완료: {numeric_sheet_id}")

                # 2단계: 셀 주소 파싱
                col_letter, row_number, col_index, row_index = self._parse_cell_address_for_note(cell_address)

                # 3단계: batchUpdate 요청 생성
                requests = self._create_note_batch_request(numeric_sheet_id, row_index, col_index, note_text)

                # 4단계: Google Sheets API 호출 (메모리 모니터링 포함)
                result = self._execute_note_api_with_monitoring(sheet_id, requests, cell_address, process)

                # 5단계: 완료 처리
                elapsed_time = time.time() - start_time
                action = "삭제" if not note_text else "저장"
                note_preview = note_text[:self.NOTE_PREVIEW_LENGTH] + '...' if note_text and len(note_text) > self.NOTE_PREVIEW_LENGTH else (note_text or 'N/A')

                final_memory = process.memory_info().rss / 1024 / 1024
                total_memory_change = final_memory - initial_memory

                logger.info(f"[MEMO_SAVE_SUCCESS] {cell_address}: {note_preview}")
                logger.info(f"[MEMO_SAVE_PERF] 소요시간: {elapsed_time:.2f}초, 메모리 변화: {total_memory_change:+.1f}MB")

                # 메모리 크래시 방지: API 응답 객체 명시적 정리
                try:
                    del result
                    import gc
                    gc.collect()
                    logger.debug(f"[MEMO_SAVE_CLEANUP] API 응답 객체 정리 완료, 가비지 컬렉션 실행")
                except Exception as cleanup_error:
                    logger.debug(f"[MEMO_SAVE_CLEANUP] 정리 중 경고: {cleanup_error}")

                return True

            except Exception as e:
                elapsed_time = time.time() - start_time
                final_memory = process.memory_info().rss / 1024 / 1024
                memory_change = final_memory - initial_memory

                self._log_note_operation_error(e, cell_address, elapsed_time, memory_change, process)
                return False

    def get_cell_note(self, sheet_id, sheet_name, cell_address):
        """
        특정 셀의 노트(댓글) 조회

        Args:
            sheet_id: 구글 시트 ID
            sheet_name: 시트 이름
            cell_address: 셀 주소 (예: 'T5')

        Returns:
            str: 노트 내용 (없으면 None)
        """
        try:
            col_letter = ''.join(filter(str.isalpha, cell_address))

            # 해당 컬럼의 모든 노트 가져오기
            notes = self.get_cell_notes(sheet_id, sheet_name, [col_letter])

            # 행 번호 추출
            row_number = int(''.join(filter(str.isdigit, cell_address)))

            # 해당 행의 노트 반환
            if row_number in notes and col_letter in notes[row_number]:
                return notes[row_number][col_letter]

            return None

        except Exception as e:
            logger.warning(f"셀 노트 조회 실패 ({cell_address}): {e}")
            return None

    def update_row_with_notes_batch(self, sheet_id, sheet_name, row_number, row_values, cell_notes=None):
        """
        행 데이터와 여러 셀 노트를 Batch로 업데이트 (통합편집 최적화)

        Args:
            sheet_id: 구글 시트 ID
            sheet_name: 시트 이름
            row_number: 행 번호 (1부터 시작)
            row_values: 행 데이터 리스트 (39개 컬럼)
            cell_notes: 셀 노트 딕셔너리 {cell_address: note_text}
                       예: {'T5': '메모1', 'U5': '메모2', 'V5': None}

        Returns:
            dict: {
                'row_updated': bool,
                'notes_updated': int,
                'total_api_calls': int
            }

        Example:
            manager.update_row_with_notes_batch(
                sheet_id='abc123',
                sheet_name='공사 현황의 사본',
                row_number=5,
                row_values=['G0001-IT', '강성환', ...],
                cell_notes={'T5': '계약금 메모', 'U5': '중도금 메모', 'V5': None}
            )
        """
        import time
        start_time = time.time()
        api_calls = 0

        try:
            # 1. 행 데이터 업데이트 (1번 API 호출)
            range_name = f'{sheet_name}!A{row_number}:AM{row_number}'
            body = {'values': [row_values]}

            request = self.service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',
                body=body
            )
            row_result = self._execute_with_retry(request, f"update_row_batch(row={row_number})")
            api_calls += 1

            updated_cells = row_result.get('updatedCells', 0)
            logger.info(f"[BATCH_UPDATE] 행 업데이트 완료: {row_number}행, {updated_cells}셀")

            # 2. 셀 노트 Batch 업데이트 (1번 API 호출로 모든 메모 처리)
            notes_updated = 0
            if cell_notes:
                # 시트 ID 조회
                numeric_sheet_id = self.get_sheet_id_by_name(sheet_id, sheet_name)

                # batchUpdate 요청 구성
                batch_requests = []
                for cell_address, note_text in cell_notes.items():
                    # 셀 주소 파싱
                    col_letter = ''.join(filter(str.isalpha, cell_address))
                    row_num = int(''.join(filter(str.isdigit, cell_address)))
                    col_index = self._column_letter_to_number(col_letter)
                    row_index = row_num - 1  # 0-based

                    # 메모 업데이트 요청 생성
                    batch_requests.append({
                        'repeatCell': {
                            'range': {
                                'sheetId': numeric_sheet_id,
                                'startRowIndex': row_index,
                                'endRowIndex': row_index + 1,
                                'startColumnIndex': col_index,
                                'endColumnIndex': col_index + 1
                            },
                            'cell': {
                                'note': note_text if note_text else None
                            },
                            'fields': 'note'
                        }
                    })
                    notes_updated += 1

                # Batch Update 실행 (1번 API 호출로 모든 메모 처리)
                if batch_requests:
                    body = {'requests': batch_requests}
                    request = self.service.spreadsheets().batchUpdate(
                        spreadsheetId=sheet_id,
                        body=body
                    )
                    notes_result = self._execute_with_retry(
                        request,
                        f"update_notes_batch({len(batch_requests)} notes)"
                    )
                    api_calls += 1

                    logger.info(f"[BATCH_UPDATE] 메모 업데이트 완료: {notes_updated}개 메모")

            elapsed_time = time.time() - start_time
            logger.info(
                f"[BATCH_UPDATE] 완료: 행 {row_number} + {notes_updated}개 메모, "
                f"API 호출 {api_calls}회, 소요시간 {elapsed_time:.2f}초"
            )

            return {
                'row_updated': True,
                'notes_updated': notes_updated,
                'total_api_calls': api_calls
            }

        except Exception as e:
            logger.error(f"[BATCH_UPDATE] 실패: {str(e)}")
            raise

    def clear_row_cache(self):
        """프로젝트 코드 → 행번호 캐시 초기화"""
        self._row_cache.clear()
        logger.debug("[ROW_CACHE] 캐시 초기화")

def test_google_sheets_connection():
    """구글 시트 연결 테스트 함수"""
    from dotenv import load_dotenv
    load_dotenv()
    
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("GOOGLE_SHEET_ID가 .env 파일에 설정되지 않았습니다.")
        return
    
    try:
        manager = GoogleSheetsManager()
        if manager.validate_connection(sheet_id):
            print("[SUCCESS] 구글 시트 연결 성공!")
            
            # 샘플 데이터 가져오기
            df = manager.get_sheet_data(sheet_id)
            print(f"[STATS] 데이터 크기: {df.shape}")
            print(f"[INFO] 컬럼 수: {len(df.columns)}")
        else:
            print("[ERROR] 구글 시트 연결 실패")
    except Exception as e:
        print(f"[ERROR] 오류 발생: {str(e)}")

if __name__ == "__main__":
    test_google_sheets_connection()
