"""
환경변수 검증 시스템
필수 환경변수 체크 및 타입 검증
"""

import os
import logging
from typing import Dict, List, Optional, Any, Union
from dotenv import load_dotenv

logger = logging.getLogger(__name__)

class EnvValidator:
    """환경변수 검증 클래스"""
    
    def __init__(self, env_file='.env', strict=True):
        """환경변수 검증기 초기화"""
        self.env_file = env_file
        self.strict = strict
        self.errors: List[str] = []
        self.warnings: List[str] = []
        # .env 파일 로드
        if os.path.exists(env_file):
            load_dotenv(env_file)
            logger.info(f"환경변수 설정 로드: {env_file}")
        else:
            self.warnings.append(f"환경변수 파일을 찾을 수 없음: {env_file}")
    def validate_required(self, var_name: str, description: str = "") -> bool:
        """필수 환경변수 검증"""
        value = os.getenv(var_name)
        if not value or value.strip() == "":
            message = f"필수 환경변수 누락: {var_name}" + (f" ({description})" if description else "")
            if self.strict:
                self.errors.append(message)
            else:
                self.warnings.append(message)
            return False
        logger.debug(f"환경변수 확인: {var_name} = {'*' * len(value)}")  # 값 마스킹
        return True
    def validate_optional(self, var_name: str, default_value: str = "", 
                         description: str = "") -> str:
        """선택적 환경변수 검증 (기본값 제공)"""
        value = os.getenv(var_name, default_value)
        
        if value == default_value and default_value:
            self.warnings.append(f"환경변수 기본값 사용: {var_name} = {default_value}" +
                               (f" ({description})" if description else ""))
        
        return value
    
    def validate_integer(self, var_name: str, min_value: Optional[int] = None,
                        max_value: Optional[int] = None, default: int = 0) -> int:
        """정수 환경변수 검증"""
        value_str = os.getenv(var_name, str(default))
        
        try:
            value = int(value_str)
            
            if min_value is not None and value < min_value:
                message = f"환경변수 {var_name} 값이 너무 작음: {value} (최소: {min_value})"
                if self.strict:
                    self.errors.append(message)
                else:
                    self.warnings.append(message)
                return default
            
            if max_value is not None and value > max_value:
                message = f"환경변수 {var_name} 값이 너무 큼: {value} (최대: {max_value})"
                if self.strict:
                    self.errors.append(message)
                else:
                    self.warnings.append(message)
                return default
            
            return value
            
        except ValueError:
            message = f"환경변수 {var_name} 정수 변환 실패: {value_str}"
            if self.strict:
                self.errors.append(message)
            else:
                self.warnings.append(message)
            return default
    
    def validate_boolean(self, var_name: str, default: bool = False) -> bool:
        """불린 환경변수 검증"""
        value_str = os.getenv(var_name, str(default)).lower()
        
        if value_str in ('true', '1', 'yes', 'on'):
            return True
        elif value_str in ('false', '0', 'no', 'off'):
            return False
        else:
            self.warnings.append(f"환경변수 {var_name} 불명확한 불린 값: {value_str}")
            return default
    
    def validate_file_exists(self, var_name: str, description: str = "",
                           fallback_paths: List[str] = None) -> bool:
        """파일 경로 환경 변수 검증 (대체 경로 시도)"""
        file_path = os.getenv(var_name)
        if file_path:
            if os.path.exists(file_path):
                return True
            self.warnings.append(f"환경변수 참조 파일을 찾을 수 없음: {file_path} (환경변수: {var_name})")
        if fallback_paths:
            for fallback_path in fallback_paths:
                if os.path.exists(fallback_path):
                    os.environ[var_name] = fallback_path
                    self.warnings.append(f"대체 경로 사용 중: {fallback_path}")
                    return True
        message = None
        if not file_path:
            message = f"환경변수 누락: {var_name}" + (f" ({description})" if description else "")
        else:
            message = f"파일을 찾을 수 없음: {file_path} (환경변수: {var_name})"
        if message:
            if self.strict:
                self.errors.append(message)
            else:
                self.warnings.append(message)
        return False
    def validate_url(self, var_name: str, schemes: List[str] = None) -> Optional[str]:
        """URL 형식 검증"""
        url = os.getenv(var_name)
        
        if not url:
            return None
        
        if schemes is None:
            schemes = ['http', 'https']
        
        if not any(url.startswith(f"{scheme}://") for scheme in schemes):
            self.errors.append(f"잘못된 URL 형식: {var_name} = {url}")
            return None
        
        return url
    
    def validate_email(self, var_name: str) -> Optional[str]:
        """이메일 형식 검증"""
        email = os.getenv(var_name)
        
        if not email:
            return None
        
        import re
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        
        if not re.match(email_pattern, email):
            self.errors.append(f"잘못된 이메일 형식: {var_name} = {email}")
            return None
        
        return email
    
    def get_validation_result(self) -> Dict[str, Any]:
        """검증 결과 반환"""
        return {
            'valid': len(self.errors) == 0,
            'errors': self.errors,
            'warnings': self.warnings,
            'error_count': len(self.errors),
            'warning_count': len(self.warnings)
        }
    
    def print_summary(self) -> None:
        """검증 결과 요약 출력"""
        result = self.get_validation_result()
        
        if result['valid']:
            logger.info("[SUCCESS] 환경변수 검증 성공")
        else:
            logger.error("[ERROR] 환경변수 검증 실패")
        
        if self.errors:
            logger.error("오류:")
            for error in self.errors:
                logger.error(f"  - {error}")
        
        if self.warnings:
            logger.warning("경고:")
            for warning in self.warnings:
                logger.warning(f"  - {warning}")

def validate_dashboard_env(strict: bool = True) -> Dict[str, Any]:
    """대시보드 환경변수 검증"""
    validator = EnvValidator(strict=strict)
    
    # 필수 환경변수
    validator.validate_required('GOOGLE_SHEET_ID', '구글 시트 ID')
    
    # credentials 파일 경로 (여러 위치 시도)
    credentials_fallback_paths = [
        'credentials.json',  # 프로젝트 루트
        'dashboard/credentials.json',  # dashboard 폴더
        '../credentials.json',  # 상위 폴더
        os.path.join(os.path.dirname(__file__), '..', '..', 'credentials.json'),  # 상대 경로
    ]
    validator.validate_file_exists('GOOGLE_APPLICATION_CREDENTIALS', '구글 서비스 계정 키 파일', 
                                  credentials_fallback_paths)
    
    # 선택적 환경변수
    validator.validate_optional('SECRET_KEY', 'default-secret-key', 'Flask 비밀 키')
    
    # 서버 설정
    port = validator.validate_integer('PORT', min_value=1000, max_value=65535, default=5000)
    debug = validator.validate_boolean('DEBUG', default=False)
    
    # OAuth 설정 (선택적)
    oauth_client_id = validator.validate_optional('GOOGLE_OAUTH_CLIENT_ID', '', 'OAuth 클라이언트 ID')
    oauth_client_secret = validator.validate_optional('GOOGLE_OAUTH_CLIENT_SECRET', '', 'OAuth 클라이언트 시크릿')
    
    # URL 검증
    if os.getenv('CALLBACK_URL'):
        validator.validate_url('CALLBACK_URL', ['http', 'https'])
    
    # OAuth 인증으로 관리자 권한 관리 (ADMIN_EMAILS 더 이상 불필요)
    
    # 결과 반환
    result = validator.get_validation_result()
    result.update({
        'config': {
            'port': port,
            'debug': debug,
            'oauth_enabled': bool(oauth_client_id and oauth_client_secret)
        }
    })
    
    result["strict"] = strict
    return result

def check_environment_health() -> bool:
    """환경 상태 체크"""
    result = validate_dashboard_env()
    
    # 결과 출력
    if result['valid']:
        logger.info("[PASS] 환경변수 검증 통과")
        if result['warnings']:
            logger.warning(f"[WARN] 경고 {len(result['warnings'])}개")
    else:
        logger.error(f"[FAIL] 환경변수 검증 실패 - 오류 {len(result['errors'])}개")
        
    return result['valid']

if __name__ == "__main__":
    # 환경변수 검증 테스트
    check_environment_health()