"""
API 응답 표준화 시스템
일관된 API 응답 형식과 에러 코드 관리
"""

from typing import Any, Dict, List, Optional, Union
from flask import jsonify, make_response
from datetime import datetime
import logging

logger = logging.getLogger(__name__)

class APIErrorCode:
    """표준 API 에러 코드"""
    
    # 인증 관련
    UNAUTHORIZED = 'UNAUTHORIZED'
    FORBIDDEN = 'FORBIDDEN'
    INVALID_TOKEN = 'INVALID_TOKEN'
    TOKEN_EXPIRED = 'TOKEN_EXPIRED'
    INVALID_CREDENTIALS = 'INVALID_CREDENTIALS'
    
    # 요청 관련
    BAD_REQUEST = 'BAD_REQUEST'
    VALIDATION_FAILED = 'VALIDATION_FAILED'
    MISSING_PARAMETER = 'MISSING_PARAMETER'
    INVALID_PARAMETER = 'INVALID_PARAMETER'
    
    # 리소스 관련
    NOT_FOUND = 'NOT_FOUND'
    ALREADY_EXISTS = 'ALREADY_EXISTS'
    CONFLICT = 'CONFLICT'
    
    # 서버 관련
    INTERNAL_ERROR = 'INTERNAL_ERROR'
    SERVICE_UNAVAILABLE = 'SERVICE_UNAVAILABLE'
    DATABASE_ERROR = 'DATABASE_ERROR'
    EXTERNAL_API_ERROR = 'EXTERNAL_API_ERROR'
    
    # 비즈니스 로직 관련
    INSUFFICIENT_PERMISSIONS = 'INSUFFICIENT_PERMISSIONS'
    OPERATION_FAILED = 'OPERATION_FAILED'
    INVALID_STATE = 'INVALID_STATE'

class APIResponse:
    """표준화된 API 응답"""
    
    @staticmethod
    def success(
        data: Any = None, 
        message: str = "Success", 
        status_code: int = 200,
        metadata: Optional[Dict[str, Any]] = None
    ):
        """성공 응답"""
        response_data = {
            'success': True,
            'message': message,
            'timestamp': datetime.now().isoformat(),
            'data': data
        }
        
        if metadata:
            response_data['metadata'] = metadata
        
        response = make_response(jsonify(response_data), status_code)
        return response
    
    @staticmethod
    def error(
        message: str,
        error_code: str = APIErrorCode.INTERNAL_ERROR,
        status_code: int = 500,
        details: Optional[Union[str, List[str], Dict[str, Any]]] = None,
        data: Any = None
    ):
        """에러 응답"""
        response_data = {
            'success': False,
            'error': {
                'code': error_code,
                'message': message,
                'details': details
            },
            'timestamp': datetime.now().isoformat()
        }
        
        if data is not None:
            response_data['data'] = data
        
        # 에러 로깅
        logger.error(f"API Error: {error_code} - {message}")
        if details:
            logger.error(f"Error details: {details}")
        
        response = make_response(jsonify(response_data), status_code)
        return response
    
    @staticmethod
    def validation_error(errors: Union[str, List[str], Dict[str, str]]):
        """유효성 검사 실패 응답"""
        return APIResponse.error(
            message="Validation failed",
            error_code=APIErrorCode.VALIDATION_FAILED,
            status_code=400,
            details=errors
        )
    
    @staticmethod
    def not_found(resource: str = "Resource"):
        """리소스를 찾을 수 없음 응답"""
        return APIResponse.error(
            message=f"{resource} not found",
            error_code=APIErrorCode.NOT_FOUND,
            status_code=404
        )
    
    @staticmethod
    def unauthorized(message: str = "Authentication required"):
        """인증 필요 응답"""
        return APIResponse.error(
            message=message,
            error_code=APIErrorCode.UNAUTHORIZED,
            status_code=401
        )
    
    @staticmethod
    def forbidden(message: str = "Insufficient permissions"):
        """권한 부족 응답"""
        return APIResponse.error(
            message=message,
            error_code=APIErrorCode.FORBIDDEN,
            status_code=403
        )
    
    @staticmethod
    def paginated(
        data: List[Any],
        page: int = 1,
        per_page: int = 20,
        total: int = 0,
        message: str = "Success"
    ):
        """페이지네이션 응답"""
        total_pages = (total + per_page - 1) // per_page if total > 0 else 1
        
        metadata = {
            'pagination': {
                'page': page,
                'per_page': per_page,
                'total': total,
                'total_pages': total_pages,
                'has_next': page < total_pages,
                'has_prev': page > 1
            }
        }
        
        return APIResponse.success(
            data=data,
            message=message,
            metadata=metadata
        )
    
    @staticmethod
    def created(data: Any = None, message: str = "Created successfully"):
        """리소스 생성 성공 응답"""
        return APIResponse.success(
            data=data,
            message=message,
            status_code=201
        )
    
    @staticmethod
    def updated(data: Any = None, message: str = "Updated successfully"):
        """리소스 업데이트 성공 응답"""
        return APIResponse.success(
            data=data,
            message=message,
            status_code=200
        )
    
    @staticmethod
    def deleted(message: str = "Deleted successfully"):
        """리소스 삭제 성공 응답"""
        return APIResponse.success(
            message=message,
            status_code=200
        )

class APIValidator:
    """API 요청 유효성 검사"""
    
    @staticmethod
    def validate_required_fields(data: Dict[str, Any], required_fields: List[str]) -> List[str]:
        """필수 필드 검증"""
        missing_fields = []
        for field in required_fields:
            if field not in data or data[field] is None or data[field] == '':
                missing_fields.append(f"'{field}' is required")
        return missing_fields
    
    @staticmethod
    def validate_field_types(data: Dict[str, Any], field_types: Dict[str, type]) -> List[str]:
        """필드 타입 검증"""
        type_errors = []
        for field, expected_type in field_types.items():
            if field in data and data[field] is not None:
                if not isinstance(data[field], expected_type):
                    type_errors.append(f"'{field}' must be of type {expected_type.__name__}")
        return type_errors
    
    @staticmethod
    def validate_field_lengths(data: Dict[str, Any], field_lengths: Dict[str, Dict[str, int]]) -> List[str]:
        """필드 길이 검증"""
        length_errors = []
        for field, constraints in field_lengths.items():
            if field in data and data[field] is not None:
                value = str(data[field])
                min_len = constraints.get('min', 0)
                max_len = constraints.get('max')
                
                if len(value) < min_len:
                    length_errors.append(f"'{field}' must be at least {min_len} characters")
                
                if max_len and len(value) > max_len:
                    length_errors.append(f"'{field}' must be at most {max_len} characters")
        
        return length_errors

# 데코레이터들
from functools import wraps

def api_response(f):
    """API 응답 표준화 데코레이터"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        try:
            result = f(*args, **kwargs)
            
            # 이미 Flask Response 객체인 경우 그대로 반환
            if hasattr(result, 'status_code'):
                return result
            
            # 딕셔너리인 경우 표준 형식으로 변환
            if isinstance(result, dict):
                if 'error' in result:
                    return APIResponse.error(
                        message=result.get('message', 'Unknown error'),
                        error_code=result.get('error_code', APIErrorCode.INTERNAL_ERROR),
                        status_code=result.get('status_code', 500),
                        details=result.get('details')
                    )
                else:
                    return APIResponse.success(
                        data=result.get('data'),
                        message=result.get('message', 'Success'),
                        status_code=result.get('status_code', 200)
                    )
            
            # 기타 경우 data로 감싸서 반환
            return APIResponse.success(data=result)
            
        except Exception as e:
            logger.error(f"API endpoint error: {str(e)}", exc_info=True)
            return APIResponse.error(
                message="Internal server error",
                error_code=APIErrorCode.INTERNAL_ERROR,
                status_code=500
            )
    
    return decorated_function

def validate_json_request(*required_fields, **field_types):
    """JSON 요청 검증 데코레이터"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            from flask import request
            
            if not request.is_json:
                return APIResponse.error(
                    message="Content-Type must be application/json",
                    error_code=APIErrorCode.BAD_REQUEST,
                    status_code=400
                )
            
            data = request.get_json()
            if not data:
                return APIResponse.error(
                    message="JSON data is required",
                    error_code=APIErrorCode.BAD_REQUEST,
                    status_code=400
                )
            
            # 필수 필드 검증
            validation_errors = []
            if required_fields:
                validation_errors.extend(
                    APIValidator.validate_required_fields(data, list(required_fields))
                )
            
            # 타입 검증
            if field_types:
                validation_errors.extend(
                    APIValidator.validate_field_types(data, field_types)
                )
            
            if validation_errors:
                return APIResponse.validation_error(validation_errors)
            
            return f(*args, **kwargs)
        
        return decorated_function
    return decorator