"""
API 버저닝 시스템 구현
URL 경로 기반 버저닝 및 하위 호환성 관리
"""

from flask import Blueprint, request, current_app
from typing import Dict, List, Optional, Callable, Any
from functools import wraps
import re

from .responses import APIResponse, APIErrorCode


class APIVersionManager:
    """API 버전 관리 클래스"""

    def __init__(self):
        self.versions: Dict[str, Dict] = {}
        self.default_version = "v1"
        self.deprecated_versions: List[str] = []
        self.version_mapping: Dict[str, str] = {}

    def register_version(
        self,
        version: str,
        blueprint: Blueprint,
        description: str = "",
        deprecated: bool = False,
        sunset_date: Optional[str] = None
    ):
        """버전 등록"""
        self.versions[version] = {
            "blueprint": blueprint,
            "description": description,
            "deprecated": deprecated,
            "sunset_date": sunset_date,
            "endpoints": []
        }

        if deprecated:
            self.deprecated_versions.append(version)

    def set_default_version(self, version: str):
        """기본 버전 설정"""
        if version in self.versions:
            self.default_version = version
        else:
            raise ValueError(f"Version {version} not registered")

    def add_version_mapping(self, old_version: str, new_version: str):
        """버전 매핑 추가 (리다이렉션용)"""
        self.version_mapping[old_version] = new_version

    def get_version_info(self, version: str) -> Optional[Dict]:
        """버전 정보 조회"""
        return self.versions.get(version)

    def is_version_deprecated(self, version: str) -> bool:
        """버전 deprecation 여부 확인"""
        return version in self.deprecated_versions

    def get_available_versions(self) -> List[str]:
        """사용 가능한 버전 목록"""
        return list(self.versions.keys())


# 전역 버전 매니저 인스턴스
version_manager = APIVersionManager()


def versioned_route(version: str, rule: str, **options):
    """버전별 라우트 데코레이터"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            # 버전 정보를 헤더에 추가
            response = func(*args, **kwargs)

            # Flask Response 객체인 경우 헤더 추가
            if hasattr(response, 'headers'):
                response.headers['API-Version'] = version

                # Deprecation 경고
                if version_manager.is_version_deprecated(version):
                    response.headers['Deprecation'] = 'true'
                    version_info = version_manager.get_version_info(version)
                    if version_info and version_info.get('sunset_date'):
                        response.headers['Sunset'] = version_info['sunset_date']

            return response

        # 함수에 버전 정보 메타데이터 추가
        wrapper._api_version = version
        wrapper._api_rule = rule

        return wrapper
    return decorator


def version_compatibility_check(supported_versions: List[str]):
    """버전 호환성 체크 데코레이터"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            # 요청된 버전 추출
            requested_version = extract_version_from_request()

            if requested_version not in supported_versions:
                return APIResponse.error(
                    APIErrorCode.BAD_REQUEST,
                    f"API version '{requested_version}' is not supported for this endpoint. "
                    f"Supported versions: {', '.join(supported_versions)}",
                    details={
                        "requested_version": requested_version,
                        "supported_versions": supported_versions
                    },
                    status_code=400
                )

            return func(*args, **kwargs)
        return wrapper
    return decorator


def extract_version_from_request() -> str:
    """요청에서 API 버전 추출"""
    # 1. URL 경로에서 버전 추출 (/api/v1/...)
    path = request.path
    version_match = re.match(r'^/api/(v\d+)/', path)
    if version_match:
        return version_match.group(1)

    # 2. Accept 헤더에서 버전 추출 (application/vnd.api.v1+json)
    accept_header = request.headers.get('Accept', '')
    version_match = re.search(r'application/vnd\.api\.(v\d+)\+json', accept_header)
    if version_match:
        return version_match.group(1)

    # 3. 커스텀 헤더에서 버전 추출
    api_version_header = request.headers.get('API-Version')
    if api_version_header:
        return api_version_header

    # 4. 쿼리 파라미터에서 버전 추출
    version_param = request.args.get('version')
    if version_param:
        return version_param

    # 기본 버전 반환
    return version_manager.default_version


def create_versioned_blueprint(
    name: str,
    version: str,
    url_prefix: str = None,
    description: str = "",
    deprecated: bool = False,
    sunset_date: Optional[str] = None
) -> Blueprint:
    """버전별 Blueprint 생성"""

    if not url_prefix:
        url_prefix = f'/api/{version}'

    # Blueprint 생성
    bp = Blueprint(
        f'{name}_{version}',
        __name__,
        url_prefix=url_prefix
    )

    # 버전 매니저에 등록
    version_manager.register_version(
        version=version,
        blueprint=bp,
        description=description,
        deprecated=deprecated,
        sunset_date=sunset_date
    )

    # 버전 정보 엔드포인트 자동 추가
    @bp.route('/version', methods=['GET'])
    def version_info():
        """현재 API 버전 정보"""
        version_data = version_manager.get_version_info(version)

        return APIResponse.success(
            data={
                "version": version,
                "description": version_data.get("description", ""),
                "deprecated": version_data.get("deprecated", False),
                "sunset_date": version_data.get("sunset_date"),
                "endpoints_count": len(version_data.get("endpoints", [])),
                "available_versions": version_manager.get_available_versions()
            },
            message=f"API {version} version information"
        )

    return bp


def migrate_endpoint(old_version: str, new_version: str, breaking_changes: List[str] = None):
    """엔드포인트 마이그레이션 데코레이터"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            requested_version = extract_version_from_request()

            # 구 버전 요청인 경우 경고 헤더 추가
            if requested_version == old_version:
                current_app.logger.warning(
                    f"Client using deprecated API version {old_version}. "
                    f"Migration to {new_version} recommended."
                )

            response = func(*args, **kwargs)

            # 마이그레이션 정보 헤더 추가
            if hasattr(response, 'headers'):
                if requested_version == old_version:
                    response.headers['Migration-Info'] = f"Please migrate to {new_version}"
                    if breaking_changes:
                        response.headers['Breaking-Changes'] = ';'.join(breaking_changes)

            return response
        return wrapper
    return decorator


class BackwardCompatibilityHandler:
    """하위 호환성 처리 클래스"""

    def __init__(self):
        self.transformations: Dict[str, Callable] = {}

    def register_transformation(self, from_version: str, to_version: str, transformer: Callable):
        """데이터 변환 함수 등록"""
        key = f"{from_version}->{to_version}"
        self.transformations[key] = transformer

    def transform_request(self, data: Dict, from_version: str, to_version: str) -> Dict:
        """요청 데이터 변환"""
        key = f"{from_version}->{to_version}"
        if key in self.transformations:
            return self.transformations[key](data)
        return data

    def transform_response(self, data: Dict, from_version: str, to_version: str) -> Dict:
        """응답 데이터 변환"""
        # 역방향 변환
        key = f"{to_version}->{from_version}"
        if key in self.transformations:
            return self.transformations[key](data)
        return data


# 전역 호환성 핸들러
compatibility_handler = BackwardCompatibilityHandler()


def with_backward_compatibility(target_version: str):
    """하위 호환성 적용 데코레이터"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            requested_version = extract_version_from_request()

            # 요청 데이터 변환 (구버전 -> 신버전)
            if hasattr(request, 'json') and request.json:
                transformed_request = compatibility_handler.transform_request(
                    request.json, requested_version, target_version
                )
                # 변환된 데이터로 요청 처리
                # 실제로는 request.json을 직접 수정할 수 없으므로 다른 방법 필요

            # 원본 함수 실행
            result = func(*args, **kwargs)

            # 응답 데이터 변환 (신버전 -> 구버전)
            if requested_version != target_version:
                # Flask Response 객체에서 JSON 추출 후 변환
                if hasattr(result, 'get_json'):
                    response_data = result.get_json()
                    if response_data and 'data' in response_data:
                        transformed_data = compatibility_handler.transform_response(
                            response_data['data'], target_version, requested_version
                        )
                        response_data['data'] = transformed_data
                        # 새로운 응답 생성
                        from flask import jsonify
                        result = jsonify(response_data)
                        result.status_code = result.status_code or 200

            return result
        return wrapper
    return decorator


def create_api_version_routes(app):
    """API 버전 관리 라우트 생성"""

    @app.route('/api/versions', methods=['GET'])
    def list_api_versions():
        """
        모든 API 버전 목록 조회
        ---
        tags:
          - API Versioning
        summary: List all API versions
        description: Get information about all available API versions
        responses:
          200:
            description: List of API versions
        """
        versions_info = []

        for version, info in version_manager.versions.items():
            versions_info.append({
                "version": version,
                "description": info.get("description", ""),
                "deprecated": info.get("deprecated", False),
                "sunset_date": info.get("sunset_date"),
                "url_prefix": f"/api/{version}",
                "is_default": version == version_manager.default_version
            })

        return APIResponse.success(
            data={
                "versions": versions_info,
                "default_version": version_manager.default_version,
                "total_versions": len(versions_info)
            },
            message="Available API versions retrieved"
        )

    @app.route('/api/version/current', methods=['GET'])
    def current_api_version():
        """
        현재 요청의 API 버전 정보
        ---
        tags:
          - API Versioning
        summary: Get current API version
        description: Get information about the API version being used for this request
        responses:
          200:
            description: Current API version information
        """
        current_version = extract_version_from_request()
        version_info = version_manager.get_version_info(current_version)

        return APIResponse.success(
            data={
                "current_version": current_version,
                "version_info": version_info,
                "extraction_method": "URL path",  # 실제로는 어떤 방법으로 추출했는지
                "is_deprecated": version_manager.is_version_deprecated(current_version)
            },
            message=f"Current API version: {current_version}"
        )


# 예제 버전 변환 함수들
def transform_v1_to_v2_project(data: Dict) -> Dict:
    """프로젝트 데이터 v1 -> v2 변환 예제"""
    transformed = data.copy()

    # v2에서는 'status' 필드가 'state'로 변경됨
    if 'status' in transformed:
        transformed['state'] = transformed.pop('status')

    # v2에서는 날짜 형식이 변경됨
    if 'created_at' in transformed:
        # ISO 형식으로 변환
        pass

    return transformed


def transform_v2_to_v1_project(data: Dict) -> Dict:
    """프로젝트 데이터 v2 -> v1 변환 예제"""
    transformed = data.copy()

    # v1에서는 'state' 필드가 'status'였음
    if 'state' in transformed:
        transformed['status'] = transformed.pop('state')

    return transformed


# 변환 함수 등록
compatibility_handler.register_transformation('v1', 'v2', transform_v1_to_v2_project)
compatibility_handler.register_transformation('v2', 'v1', transform_v2_to_v1_project)