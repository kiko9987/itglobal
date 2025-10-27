"""
자동 OpenAPI 스펙 생성기
라우트 분석기와 스키마 분석기의 결과를 조합하여 완전한 OpenAPI 스펙 생성
"""

import re
from typing import Dict, List, Optional, Any, Set
from datetime import datetime
from flask import Flask

from ..analyzers.route_analyzer import RouteAnalyzer, RouteInfo
from ..analyzers.schema_analyzer import SchemaAnalyzer, SchemaInfo


class OpenAPIGenerator:
    """OpenAPI 3.0 스펙을 자동으로 생성하는 클래스"""

    def __init__(self, app: Flask):
        self.app = app
        self.route_analyzer = RouteAnalyzer(app)
        self.schema_analyzer = SchemaAnalyzer()
        self.openapi_spec: Dict[str, Any] = {}
        self.generated_schemas: Set[str] = set()

    def generate_complete_spec(self) -> Dict[str, Any]:
        """완전한 OpenAPI 스펙 생성"""
        # 1. 라우트 분석
        routes = self.route_analyzer.analyze_all_routes()

        # 2. 스키마 분석 (validation 모듈에서)
        self._analyze_schemas_from_modules()

        # 3. OpenAPI 스펙 기본 구조 생성
        self.openapi_spec = self._create_base_spec()

        # 4. 경로별 스펙 생성
        self._generate_paths_spec(routes)

        # 5. 컴포넌트 스펙 생성
        self._generate_components_spec()

        # 6. 태그 정의 생성
        self._generate_tags_spec(routes)

        return self.openapi_spec

    def _create_base_spec(self) -> Dict[str, Any]:
        """OpenAPI 기본 구조 생성"""
        return {
            "openapi": "3.0.3",
            "info": {
                "title": "IT Global Dashboard API",
                "description": self._generate_api_description(),
                "version": "1.0.0",
                "contact": {
                    "name": "IT Global Development Team",
                    "email": "dev@itglobal.com"
                },
                "license": {
                    "name": "MIT",
                    "url": "https://opensource.org/licenses/MIT"
                }
            },
            "servers": [
                {
                    "url": "http://localhost:5000",
                    "description": "Development server"
                },
                {
                    "url": "https://dashboard.itglobal.com",
                    "description": "Production server"
                }
            ],
            "paths": {},
            "components": {
                "schemas": {},
                "responses": {},
                "parameters": {},
                "securitySchemes": {
                    "bearerAuth": {
                        "type": "http",
                        "scheme": "bearer",
                        "bearerFormat": "JWT"
                    },
                    "sessionAuth": {
                        "type": "apiKey",
                        "in": "cookie",
                        "name": "session"
                    }
                }
            },
            "tags": []
        }

    def _generate_api_description(self) -> str:
        """API 전체 설명 생성"""
        return """
IT Global 비즈니스 관리 대시보드를 위한 종합 API입니다.

## 주요 기능
- 프로젝트 관리 및 추적
- 매출 및 재무 분석
- 실시간 모니터링 및 메트릭
- 사용자 관리 및 인증
- 캐시 및 시스템 관리

## 인증 방식
이 API는 여러 인증 방법을 지원합니다:
- 세션 기반 인증 (웹 인터페이스)
- Bearer 토큰 인증 (API 접근)
- Google OAuth 2.0 (웹 인터페이스)

## 표준 응답 형식
모든 API 응답은 다음과 같은 표준 형식을 따릅니다:
```json
{
  "success": true|false,
  "data": any,
  "error": {...},
  "meta": {...}
}
```

## 버전 관리
API는 URL 경로 기반 버전 관리를 사용합니다 (/api/v1/, /api/v2/ 등).
        """.strip()

    def _analyze_schemas_from_modules(self):
        """애플리케이션 모듈에서 스키마 분석"""
        try:
            # validation 모듈에서 스키마 분석
            from ..api.validation_simple import (
                CreateProjectSchema, PatchProjectSchema,
                PaginationSchema, CreateUserSchema
            )

            schemas_to_analyze = [
                CreateProjectSchema,
                PatchProjectSchema,
                PaginationSchema,
                CreateUserSchema
            ]

            for schema_class in schemas_to_analyze:
                try:
                    self.schema_analyzer.analyze_schema(schema_class)
                except Exception as e:
                    print(f"스키마 분석 오류 {schema_class.__name__}: {str(e)}")

        except ImportError as e:
            print(f"스키마 모듈 import 오류: {str(e)}")

    def _generate_paths_spec(self, routes: List[RouteInfo]):
        """모든 라우트에 대한 경로 스펙 생성"""
        paths = {}

        # API 라우트만 필터링
        api_routes = [r for r in routes if r.rule.startswith('/api/')]

        for route in api_routes:
            try:
                path_spec = self._generate_path_spec(route)
                if path_spec:
                    if route.rule not in paths:
                        paths[route.rule] = {}

                    # HTTP 메서드별로 스펙 추가
                    for method in route.methods:
                        method_lower = method.lower()
                        if method_lower not in ['head', 'options']:
                            paths[route.rule][method_lower] = path_spec

            except Exception as e:
                print(f"경로 스펙 생성 오류 {route.rule}: {str(e)}")

        self.openapi_spec["paths"] = paths

    def _generate_path_spec(self, route: RouteInfo) -> Dict[str, Any]:
        """개별 경로의 OpenAPI 스펙 생성"""
        spec = {
            "operationId": f"{route.endpoint}",
            "summary": route.summary or self._generate_auto_summary(route),
            "description": route.description or self._generate_auto_description(route),
            "tags": route.tags or self._infer_tags(route),
        }

        # 파라미터 추가
        if route.parameters:
            spec["parameters"] = self._format_parameters(route.parameters)

        # 요청 본문 스키마 추출
        request_schema = self._extract_request_schema(route)
        if request_schema:
            spec["requestBody"] = {
                "required": True,
                "content": {
                    "application/json": {
                        "schema": request_schema
                    }
                }
            }

        # 응답 스펙 생성
        spec["responses"] = self._generate_responses_spec(route)

        # 보안 요구사항
        if self._requires_authentication(route):
            spec["security"] = [
                {"bearerAuth": []},
                {"sessionAuth": []}
            ]

        # Deprecated 표시
        if route.deprecated:
            spec["deprecated"] = True

        return spec

    def _generate_auto_summary(self, route: RouteInfo) -> str:
        """라우트 정보를 기반으로 자동 요약 생성"""
        # URL과 메서드를 기반으로 의미있는 요약 생성
        path_parts = route.rule.split('/')

        if 'projects' in path_parts:
            if 'POST' in route.methods:
                return "프로젝트 생성"
            elif 'GET' in route.methods and '<' in route.rule:
                return "프로젝트 상세 조회"
            elif 'GET' in route.methods:
                return "프로젝트 목록 조회"
            elif 'PUT' in route.methods or 'PATCH' in route.methods:
                return "프로젝트 업데이트"
            elif 'DELETE' in route.methods:
                return "프로젝트 삭제"

        elif 'users' in path_parts:
            if 'POST' in route.methods:
                return "사용자 생성"
            elif 'GET' in route.methods:
                return "사용자 조회"

        elif 'health' in path_parts:
            return "시스템 상태 확인"

        # 기본 패턴
        return f"{route.endpoint} 엔드포인트"

    def _generate_auto_description(self, route: RouteInfo) -> str:
        """라우트 정보를 기반으로 자동 설명 생성"""
        if route.description:
            return route.description

        # 기본 설명 패턴
        method_descriptions = {
            'GET': '데이터를 조회합니다',
            'POST': '새로운 데이터를 생성합니다',
            'PUT': '데이터를 업데이트합니다',
            'PATCH': '데이터를 부분적으로 업데이트합니다',
            'DELETE': '데이터를 삭제합니다'
        }

        primary_method = route.methods[0] if route.methods else 'GET'
        base_description = method_descriptions.get(primary_method, '작업을 수행합니다')

        return f"{route.rule} 경로에서 {base_description}."

    def _infer_tags(self, route: RouteInfo) -> List[str]:
        """라우트에서 태그 자동 추론"""
        if route.tags:
            return route.tags

        path_parts = route.rule.split('/')

        # 경로 기반 태그 추론
        if 'projects' in path_parts:
            return ['Projects']
        elif 'users' in path_parts:
            return ['Users']
        elif 'monitoring' in path_parts or 'metrics' in path_parts:
            return ['Monitoring']
        elif 'test' in path_parts:
            return ['Test']
        elif route.version:
            return [f'API {route.version.upper()}']

        return ['General']

    def _extract_request_schema(self, route: RouteInfo) -> Optional[Dict[str, Any]]:
        """라우트에서 요청 스키마 추출"""
        # 데코레이터에서 스키마 정보 추출 시도
        if hasattr(route.function, '__wrapped__'):
            func = route.function
            while hasattr(func, '__wrapped__'):
                if hasattr(func, '_validation_schema'):
                    schema_name = func._validation_schema
                    if schema_name in self.schema_analyzer.schemas_info:
                        return {"$ref": f"#/components/schemas/{schema_name}"}
                func = func.__wrapped__

        # 함수명 기반 스키마 추론
        if 'create' in route.endpoint.lower():
            if 'project' in route.endpoint.lower():
                return {"$ref": "#/components/schemas/CreateProjectSchema"}
            elif 'user' in route.endpoint.lower():
                return {"$ref": "#/components/schemas/CreateUserSchema"}

        return None

    def _format_parameters(self, parameters: List[Dict]) -> List[Dict]:
        """파라미터를 OpenAPI 형식으로 포맷"""
        formatted = []

        for param in parameters:
            formatted_param = {
                "name": param["name"],
                "in": param["in"],
                "required": param["required"],
                "schema": {
                    "type": param["type"]
                }
            }

            if "description" in param:
                formatted_param["description"] = param["description"]

            if "format" in param:
                formatted_param["schema"]["format"] = param["format"]

            formatted.append(formatted_param)

        return formatted

    def _generate_responses_spec(self, route: RouteInfo) -> Dict[str, Any]:
        """응답 스펙 생성"""
        responses = {
            "200": {
                "description": "성공",
                "content": {
                    "application/json": {
                        "schema": {"$ref": "#/components/schemas/StandardResponse"}
                    }
                }
            }
        }

        # HTTP 메서드별 추가 응답
        if 'POST' in route.methods:
            responses["201"] = {
                "description": "생성됨",
                "content": {
                    "application/json": {
                        "schema": {"$ref": "#/components/schemas/StandardResponse"}
                    }
                }
            }

        # 공통 에러 응답
        responses.update({
            "400": {
                "description": "잘못된 요청",
                "content": {
                    "application/json": {
                        "schema": {"$ref": "#/components/schemas/ErrorResponse"}
                    }
                }
            },
            "404": {
                "description": "찾을 수 없음",
                "content": {
                    "application/json": {
                        "schema": {"$ref": "#/components/schemas/ErrorResponse"}
                    }
                }
            },
            "500": {
                "description": "서버 오류",
                "content": {
                    "application/json": {
                        "schema": {"$ref": "#/components/schemas/ErrorResponse"}
                    }
                }
            }
        })

        # 인증이 필요한 경우
        if self._requires_authentication(route):
            responses["401"] = {
                "description": "인증 필요",
                "content": {
                    "application/json": {
                        "schema": {"$ref": "#/components/schemas/ErrorResponse"}
                    }
                }
            }
            responses["403"] = {
                "description": "권한 없음",
                "content": {
                    "application/json": {
                        "schema": {"$ref": "#/components/schemas/ErrorResponse"}
                    }
                }
            }

        return responses

    def _requires_authentication(self, route: RouteInfo) -> bool:
        """인증이 필요한 라우트인지 확인"""
        # 테스트나 public 엔드포인트는 인증 불필요
        if ('/test/' in route.rule or
            '/health' in route.rule or
            '/versions' in route.rule):
            return False

        # API 라우트는 기본적으로 인증 필요
        return route.rule.startswith('/api/')

    def _generate_components_spec(self):
        """컴포넌트 스펙 생성"""
        components = self.openapi_spec["components"]

        # 스키마 컴포넌트 추가
        schema_components = self.schema_analyzer.get_openapi_components()
        components["schemas"].update(schema_components["schemas"])

        # 표준 응답 스키마 추가
        self._add_standard_response_schemas(components)

        # 공통 파라미터 추가
        self._add_common_parameters(components)

    def _add_standard_response_schemas(self, components: Dict[str, Any]):
        """표준 응답 스키마 추가"""
        components["schemas"]["StandardResponse"] = {
            "type": "object",
            "properties": {
                "success": {
                    "type": "boolean",
                    "description": "요청 성공 여부"
                },
                "data": {
                    "description": "응답 데이터"
                },
                "error": {
                    "anyOf": [
                        {"$ref": "#/components/schemas/ErrorDetails"},
                        {"type": "null"}
                    ]
                },
                "meta": {
                    "$ref": "#/components/schemas/ResponseMeta"
                }
            },
            "required": ["success", "data", "error", "meta"]
        }

        components["schemas"]["ErrorResponse"] = {
            "allOf": [
                {"$ref": "#/components/schemas/StandardResponse"},
                {
                    "type": "object",
                    "properties": {
                        "success": {
                            "type": "boolean",
                            "example": False
                        },
                        "data": {
                            "type": "null"
                        }
                    }
                }
            ]
        }

        components["schemas"]["ErrorDetails"] = {
            "type": "object",
            "properties": {
                "code": {
                    "type": "string",
                    "description": "에러 코드"
                },
                "message": {
                    "type": "string",
                    "description": "에러 메시지"
                },
                "details": {
                    "type": "object",
                    "description": "추가 에러 정보"
                }
            },
            "required": ["code", "message"]
        }

        components["schemas"]["ResponseMeta"] = {
            "type": "object",
            "properties": {
                "timestamp": {
                    "type": "string",
                    "format": "date-time",
                    "description": "응답 시간"
                },
                "version": {
                    "type": "string",
                    "description": "API 버전"
                },
                "request_id": {
                    "type": "string",
                    "format": "uuid",
                    "description": "요청 ID"
                }
            },
            "required": ["timestamp", "version", "request_id"]
        }

    def _add_common_parameters(self, components: Dict[str, Any]):
        """공통 파라미터 추가"""
        components["parameters"]["PageParam"] = {
            "name": "page",
            "in": "query",
            "description": "페이지 번호",
            "schema": {
                "type": "integer",
                "minimum": 1,
                "default": 1
            }
        }

        components["parameters"]["LimitParam"] = {
            "name": "limit",
            "in": "query",
            "description": "페이지당 항목 수",
            "schema": {
                "type": "integer",
                "minimum": 1,
                "maximum": 100,
                "default": 20
            }
        }

    def _generate_tags_spec(self, routes: List[RouteInfo]):
        """태그 정의 생성"""
        tags_info = {}

        for route in routes:
            for tag in self._infer_tags(route):
                if tag not in tags_info:
                    tags_info[tag] = {
                        "name": tag,
                        "description": self._get_tag_description(tag)
                    }

        self.openapi_spec["tags"] = list(tags_info.values())

    def _get_tag_description(self, tag: str) -> str:
        """태그별 설명 반환"""
        descriptions = {
            "Projects": "프로젝트 관리 관련 API",
            "Users": "사용자 관리 관련 API",
            "Monitoring": "시스템 모니터링 관련 API",
            "Test": "테스트 및 데모용 API",
            "Analytics": "분석 및 리포트 관련 API",
            "System": "시스템 관리 관련 API",
            "General": "기타 API"
        }
        return descriptions.get(tag, f"{tag} 관련 API")

    def save_spec_to_file(self, file_path: str):
        """OpenAPI 스펙을 파일로 저장"""
        import json

        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(self.openapi_spec, f, ensure_ascii=False, indent=2)

    def get_generation_stats(self) -> Dict[str, Any]:
        """생성 통계 반환"""
        return {
            "total_paths": len(self.openapi_spec.get("paths", {})),
            "total_schemas": len(self.openapi_spec.get("components", {}).get("schemas", {})),
            "total_tags": len(self.openapi_spec.get("tags", [])),
            "generation_timestamp": datetime.utcnow().isoformat(),
            "route_analysis": self.route_analyzer.get_routes_summary(),
            "schema_analysis": self.schema_analyzer.export_analysis_report()
        }


# 사용 예시
if __name__ == "__main__":
    from flask import Flask

    app = Flask(__name__)

    @app.route('/api/v1/test')
    def test_endpoint():
        """테스트 엔드포인트"""
        return {'message': 'test'}

    generator = OpenAPIGenerator(app)
    spec = generator.generate_complete_spec()

    print("생성된 OpenAPI 스펙:")
    import json
    print(json.dumps(spec, ensure_ascii=False, indent=2))