"""
Flask 라우트 자동 분석기
애플리케이션의 모든 라우트를 분석하여 API 문서 생성을 위한 메타데이터 추출
"""

import inspect
import re
from typing import Dict, List, Optional, Any, Tuple
from flask import Flask
from dataclasses import dataclass
from datetime import datetime


@dataclass
class RouteInfo:
    """라우트 정보를 담는 데이터 클래스"""
    rule: str                    # URL 경로 (/api/v1/projects)
    methods: List[str]           # HTTP 메서드 ['GET', 'POST']
    endpoint: str               # 엔드포인트 명 (function name)
    function: callable          # 실제 함수 객체
    module: str                 # 모듈명 (api.v1.projects)
    docstring: Optional[str]    # 함수 docstring
    parameters: List[Dict]      # URL 파라미터 정보
    decorators: List[str]       # 적용된 데코레이터들
    version: Optional[str]      # API 버전 (v1, v2 등)
    tags: List[str]            # 문서화 태그
    deprecated: bool           # deprecated 여부
    summary: Optional[str]     # 간단한 설명
    description: Optional[str] # 상세 설명


class RouteAnalyzer:
    """Flask 애플리케이션의 라우트를 분석하는 클래스"""

    def __init__(self, app: Flask):
        self.app = app
        self.routes_info: List[RouteInfo] = []

    def analyze_all_routes(self) -> List[RouteInfo]:
        """모든 라우트를 분석하여 RouteInfo 리스트 반환"""
        self.routes_info = []

        for rule in self.app.url_map.iter_rules():
            # OPTIONS 메서드는 제외 (CORS 등에서 자동 생성)
            methods = [m for m in rule.methods if m not in ['HEAD', 'OPTIONS']]

            if not methods:
                continue

            # 엔드포인트 함수 가져오기
            try:
                endpoint_func = self.app.view_functions[rule.endpoint]
            except KeyError:
                continue

            # 라우트 정보 추출
            route_info = self._extract_route_info(rule, methods, endpoint_func)

            if route_info:
                self.routes_info.append(route_info)

        return self.routes_info

    def _extract_route_info(self, rule, methods: List[str], func: callable) -> Optional[RouteInfo]:
        """개별 라우트의 상세 정보 추출"""
        try:
            # 모듈 정보 추출
            module_name = getattr(func, '__module__', 'unknown')

            # Docstring 분석
            docstring = inspect.getdoc(func) or ""
            summary, description, tags = self._parse_docstring(docstring)

            # URL 파라미터 분석
            parameters = self._extract_url_parameters(rule)

            # 데코레이터 분석
            decorators = self._extract_decorators(func)

            # API 버전 추출
            version = self._extract_api_version(rule.rule)

            # Deprecated 체크
            deprecated = self._check_deprecated(func, docstring)

            route_info = RouteInfo(
                rule=rule.rule,
                methods=methods,
                endpoint=rule.endpoint,
                function=func,
                module=module_name,
                docstring=docstring,
                parameters=parameters,
                decorators=decorators,
                version=version,
                tags=tags,
                deprecated=deprecated,
                summary=summary,
                description=description
            )

            return route_info

        except Exception as e:
            print(f"라우트 분석 중 오류: {rule.rule} - {str(e)}")
            return None

    def _parse_docstring(self, docstring: str) -> Tuple[Optional[str], Optional[str], List[str]]:
        """Docstring을 파싱하여 summary, description, tags 추출"""
        if not docstring:
            return None, None, []

        lines = docstring.strip().split('\n')

        # 첫 번째 줄을 summary로 사용
        summary = lines[0].strip() if lines else None

        # description과 tags 추출
        description_lines = []
        tags = []
        in_description = False

        for line in lines[1:]:
            line = line.strip()

            # YAML front matter 스타일의 tags 파싱
            if line.startswith('tags:'):
                continue
            elif line.startswith('- ') and any('tags:' in prev_line for prev_line in lines):
                tag = line[2:].strip()
                if tag:
                    tags.append(tag)
            elif line.startswith('---'):
                # Swagger 스타일 docstring 구분자
                in_description = False
            elif line and not line.startswith('---'):
                if not in_description and line:
                    in_description = True
                if in_description:
                    description_lines.append(line)

        description = '\n'.join(description_lines).strip() if description_lines else None

        return summary, description, tags

    def _extract_url_parameters(self, rule) -> List[Dict]:
        """URL 경로에서 파라미터 정보 추출"""
        parameters = []

        for arg in rule.arguments:
            param_info = {
                'name': arg,
                'in': 'path',
                'required': True,
                'type': 'string',  # 기본값, 나중에 타입 힌트로 개선 가능
                'description': f'{arg} 파라미터'
            }

            # URL 컨버터 정보 활용하여 타입 결정
            if rule._converters and arg in rule._converters:
                converter = rule._converters[arg]
                if hasattr(converter, '__class__'):
                    converter_name = converter.__class__.__name__
                    if 'Int' in converter_name:
                        param_info['type'] = 'integer'
                    elif 'Float' in converter_name:
                        param_info['type'] = 'number'
                    elif 'UUID' in converter_name:
                        param_info['type'] = 'string'
                        param_info['format'] = 'uuid'

            parameters.append(param_info)

        return parameters

    def _extract_decorators(self, func: callable) -> List[str]:
        """함수에 적용된 데코레이터 추출"""
        decorators = []

        # 일반적인 데코레이터 패턴 확인
        if hasattr(func, '__wrapped__'):
            decorators.append('decorator_wrapped')

        # 커스텀 속성으로 저장된 데코레이터 정보 확인
        if hasattr(func, '_api_version'):
            decorators.append(f'versioned_route({func._api_version})')

        if hasattr(func, '_validation_schema'):
            decorators.append('validation_decorator')

        # 함수명으로 추정 가능한 데코레이터들
        func_name = getattr(func, '__name__', '')
        if 'wrapper' in func_name or 'decorated' in func_name:
            decorators.append('generic_decorator')

        return decorators

    def _extract_api_version(self, rule_path: str) -> Optional[str]:
        """URL 경로에서 API 버전 추출"""
        # /api/v1/projects -> v1
        version_match = re.search(r'/api/(v\d+)/', rule_path)
        if version_match:
            return version_match.group(1)

        # Legacy 패턴이나 다른 버전 패턴 체크 가능
        return None

    def _check_deprecated(self, func: callable, docstring: str) -> bool:
        """함수나 docstring에서 deprecated 여부 확인"""
        # Docstring에서 deprecated 키워드 확인
        if docstring and 'deprecated' in docstring.lower():
            return True

        # 데코레이터로 deprecated 표시된 경우
        if hasattr(func, '_deprecated'):
            return getattr(func, '_deprecated', False)

        return False

    def get_routes_by_version(self, version: str) -> List[RouteInfo]:
        """특정 버전의 라우트만 필터링"""
        return [route for route in self.routes_info if route.version == version]

    def get_routes_by_tags(self, tags: List[str]) -> List[RouteInfo]:
        """특정 태그를 가진 라우트만 필터링"""
        return [
            route for route in self.routes_info
            if any(tag in route.tags for tag in tags)
        ]

    def get_api_routes_only(self) -> List[RouteInfo]:
        """API 라우트만 필터링 (/api/ 경로)"""
        return [
            route for route in self.routes_info
            if route.rule.startswith('/api/')
        ]

    def get_routes_summary(self) -> Dict[str, Any]:
        """라우트 분석 요약 정보"""
        total_routes = len(self.routes_info)
        api_routes = len(self.get_api_routes_only())

        # 버전별 통계
        version_stats = {}
        for route in self.routes_info:
            version = route.version or 'unversioned'
            version_stats[version] = version_stats.get(version, 0) + 1

        # 메서드별 통계
        method_stats = {}
        for route in self.routes_info:
            for method in route.methods:
                method_stats[method] = method_stats.get(method, 0) + 1

        # 태그별 통계
        tag_stats = {}
        for route in self.routes_info:
            for tag in route.tags:
                tag_stats[tag] = tag_stats.get(tag, 0) + 1

        return {
            'total_routes': total_routes,
            'api_routes': api_routes,
            'version_stats': version_stats,
            'method_stats': method_stats,
            'tag_stats': tag_stats,
            'deprecated_routes': len([r for r in self.routes_info if r.deprecated]),
            'analysis_timestamp': datetime.utcnow().isoformat()
        }

    def export_routes_info(self) -> List[Dict]:
        """RouteInfo를 JSON 직렬화 가능한 형태로 변환"""
        routes_data = []

        for route in self.routes_info:
            route_data = {
                'rule': route.rule,
                'methods': route.methods,
                'endpoint': route.endpoint,
                'module': route.module,
                'docstring': route.docstring,
                'parameters': route.parameters,
                'decorators': route.decorators,
                'version': route.version,
                'tags': route.tags,
                'deprecated': route.deprecated,
                'summary': route.summary,
                'description': route.description,
                'function_name': getattr(route.function, '__name__', 'unknown')
            }
            routes_data.append(route_data)

        return routes_data


def analyze_flask_app(app: Flask) -> RouteAnalyzer:
    """Flask 앱을 분석하고 RouteAnalyzer 인스턴스 반환"""
    analyzer = RouteAnalyzer(app)
    analyzer.analyze_all_routes()
    return analyzer


# 사용 예시
if __name__ == "__main__":
    # 테스트용 간단한 사용법
    from flask import Flask

    app = Flask(__name__)

    @app.route('/test')
    def test_route():
        """
        테스트 라우트
        ---
        tags:
          - Test
        summary: 테스트 엔드포인트
        description: 간단한 테스트를 위한 엔드포인트입니다
        """
        return {'message': 'test'}

    analyzer = analyze_flask_app(app)
    print(f"분석된 라우트 수: {len(analyzer.routes_info)}")

    for route in analyzer.routes_info:
        print(f"- {route.rule} [{', '.join(route.methods)}] - {route.summary}")