"""
자동 문서 생성 시스템 통합 관리자
Flask 애플리케이션과 완전 통합된 실시간 문서 생성 및 관리 시스템
"""

import os
import json
import yaml
import logging
from datetime import datetime
from typing import Dict, List, Optional, Any
from pathlib import Path

from flask import Flask, current_app
from .generators.openapi_generator import OpenAPIGenerator
from .analyzers.route_analyzer import RouteAnalyzer
from .analyzers.schema_analyzer import SchemaAnalyzer

logger = logging.getLogger(__name__)


class DocumentationManager:
    """문서 생성 및 관리를 위한 중앙 관리자"""

    def __init__(self, app: Optional[Flask] = None):
        self.app = app
        self.docs_dir = Path(__file__).parent
        self.output_dir = self.docs_dir / 'generated'
        self.config = {}

        if app is not None:
            self.init_app(app)

    def init_app(self, app: Flask):
        """Flask 앱 초기화"""
        self.app = app

        # 설정값 로드
        self.config = {
            'AUTO_DOCS_ENABLED': app.config.get('AUTO_DOCS_ENABLED', True),
            'DOCS_UPDATE_INTERVAL': app.config.get('DOCS_UPDATE_INTERVAL', 300),  # 5분
            'DOCS_OUTPUT_DIR': app.config.get('DOCS_OUTPUT_DIR', str(self.output_dir)),
            'GENERATE_MARKDOWN': app.config.get('GENERATE_MARKDOWN', True),
            'GENERATE_POSTMAN': app.config.get('GENERATE_POSTMAN', True),
            'VALIDATE_DOCS': app.config.get('VALIDATE_DOCS', True),
        }

        # 출력 디렉터리 생성
        self.output_dir = Path(self.config['DOCS_OUTPUT_DIR'])
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # CLI 명령어 등록
        self._register_cli_commands(app)

        # 개발 모드에서 자동 업데이트 활성화
        if app.debug and self.config['AUTO_DOCS_ENABLED']:
            self._setup_auto_update()

    def _register_cli_commands(self, app: Flask):
        """Flask CLI 명령어 등록"""

        @app.cli.command('generate-docs')
        def generate_docs_command():
            """API 문서 생성 명령어"""
            logger.info("API 문서 생성을 시작합니다...")

            try:
                result = self.generate_all_documentation()

                logger.info("\n[SUCCESS] 문서 생성 완료!")
                logger.info("📁 출력 디렉터리: %s", self.output_dir)
                logger.info("[STATS] 생성된 파일:")

                for format_type, file_path in result['generated_files'].items():
                    logger.info("   - %s: %s", format_type, file_path)

                logger.info("\n📈 통계:")
                stats = result['stats']
                logger.info("   - 총 경로 수: %d", stats['total_paths'])
                logger.info("   - 총 스키마 수: %d", stats['total_schemas'])
                logger.info("   - 태그 수: %d", stats['total_tags'])

            except Exception as e:
                logger.error("[ERROR] 문서 생성 중 오류 발생: %s", str(e))
                return 1

        @app.cli.command('validate-docs')
        def validate_docs_command():
            """문서 검증 명령어"""
            logger.info("API 문서 검증을 시작합니다...")

            try:
                validation_result = self.validate_documentation()

                if validation_result['is_valid']:
                    logger.info("[SUCCESS] 문서 검증 통과!")
                else:
                    logger.error("[ERROR] 문서 검증 실패:")
                    for error in validation_result['errors']:
                        logger.error("   - %s", error)

            except Exception as e:
                logger.error("[ERROR] 문서 검증 중 오류 발생: %s", str(e))
                return 1

        @app.cli.command('docs-stats')
        def docs_stats_command():
            """문서 통계 조회"""
            try:
                stats = self.get_documentation_stats()

                logger.info("[STATS] API 문서 통계:")
                logger.info("   - 총 라우트 수: %d", stats['route_stats']['total_routes'])
                logger.info("   - API 라우트 수: %d", stats['route_stats']['api_routes'])
                logger.info("   - 스키마 수: %d", stats['schema_stats']['total_schemas'])

                logger.info("\n[INFO] 버전별 분포:")
                for version, count in stats['route_stats']['version_stats'].items():
                    logger.info("   - %s: %d개", version, count)

                logger.info("\n🏷️ 태그별 분포:")
                for tag, count in stats['route_stats']['tag_stats'].items():
                    logger.info("   - %s: %d개", tag, count)

            except Exception as e:
                logger.error("[ERROR] 통계 조회 중 오류 발생: %s", str(e))
                return 1

    def _setup_auto_update(self):
        """개발 모드에서 자동 문서 업데이트 설정"""
        try:
            from flask import request

            @self.app.before_request
            def check_docs_update():
                """요청 전 문서 업데이트 체크"""
                # 특정 조건에서만 업데이트 (예: 새로운 라우트 추가 감지)
                if hasattr(request, 'endpoint') and request.endpoint:
                    self._maybe_update_docs()

        except ImportError:
            logger.warning("[WARN] 자동 업데이트를 위한 모듈을 찾을 수 없습니다.")

    def _maybe_update_docs(self):
        """조건부 문서 업데이트"""
        # 캐시된 라우트 수와 현재 라우트 수 비교
        current_routes_count = len(list(self.app.url_map.iter_rules()))
        cache_file = self.output_dir / '.routes_cache'

        last_count = 0
        if cache_file.exists():
            try:
                last_count = int(cache_file.read_text())
            except (ValueError, FileNotFoundError):
                pass

        if current_routes_count != last_count:
            logger.info("[UPDATE] 라우트 변경 감지 (%d -> %d), 문서 업데이트 중...", last_count, current_routes_count)
            self.generate_all_documentation()
            cache_file.write_text(str(current_routes_count))

    def generate_all_documentation(self) -> Dict[str, Any]:
        """모든 형태의 문서 생성"""
        if not self.app:
            raise RuntimeError("Flask 앱이 초기화되지 않았습니다.")

        logger.info("[SEARCH] 애플리케이션 분석 중...")

        # OpenAPI 스펙 생성
        generator = OpenAPIGenerator(self.app)
        openapi_spec = generator.generate_complete_spec()

        # 생성 결과
        result = {
            'generated_files': {},
            'stats': generator.get_generation_stats(),
            'timestamp': datetime.utcnow().isoformat()
        }

        logger.info("[DOCS] 문서 파일 생성 중...")

        # 1. OpenAPI JSON 파일 생성
        json_file = self.output_dir / 'openapi.json'
        self._save_json(openapi_spec, json_file)
        result['generated_files']['OpenAPI JSON'] = str(json_file)

        # 2. OpenAPI YAML 파일 생성
        yaml_file = self.output_dir / 'openapi.yaml'
        self._save_yaml(openapi_spec, yaml_file)
        result['generated_files']['OpenAPI YAML'] = str(yaml_file)

        # 3. Markdown 문서 생성
        if self.config['GENERATE_MARKDOWN']:
            md_file = self.output_dir / 'api_documentation.md'
            self._generate_markdown_docs(openapi_spec, md_file, generator)
            result['generated_files']['Markdown'] = str(md_file)

        # 4. Postman 컬렉션 생성
        if self.config['GENERATE_POSTMAN']:
            postman_file = self.output_dir / 'postman_collection.json'
            self._generate_postman_collection(openapi_spec, postman_file)
            result['generated_files']['Postman Collection'] = str(postman_file)

        # 5. 개발자 가이드 생성
        guide_file = self.output_dir / 'developer_guide.md'
        self._generate_developer_guide(openapi_spec, guide_file, generator)
        result['generated_files']['Developer Guide'] = str(guide_file)

        # 6. 메타데이터 파일 생성
        meta_file = self.output_dir / 'generation_metadata.json'
        self._save_json(result, meta_file)
        result['generated_files']['Metadata'] = str(meta_file)

        logger.info("[SUCCESS] 모든 문서 생성 완료!")
        return result

    def _save_json(self, data: Dict, file_path: Path):
        """JSON 파일 저장"""
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _save_yaml(self, data: Dict, file_path: Path):
        """YAML 파일 저장"""
        with open(file_path, 'w', encoding='utf-8') as f:
            yaml.dump(data, f, allow_unicode=True, default_flow_style=False, indent=2)

    def _generate_markdown_docs(self, openapi_spec: Dict, file_path: Path, generator: OpenAPIGenerator):
        """Markdown 형태의 API 문서 생성"""

        markdown_content = f"""# {openapi_spec['info']['title']}

> **버전**: {openapi_spec['info']['version']}
> **생성일**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

## 개요

{openapi_spec['info']['description']}

## 서버 정보

"""

        for server in openapi_spec.get('servers', []):
            markdown_content += f"- **{server['description']}**: `{server['url']}`\n"

        markdown_content += "\n## 인증\n\n"
        security_schemes = openapi_spec.get('components', {}).get('securitySchemes', {})

        for scheme_name, scheme_info in security_schemes.items():
            markdown_content += f"### {scheme_name}\n"
            markdown_content += f"- **타입**: {scheme_info['type']}\n"
            if 'scheme' in scheme_info:
                markdown_content += f"- **스키마**: {scheme_info['scheme']}\n"
            if 'in' in scheme_info:
                markdown_content += f"- **위치**: {scheme_info['in']}\n"
            markdown_content += "\n"

        # 태그별로 엔드포인트 그룹화
        tags = {tag['name']: tag for tag in openapi_spec.get('tags', [])}
        paths_by_tag = {}

        for path, path_info in openapi_spec.get('paths', {}).items():
            for method, method_info in path_info.items():
                for tag in method_info.get('tags', ['General']):
                    if tag not in paths_by_tag:
                        paths_by_tag[tag] = []
                    paths_by_tag[tag].append({
                        'path': path,
                        'method': method.upper(),
                        'info': method_info
                    })

        # 태그별 섹션 생성
        for tag_name, endpoints in paths_by_tag.items():
            tag_info = tags.get(tag_name, {})

            markdown_content += f"\n## {tag_name}\n\n"
            if tag_info.get('description'):
                markdown_content += f"{tag_info['description']}\n\n"

            for endpoint in endpoints:
                markdown_content += f"### {endpoint['method']} {endpoint['path']}\n\n"

                info = endpoint['info']
                if info.get('summary'):
                    markdown_content += f"**{info['summary']}**\n\n"

                if info.get('description'):
                    markdown_content += f"{info['description']}\n\n"

                # 파라미터
                if info.get('parameters'):
                    markdown_content += "**파라미터:**\n\n"
                    for param in info['parameters']:
                        required = " (필수)" if param.get('required') else " (선택)"
                        markdown_content += f"- `{param['name']}` ({param['in']}){required}: {param.get('description', '')}\n"
                    markdown_content += "\n"

                # 요청 본문
                if info.get('requestBody'):
                    markdown_content += "**요청 본문:**\n\n"
                    markdown_content += "```json\n"
                    markdown_content += "// Content-Type: application/json\n"
                    markdown_content += "{\n  // 요청 데이터 구조\n}\n"
                    markdown_content += "```\n\n"

                # 응답
                if info.get('responses'):
                    markdown_content += "**응답:**\n\n"
                    for status_code, response_info in info['responses'].items():
                        markdown_content += f"- **{status_code}**: {response_info.get('description', '')}\n"
                    markdown_content += "\n"

                markdown_content += "---\n\n"

        # 스키마 정보
        schemas = openapi_spec.get('components', {}).get('schemas', {})
        if schemas:
            markdown_content += "\n## 데이터 스키마\n\n"
            for schema_name, schema_info in schemas.items():
                markdown_content += f"### {schema_name}\n\n"
                if schema_info.get('description'):
                    markdown_content += f"{schema_info['description']}\n\n"

                if schema_info.get('properties'):
                    markdown_content += "**속성:**\n\n"
                    for prop_name, prop_info in schema_info['properties'].items():
                        prop_type = prop_info.get('type', 'unknown')
                        prop_desc = prop_info.get('description', '')
                        required = prop_name in schema_info.get('required', [])
                        req_text = " (필수)" if required else ""
                        markdown_content += f"- `{prop_name}` ({prop_type}){req_text}: {prop_desc}\n"

                markdown_content += "\n"

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

    def _generate_postman_collection(self, openapi_spec: Dict, file_path: Path):
        """Postman 컬렉션 생성"""

        collection = {
            "info": {
                "name": openapi_spec['info']['title'],
                "description": openapi_spec['info']['description'],
                "schema": "https://schema.getpostman.com/json/collection/v2.1.0/collection.json",
                "_postman_id": f"generated-{datetime.utcnow().strftime('%Y%m%d-%H%M%S')}"
            },
            "auth": {
                "type": "bearer",
                "bearer": [
                    {
                        "key": "token",
                        "value": "{{bearer_token}}",
                        "type": "string"
                    }
                ]
            },
            "variable": [
                {
                    "key": "base_url",
                    "value": openapi_spec.get('servers', [{}])[0].get('url', 'http://localhost:5000'),
                    "type": "string"
                },
                {
                    "key": "bearer_token",
                    "value": "your_token_here",
                    "type": "string"
                }
            ],
            "item": []
        }

        # 경로별 요청 생성
        for path, path_info in openapi_spec.get('paths', {}).items():
            for method, method_info in path_info.items():

                # URL에서 파라미터 변환 ({id} -> :id)
                postman_url = path.replace('{', ':').replace('}', '')

                request_item = {
                    "name": method_info.get('summary', f"{method.upper()} {path}"),
                    "request": {
                        "method": method.upper(),
                        "header": [
                            {
                                "key": "Content-Type",
                                "value": "application/json",
                                "type": "text"
                            }
                        ],
                        "url": {
                            "raw": f"{{{{base_url}}}}{postman_url}",
                            "host": ["{{base_url}}"],
                            "path": [p for p in postman_url.split('/') if p]
                        },
                        "description": method_info.get('description', '')
                    }
                }

                # 요청 본문 추가
                if method_info.get('requestBody'):
                    request_item['request']['body'] = {
                        "mode": "raw",
                        "raw": json.dumps({"example": "data"}, indent=2),
                        "options": {
                            "raw": {
                                "language": "json"
                            }
                        }
                    }

                collection['item'].append(request_item)

        self._save_json(collection, file_path)

    def _generate_developer_guide(self, openapi_spec: Dict, file_path: Path, generator: OpenAPIGenerator):
        """개발자 가이드 생성"""

        guide_content = f"""# {openapi_spec['info']['title']} 개발자 가이드

## 시작하기

이 가이드는 {openapi_spec['info']['title']} API를 사용하여 애플리케이션을 개발하는 방법을 설명합니다.

## 빠른 시작

### 1. 인증 설정

```bash
# Bearer 토큰 인증
curl -H "Authorization: Bearer YOUR_TOKEN" \\
     -H "Content-Type: application/json" \\
     {{base_url}}/api/v1/test
```

### 2. 기본 요청 예시

```python
import requests

# API 기본 URL
BASE_URL = "http://localhost:5000"

# 인증 헤더
headers = {{
    "Authorization": "Bearer YOUR_TOKEN",
    "Content-Type": "application/json"
}}

# GET 요청
response = requests.get(f"{{BASE_URL}}/api/v1/projects", headers=headers)
data = response.json()

print(f"성공: {{data['success']}}")
print(f"데이터: {{data['data']}}")
```

## 표준 응답 형식

모든 API 응답은 다음과 같은 표준 형식을 따릅니다:

```json
{{
  "success": true,
  "data": "실제 응답 데이터",
  "error": null,
  "meta": {{
    "timestamp": "2025-01-01T00:00:00Z",
    "version": "v1",
    "request_id": "uuid-string"
  }}
}}
```

### 성공 응답
- `success`: `true`
- `data`: 요청한 데이터 또는 작업 결과
- `error`: `null`

### 오류 응답
- `success`: `false`
- `data`: `null`
- `error`: 오류 상세 정보

## 오류 코드

| HTTP 코드 | 설명 | 해결 방법 |
|-----------|------|-----------|
| 400 | 잘못된 요청 | 요청 파라미터나 본문을 확인하세요 |
| 401 | 인증 필요 | 유효한 인증 토큰을 제공하세요 |
| 403 | 권한 없음 | 해당 리소스에 접근할 권한이 있는지 확인하세요 |
| 404 | 찾을 수 없음 | URL이나 리소스 ID를 확인하세요 |
| 500 | 서버 오류 | 서버 관리자에게 문의하세요 |

## 페이지네이션

목록 조회 API는 페이지네이션을 지원합니다:

```python
# 페이지네이션 파라미터
params = {{
    "page": 1,      # 페이지 번호 (기본: 1)
    "limit": 20     # 페이지당 항목 수 (기본: 20, 최대: 100)
}}

response = requests.get(f"{{BASE_URL}}/api/v1/projects", headers=headers, params=params)
```

## 코드 예시

### 프로젝트 생성
```python
project_data = {{
    "name": "새 프로젝트",
    "description": "프로젝트 설명",
    "start_date": "2025-01-01",
    "budget": 10000.00
}}

response = requests.post(
    f"{{BASE_URL}}/api/v1/projects",
    headers=headers,
    json=project_data
)

if response.json()['success']:
    print("프로젝트 생성 성공!")
else:
    print(f"오류: {{response.json()['error']}}")
```

### 프로젝트 목록 조회
```python
response = requests.get(f"{{BASE_URL}}/api/v1/projects", headers=headers)
projects = response.json()['data']

for project in projects:
    print(f"프로젝트: {{project['name']}} (ID: {{project['id']}})")
```

## SDK 및 클라이언트 라이브러리

### Python 클라이언트
```python
# pip install requests

class ITGlobalAPIClient:
    def __init__(self, base_url, token):
        self.base_url = base_url
        self.headers = {{
            "Authorization": f"Bearer {{token}}",
            "Content-Type": "application/json"
        }}

    def get_projects(self, page=1, limit=20):
        response = requests.get(
            f"{{self.base_url}}/api/v1/projects",
            headers=self.headers,
            params={{"page": page, "limit": limit}}
        )
        return response.json()

    def create_project(self, project_data):
        response = requests.post(
            f"{{self.base_url}}/api/v1/projects",
            headers=self.headers,
            json=project_data
        )
        return response.json()

# 사용법
client = ITGlobalAPIClient("http://localhost:5000", "your_token")
projects = client.get_projects()
```

### JavaScript 클라이언트
```javascript
class ITGlobalAPIClient {{
    constructor(baseUrl, token) {{
        this.baseUrl = baseUrl;
        this.headers = {{
            'Authorization': `Bearer ${{token}}`,
            'Content-Type': 'application/json'
        }};
    }}

    async getProjects(page = 1, limit = 20) {{
        const response = await fetch(
            `${{this.baseUrl}}/api/v1/projects?page=${{page}}&limit=${{limit}}`,
            {{ headers: this.headers }}
        );
        return await response.json();
    }}

    async createProject(projectData) {{
        const response = await fetch(`${{this.baseUrl}}/api/v1/projects`, {{
            method: 'POST',
            headers: this.headers,
            body: JSON.stringify(projectData)
        }});
        return await response.json();
    }}
}}

// 사용법
const client = new ITGlobalAPIClient('http://localhost:5000', 'your_token');
client.getProjects().then(data => console.log(data));
```

## 개발 환경 설정

### 로컬 개발
1. 개발 서버 시작: `http://localhost:5000`
2. API 문서: `http://localhost:5000/docs`
3. 테스트 엔드포인트: `http://localhost:5000/api/v1/test`

### 프로덕션 환경
- URL: `https://dashboard.itglobal.com`
- SSL 인증서 필요
- Rate Limiting 적용

## 문의 및 지원

- 개발팀 이메일: {openapi_spec['info'].get('contact', {}).get('email', 'dev@itglobal.com')}
- API 문서: [Swagger UI](/docs)
- GitHub Issues: [프로젝트 저장소](https://github.com/itglobal/dashboard)

---

*이 가이드는 자동으로 생성되었습니다. 최종 업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*
"""

        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(guide_content)

    def validate_documentation(self) -> Dict[str, Any]:
        """생성된 문서의 유효성 검증"""
        if not self.app:
            raise RuntimeError("Flask 앱이 초기화되지 않았습니다.")

        errors = []
        warnings = []

        # OpenAPI 스펙 검증
        spec_file = self.output_dir / 'openapi.json'
        if not spec_file.exists():
            errors.append("OpenAPI 스펙 파일이 존재하지 않습니다.")
        else:
            try:
                with open(spec_file, 'r', encoding='utf-8') as f:
                    spec = json.load(f)

                # 필수 필드 검증
                required_fields = ['openapi', 'info', 'paths']
                for field in required_fields:
                    if field not in spec:
                        errors.append(f"OpenAPI 스펙에서 필수 필드 '{field}'이 누락되었습니다.")

                # 경로 검증
                actual_routes = set()
                for rule in self.app.url_map.iter_rules():
                    if rule.rule.startswith('/api/'):
                        actual_routes.add(rule.rule)

                documented_routes = set(spec.get('paths', {}).keys())

                # 문서화되지 않은 라우트
                undocumented = actual_routes - documented_routes
                if undocumented:
                    warnings.extend([f"문서화되지 않은 라우트: {route}" for route in undocumented])

                # 존재하지 않는 라우트 문서
                non_existent = documented_routes - actual_routes
                if non_existent:
                    errors.extend([f"존재하지 않는 라우트가 문서화됨: {route}" for route in non_existent])

            except Exception as e:
                errors.append(f"OpenAPI 스펙 파싱 오류: {str(e)}")

        return {
            'is_valid': len(errors) == 0,
            'errors': errors,
            'warnings': warnings,
            'validation_timestamp': datetime.utcnow().isoformat()
        }

    def get_documentation_stats(self) -> Dict[str, Any]:
        """문서 생성 통계 반환"""
        if not self.app:
            raise RuntimeError("Flask 앱이 초기화되지 않았습니다.")

        # 라우트 분석
        route_analyzer = RouteAnalyzer(self.app)
        route_stats = route_analyzer.analyze_all_routes()

        # 스키마 분석
        schema_analyzer = SchemaAnalyzer()
        try:
            from ..api.validation_simple import (
                CreateProjectSchema, PatchProjectSchema,
                PaginationSchema, CreateUserSchema
            )
            for schema in [CreateProjectSchema, PatchProjectSchema, PaginationSchema, CreateUserSchema]:
                schema_analyzer.analyze_schema(schema)
        except ImportError:
            pass

        return {
            'route_stats': route_analyzer.get_routes_summary(),
            'schema_stats': schema_analyzer.export_analysis_report(),
            'file_stats': self._get_file_stats(),
            'last_generation': self._get_last_generation_time()
        }

    def _get_file_stats(self) -> Dict[str, Any]:
        """생성된 파일 통계"""
        stats = {
            'total_files': 0,
            'file_sizes': {},
            'last_modified': {}
        }

        for file_path in self.output_dir.glob('*'):
            if file_path.is_file():
                stats['total_files'] += 1
                stats['file_sizes'][file_path.name] = file_path.stat().st_size
                stats['last_modified'][file_path.name] = datetime.fromtimestamp(
                    file_path.stat().st_mtime
                ).isoformat()

        return stats

    def _get_last_generation_time(self) -> Optional[str]:
        """마지막 생성 시간 반환"""
        meta_file = self.output_dir / 'generation_metadata.json'
        if meta_file.exists():
            try:
                with open(meta_file, 'r', encoding='utf-8') as f:
                    meta = json.load(f)
                    return meta.get('timestamp')
            except Exception:
                pass
        return None


# 전역 인스턴스
docs_manager = DocumentationManager()

def init_documentation(app: Flask):
    """Flask 앱에 문서 시스템 초기화"""
    docs_manager.init_app(app)
    return docs_manager