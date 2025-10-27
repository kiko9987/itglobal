"""
자동 문서 생성 시스템 API 엔드포인트
웹 인터페이스를 통한 실시간 문서 생성 및 관리
"""

import json
from datetime import datetime
from flask import Blueprint, jsonify, request, send_file, current_app
from pathlib import Path
import tempfile
import zipfile

from ..api.responses import APIResponse
from .doc_generator import docs_manager


docs_api = Blueprint('docs_api', __name__, url_prefix='/api/v1/docs')


@docs_api.route('/generate', methods=['POST'])
def generate_documentation():
    """
    API 문서 생성
    ---
    tags:
      - Documentation
    summary: API 문서 자동 생성
    description: 현재 애플리케이션 상태를 분석하여 모든 형태의 API 문서를 생성합니다
    requestBody:
      content:
        application/json:
          schema:
            type: object
            properties:
              formats:
                type: array
                items:
                  type: string
                  enum: [json, yaml, markdown, postman, guide]
                description: 생성할 문서 형식 (미지정시 모든 형식)
              force:
                type: boolean
                description: 강제 재생성 여부
                default: false
    responses:
      200:
        description: 문서 생성 성공
      500:
        description: 문서 생성 실패
    """
    try:
        data = request.get_json() or {}
        requested_formats = data.get('formats', [])
        force_regenerate = data.get('force', False)

        # 기존 문서 존재 여부 확인
        if not force_regenerate:
            meta_file = Path(docs_manager.output_dir) / 'generation_metadata.json'
            if meta_file.exists():
                try:
                    with open(meta_file, 'r', encoding='utf-8') as f:
                        meta = json.load(f)
                        last_gen = datetime.fromisoformat(meta['timestamp'])
                        if (datetime.utcnow() - last_gen).total_seconds() < 60:
                            return APIResponse.success({
                                'message': '최근에 생성된 문서가 존재합니다',
                                'last_generation': meta['timestamp'],
                                'files': meta.get('generated_files', {})
                            })
                except Exception:
                    pass

        # 문서 생성
        result = docs_manager.generate_all_documentation()

        # 요청된 형식 필터링
        if requested_formats:
            filtered_files = {
                k: v for k, v in result['generated_files'].items()
                if any(fmt.lower() in k.lower() for fmt in requested_formats)
            }
            result['generated_files'] = filtered_files

        return APIResponse.success({
            'message': '문서 생성 완료',
            'generated_files': result['generated_files'],
            'stats': result['stats'],
            'generation_time': result['timestamp']
        })

    except Exception as e:
        current_app.logger.error(f"문서 생성 중 오류: {str(e)}")
        return APIResponse.error(
            message="문서 생성 중 오류가 발생했습니다",
            details={'error': str(e)},
            status_code=500
        )


@docs_api.route('/status', methods=['GET'])
def get_documentation_status():
    """
    문서 생성 상태 조회
    ---
    tags:
      - Documentation
    summary: 문서 생성 시스템 상태 확인
    description: 현재 문서 생성 시스템의 상태와 통계 정보를 반환합니다
    responses:
      200:
        description: 상태 조회 성공
    """
    try:
        stats = docs_manager.get_documentation_stats()

        # 출력 디렉터리 상태 확인
        output_dir = Path(docs_manager.output_dir)
        output_exists = output_dir.exists()

        # 생성된 파일 목록
        generated_files = {}
        if output_exists:
            for file_path in output_dir.glob('*'):
                if file_path.is_file():
                    generated_files[file_path.name] = {
                        'size': file_path.stat().st_size,
                        'modified': datetime.fromtimestamp(
                            file_path.stat().st_mtime
                        ).isoformat()
                    }

        return APIResponse.success({
            'system_status': 'active' if docs_manager.config.get('AUTO_DOCS_ENABLED') else 'inactive',
            'output_directory': str(output_dir),
            'output_directory_exists': output_exists,
            'generated_files': generated_files,
            'statistics': stats,
            'config': {
                'auto_docs_enabled': docs_manager.config.get('AUTO_DOCS_ENABLED'),
                'update_interval': docs_manager.config.get('DOCS_UPDATE_INTERVAL'),
                'generate_markdown': docs_manager.config.get('GENERATE_MARKDOWN'),
                'generate_postman': docs_manager.config.get('GENERATE_POSTMAN')
            }
        })

    except Exception as e:
        current_app.logger.error(f"문서 상태 조회 중 오류: {str(e)}")
        return APIResponse.error(
            message="문서 상태 조회 중 오류가 발생했습니다",
            details={'error': str(e)},
            status_code=500
        )


@docs_api.route('/validate', methods=['POST'])
def validate_documentation():
    """
    문서 유효성 검증
    ---
    tags:
      - Documentation
    summary: 생성된 문서의 유효성 검증
    description: 현재 애플리케이션 상태와 생성된 문서의 일치성을 검증합니다
    responses:
      200:
        description: 검증 완료
    """
    try:
        validation_result = docs_manager.validate_documentation()

        return APIResponse.success({
            'validation_result': validation_result,
            'is_valid': validation_result['is_valid'],
            'error_count': len(validation_result.get('errors', [])),
            'warning_count': len(validation_result.get('warnings', []))
        })

    except Exception as e:
        current_app.logger.error(f"문서 검증 중 오류: {str(e)}")
        return APIResponse.error(
            message="문서 검증 중 오류가 발생했습니다",
            details={'error': str(e)},
            status_code=500
        )


@docs_api.route('/download/<file_type>', methods=['GET'])
def download_documentation(file_type):
    """
    문서 파일 다운로드
    ---
    tags:
      - Documentation
    summary: 생성된 문서 파일 다운로드
    description: 지정된 형식의 문서 파일을 다운로드합니다
    parameters:
      - name: file_type
        in: path
        required: true
        schema:
          type: string
          enum: [json, yaml, markdown, postman, guide, all]
    responses:
      200:
        description: 파일 다운로드 성공
      404:
        description: 파일을 찾을 수 없음
    """
    try:
        output_dir = Path(docs_manager.output_dir)

        if file_type == 'all':
            # 모든 파일을 ZIP으로 압축하여 다운로드
            temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')

            with zipfile.ZipFile(temp_zip.name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path in output_dir.glob('*'):
                    if file_path.is_file():
                        zipf.write(file_path, file_path.name)

            return send_file(
                temp_zip.name,
                as_attachment=True,
                download_name=f'api_docs_{datetime.now().strftime("%Y%m%d_%H%M%S")}.zip',
                mimetype='application/zip'
            )

        # 개별 파일 타입별 다운로드
        file_mappings = {
            'json': 'openapi.json',
            'yaml': 'openapi.yaml',
            'markdown': 'api_documentation.md',
            'postman': 'postman_collection.json',
            'guide': 'developer_guide.md'
        }

        filename = file_mappings.get(file_type)
        if not filename:
            return APIResponse.error(
                message="지원하지 않는 파일 형식입니다",
                details={'supported_types': list(file_mappings.keys())},
                status_code=400
            )

        file_path = output_dir / filename
        if not file_path.exists():
            return APIResponse.error(
                message="요청한 문서 파일이 존재하지 않습니다",
                details={'file_path': str(file_path)},
                status_code=404
            )

        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        current_app.logger.error(f"문서 다운로드 중 오류: {str(e)}")
        return APIResponse.error(
            message="문서 다운로드 중 오류가 발생했습니다",
            details={'error': str(e)},
            status_code=500
        )


@docs_api.route('/preview/<file_type>', methods=['GET'])
def preview_documentation(file_type):
    """
    문서 미리보기
    ---
    tags:
      - Documentation
    summary: 생성된 문서 파일 미리보기
    description: 지정된 형식의 문서 내용을 JSON으로 반환합니다
    parameters:
      - name: file_type
        in: path
        required: true
        schema:
          type: string
          enum: [json, yaml, markdown, postman, guide]
    responses:
      200:
        description: 미리보기 성공
      404:
        description: 파일을 찾을 수 없음
    """
    try:
        output_dir = Path(docs_manager.output_dir)

        file_mappings = {
            'json': 'openapi.json',
            'yaml': 'openapi.yaml',
            'markdown': 'api_documentation.md',
            'postman': 'postman_collection.json',
            'guide': 'developer_guide.md'
        }

        filename = file_mappings.get(file_type)
        if not filename:
            return APIResponse.error(
                message="지원하지 않는 파일 형식입니다",
                details={'supported_types': list(file_mappings.keys())},
                status_code=400
            )

        file_path = output_dir / filename
        if not file_path.exists():
            return APIResponse.error(
                message="요청한 문서 파일이 존재하지 않습니다",
                details={'file_path': str(file_path)},
                status_code=404
            )

        # 파일 내용 읽기
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # JSON 파일의 경우 파싱해서 반환
        if filename.endswith('.json'):
            try:
                content = json.loads(content)
            except json.JSONDecodeError:
                pass

        return APIResponse.success({
            'file_type': file_type,
            'filename': filename,
            'content': content,
            'size': file_path.stat().st_size,
            'modified': datetime.fromtimestamp(file_path.stat().st_mtime).isoformat()
        })

    except Exception as e:
        current_app.logger.error(f"문서 미리보기 중 오류: {str(e)}")
        return APIResponse.error(
            message="문서 미리보기 중 오류가 발생했습니다",
            details={'error': str(e)},
            status_code=500
        )


@docs_api.route('/analytics', methods=['GET'])
def get_documentation_analytics():
    """
    문서 분석 정보 조회
    ---
    tags:
      - Documentation
    summary: API 문서 분석 및 통계 정보
    description: 현재 API의 구조 분석 결과와 문서화 통계를 반환합니다
    responses:
      200:
        description: 분석 정보 조회 성공
    """
    try:
        stats = docs_manager.get_documentation_stats()

        # 추가 분석 정보 생성
        route_stats = stats.get('route_stats', {})
        schema_stats = stats.get('schema_stats', {})

        analytics = {
            'api_overview': {
                'total_endpoints': route_stats.get('total_routes', 0),
                'api_endpoints': route_stats.get('api_routes', 0),
                'deprecated_endpoints': route_stats.get('deprecated_routes', 0),
                'coverage_percentage': round(
                    (route_stats.get('api_routes', 0) / max(route_stats.get('total_routes', 1), 1)) * 100, 2
                )
            },
            'version_distribution': route_stats.get('version_stats', {}),
            'method_distribution': route_stats.get('method_stats', {}),
            'tag_distribution': route_stats.get('tag_stats', {}),
            'schema_analysis': {
                'total_schemas': schema_stats.get('total_schemas', 0),
                'schemas_detail': schema_stats.get('schemas_summary', {})
            },
            'documentation_health': {
                'has_openapi_spec': (Path(docs_manager.output_dir) / 'openapi.json').exists(),
                'has_markdown_docs': (Path(docs_manager.output_dir) / 'api_documentation.md').exists(),
                'has_postman_collection': (Path(docs_manager.output_dir) / 'postman_collection.json').exists(),
                'last_update': stats.get('last_generation')
            },
            'recommendations': []
        }

        # 개선 권장 사항 생성
        if analytics['api_overview']['deprecated_endpoints'] > 0:
            analytics['recommendations'].append({
                'type': 'warning',
                'message': f"{analytics['api_overview']['deprecated_endpoints']}개의 deprecated API가 있습니다. 마이그레이션을 고려하세요."
            })

        if analytics['api_overview']['coverage_percentage'] < 50:
            analytics['recommendations'].append({
                'type': 'info',
                'message': "API 엔드포인트 중 일부만 /api/ 경로를 사용하고 있습니다. 표준화를 고려하세요."
            })

        if not analytics['documentation_health']['last_update']:
            analytics['recommendations'].append({
                'type': 'error',
                'message': "문서가 생성되지 않았습니다. 문서 생성을 실행하세요."
            })

        return APIResponse.success(analytics)

    except Exception as e:
        current_app.logger.error(f"문서 분석 정보 조회 중 오류: {str(e)}")
        return APIResponse.error(
            message="문서 분석 정보 조회 중 오류가 발생했습니다",
            details={'error': str(e)},
            status_code=500
        )


@docs_api.route('/config', methods=['GET', 'PUT'])
def manage_documentation_config():
    """
    문서 시스템 설정 관리
    ---
    tags:
      - Documentation
    summary: 문서 생성 시스템 설정 조회/변경
    description: 자동 문서 생성 시스템의 설정을 조회하거나 변경합니다
    """
    if request.method == 'GET':
        return APIResponse.success({
            'current_config': docs_manager.config,
            'available_options': {
                'AUTO_DOCS_ENABLED': 'boolean - 자동 문서 생성 활성화',
                'DOCS_UPDATE_INTERVAL': 'integer - 업데이트 간격 (초)',
                'GENERATE_MARKDOWN': 'boolean - Markdown 문서 생성',
                'GENERATE_POSTMAN': 'boolean - Postman 컬렉션 생성',
                'VALIDATE_DOCS': 'boolean - 문서 유효성 검증'
            }
        })

    elif request.method == 'PUT':
        try:
            new_config = request.get_json() or {}

            # 허용된 설정만 업데이트
            allowed_keys = [
                'AUTO_DOCS_ENABLED', 'DOCS_UPDATE_INTERVAL',
                'GENERATE_MARKDOWN', 'GENERATE_POSTMAN', 'VALIDATE_DOCS'
            ]

            updated_keys = []
            for key, value in new_config.items():
                if key in allowed_keys:
                    docs_manager.config[key] = value
                    current_app.config[key] = value
                    updated_keys.append(key)

            return APIResponse.success({
                'message': f'{len(updated_keys)}개 설정이 업데이트되었습니다',
                'updated_keys': updated_keys,
                'current_config': docs_manager.config
            })

        except Exception as e:
            return APIResponse.error(
                message="설정 업데이트 중 오류가 발생했습니다",
                details={'error': str(e)},
                status_code=500
            )


# 에러 핸들러
@docs_api.errorhandler(404)
def docs_not_found(error):
    return APIResponse.error(
        message="요청한 문서 API를 찾을 수 없습니다",
        status_code=404
    )


@docs_api.errorhandler(500)
def docs_internal_error(error):
    return APIResponse.error(
        message="문서 시스템 내부 오류가 발생했습니다",
        status_code=500
    )