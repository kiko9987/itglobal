"""
Marshmallow 스키마 자동 분석기
Marshmallow 스키마를 OpenAPI 스키마로 변환하고 예제 데이터 생성
"""

import inspect
from typing import Dict, List, Optional, Any, Type, Union
from dataclasses import dataclass
from datetime import datetime, date
from decimal import Decimal
import re

from marshmallow import Schema, fields
from marshmallow.validate import Length, Range, OneOf, Email, URL


@dataclass
class SchemaInfo:
    """스키마 정보를 담는 데이터 클래스"""
    name: str                           # 스키마 클래스명
    schema_class: Type[Schema]          # 스키마 클래스 객체
    fields_info: Dict[str, Dict]        # 필드별 상세 정보
    openapi_schema: Dict[str, Any]      # OpenAPI 스키마 형태
    example_data: Dict[str, Any]        # 예제 데이터
    description: Optional[str]          # 스키마 설명
    required_fields: List[str]          # 필수 필드 목록
    module: str                         # 모듈명


class SchemaAnalyzer:
    """Marshmallow 스키마를 분석하는 클래스"""

    def __init__(self):
        self.schemas_info: Dict[str, SchemaInfo] = {}
        self.type_mapping = self._build_type_mapping()

    def _build_type_mapping(self) -> Dict[Type, Dict[str, Any]]:
        """Marshmallow 필드 타입을 OpenAPI 타입으로 매핑"""
        return {
            fields.String: {'type': 'string'},
            fields.Integer: {'type': 'integer'},
            fields.Float: {'type': 'number'},
            fields.Boolean: {'type': 'boolean'},
            fields.DateTime: {'type': 'string', 'format': 'date-time'},
            fields.Date: {'type': 'string', 'format': 'date'},
            fields.Time: {'type': 'string', 'format': 'time'},
            fields.Email: {'type': 'string', 'format': 'email'},
            fields.Url: {'type': 'string', 'format': 'uri'},
            fields.UUID: {'type': 'string', 'format': 'uuid'},
            fields.Decimal: {'type': 'number'},
            fields.List: {'type': 'array'},
            fields.Dict: {'type': 'object'},
            fields.Raw: {},  # 타입 제한 없음
        }

    def analyze_schema(self, schema_class: Type[Schema]) -> SchemaInfo:
        """개별 스키마 클래스 분석"""
        schema_name = schema_class.__name__
        schema_instance = schema_class()

        # 필드 정보 분석
        fields_info = {}
        required_fields = []
        openapi_properties = {}

        for field_name, field_obj in schema_instance.fields.items():
            field_info = self._analyze_field(field_name, field_obj)
            fields_info[field_name] = field_info

            if field_info['required']:
                required_fields.append(field_name)

            openapi_properties[field_name] = field_info['openapi_schema']

        # OpenAPI 스키마 구성
        openapi_schema = {
            'type': 'object',
            'properties': openapi_properties
        }

        if required_fields:
            openapi_schema['required'] = required_fields

        # 예제 데이터 생성
        example_data = self._generate_example_data(fields_info)

        # 스키마 설명 추출
        description = self._extract_schema_description(schema_class)

        schema_info = SchemaInfo(
            name=schema_name,
            schema_class=schema_class,
            fields_info=fields_info,
            openapi_schema=openapi_schema,
            example_data=example_data,
            description=description,
            required_fields=required_fields,
            module=schema_class.__module__
        )

        self.schemas_info[schema_name] = schema_info
        return schema_info

    def _analyze_field(self, field_name: str, field_obj: fields.Field) -> Dict[str, Any]:
        """개별 필드 분석"""
        field_type = type(field_obj)

        # 기본 타입 정보
        openapi_info = self.type_mapping.get(field_type, {'type': 'string'}).copy()

        # 필수 여부
        required = field_obj.required

        # 기본값
        default_value = getattr(field_obj, 'load_default', None)
        if default_value is not None and default_value != fields.missing_:
            openapi_info['default'] = default_value

        # null 허용 여부
        if field_obj.allow_none:
            openapi_info['nullable'] = True

        # 검증 규칙 분석
        self._analyze_validators(field_obj, openapi_info)

        # 설명 추출
        description = self._extract_field_description(field_obj)
        if description:
            openapi_info['description'] = description

        # 특수 필드 타입별 처리
        if isinstance(field_obj, fields.List):
            openapi_info['items'] = self._analyze_list_items(field_obj)
        elif isinstance(field_obj, fields.Nested):
            openapi_info = self._analyze_nested_field(field_obj)

        return {
            'field_type': field_type.__name__,
            'required': required,
            'openapi_schema': openapi_info,
            'validators': [v.__class__.__name__ for v in field_obj.validators],
            'default': default_value
        }

    def _analyze_validators(self, field_obj: fields.Field, openapi_info: Dict):
        """필드의 검증 규칙을 OpenAPI 스키마에 반영"""
        for validator in field_obj.validators:
            if isinstance(validator, Length):
                if validator.min is not None:
                    openapi_info['minLength'] = validator.min
                if validator.max is not None:
                    openapi_info['maxLength'] = validator.max

            elif isinstance(validator, Range):
                if validator.min is not None:
                    openapi_info['minimum'] = validator.min
                if validator.max is not None:
                    openapi_info['maximum'] = validator.max

            elif isinstance(validator, OneOf):
                openapi_info['enum'] = list(validator.choices)

            elif isinstance(validator, Email):
                openapi_info['format'] = 'email'

            elif isinstance(validator, URL):
                openapi_info['format'] = 'uri'

    def _analyze_list_items(self, list_field: fields.List) -> Dict[str, Any]:
        """List 필드의 아이템 타입 분석"""
        inner_field = list_field.inner

        if inner_field:
            inner_type = type(inner_field)
            return self.type_mapping.get(inner_type, {'type': 'string'}).copy()

        return {'type': 'string'}  # 기본값

    def _analyze_nested_field(self, nested_field: fields.Nested) -> Dict[str, Any]:
        """Nested 필드 분석"""
        nested_schema = nested_field.inner

        if inspect.isclass(nested_schema) and issubclass(nested_schema, Schema):
            # 중첩된 스키마도 분석
            nested_info = self.analyze_schema(nested_schema)
            return {
                '$ref': f'#/components/schemas/{nested_info.name}'
            }

        return {'type': 'object'}

    def _extract_field_description(self, field_obj: fields.Field) -> Optional[str]:
        """필드에서 설명 추출"""
        # metadata에서 description 찾기
        if hasattr(field_obj, 'metadata') and 'description' in field_obj.metadata:
            return field_obj.metadata['description']

        # doc 속성 확인
        if hasattr(field_obj, 'doc') and field_obj.doc:
            return field_obj.doc

        return None

    def _extract_schema_description(self, schema_class: Type[Schema]) -> Optional[str]:
        """스키마 클래스에서 설명 추출"""
        docstring = inspect.getdoc(schema_class)
        if docstring:
            # 첫 번째 줄을 설명으로 사용
            return docstring.split('\n')[0].strip()

        return None

    def _generate_example_data(self, fields_info: Dict[str, Dict]) -> Dict[str, Any]:
        """필드 정보를 바탕으로 예제 데이터 생성"""
        example = {}

        for field_name, field_info in fields_info.items():
            openapi_schema = field_info['openapi_schema']
            example_value = self._generate_field_example(field_name, openapi_schema)

            if example_value is not None:
                example[field_name] = example_value

        return example

    def _generate_field_example(self, field_name: str, schema: Dict[str, Any]) -> Any:
        """개별 필드의 예제 값 생성"""
        # enum이 있으면 첫 번째 값 사용
        if 'enum' in schema:
            return schema['enum'][0]

        # default 값이 있으면 사용
        if 'default' in schema:
            return schema['default']

        # 타입별 예제 생성
        field_type = schema.get('type', 'string')
        field_format = schema.get('format')

        if field_type == 'string':
            if field_format == 'email':
                return 'user@example.com'
            elif field_format == 'date':
                return '2025-01-01'
            elif field_format == 'date-time':
                return '2025-01-01T00:00:00Z'
            elif field_format == 'uuid':
                return '123e4567-e89b-12d3-a456-426614174000'
            elif field_format == 'uri':
                return 'https://example.com'
            else:
                # 필드명을 기반으로 의미있는 예제 생성
                return self._generate_semantic_string_example(field_name)

        elif field_type == 'integer':
            min_val = schema.get('minimum', 1)
            max_val = schema.get('maximum', 100)
            return min(max_val, max(min_val, 42))

        elif field_type == 'number':
            min_val = schema.get('minimum', 0.0)
            max_val = schema.get('maximum', 1000.0)
            return min(max_val, max(min_val, 123.45))

        elif field_type == 'boolean':
            return True

        elif field_type == 'array':
            items_schema = schema.get('items', {'type': 'string'})
            example_item = self._generate_field_example('item', items_schema)
            return [example_item] if example_item is not None else []

        elif field_type == 'object':
            return {}

        return None

    def _generate_semantic_string_example(self, field_name: str) -> str:
        """필드명을 기반으로 의미있는 문자열 예제 생성"""
        field_name_lower = field_name.lower()

        # 일반적인 필드명 패턴별 예제
        if 'name' in field_name_lower:
            return '프로젝트명 예시'
        elif 'title' in field_name_lower:
            return '제목 예시'
        elif 'description' in field_name_lower:
            return '설명 예시입니다'
        elif 'code' in field_name_lower:
            return 'ABC123'
        elif 'id' in field_name_lower:
            return 'example_id'
        elif 'status' in field_name_lower:
            return 'active'
        elif 'type' in field_name_lower:
            return 'default'
        elif 'phone' in field_name_lower:
            return '010-1234-5678'
        elif 'address' in field_name_lower:
            return '서울시 강남구'
        elif 'company' in field_name_lower:
            return '회사명 예시'
        else:
            return f'{field_name} 예시'

    def analyze_module_schemas(self, module) -> List[SchemaInfo]:
        """모듈에서 모든 Schema 클래스 찾아서 분석"""
        schema_infos = []

        for name in dir(module):
            obj = getattr(module, name)

            # Schema 클래스인지 확인
            if (inspect.isclass(obj) and
                issubclass(obj, Schema) and
                obj is not Schema):

                try:
                    schema_info = self.analyze_schema(obj)
                    schema_infos.append(schema_info)
                except Exception as e:
                    print(f"스키마 분석 중 오류: {name} - {str(e)}")

        return schema_infos

    def get_openapi_components(self) -> Dict[str, Any]:
        """모든 분석된 스키마를 OpenAPI components 형태로 반환"""
        components = {
            'schemas': {}
        }

        for schema_name, schema_info in self.schemas_info.items():
            components['schemas'][schema_name] = schema_info.openapi_schema

        return components

    def get_schema_examples(self) -> Dict[str, Any]:
        """모든 스키마의 예제 데이터 반환"""
        examples = {}

        for schema_name, schema_info in self.schemas_info.items():
            examples[schema_name] = schema_info.example_data

        return examples

    def export_analysis_report(self) -> Dict[str, Any]:
        """스키마 분석 리포트 생성"""
        return {
            'total_schemas': len(self.schemas_info),
            'schemas_summary': {
                name: {
                    'fields_count': len(info.fields_info),
                    'required_fields_count': len(info.required_fields),
                    'module': info.module,
                    'description': info.description
                }
                for name, info in self.schemas_info.items()
            },
            'analysis_timestamp': datetime.utcnow().isoformat()
        }


# 사용 예시
if __name__ == "__main__":
    from marshmallow import Schema, fields

    # 테스트용 스키마
    class TestProjectSchema(Schema):
        """프로젝트 생성 스키마"""
        name = fields.String(required=True, validate=Length(min=1, max=100))
        description = fields.String(validate=Length(max=500))
        start_date = fields.Date(required=True)
        budget = fields.Decimal(validate=Range(min=0))
        status = fields.String(validate=OneOf(['active', 'inactive']))

    analyzer = SchemaAnalyzer()
    schema_info = analyzer.analyze_schema(TestProjectSchema)

    print(f"스키마명: {schema_info.name}")
    print(f"필수 필드: {schema_info.required_fields}")
    print(f"예제 데이터: {schema_info.example_data}")
    print(f"OpenAPI 스키마: {schema_info.openapi_schema}")