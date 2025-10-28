"""
프로젝트 관련 Marshmallow 검증 스키마
"""
from marshmallow import Schema, fields, validate, validates, ValidationError, EXCLUDE
from datetime import datetime
import re


class MoneyField(fields.Field):
    """
    금액 필드 (통화 기호와 콤마를 자동으로 처리)

    지원 형식:
    - "₩1,000,000" → 1000000.0
    - "$1,000.00" → 1000.0
    - "1000000" → 1000000.0
    - 1000000 → 1000000.0

    ₩ 기호만 허용하며, 다른 통화는 거부합니다.
    """

    def _deserialize(self, value, attr, data, **kwargs):
        if value is None or value == '':
            return None

        # 이미 숫자인 경우
        if isinstance(value, (int, float)):
            if value < 0:
                raise ValidationError("금액은 0 이상이어야 합니다")
            return float(value)

        # 문자열 처리
        if isinstance(value, str):
            value = value.strip()

            # 빈 문자열
            if not value:
                return None

            # 통화 기호 검증 (₩만 허용)
            if any(symbol in value for symbol in ['$', '€', '£', '¥', '元', '¢']):
                raise ValidationError("₩ 기호만 사용할 수 있습니다 (다른 통화는 지원하지 않습니다)")

            # ₩ 기호와 콤마, 공백 제거
            clean_value = re.sub(r'[₩,\s]', '', value)

            # 숫자만 남았는지 확인
            if not re.match(r'^\d+(\.\d+)?$', clean_value):
                raise ValidationError(f"올바르지 않은 금액 형식입니다: {value}")

            try:
                amount = float(clean_value)
                if amount < 0:
                    raise ValidationError("금액은 0 이상이어야 합니다")
                return amount
            except ValueError:
                raise ValidationError(f"숫자로 변환할 수 없습니다: {value}")

        raise ValidationError(f"지원하지 않는 타입입니다: {type(value)}")

    def _serialize(self, value, attr, obj, **kwargs):
        if value is None:
            return None
        return float(value)


class ProjectCreateSchema(Schema):
    """프로젝트 생성 스키마"""

    class Meta:
        unknown = EXCLUDE  # 알 수 없는 필드는 무시

    # 필수 필드
    project_code = fields.Str(
        required=True,
        validate=validate.Length(min=1, max=20, error="프로젝트 코드는 1-20자여야 합니다")
    )
    manager = fields.Str(
        required=True,
        validate=validate.Length(min=1, max=50, error="담당자는 1-50자여야 합니다")
    )
    company = fields.Str(
        required=True,
        validate=validate.Length(min=1, max=100, error="거래처는 1-100자여야 합니다")
    )

    # 선택 필드
    address = fields.Str(
        allow_none=True,
        validate=validate.Length(max=500, error="주소는 500자 이하여야 합니다")
    )
    work_content = fields.Str(
        allow_none=True,
        validate=validate.Length(max=500, error="공사 내용은 500자 이하여야 합니다")
    )

    # 날짜 필드 (문자열로 받아서 검증)
    start_date = fields.Str(allow_none=True)
    end_date = fields.Str(allow_none=True)

    # 금액 필드 (커스텀 MoneyField 사용 - ₩ 기호와 콤마 자동 처리)
    down_payment = MoneyField(allow_none=True)
    mid_payment = MoneyField(allow_none=True)
    final_payment = MoneyField(allow_none=True)
    total_amount = MoneyField(allow_none=True)

    @validates('start_date')
    def validate_start_date(self, value, **kwargs):
        """시작일 유효성 검증"""
        if value:
            try:
                datetime.strptime(value, '%Y-%m-%d')
            except ValueError:
                raise ValidationError("시작일 형식이 올바르지 않습니다 (YYYY-MM-DD)")

    @validates('end_date')
    def validate_end_date(self, value, **kwargs):
        """종료일 유효성 검증"""
        if value:
            try:
                datetime.strptime(value, '%Y-%m-%d')
            except ValueError:
                raise ValidationError("종료일 형식이 올바르지 않습니다 (YYYY-MM-DD)")


class ProjectAutoCreateSchema(Schema):
    """프로젝트 자동 생성 스키마 (프로젝트 코드 자동 생성용)"""

    class Meta:
        unknown = EXCLUDE

    # 필수 필드 (project_code는 자동 생성되므로 제외)
    manager = fields.Str(
        required=True,
        validate=validate.Length(min=1, max=50, error="담당자는 1-50자여야 합니다")
    )
    company = fields.Str(
        required=True,
        validate=validate.Length(min=1, max=100, error="거래처는 1-100자여야 합니다")
    )

    # 선택 필드
    address = fields.Str(
        allow_none=True,
        validate=validate.Length(max=500, error="주소는 500자 이하여야 합니다")
    )
    work_content = fields.Str(
        allow_none=True,
        validate=validate.Length(max=500, error="공사 내용은 500자 이하여야 합니다")
    )

    # 날짜 필드
    start_date = fields.Str(allow_none=True)
    end_date = fields.Str(allow_none=True)

    # 금액 필드 (커스텀 MoneyField 사용 - ₩ 기호와 콤마 자동 처리)
    down_payment = MoneyField(allow_none=True)
    mid_payment = MoneyField(allow_none=True)
    final_payment = MoneyField(allow_none=True)
    total_amount = MoneyField(allow_none=True)

    @validates('start_date')
    def validate_start_date(self, value, **kwargs):
        """시작일 유효성 검증"""
        if value:
            try:
                datetime.strptime(value, '%Y-%m-%d')
            except ValueError:
                raise ValidationError("시작일 형식이 올바르지 않습니다 (YYYY-MM-DD)")

    @validates('end_date')
    def validate_end_date(self, value, **kwargs):
        """종료일 유효성 검증"""
        if value:
            try:
                datetime.strptime(value, '%Y-%m-%d')
            except ValueError:
                raise ValidationError("종료일 형식이 올바르지 않습니다 (YYYY-MM-DD)")


class ProjectUpdateSchema(Schema):
    """프로젝트 수정 스키마 (모든 필드 선택)"""

    class Meta:
        unknown = EXCLUDE

    project_code = fields.Str(
        validate=validate.Length(min=1, max=20, error="프로젝트 코드는 1-20자여야 합니다")
    )
    manager = fields.Str(
        validate=validate.Length(min=1, max=50, error="담당자는 1-50자여야 합니다")
    )
    company = fields.Str(
        validate=validate.Length(min=1, max=100, error="거래처는 1-100자여야 합니다")
    )
    address = fields.Str(
        allow_none=True,
        validate=validate.Length(max=500, error="주소는 500자 이하여야 합니다")
    )
    work_content = fields.Str(
        allow_none=True,
        validate=validate.Length(max=500, error="공사 내용은 500자 이하여야 합니다")
    )

    start_date = fields.Str(allow_none=True)
    end_date = fields.Str(allow_none=True)

    down_payment = fields.Float(
        allow_none=True,
        validate=validate.Range(min=0, error="계약금은 0 이상이어야 합니다")
    )
    mid_payment = fields.Float(
        allow_none=True,
        validate=validate.Range(min=0, error="중도금은 0 이상이어야 합니다")
    )
    final_payment = fields.Float(
        allow_none=True,
        validate=validate.Range(min=0, error="잔금은 0 이상이어야 합니다")
    )
    total_amount = fields.Float(
        allow_none=True,
        validate=validate.Range(min=0, error="총액은 0 이상이어야 합니다")
    )


class CellMemoSchema(Schema):
    """셀 메모 저장 스키마"""

    class Meta:
        unknown = EXCLUDE

    # 필수 필드
    row = fields.Integer(
        required=True,
        validate=validate.Range(min=1, error="행 번호는 1 이상이어야 합니다")
    )
    column = fields.Str(
        required=True,
        validate=validate.Regexp(
            r'^[A-Z]{1,3}$',
            error="컬럼은 A-ZZZ 형식이어야 합니다"
        )
    )
    memo = fields.Str(
        required=True,
        validate=validate.Length(max=2000, error="메모는 2000자 이하여야 합니다")
    )

    # 선택 필드
    project_code = fields.Str(
        allow_none=True,
        validate=validate.Length(max=20, error="프로젝트 코드는 20자 이하여야 합니다")
    )
    field_name = fields.Str(
        allow_none=True,
        validate=validate.OneOf(
            ['계약금', '중도금', '잔금'],
            error="field_name은 '계약금', '중도금', '잔금' 중 하나여야 합니다"
        )
    )


# 검증 헬퍼 함수
def validate_request_data(schema_class, data):
    """
    요청 데이터 검증 헬퍼

    Args:
        schema_class: Marshmallow Schema 클래스
        data: 검증할 데이터 dict

    Returns:
        tuple: (validated_data, errors)
            - validated_data: 검증된 데이터 (성공 시)
            - errors: 에러 메시지 dict (실패 시)
    """
    schema = schema_class()
    try:
        validated_data = schema.load(data)
        return validated_data, None
    except ValidationError as err:
        return None, err.messages


def format_validation_errors(errors):
    """
    검증 에러를 사용자 친화적 형식으로 변환

    Args:
        errors: Marshmallow ValidationError.messages

    Returns:
        list: 에러 메시지 리스트
    """
    error_messages = []
    for field, messages in errors.items():
        if isinstance(messages, list):
            for msg in messages:
                error_messages.append(f"{field}: {msg}")
        elif isinstance(messages, dict):
            # 중첩된 에러
            for key, msg in messages.items():
                error_messages.append(f"{field}.{key}: {msg}")
        else:
            error_messages.append(f"{field}: {messages}")
    return error_messages
