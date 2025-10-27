"""
Request/Response Validation Framework
Using Marshmallow for schema validation with standard error handling
"""

from functools import wraps
from flask import request, g
from marshmallow import Schema, fields, validate, ValidationError, pre_load, post_dump
from typing import Dict, Any, Optional, Union, Type
import re

from .responses import APIResponse, ValidationException


# Custom validation functions
def validate_project_code(value: str) -> str:
    """Validate project code format"""
    if not re.match(r'^[A-Z0-9]{3,10}$', value):
        raise ValidationError('Project code must be 3-10 characters of uppercase letters and numbers')
    return value


def validate_email_format(value: str) -> str:
    """Enhanced email validation"""
    if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', value):
        raise ValidationError('Invalid email format')
    return value


def validate_phone_number(value: str) -> str:
    """Validate phone number format"""
    # Remove all non-digit characters
    digits_only = re.sub(r'\D', '', value)
    if len(digits_only) < 10 or len(digits_only) > 15:
        raise ValidationError('Phone number must be 10-15 digits')
    return value


# Base schemas
class BaseSchema(Schema):
    """Base schema with common functionality"""

    @pre_load
    def strip_strings(self, data, **kwargs):
        """Strip whitespace from string fields"""
        if isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, str):
                    data[key] = value.strip()
        return data

    @post_dump
    def remove_nulls(self, data, **kwargs):
        """Remove null values from output"""
        return {key: value for key, value in data.items() if value is not None}

    class Meta:
        unknown = 'EXCLUDE'  # Ignore unknown fields


class PaginationSchema(BaseSchema):
    """Schema for pagination parameters"""
    page = fields.Integer(
        load_default=1,
        validate=validate.Range(min=1),
        description="Page number"
    )
    limit = fields.Integer(
        load_default=20,
        validate=validate.Range(min=1, max=100),
        description="Items per page"
    )
    sort = fields.String(
        load_default='created_at',
        validate=validate.Length(min=1, max=50),
        description="Field to sort by"
    )
    order = fields.String(
        load_default='desc',
        validate=validate.OneOf(['asc', 'desc']),
        description="Sort order"
    )


# Project schemas
class ProjectSchema(BaseSchema):
    """Schema for project data"""
    id = fields.String(
        required=True,
        validate=validate_project_code,
        description="Project code"
    )
    name = fields.String(
        required=True,
        validate=validate.Length(min=1, max=100),
        description="Project name"
    )
    description = fields.String(
        validate=validate.Length(max=500),
        allow_none=True,
        description="Project description"
    )
    status = fields.String(
        required=True,
        validate=validate.OneOf(['active', 'completed', 'cancelled', 'on_hold']),
        description="Project status"
    )
    start_date = fields.Date(
        required=True,
        description="Project start date"
    )
    end_date = fields.Date(
        allow_none=True,
        description="Project end date"
    )
    budget = fields.Decimal(
        places=2,
        validate=validate.Range(min=0),
        allow_none=True,
        description="Project budget"
    )
    progress = fields.Integer(
        validate=validate.Range(min=0, max=100),
        load_default=0,
        description="Completion percentage"
    )
    created_at = fields.DateTime(
        dump_only=True,
        description="Creation timestamp"
    )
    updated_at = fields.DateTime(
        dump_only=True,
        description="Last update timestamp"
    )
    created_by = fields.String(
        dump_only=True,
        description="User who created the project"
    )


class CreateProjectSchema(BaseSchema):
    """Schema for creating a project"""
    name = fields.String(
        required=True,
        validate=validate.Length(min=1, max=100),
        description="Project name"
    )
    description = fields.String(
        validate=validate.Length(max=500),
        allow_none=True,
        description="Project description"
    )
    start_date = fields.Date(
        required=True,
        description="Project start date"
    )
    end_date = fields.Date(
        allow_none=True,
        description="Project end date"
    )
    budget = fields.Decimal(
        places=2,
        validate=validate.Range(min=0),
        allow_none=True,
        description="Project budget"
    )


class UpdateProjectSchema(CreateProjectSchema):
    """Schema for updating a project"""
    status = fields.String(
        validate=validate.OneOf(['active', 'completed', 'cancelled', 'on_hold']),
        description="Project status"
    )
    progress = fields.Integer(
        validate=validate.Range(min=0, max=100),
        description="Completion percentage"
    )


class PatchProjectSchema(BaseSchema):
    """Schema for partial project updates"""
    name = fields.String(
        validate=validate.Length(min=1, max=100),
        description="Project name"
    )
    description = fields.String(
        validate=validate.Length(max=500),
        allow_none=True,
        description="Project description"
    )
    status = fields.String(
        validate=validate.OneOf(['active', 'completed', 'cancelled', 'on_hold']),
        description="Project status"
    )
    end_date = fields.Date(
        allow_none=True,
        description="Project end date"
    )
    budget = fields.Decimal(
        places=2,
        validate=validate.Range(min=0),
        allow_none=True,
        description="Project budget"
    )
    progress = fields.Integer(
        validate=validate.Range(min=0, max=100),
        description="Completion percentage"
    )


# User schemas
class UserSchema(BaseSchema):
    """Schema for user data"""
    id = fields.String(
        dump_only=True,
        description="User ID"
    )
    email = fields.Email(
        required=True,
        description="Email address"
    )
    name = fields.String(
        required=True,
        validate=validate.Length(min=1, max=100),
        description="Full name"
    )
    role = fields.String(
        required=True,
        validate=validate.OneOf(['admin', 'manager', 'user']),
        description="User role"
    )
    status = fields.String(
        validate=validate.OneOf(['active', 'inactive', 'suspended']),
        load_default='active',
        description="User status"
    )
    last_login = fields.DateTime(
        dump_only=True,
        allow_none=True,
        description="Last login timestamp"
    )
    created_at = fields.DateTime(
        dump_only=True,
        description="Account creation timestamp"
    )


class CreateUserSchema(BaseSchema):
    """Schema for creating a user"""
    email = fields.Email(
        required=True,
        description="Email address"
    )
    name = fields.String(
        required=True,
        validate=validate.Length(min=1, max=100),
        description="Full name"
    )
    role = fields.String(
        validate=validate.OneOf(['admin', 'manager', 'user']),
        load_default='user',
        description="User role"
    )


# Analytics schemas
class MonthlySalesQuerySchema(BaseSchema):
    """Schema for monthly sales query parameters"""
    year = fields.Integer(
        validate=validate.Range(min=2020, max=2030),
        allow_none=True,
        description="Year for analysis"
    )
    month = fields.Integer(
        validate=validate.Range(min=1, max=12),
        allow_none=True,
        description="Specific month (1-12)"
    )


# Validation decorators
def validate_json(schema_class: Type[Schema], location: str = 'json'):
    """Decorator to validate request data against a schema"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            schema = schema_class()

            try:
                if location == 'json':
                    data = schema.load(request.get_json() or {})
                elif location == 'args':
                    data = schema.load(request.args.to_dict())
                elif location == 'form':
                    data = schema.load(request.form.to_dict())
                else:
                    raise ValueError(f"Unknown location: {location}")

                # Store validated data in Flask g for access in the route
                g.validated_data = data
                return func(*args, **kwargs)

            except ValidationError as err:
                raise ValidationException(
                    details=err.messages,
                    message="Request validation failed"
                )

        return wrapper
    return decorator


def validate_query_params(schema_class: Type[Schema]):
    """Decorator to validate query parameters"""
    return validate_json(schema_class, location='args')


def validate_form_data(schema_class: Type[Schema]):
    """Decorator to validate form data"""
    return validate_json(schema_class, location='form')


def validate_response(schema_class: Type[Schema]):
    """Decorator to validate response data"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            result = func(*args, **kwargs)

            # Only validate if we have data to validate
            if hasattr(result, 'get_json'):
                # It's a Response object
                try:
                    response_data = result.get_json()
                    if response_data and 'data' in response_data and response_data['data']:
                        schema = schema_class()
                        schema.load(response_data['data'])
                except Exception as e:
                    # Log validation error but don't fail the response
                    import logging
                    logging.warning(f"Response validation failed: {str(e)}")

            return result

        return wrapper
    return decorator


# Query parameter extraction helpers
def get_pagination_params() -> Dict[str, Any]:
    """Extract and validate pagination parameters from request"""
    schema = PaginationSchema()
    try:
        return schema.load(request.args.to_dict())
    except ValidationError as err:
        raise ValidationException(
            details=err.messages,
            message="Invalid pagination parameters"
        )


def get_validated_data() -> Dict[str, Any]:
    """Get validated data from Flask g (set by validation decorators)"""
    return getattr(g, 'validated_data', {})


# Schema registry for dynamic access
SCHEMA_REGISTRY = {
    'project': ProjectSchema,
    'create_project': CreateProjectSchema,
    'update_project': UpdateProjectSchema,
    'patch_project': PatchProjectSchema,
    'user': UserSchema,
    'create_user': CreateUserSchema,
    'monthly_sales_query': MonthlySalesQuerySchema,
    'pagination': PaginationSchema
}


def get_schema(schema_name: str) -> Schema:
    """Get schema instance by name"""
    schema_class = SCHEMA_REGISTRY.get(schema_name)
    if not schema_class:
        raise ValueError(f"Unknown schema: {schema_name}")
    return schema_class()


# Field validation helpers
class CustomFields:
    """Custom field types for common use cases"""

    @staticmethod
    def ProjectCode(**kwargs):
        """Project code field with validation"""
        return fields.String(
            validate=validate_project_code,
            **kwargs
        )

    @staticmethod
    def Email(**kwargs):
        """Email field with enhanced validation"""
        return fields.String(
            validate=validate_email_format,
            **kwargs
        )

    @staticmethod
    def PhoneNumber(**kwargs):
        """Phone number field with validation"""
        return fields.String(
            validate=validate_phone_number,
            **kwargs
        )

    @staticmethod
    def StatusField(choices: list, **kwargs):
        """Generic status field"""
        return fields.String(
            validate=validate.OneOf(choices),
            **kwargs
        )

    @staticmethod
    def DateRange(start_field: str = 'start_date', end_field: str = 'end_date'):
        """Validator for date ranges"""
        def validate_date_range(data, **kwargs):
            start = data.get(start_field)
            end = data.get(end_field)

            if start and end and start > end:
                raise ValidationError(
                    f"{end_field} must be after {start_field}",
                    field_name=end_field
                )

        return validate_date_range