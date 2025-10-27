"""
Simplified Request/Response Validation Framework
Compatible with current Marshmallow version
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

    class Meta:
        unknown = 'EXCLUDE'  # Ignore unknown fields


class PaginationSchema(BaseSchema):
    """Schema for pagination parameters"""
    page = fields.Integer(load_default=1, validate=validate.Range(min=1))
    limit = fields.Integer(load_default=20, validate=validate.Range(min=1, max=100))
    sort = fields.String(load_default='created_at', validate=validate.Length(min=1, max=50))
    order = fields.String(load_default='desc', validate=validate.OneOf(['asc', 'desc']))


# Project schemas
class CreateProjectSchema(BaseSchema):
    """Schema for creating a project"""
    name = fields.String(required=True, validate=validate.Length(min=1, max=100))
    description = fields.String(validate=validate.Length(max=500), allow_none=True)
    start_date = fields.Date(required=True)
    end_date = fields.Date(allow_none=True)
    budget = fields.Decimal(places=2, validate=validate.Range(min=0), allow_none=True)


class PatchProjectSchema(BaseSchema):
    """Schema for partial project updates"""
    name = fields.String(validate=validate.Length(min=1, max=100))
    description = fields.String(validate=validate.Length(max=500), allow_none=True)
    status = fields.String(validate=validate.OneOf(['active', 'completed', 'cancelled', 'on_hold']))
    end_date = fields.Date(allow_none=True)
    budget = fields.Decimal(places=2, validate=validate.Range(min=0), allow_none=True)
    progress = fields.Integer(validate=validate.Range(min=0, max=100))


# User schemas
class CreateUserSchema(BaseSchema):
    """Schema for creating a user"""
    email = fields.Email(required=True)
    name = fields.String(required=True, validate=validate.Length(min=1, max=100))
    role = fields.String(load_default='user', validate=validate.OneOf(['admin', 'manager', 'user']))


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