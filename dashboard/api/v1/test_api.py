"""
Test API endpoints to demonstrate standardized response format
These endpoints showcase the new API standardization features
"""

from flask import Blueprint, request
from datetime import datetime
import time

from ..responses import APIResponse, APIException, NotFoundException, ValidationException
from ..validation_simple import (
    validate_json, validate_query_params, get_pagination_params,
    PaginationSchema, CreateProjectSchema, PatchProjectSchema
)

# Create blueprint for test API endpoints
test_api_bp = Blueprint('test_api_v1', __name__, url_prefix='/api/v1/test')


@test_api_bp.route('/health', methods=['GET'])
def health_check():
    """
    Health check endpoint demonstrating standard success response
    ---
    tags:
      - Test API
    summary: Health check
    description: Simple health check endpoint that returns system status
    responses:
      200:
        description: System is healthy
        schema:
          $ref: '#/definitions/StandardResponse'
    """
    health_data = {
        "status": "healthy",
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "uptime": "24 hours",
        "version": "1.0.0",
        "checks": {
            "database": "healthy",
            "cache": "healthy",
            "storage": "healthy"
        }
    }

    return APIResponse.success(
        data=health_data,
        message="System is operating normally"
    )


@test_api_bp.route('/error-demo/<error_type>', methods=['GET'])
def error_demonstration(error_type):
    """
    Demonstrate different types of API errors
    ---
    tags:
      - Test API
    summary: Error demonstration
    description: Endpoint to test different error response formats
    parameters:
      - name: error_type
        in: path
        required: true
        type: string
        enum: [validation, not_found, unauthorized, forbidden, conflict, internal]
    responses:
      400:
        description: Validation error
      401:
        description: Unauthorized error
      403:
        description: Forbidden error
      404:
        description: Not found error
      409:
        description: Conflict error
      500:
        description: Internal server error
    """
    if error_type == 'validation':
        raise ValidationException(
            details={'field': ['This field is required']},
            message="Example validation error"
        )
    elif error_type == 'not_found':
        raise NotFoundException("Test Resource", "123")
    elif error_type == 'unauthorized':
        return APIResponse.unauthorized("Example unauthorized error")
    elif error_type == 'forbidden':
        return APIResponse.forbidden("Example forbidden error")
    elif error_type == 'conflict':
        return APIResponse.conflict("Test Resource", "Example conflict error")
    elif error_type == 'internal':
        raise Exception("Example internal server error")
    else:
        return APIResponse.not_found("Error Type", error_type)


@test_api_bp.route('/validation-demo', methods=['POST'])
@validate_json(CreateProjectSchema)
def validation_demonstration():
    """
    Demonstrate request validation with standard error responses
    ---
    tags:
      - Test API
    summary: Validation demonstration
    description: Endpoint to test request validation and error handling
    parameters:
      - name: body
        in: body
        required: true
        schema:
          type: object
          properties:
            name:
              type: string
              description: Project name (required)
            description:
              type: string
              description: Project description
            start_date:
              type: string
              format: date
              description: Start date (required)
            budget:
              type: number
              description: Project budget
    responses:
      200:
        description: Validation successful
      400:
        description: Validation failed
    """
    from flask import g

    validated_data = g.validated_data

    return APIResponse.success(
        data={
            "message": "Validation successful",
            "received_data": validated_data,
            "validation_info": {
                "schema": "CreateProjectSchema",
                "fields_validated": list(validated_data.keys())
            }
        },
        message="Request validation passed"
    )


@test_api_bp.route('/pagination-demo', methods=['GET'])
@validate_query_params(PaginationSchema)
def pagination_demonstration():
    """
    Demonstrate pagination with standard response format
    ---
    tags:
      - Test API
    summary: Pagination demonstration
    description: Endpoint to test pagination parameters and response format
    parameters:
      - name: page
        in: query
        type: integer
        minimum: 1
        default: 1
      - name: limit
        in: query
        type: integer
        minimum: 1
        maximum: 100
        default: 20
      - name: sort
        in: query
        type: string
        default: created_at
      - name: order
        in: query
        type: string
        enum: [asc, desc]
        default: desc
    responses:
      200:
        description: Paginated data
    """
    from flask import g
    from ..responses import PaginationHelper

    pagination_params = g.validated_data

    # Simulate data
    total_items = 157
    page = pagination_params['page']
    limit = pagination_params['limit']

    # Generate mock items for current page
    start_index = (page - 1) * limit
    mock_items = []

    for i in range(limit):
        item_index = start_index + i + 1
        if item_index <= total_items:
            mock_items.append({
                "id": f"item_{item_index:03d}",
                "name": f"Test Item {item_index}",
                "value": item_index * 10,
                "created_at": datetime.utcnow().isoformat() + "Z"
            })

    return PaginationHelper.paginated_response(
        items=mock_items,
        page=page,
        limit=limit,
        total=total_items,
        message=f"Retrieved {len(mock_items)} items"
    )


@test_api_bp.route('/performance-demo', methods=['GET'])
def performance_demonstration():
    """
    Demonstrate performance monitoring in response metadata
    ---
    tags:
      - Test API
    summary: Performance demonstration
    description: Endpoint to test performance tracking in response metadata
    parameters:
      - name: delay
        in: query
        type: integer
        minimum: 0
        maximum: 5
        default: 0
        description: Artificial delay in seconds
    responses:
      200:
        description: Performance data with timing information
    """
    # Add artificial delay if requested
    delay = request.args.get('delay', 0, type=int)
    if delay > 0:
        time.sleep(min(delay, 5))  # Cap at 5 seconds

    performance_data = {
        "operation": "performance_test",
        "artificial_delay": f"{delay} seconds",
        "server_info": {
            "environment": "development",
            "version": "1.0.0"
        }
    }

    return APIResponse.success(
        data=performance_data,
        message="Performance test completed",
        performance_test=True,
        artificial_delay=delay
    )


@test_api_bp.route('/async-demo', methods=['POST'])
def async_operation_demonstration():
    """
    Demonstrate async operation with 202 Accepted response
    ---
    tags:
      - Test API
    summary: Async operation demonstration
    description: Endpoint to demonstrate asynchronous operation handling
    responses:
      202:
        description: Operation accepted for processing
    """
    operation_id = f"op_{int(time.time())}"

    async_data = {
        "operation_id": operation_id,
        "status": "accepted",
        "estimated_completion": "2-3 minutes",
        "status_check_url": f"/api/v1/test/status/{operation_id}"
    }

    return APIResponse.success(
        data=async_data,
        message="Operation accepted for processing",
        status_code=202
    )


@test_api_bp.route('/status/<operation_id>', methods=['GET'])
def operation_status(operation_id):
    """
    Check status of async operation
    ---
    tags:
      - Test API
    summary: Operation status check
    description: Check the status of an asynchronous operation
    parameters:
      - name: operation_id
        in: path
        required: true
        type: string
    responses:
      200:
        description: Operation status
      404:
        description: Operation not found
    """
    # Simulate operation status
    status_data = {
        "operation_id": operation_id,
        "status": "completed",
        "progress": 100,
        "started_at": datetime.utcnow().isoformat() + "Z",
        "completed_at": datetime.utcnow().isoformat() + "Z",
        "result": {
            "message": "Operation completed successfully",
            "processed_items": 42
        }
    }

    return APIResponse.success(
        data=status_data,
        message="Operation status retrieved"
    )


@test_api_bp.route('/bulk-demo', methods=['POST'])
def bulk_operation_demonstration():
    """
    Demonstrate bulk operation with detailed results
    ---
    tags:
      - Test API
    summary: Bulk operation demonstration
    description: Endpoint to demonstrate bulk operations with detailed success/error reporting
    parameters:
      - name: body
        in: body
        required: true
        schema:
          type: object
          properties:
            items:
              type: array
              items:
                type: object
                properties:
                  id:
                    type: string
                  action:
                    type: string
                    enum: [create, update, delete]
    responses:
      200:
        description: Bulk operation results
    """
    request_data = request.get_json() or {}
    items = request_data.get('items', [])

    results = {
        "total_items": len(items),
        "successful": 0,
        "failed": 0,
        "results": []
    }

    for i, item in enumerate(items):
        # Simulate some failures
        if i % 7 == 0:  # Every 7th item fails
            results["failed"] += 1
            results["results"].append({
                "item_id": item.get('id', f'item_{i}'),
                "status": "failed",
                "error": "Simulated failure for demonstration"
            })
        else:
            results["successful"] += 1
            results["results"].append({
                "item_id": item.get('id', f'item_{i}'),
                "status": "success",
                "action": item.get('action', 'processed')
            })

    return APIResponse.success(
        data=results,
        message=f"Bulk operation completed: {results['successful']} successful, {results['failed']} failed"
    )