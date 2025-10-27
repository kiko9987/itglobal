# Phase 4A: API Standardization Strategy & Architecture Design

## 1. API Versioning Strategy

### Current State
- No versioning system in place
- All endpoints are unversioned
- Risk of breaking changes affecting clients

### Proposed Solution: URL Path Versioning
```
Current: /api/projects/list
Proposed: /api/v1/projects
```

**Benefits:**
- Clear version identification
- Easy to implement
- Backward compatibility support
- Client can choose version explicitly

**Implementation Plan:**
1. Create versioned blueprints (`v1_bp`, `v2_bp`)
2. Gradually migrate endpoints to `/api/v1/`
3. Maintain legacy routes with deprecation warnings
4. Default unversioned routes to latest version

## 2. Standard Response Format

### Current Issues
- Mixed JSON structures
- Inconsistent error responses
- No standard metadata
- Mixed success/error indicators

### Proposed Standard Response Wrapper
```json
{
  "success": true|false,
  "data": any,
  "error": {
    "code": "ERROR_CODE",
    "message": "Human readable message",
    "details": {}
  },
  "meta": {
    "timestamp": "2025-09-17T10:30:00Z",
    "version": "v1",
    "request_id": "uuid4",
    "pagination": {
      "page": 1,
      "limit": 20,
      "total": 100,
      "has_next": true
    }
  }
}
```

**Success Response Example:**
```json
{
  "success": true,
  "data": {
    "projects": [...],
    "summary": {...}
  },
  "error": null,
  "meta": {
    "timestamp": "2025-09-17T10:30:00Z",
    "version": "v1",
    "request_id": "a1b2c3d4-e5f6-7890"
  }
}
```

**Error Response Example:**
```json
{
  "success": false,
  "data": null,
  "error": {
    "code": "PROJECT_NOT_FOUND",
    "message": "Project with code 'ABC123' was not found",
    "details": {
      "project_code": "ABC123",
      "suggestions": ["ABC124", "ABC125"]
    }
  },
  "meta": {
    "timestamp": "2025-09-17T10:30:00Z",
    "version": "v1",
    "request_id": "a1b2c3d4-e5f6-7890"
  }
}
```

## 3. URL Pattern Standardization

### Current Issues
- Mixed patterns: `/api/*`, `/projects/*`, bare routes
- Inconsistent resource naming
- Non-RESTful patterns

### Proposed RESTful URL Structure
```
/api/v1/resources[/resource_id][/sub_resources[/sub_resource_id]]
```

**Resource Mapping:**
```
Old                               New
/api/projects/list               → /api/v1/projects
/api/projects/<code>             → /api/v1/projects/<code>
/api/update-project-inline       → /api/v1/projects/<code>
/api/monthly-sales               → /api/v1/analytics/sales/monthly
/api/regional-analysis           → /api/v1/analytics/regions
/api/outstanding-analysis        → /api/v1/analytics/outstanding
/api/brand-analysis              → /api/v1/analytics/brands
/api/cache-stats                 → /api/v1/system/cache/stats
/api/cache-clear                 → /api/v1/system/cache
/api/users                       → /api/v1/users
/projects/cancel/<code>          → /api/v1/projects/<code>/cancel
/projects/resume/<code>          → /api/v1/projects/<code>/resume
```

## 4. HTTP Method Standardization

### RESTful Method Mapping
| Resource Operation | HTTP Method | URL Pattern |
|-------------------|-------------|-------------|
| List resources | GET | `/api/v1/projects` |
| Get single resource | GET | `/api/v1/projects/{id}` |
| Create resource | POST | `/api/v1/projects` |
| Update resource (full) | PUT | `/api/v1/projects/{id}` |
| Update resource (partial) | PATCH | `/api/v1/projects/{id}` |
| Delete resource | DELETE | `/api/v1/projects/{id}` |
| Resource actions | POST | `/api/v1/projects/{id}/cancel` |

## 5. Error Handling Standardization

### Standard Error Codes
```python
class APIErrorCodes:
    # Client Errors (4xx)
    VALIDATION_ERROR = "VALIDATION_ERROR"
    RESOURCE_NOT_FOUND = "RESOURCE_NOT_FOUND"
    UNAUTHORIZED = "UNAUTHORIZED"
    FORBIDDEN = "FORBIDDEN"
    CONFLICT = "CONFLICT"
    RATE_LIMITED = "RATE_LIMITED"

    # Server Errors (5xx)
    INTERNAL_ERROR = "INTERNAL_ERROR"
    SERVICE_UNAVAILABLE = "SERVICE_UNAVAILABLE"
    DATABASE_ERROR = "DATABASE_ERROR"
    EXTERNAL_SERVICE_ERROR = "EXTERNAL_SERVICE_ERROR"

    # Business Logic Errors
    PROJECT_ALREADY_EXISTS = "PROJECT_ALREADY_EXISTS"
    PROJECT_LOCKED = "PROJECT_LOCKED"
    INSUFFICIENT_PERMISSIONS = "INSUFFICIENT_PERMISSIONS"
```

### Error Response Middleware
```python
@app.errorhandler(Exception)
def handle_error(error):
    error_response = {
        "success": False,
        "data": None,
        "error": {
            "code": error.code if hasattr(error, 'code') else "INTERNAL_ERROR",
            "message": str(error),
            "details": getattr(error, 'details', {})
        },
        "meta": {
            "timestamp": datetime.utcnow().isoformat() + "Z",
            "version": "v1",
            "request_id": g.request_id
        }
    }

    status_code = getattr(error, 'status_code', 500)
    return jsonify(error_response), status_code
```

## 6. Request/Response Validation Strategy

### Schema Definition with Marshmallow
```python
# Request Schemas
class CreateProjectSchema(Schema):
    name = fields.Str(required=True, validate=Length(min=1, max=100))
    description = fields.Str(validate=Length(max=500))
    start_date = fields.Date(required=True)
    end_date = fields.Date()
    budget = fields.Decimal(places=2, validate=Range(min=0))

# Response Schemas
class ProjectResponseSchema(Schema):
    id = fields.Str(required=True)
    name = fields.Str(required=True)
    description = fields.Str()
    status = fields.Str()
    created_at = fields.DateTime()
    updated_at = fields.DateTime()
```

### Validation Decorators
```python
def validate_json(schema_class):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            schema = schema_class()
            try:
                data = schema.load(request.json)
                g.validated_data = data
                return f(*args, **kwargs)
            except ValidationError as err:
                return create_error_response(
                    "VALIDATION_ERROR",
                    "Request validation failed",
                    err.messages
                ), 400
        return decorated_function
    return decorator
```

## 7. Authentication & Authorization Strategy

### Current State
- Google OAuth for UI authentication
- No API key authentication for programmatic access
- No fine-grained permissions

### Proposed Enhancement
```python
# API Key Authentication for programmatic access
@app.before_request
def authenticate_request():
    if request.path.startswith('/api/'):
        auth_header = request.headers.get('Authorization')
        if auth_header and auth_header.startswith('Bearer '):
            token = auth_header.split(' ')[1]
            # Validate API key or JWT token
        elif 'session' in session:
            # Use existing session auth
        else:
            return create_error_response("UNAUTHORIZED", "Authentication required"), 401
```

## 8. Pagination Strategy

### Standard Pagination Parameters
```
?page=1&limit=20&sort=created_at&order=desc
```

### Pagination Response
```json
{
  "success": true,
  "data": [...],
  "meta": {
    "pagination": {
      "page": 1,
      "limit": 20,
      "total": 150,
      "pages": 8,
      "has_next": true,
      "has_prev": false,
      "next_page": 2,
      "prev_page": null
    }
  }
}
```

## 9. Implementation Architecture

### File Structure
```
dashboard/
├── api/
│   ├── __init__.py
│   ├── v1/
│   │   ├── __init__.py
│   │   ├── projects.py
│   │   ├── analytics.py
│   │   ├── users.py
│   │   ├── system.py
│   │   └── monitoring.py
│   ├── middleware/
│   │   ├── __init__.py
│   │   ├── response_wrapper.py
│   │   ├── error_handler.py
│   │   ├── validation.py
│   │   └── auth.py
│   ├── schemas/
│   │   ├── __init__.py
│   │   ├── projects.py
│   │   ├── users.py
│   │   └── common.py
│   └── utils/
│       ├── __init__.py
│       ├── responses.py
│       └── pagination.py
```

### Blueprint Registration
```python
from api.v1 import projects_v1_bp, analytics_v1_bp, users_v1_bp

app.register_blueprint(projects_v1_bp, url_prefix='/api/v1')
app.register_blueprint(analytics_v1_bp, url_prefix='/api/v1')
app.register_blueprint(users_v1_bp, url_prefix='/api/v1')
```

## 10. Migration Strategy

### Phase 1: Infrastructure Setup
1. Create API middleware components
2. Implement response wrapper
3. Set up error handling
4. Create validation framework

### Phase 2: Gradual Migration
1. Migrate high-priority endpoints first
2. Maintain backward compatibility
3. Add deprecation warnings to old endpoints
4. Monitor usage and performance

### Phase 3: Documentation & Testing
1. Generate OpenAPI specs
2. Create comprehensive tests
3. Build API documentation site
4. Implement monitoring

### Phase 4: Cleanup
1. Remove deprecated endpoints
2. Optimize performance
3. Final testing and validation

## 11. Success Metrics

### Technical Metrics
- API response time consistency
- Error rate reduction
- Documentation coverage
- Test coverage improvement

### Developer Experience Metrics
- Time to integrate new endpoints
- Developer feedback scores
- API discovery time
- Error resolution time

## 12. Next Steps

1. **Immediate**: Implement core middleware components
2. **Week 1**: Migrate 5-10 high-priority endpoints
3. **Week 2**: OpenAPI specification and documentation
4. **Week 3**: Complete migration of remaining endpoints
5. **Week 4**: Testing, optimization, and cleanup