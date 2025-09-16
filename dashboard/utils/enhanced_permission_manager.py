"""
향상된 권한 관리 시스템
필드별, 프로젝트별 세분화된 권한 제어
"""

from enum import Enum
from typing import Dict, List, Set, Optional, Any
from dataclasses import dataclass
from functools import wraps
from flask import session, abort, jsonify
import logging

logger = logging.getLogger(__name__)

class Permission(Enum):
    """권한 종류"""
    # 기본 권한
    PROJECT_VIEW = "project_view"
    PROJECT_EDIT = "project_edit"
    PROJECT_CREATE = "project_create"
    PROJECT_DELETE = "project_delete"

    # 필드별 권한
    BASIC_INFO_EDIT = "basic_info_edit"          # 기본정보 (주소, 담당자 등)
    CONSTRUCTION_INFO_EDIT = "construction_info_edit"  # 공사정보
    FINANCIAL_VIEW = "financial_view"            # 금액정보 조회
    FINANCIAL_EDIT = "financial_edit"            # 금액정보 수정
    PAYMENT_VIEW = "payment_view"                # 수금정보 조회
    PAYMENT_EDIT = "payment_edit"                # 수금정보 수정
    PROFIT_VIEW = "profit_view"                  # 손익정보 조회
    PROFIT_EDIT = "profit_edit"                  # 손익정보 수정

    # 관리 권한
    USER_MANAGEMENT = "user_management"
    AUDIT_LOG_VIEW = "audit_log_view"
    SYSTEM_ADMIN = "system_admin"
    CACHE_MANAGEMENT = "cache_management"

    # 지역별 권한
    REGION_ALL = "region_all"
    REGION_SEOUL = "region_seoul"
    REGION_BUSAN = "region_busan"
    REGION_DAEGU = "region_daegu"

class Role(Enum):
    """역할 정의"""
    ADMIN = "admin"
    MANAGER = "manager"           # 지역 관리자
    SALES_SENIOR = "sales_senior" # 수석 영업사원
    SALES_JUNIOR = "sales_junior" # 일반 영업사원
    FINANCE = "finance"           # 재무 담당자
    VIEWER = "viewer"             # 조회만 가능

@dataclass
class UserPermission:
    """사용자 권한 정보"""
    user_id: str
    role: Role
    permissions: Set[Permission]
    regions: Set[str]  # 접근 가능한 지역
    projects: Set[str] = None  # 특정 프로젝트 접근 권한 (None이면 지역 내 모든 프로젝트)

class EnhancedPermissionManager:
    """향상된 권한 관리자"""

    # 역할별 기본 권한 매트릭스
    ROLE_PERMISSIONS = {
        Role.ADMIN: {
            Permission.PROJECT_VIEW, Permission.PROJECT_EDIT, Permission.PROJECT_CREATE, Permission.PROJECT_DELETE,
            Permission.BASIC_INFO_EDIT, Permission.CONSTRUCTION_INFO_EDIT,
            Permission.FINANCIAL_VIEW, Permission.FINANCIAL_EDIT,
            Permission.PAYMENT_VIEW, Permission.PAYMENT_EDIT,
            Permission.PROFIT_VIEW, Permission.PROFIT_EDIT,
            Permission.USER_MANAGEMENT, Permission.AUDIT_LOG_VIEW, Permission.SYSTEM_ADMIN,
            Permission.CACHE_MANAGEMENT, Permission.REGION_ALL
        },

        Role.MANAGER: {
            Permission.PROJECT_VIEW, Permission.PROJECT_EDIT, Permission.PROJECT_CREATE,
            Permission.BASIC_INFO_EDIT, Permission.CONSTRUCTION_INFO_EDIT,
            Permission.FINANCIAL_VIEW, Permission.FINANCIAL_EDIT,
            Permission.PAYMENT_VIEW, Permission.PAYMENT_EDIT,
            Permission.PROFIT_VIEW, Permission.PROFIT_EDIT,
            Permission.AUDIT_LOG_VIEW
        },

        Role.SALES_SENIOR: {
            Permission.PROJECT_VIEW, Permission.PROJECT_EDIT, Permission.PROJECT_CREATE,
            Permission.BASIC_INFO_EDIT, Permission.CONSTRUCTION_INFO_EDIT,
            Permission.FINANCIAL_VIEW, Permission.PAYMENT_VIEW,
            Permission.PROFIT_VIEW
        },

        Role.SALES_JUNIOR: {
            Permission.PROJECT_VIEW, Permission.PROJECT_EDIT,
            Permission.BASIC_INFO_EDIT, Permission.CONSTRUCTION_INFO_EDIT,
            Permission.FINANCIAL_VIEW  # 조회만 가능
        },

        Role.FINANCE: {
            Permission.PROJECT_VIEW,
            Permission.FINANCIAL_VIEW, Permission.FINANCIAL_EDIT,
            Permission.PAYMENT_VIEW, Permission.PAYMENT_EDIT,
            Permission.PROFIT_VIEW, Permission.PROFIT_EDIT,
            Permission.REGION_ALL  # 모든 지역 재무 데이터 접근
        },

        Role.VIEWER: {
            Permission.PROJECT_VIEW
        }
    }

    # 필드별 필요 권한 매핑
    FIELD_PERMISSIONS = {
        # 기본 정보
        '현장명': Permission.BASIC_INFO_EDIT,
        '현장 주소': Permission.BASIC_INFO_EDIT,
        '현장 담당자': Permission.BASIC_INFO_EDIT,
        '담당자 연락처': Permission.BASIC_INFO_EDIT,
        '담당자 이메일': Permission.BASIC_INFO_EDIT,
        '사업자': Permission.BASIC_INFO_EDIT,
        '담당자': Permission.BASIC_INFO_EDIT,

        # 공사 정보
        '공사 종류': Permission.CONSTRUCTION_INFO_EDIT,
        '공사상태': Permission.CONSTRUCTION_INFO_EDIT,
        '시공자': Permission.CONSTRUCTION_INFO_EDIT,
        '도급 구분': Permission.CONSTRUCTION_INFO_EDIT,
        '착공일': Permission.CONSTRUCTION_INFO_EDIT,
        '준공일': Permission.CONSTRUCTION_INFO_EDIT,

        # 금액 정보
        '계약금액': Permission.FINANCIAL_EDIT,
        '총액1': Permission.FINANCIAL_EDIT,
        '총액2': Permission.FINANCIAL_EDIT,
        '부가세': Permission.FINANCIAL_EDIT,

        # 수금 정보
        '수금액': Permission.PAYMENT_EDIT,
        '수금일': Permission.PAYMENT_EDIT,
        '미수금': Permission.PAYMENT_EDIT,

        # 손익 정보
        '제품대': Permission.PROFIT_EDIT,
        '도급비': Permission.PROFIT_EDIT,
        '자재비': Permission.PROFIT_EDIT,
        '기타비': Permission.PROFIT_EDIT,
        '순이익': Permission.PROFIT_EDIT,
        '이익률': Permission.PROFIT_EDIT
    }

    def __init__(self):
        self.user_permissions: Dict[str, UserPermission] = {}

    def get_user_permission(self, user_id: str) -> Optional[UserPermission]:
        """사용자 권한 정보 조회"""
        return self.user_permissions.get(user_id)

    def load_user_permission(self, user_id: str, user_data: Dict[str, Any]) -> UserPermission:
        """사용자 데이터로부터 권한 정보 로드"""
        try:
            role = Role(user_data.get('role', 'viewer'))
        except ValueError:
            role = Role.VIEWER

        # 역할 기반 기본 권한
        base_permissions = self.ROLE_PERMISSIONS.get(role, set())

        # 추가 권한 (사용자별 커스텀)
        additional_permissions = set()
        for perm_str in user_data.get('additional_permissions', []):
            try:
                additional_permissions.add(Permission(perm_str))
            except ValueError:
                logger.warning(f"잘못된 권한: {perm_str}")

        all_permissions = base_permissions | additional_permissions

        # 지역 권한
        regions = set(user_data.get('regions', []))
        if Permission.REGION_ALL in all_permissions:
            regions = {'all'}

        # 특정 프로젝트 권한
        specific_projects = user_data.get('specific_projects')
        projects = set(specific_projects) if specific_projects else None

        user_permission = UserPermission(
            user_id=user_id,
            role=role,
            permissions=all_permissions,
            regions=regions,
            projects=projects
        )

        self.user_permissions[user_id] = user_permission
        return user_permission

    def check_permission(self, user_id: str, permission: Permission) -> bool:
        """권한 확인"""
        user_perm = self.get_user_permission(user_id)
        if not user_perm:
            return False

        return permission in user_perm.permissions

    def can_access_project(self, user_id: str, project_code: str, project_region: str = None) -> bool:
        """프로젝트 접근 권한 확인"""
        user_perm = self.get_user_permission(user_id)
        if not user_perm:
            return False

        # 프로젝트 조회 권한이 없으면 거부
        if Permission.PROJECT_VIEW not in user_perm.permissions:
            return False

        # 특정 프로젝트 권한이 설정된 경우
        if user_perm.projects is not None:
            return project_code in user_perm.projects

        # 지역 권한 확인
        if 'all' in user_perm.regions:
            return True

        if project_region and project_region in user_perm.regions:
            return True

        return False

    def can_edit_field(self, user_id: str, field_name: str, project_code: str = None,
                      project_region: str = None) -> bool:
        """필드 편집 권한 확인"""
        # 프로젝트 접근 권한 먼저 확인
        if project_code and not self.can_access_project(user_id, project_code, project_region):
            return False

        # 필드별 권한 확인
        required_permission = self.FIELD_PERMISSIONS.get(field_name)
        if not required_permission:
            # 매핑되지 않은 필드는 기본 편집 권한으로 확인
            required_permission = Permission.PROJECT_EDIT

        return self.check_permission(user_id, required_permission)

    def get_accessible_fields(self, user_id: str, project_code: str = None,
                            project_region: str = None) -> Dict[str, Dict[str, bool]]:
        """사용자가 접근 가능한 필드 목록"""
        if project_code and not self.can_access_project(user_id, project_code, project_region):
            return {}

        user_perm = self.get_user_permission(user_id)
        if not user_perm:
            return {}

        result = {}
        field_groups = {
            'basic': ['현장명', '현장 주소', '현장 담당자', '담당자 연락처', '담당자 이메일', '사업자', '담당자'],
            'construction': ['공사 종류', '공사상태', '시공자', '도급 구분', '착공일', '준공일'],
            'financial': ['계약금액', '총액1', '총액2', '부가세'],
            'payment': ['수금액', '수금일', '미수금'],
            'profit': ['제품대', '도급비', '자재비', '기타비', '순이익', '이익률']
        }

        for group_name, fields in field_groups.items():
            result[group_name] = {}
            for field in fields:
                required_permission = self.FIELD_PERMISSIONS.get(field, Permission.PROJECT_EDIT)

                # 조회 권한
                view_permission = self._get_view_permission(required_permission)
                can_view = view_permission in user_perm.permissions

                # 편집 권한
                can_edit = required_permission in user_perm.permissions

                result[group_name][field] = {
                    'can_view': can_view,
                    'can_edit': can_edit
                }

        return result

    def _get_view_permission(self, edit_permission: Permission) -> Permission:
        """편집 권한에 대응하는 조회 권한 반환"""
        view_mapping = {
            Permission.FINANCIAL_EDIT: Permission.FINANCIAL_VIEW,
            Permission.PAYMENT_EDIT: Permission.PAYMENT_VIEW,
            Permission.PROFIT_EDIT: Permission.PROFIT_VIEW,
        }
        return view_mapping.get(edit_permission, Permission.PROJECT_VIEW)

    def get_user_role_info(self, user_id: str) -> Dict[str, Any]:
        """사용자 역할 정보 반환 (프론트엔드용)"""
        user_perm = self.get_user_permission(user_id)
        if not user_perm:
            return {'role': 'viewer', 'permissions': [], 'regions': []}

        return {
            'role': user_perm.role.value,
            'permissions': [p.value for p in user_perm.permissions],
            'regions': list(user_perm.regions),
            'projects': list(user_perm.projects) if user_perm.projects else None
        }


# 전역 권한 관리자 인스턴스
enhanced_permission_manager = EnhancedPermissionManager()


def require_permission(permission: Permission):
    """권한 확인 데코레이터"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user' not in session:
                return jsonify({'error': '로그인이 필요합니다'}), 401

            user_id = session['user']['id']
            if not enhanced_permission_manager.check_permission(user_id, permission):
                return jsonify({'error': '권한이 없습니다'}), 403

            return f(*args, **kwargs)
        return decorated_function
    return decorator


def require_project_access(project_code_param='project_code'):
    """프로젝트 접근 권한 확인 데코레이터"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user' not in session:
                return jsonify({'error': '로그인이 필요합니다'}), 401

            user_id = session['user']['id']
            project_code = kwargs.get(project_code_param)

            if not project_code:
                return jsonify({'error': '프로젝트 코드가 필요합니다'}), 400

            # 프로젝트 지역 정보 조회 (실제 구현에서는 데이터베이스에서 조회)
            project_region = get_project_region(project_code)

            if not enhanced_permission_manager.can_access_project(user_id, project_code, project_region):
                return jsonify({'error': '해당 프로젝트에 접근할 권한이 없습니다'}), 403

            return f(*args, **kwargs)
        return decorated_function
    return decorator


def get_project_region(project_code: str) -> str:
    """프로젝트 지역 정보 조회 (실제 구현 필요)"""
    # 실제로는 데이터베이스나 캐시에서 조회
    # 임시로 프로젝트 코드 패턴으로 지역 판단
    if project_code.startswith('S'):
        return 'seoul'
    elif project_code.startswith('B'):
        return 'busan'
    elif project_code.startswith('D'):
        return 'daegu'
    else:
        return 'unknown'


def load_user_permissions_from_session():
    """세션에서 사용자 권한 로드"""
    if 'user' in session:
        user_data = session['user']
        user_id = user_data['id']
        enhanced_permission_manager.load_user_permission(user_id, user_data)