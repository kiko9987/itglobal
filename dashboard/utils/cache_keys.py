"""
캐시 키 상수 정의

하드코딩된 캐시 키를 상수로 관리하여 오타 방지 및 중앙 집중 관리.
CACHE_CONTRACT.md의 네이밍 규칙을 따릅니다.

Usage:
    from dashboard.utils.cache_keys import CacheKeys

    # 기존 (하드코딩)
    df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)

    # 개선 (상수 사용)
    df = smart_get(CacheKeys.PROJECT_LIST, CacheStrategy.CRITICAL_DATA)
"""

from typing import Optional


class CacheKeys:
    """캐시 키 상수 클래스"""

    # ===== 프로젝트 관련 =====

    # 전체 프로젝트 목록 (DataFrame)
    PROJECT_LIST = "current_sheet_data"

    @staticmethod
    def project_detail(project_code: str) -> str:
        """
        개별 프로젝트 상세 캐시 키 (추후 구현)

        Args:
            project_code: 프로젝트 코드 (예: "IT-2024-001")

        Returns:
            str: "project_{project_code}"
        """
        return f"project_{project_code}"

    @staticmethod
    def projects_filter(filter_hash: str) -> str:
        """
        필터링된 프로젝트 목록 캐시 키 (추후 구현)

        Args:
            filter_hash: 필터 조건 해시값

        Returns:
            str: "projects_filter_{hash}"
        """
        return f"projects_filter_{filter_hash}"

    # ===== 셀 노트 관련 =====

    @staticmethod
    def cell_notes(sheet_id: str) -> str:
        """
        전체 셀 노트 캐시 키

        Args:
            sheet_id: Google Sheets ID

        Returns:
            str: "cell_notes_{sheet_id}"
        """
        return f"cell_notes_{sheet_id}"

    @staticmethod
    def notes_row(sheet_id: str, row_number: int) -> str:
        """
        특정 행의 노트 모음 캐시 키 (임시)

        Args:
            sheet_id: Google Sheets ID
            row_number: 행 번호

        Returns:
            str: "notes_row_{sheet_id}_{row_number}"
        """
        return f"notes_row_{sheet_id}_{row_number}"

    # ===== 메타데이터 관련 =====

    @staticmethod
    def metadata(metadata_type: str) -> str:
        """
        메타데이터 캐시 키 (담당자, 사업자, 거래처 등)

        Args:
            metadata_type: 메타데이터 타입 (예: "담당자", "사업자")

        Returns:
            str: "metadata_{type}"
        """
        return f"metadata_{metadata_type}"

    # 사용자 목록
    USER_LIST = "user_list"

    @staticmethod
    def permissions(user_email: str) -> str:
        """
        사용자 권한 정보 캐시 키 (추후 구현)

        Args:
            user_email: 사용자 이메일

        Returns:
            str: "permissions_{user_email}"
        """
        return f"permissions_{user_email}"

    # ===== 리드 관련 =====

    # 전체 리드 목록
    LEADS_LIST = "leads_list"

    @staticmethod
    def lead_detail(lead_no: str) -> str:
        """
        개별 리드 상세 캐시 키

        Args:
            lead_no: 리드 번호

        Returns:
            str: "lead_{lead_no}"
        """
        return f"lead_{lead_no}"

    # ===== 폴더 관련 =====

    @staticmethod
    def folder_id(folder_type: str) -> str:
        """
        Google Drive 폴더 ID 캐시 키

        Args:
            folder_type: 폴더 타입 (예: "projects", "documents")

        Returns:
            str: "folder_id_{folder_type}"
        """
        return f"folder_id_{folder_type}"

    # ===== 기타 =====

    # 설정 정보 (추후 구현)
    CONFIG = "config"

    @staticmethod
    def session(session_id: str) -> str:
        """
        세션 캐시 키 (추후 구현)

        Args:
            session_id: 세션 ID

        Returns:
            str: "session_{session_id}"
        """
        return f"session_{session_id}"


# Alias for backward compatibility
CACHE_KEYS = CacheKeys


# ===== 캐시 키 패턴 (와일드카드용) =====

class CacheKeyPatterns:
    """캐시 키 패턴 (invalidate_cache에서 와일드카드로 사용)"""

    # 모든 셀 노트
    ALL_CELL_NOTES = "cell_notes_*"

    # 모든 프로젝트 상세
    ALL_PROJECTS = "project_*"

    # 모든 메타데이터
    ALL_METADATA = "metadata_*"

    # 모든 폴더 ID
    ALL_FOLDER_IDS = "folder_id_*"

    # 모든 리드
    ALL_LEADS = "lead_*"


# ===== 사용 예시 =====

if __name__ == "__main__":
    # 프로젝트 목록 키
    print(CacheKeys.PROJECT_LIST)  # "current_sheet_data"

    # 프로젝트 상세 키
    print(CacheKeys.project_detail("IT-2024-001"))  # "project_IT-2024-001"

    # 셀 노트 키
    print(CacheKeys.cell_notes("1a2b3c4d5e"))  # "cell_notes_1a2b3c4d5e"

    # 메타데이터 키
    print(CacheKeys.metadata("담당자"))  # "metadata_담당자"

    # 폴더 ID 키
    print(CacheKeys.folder_id("projects"))  # "folder_id_projects"

    # 패턴 (와일드카드)
    print(CacheKeyPatterns.ALL_CELL_NOTES)  # "cell_notes_*"
