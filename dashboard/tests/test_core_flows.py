"""핵심 순수 함수 회귀 테스트 (2026-07-08 신규).

기존 conftest.py의 import 경로 이슈로 통합 테스트 실행 불가한 상태라
self-contained 방식으로 순수 함수만 검증. Flask app fixture·mocking 불필요.

실행:
    cd "C:\\Users\\SECOM\\Desktop\\ITG-Project\\Claude Project"
    .\\.venv\\Scripts\\python.exe -m pytest dashboard/tests/test_core_flows.py -v
"""

import os
import sys
from pathlib import Path

import pandas as pd
import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(PROJECT_ROOT))

# 환경변수 로드 (일부 함수가 .env 참조)
from dotenv import load_dotenv
load_dotenv(PROJECT_ROOT / ".env", override=True)


# ─────────────────────────────────────────────────────────────
# sanitize_project_for_json — pandas 타입 → JSON 직렬화 가능
# ─────────────────────────────────────────────────────────────

class TestSanitizeProjectForJson:
    def test_none_input(self):
        from dashboard.blueprints.projects import sanitize_project_for_json
        assert sanitize_project_for_json(None) is None

    def test_empty_dict(self):
        from dashboard.blueprints.projects import sanitize_project_for_json
        assert sanitize_project_for_json({}) == {}

    def test_string_pass_through(self):
        from dashboard.blueprints.projects import sanitize_project_for_json
        result = sanitize_project_for_json({'프로젝트 코드': 'G0001-JW', '담당자': '박정우'})
        assert result['프로젝트 코드'] == 'G0001-JW'
        assert result['담당자'] == '박정우'

    def test_pandas_nat_to_none(self):
        from dashboard.blueprints.projects import sanitize_project_for_json
        result = sanitize_project_for_json({'공사 시작': pd.NaT})
        assert result['공사 시작'] is None

    def test_pandas_timestamp_to_string(self):
        from dashboard.blueprints.projects import sanitize_project_for_json
        ts = pd.Timestamp('2026-07-08')
        result = sanitize_project_for_json({'공사 시작': ts})
        assert result['공사 시작'] == '2026-07-08'

    def test_numpy_int_to_python(self):
        import numpy as np
        from dashboard.blueprints.projects import sanitize_project_for_json
        result = sanitize_project_for_json({'총액 1': np.int64(1000000)})
        assert result['총액 1'] == 1000000
        assert isinstance(result['총액 1'], int)

    def test_bool_pass_through(self):
        from dashboard.blueprints.projects import sanitize_project_for_json
        result = sanitize_project_for_json({'수금 확인': False})
        assert result['수금 확인'] is False


# ─────────────────────────────────────────────────────────────
# _to_initial — 한국 이름 → 이니셜 매핑
# ─────────────────────────────────────────────────────────────

class TestToInitial:
    def test_known_name(self):
        from dashboard.blueprints.slack_helpers import _to_initial
        assert _to_initial('박정우') == 'JW'
        assert _to_initial('박용구') == 'YG'
        assert _to_initial('고광일') == 'KiKO'

    def test_english_initial_uppercase(self):
        from dashboard.blueprints.slack_helpers import _to_initial
        assert _to_initial('jw') == 'JW'
        assert _to_initial('YG') == 'YG'

    def test_kiko_special_case(self):
        from dashboard.blueprints.slack_helpers import _to_initial
        assert _to_initial('kiko') == 'KiKO'
        assert _to_initial('KIKO') == 'KiKO'

    def test_empty_input(self):
        from dashboard.blueprints.slack_helpers import _to_initial
        assert _to_initial('') == ''
        assert _to_initial(None) == ''

    def test_dash_or_undefined(self):
        from dashboard.blueprints.slack_helpers import _to_initial
        assert _to_initial('-') == ''
        assert _to_initial('미정') == ''

    def test_unknown_korean_passthrough(self):
        from dashboard.blueprints.slack_helpers import _to_initial
        # 매핑 없는 한국 이름은 원본 반환
        assert _to_initial('홍길동') == '홍길동'


# ─────────────────────────────────────────────────────────────
# folders.py URL 파싱 헬퍼
# ─────────────────────────────────────────────────────────────

class TestExtractFolderIdFromUrl:
    def test_drive_folder_url(self):
        from dashboard.blueprints.folders import _extract_url_folder_id
        url = 'https://drive.google.com/drive/folders/1abcXYZ123456789012345'
        assert _extract_url_folder_id(url) == '1abcXYZ123456789012345'

    def test_folder_url_with_query(self):
        from dashboard.blueprints.folders import _extract_url_folder_id
        url = 'https://drive.google.com/drive/folders/1abcXYZ123456789012345?usp=sharing'
        assert _extract_url_folder_id(url) == '1abcXYZ123456789012345'

    def test_non_url_input(self):
        from dashboard.blueprints.folders import _extract_url_folder_id
        assert _extract_url_folder_id('C:\\Users\\Documents') is None
        assert _extract_url_folder_id('') is None
        assert _extract_url_folder_id(None) is None


class TestExtractLeafAndParentFromWindowsPath:
    def test_typical_windows_path(self):
        from dashboard.blueprints.folders import _extract_leaf_and_parent_from_windows_path
        path = r'G:\내 드라이브\ITG\1. 공사 확정\(박S) 테스트 프로젝트'
        leaf, parent = _extract_leaf_and_parent_from_windows_path(path)
        assert leaf == '(박S) 테스트 프로젝트'
        assert parent == '1. 공사 확정'

    def test_forward_slash_path(self):
        from dashboard.blueprints.folders import _extract_leaf_and_parent_from_windows_path
        path = 'G:/내 드라이브/ITG/(박S) 테스트'
        leaf, parent = _extract_leaf_and_parent_from_windows_path(path)
        assert leaf == '(박S) 테스트'
        assert parent == 'ITG'

    def test_empty_or_none(self):
        from dashboard.blueprints.folders import _extract_leaf_and_parent_from_windows_path
        assert _extract_leaf_and_parent_from_windows_path('') == (None, None)
        assert _extract_leaf_and_parent_from_windows_path(None) == (None, None)


# ─────────────────────────────────────────────────────────────
# 이니셜 매핑 데이터 무결성
# ─────────────────────────────────────────────────────────────

class TestSalesInitialsIntegrity:
    def test_no_duplicate_initials(self):
        """같은 이니셜을 두 사람에게 매핑하면 프로젝트 코드 충돌"""
        from dashboard.blueprints.slack_helpers import SALES_INITIALS
        initials = list(SALES_INITIALS.values())
        assert len(initials) == len(set(initials)), (
            f"중복 이니셜 발견: {[i for i in initials if initials.count(i) > 1]}"
        )

    def test_all_names_korean(self):
        """SALES_INITIALS의 key는 한국 이름이어야 함"""
        from dashboard.blueprints.slack_helpers import SALES_INITIALS
        for name in SALES_INITIALS.keys():
            assert 2 <= len(name) <= 4, f"이름 길이 이상: {name}"


# ─────────────────────────────────────────────────────────────
# 캐시 부분 갱신 함수 시그니처 (import 확인)
# ─────────────────────────────────────────────────────────────

class TestCachePartialUpdate:
    def test_update_project_in_cache_importable(self):
        from dashboard.services.project_service import update_project_in_cache
        assert callable(update_project_in_cache)

    def test_update_project_in_cache_cache_miss(self):
        """캐시 없으면 False 반환 (transient 아닌 명시 실패 신호)"""
        from dashboard.services.project_service import update_project_in_cache
        from dashboard.utils.smart_cache_manager import smart_invalidate
        # 캐시 무효화 후 update 시도 → False (호출자가 fallback해야 함)
        smart_invalidate("current_sheet_data")
        result = update_project_in_cache('NONEXISTENT-XX', {'수금 관련 특이사항': '공사 취소'})
        assert result is False


# ─────────────────────────────────────────────────────────────
# Windows용 시나리오: pytest 실행 시 pytest.ini 없어도 동작
# ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # 직접 실행 시 pytest 호출
    sys.exit(pytest.main([__file__, '-v', '--tb=short']))
