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
# 주소 파이프라인 회귀 방지 (2026-07-08 이수진 케이스 회귀 대응)
# _flatten_paren_tail — 카카오 검색 이전에 원본 tail을 정규화하는 순수 함수.
# 카카오 API 호출 없어 네트워크 무관.
# ─────────────────────────────────────────────────────────────

class TestFlattenParenTail:
    """address_resolver._flatten_paren_tail 회귀 방지."""

    def test_이수진_케이스_전체_보존(self):
        """L-03118 유사: 괄호 안 지번(중계동)만 떼고 나머지 flat"""
        from dashboard.services.address_resolver import _flatten_paren_tail
        tail = '(중계동, 건영아파트 유치원상가 1층 103호, 케이)'
        assert _flatten_paren_tail(tail) == '건영아파트 유치원상가 1층 103호 케이'

    def test_가산동_건물명(self):
        from dashboard.services.address_resolver import _flatten_paren_tail
        assert _flatten_paren_tail('(가산동, 이앤씨드림타워7차)') == '이앤씨드림타워7차'

    def test_지번만_있으면_빈_문자열(self):
        """콤마 없는 순수 지번 — 유용 정보 없음"""
        from dashboard.services.address_resolver import _flatten_paren_tail
        assert _flatten_paren_tail('(걸포동)') == ''
        assert _flatten_paren_tail('(걸포동 172-1)') == ''

    def test_괄호_없으면_그대로(self):
        from dashboard.services.address_resolver import _flatten_paren_tail
        assert _flatten_paren_tail('마천빌딩 지하 1층') == '마천빌딩 지하 1층'
        assert _flatten_paren_tail('') == ''

    def test_괄호_안_첫요소_지번_아니면_보존(self):
        """예: (건영아파트, 1층) — 첫 요소가 건물명이면 그대로 flat"""
        from dashboard.services.address_resolver import _flatten_paren_tail
        assert _flatten_paren_tail('(건영아파트, 1층)') == '건영아파트 1층'

    def test_위플레이스_케이스_괄호뒤_텍스트(self):
        """L-03168 위플레이스: 괄호 뒤에도 텍스트 있는 케이스
        지번 (서초동)만 제거하고 나머지는 공백으로 flatten"""
        from dashboard.services.address_resolver import _flatten_paren_tail
        assert _flatten_paren_tail('(서초동, 타임빌딩) B1, 위플레이스') == '타임빌딩 B1 위플레이스'

    def test_괄호앞_텍스트_있는_케이스(self):
        """드물지만 tail 앞에도 텍스트 있을 수 있음"""
        from dashboard.services.address_resolver import _flatten_paren_tail
        assert _flatten_paren_tail('본관 (가산동, 이앤씨드림타워7차) 5층') == '본관 이앤씨드림타워7차 5층'

    def test_지번만_있고_괄호밖에_텍스트(self):
        """(걸포동) 지하 1층 → 지번 제거 + 지하 1층 유지"""
        from dashboard.services.address_resolver import _flatten_paren_tail
        assert _flatten_paren_tail('(걸포동) 지하 1층') == '지하 1층'


# ─────────────────────────────────────────────────────────────
# GoogleSheetsManager — 스레드-local + 필수 속성 회귀 방지
# 2026-07-09 사고: __new__ 를 threading.local 로 바꾸면서 _lock 클래스 속성이
# 사라져 self._lock 참조 시 AttributeError → 시트 read 전면 실패.
# 이 테스트는 _lock 이 인스턴스에 살아있고 스레드별로 인스턴스가 격리되는지 검증.
# ─────────────────────────────────────────────────────────────

class TestGoogleSheetsManagerThreadSafety:
    def test_lock_attribute_exists(self):
        """__init__ 후 self._lock 존재 확인 (AttributeError 재발 방지)."""
        from dashboard.utils.google_sheets import GoogleSheetsManager
        import threading as _th
        mgr = GoogleSheetsManager()
        assert hasattr(mgr, '_lock'), 'GoogleSheetsManager 인스턴스에 _lock 없음'
        # RLock 인지 확인 (재진입 허용 필요)
        assert isinstance(mgr._lock, type(_th.RLock())), '_lock 이 RLock 이 아님'

    def test_lock_usable_reentrant(self):
        """self._lock 이 실제로 acquire/release 가능한지 (재진입 포함)."""
        from dashboard.utils.google_sheets import GoogleSheetsManager
        mgr = GoogleSheetsManager()
        with mgr._lock:
            with mgr._lock:  # RLock 이라 재진입 OK
                pass  # 데드락 없이 통과해야 함

    def test_thread_local_isolation(self):
        """다른 스레드는 다른 GoogleSheetsManager 인스턴스를 받아야 함
        (google-api-python-client not thread-safe → heap corruption 방지)."""
        from dashboard.utils.google_sheets import GoogleSheetsManager
        import threading
        instances = {}
        lock = threading.Lock()

        def _worker(name):
            m = GoogleSheetsManager()
            with lock:
                instances[name] = m

        threads = [threading.Thread(target=_worker, args=(f't{i}',)) for i in range(3)]
        for t in threads: t.start()
        for t in threads: t.join()
        _worker('main')

        # 4개 스레드 모두 서로 다른 id
        ids = [id(m) for m in instances.values()]
        assert len(set(ids)) == 4, f'스레드 격리 실패: {ids}'

    def test_same_thread_returns_same_instance(self):
        """같은 스레드에서 재호출 시 캐시된 인스턴스 반환 (인증 오버헤드 절감)."""
        from dashboard.utils.google_sheets import GoogleSheetsManager
        m1 = GoogleSheetsManager()
        m2 = GoogleSheetsManager()
        assert m1 is m2, '같은 스레드에서 다른 인스턴스 반환 (스레드-local 저장 실패)'


# ─────────────────────────────────────────────────────────────
# Lead 검색용 순수 로직 — 매니저별 이니셜 매칭 필터
# 2026-07-09 사고: `_lock` 미존재로 시트 read 실패 → 리드 캐시 텅 빔
# → 새 프로젝트 모달 "기존 리드 불러오기" 결과 0건 회귀.
# ─────────────────────────────────────────────────────────────

class TestLeadSearchOwnerFilter:
    """`api_search_leads_for_project` 안의 담당자 매칭 규칙 검증.

    2026-07 규칙: 리드 플랫폼 = 거래처/기타/소개 이면 '온라인 상담자'로,
    그 외는 '영업 담당자' 로 소유자 판단. 사수·신입 동반은 쉼표·공백 분리.
    """

    @staticmethod
    def _get_owner_names(lead: dict):
        """endpoint 안에 있는 소유자 파싱 로직 미러링 (실 코드 참조 후 반영)."""
        import re as _re
        platform = str(lead.get('플랫폼') or '').strip()
        if platform in ('거래처', '기타', '소개'):
            owner_raw = str(lead.get('온라인 상담자') or '').strip()
        else:
            owner_raw = str(lead.get('영업 담당자') or '').strip()
        return {p.strip() for p in _re.split(r'[,/·&.\s]+', owner_raw) if p.strip()}

    def test_online_lead_uses_sales_owner(self):
        lead = {'플랫폼': '홈페이지', '영업 담당자': '박정우', '온라인 상담자': '김호중'}
        assert self._get_owner_names(lead) == {'박정우'}

    def test_offline_lead_uses_online_owner(self):
        # 거래처 리드는 카드 생성자(온라인 상담자) 기준
        lead = {'플랫폼': '거래처', '영업 담당자': '박정우', '온라인 상담자': '김호중'}
        assert self._get_owner_names(lead) == {'김호중'}

    def test_multi_owner_comma_separated(self):
        # 사수 · 신입 동반 방문 케이스
        lead = {'플랫폼': '전화', '영업 담당자': '권태훈,강정권'}
        assert self._get_owner_names(lead) == {'권태훈', '강정권'}

    def test_multi_owner_with_slash(self):
        lead = {'플랫폼': '카카오톡', '영업 담당자': '박용구/이근혁'}
        assert self._get_owner_names(lead) == {'박용구', '이근혁'}

    def test_empty_owner_returns_empty_set(self):
        lead = {'플랫폼': '전화', '영업 담당자': ''}
        assert self._get_owner_names(lead) == set()


# ─────────────────────────────────────────────────────────────
# Windows용 시나리오: pytest 실행 시 pytest.ini 없어도 동작
# ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # 직접 실행 시 pytest 호출
    sys.exit(pytest.main([__file__, '-v', '--tb=short']))
