# -*- coding: utf-8 -*-
"""재상담 신원 덮어쓰기 방어 회귀 테스트 (2026-07-31 L-03367).

배경: 매니저(JK)가 다른 고객(SD)의 완료 카드에서 [재상담] 을 눌러 전혀 다른 고객
정보를 입력 → 기존 리드 고객명·주소가 조용히 덮어써져 원 리드(성남 등자로 56)가
소실. _consult_identity_changed 가 신원 변경을 감지해 확인 게이트/백업을 트리거한다.
같은 고객 재상담(pre-fill 그대로)은 안 걸려야 정상 흐름에 마찰이 없다.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.blueprints.slack_bot import _consult_identity_changed


class TestConsultIdentityChanged:
    def test_same_customer_no_change(self):
        old = {'고객명': '센부동산', '고객 연락처': '010-8247-4206'}
        assert _consult_identity_changed(old, '센부동산', '010-8247-4206')['changed'] is False

    def test_l03367_name_changed_same_phone(self):
        """실제 사고: 의뢰인→센부동산, 전화 동일(pre-fill 그대로) → name 변경 감지."""
        old = {'고객명': '의뢰인', '고객 연락처': '010-8247-4206'}
        r = _consult_identity_changed(old, '센부동산', '010-8247-4206')
        assert r['changed'] is True
        assert r['name'] == ('의뢰인', '센부동산')
        assert r['contact'] is None

    def test_contact_changed(self):
        old = {'고객명': '홍길동', '고객 연락처': '010-1111-2222'}
        r = _consult_identity_changed(old, '홍길동', '010-3333-4444')
        assert r['changed'] is True
        assert r['contact'][0] == '010-1111-2222'
        assert r['name'] is None

    def test_whitespace_hyphen_only_no_change(self):
        """공백·하이픈만 다르면 동일 취급 (오탐 방지)."""
        old = {'고객명': '김 사장', '고객 연락처': '010-1234-5678'}
        assert _consult_identity_changed(old, '김사장', '01012345678')['changed'] is False

    def test_empty_old_name_no_change(self):
        """옛 이름 없으면 비교 불가 → 변경 아님(이름 채우기는 무해)."""
        old = {'고객명': '', '고객 연락처': '010-1234-5678'}
        assert _consult_identity_changed(old, '새이름', '010-1234-5678')['changed'] is False

    def test_empty_new_values_no_change(self):
        """새 값이 비면(필드 미입력) 변경 아님."""
        old = {'고객명': '홍길동', '고객 연락처': '010-1234-5678'}
        assert _consult_identity_changed(old, '', '')['changed'] is False


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
