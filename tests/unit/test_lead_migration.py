# -*- coding: utf-8 -*-
"""리드 번호 강등/승격 시 Redis 키·List 마이그레이션 회귀 테스트 (2026-08-06).

L-03491→ETC-429d99 사고: 강등 시 lead_no·시트·카드만 바뀌고 Redis 키(visit_notice_msg
등)·List 방문유형·플랫폼이 옛 번호에 남아 정합성 깨짐. webhook delete+add 불안정 →
직접 API in-place 마이그레이션으로 대체. Redis 키 이관 로직 검증(fake redis).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.blueprints.slack_bot as sb


class _FakeRedis:
    def __init__(self):
        self.store = {}
        self.ttls = {}

    def get(self, k):
        return self.store.get(k)

    def set(self, k, v, ex=None):
        self.store[k] = v
        if ex:
            self.ttls[k] = ex

    def delete(self, k):
        self.store.pop(k, None)
        self.ttls.pop(k, None)

    def ttl(self, k):
        return self.ttls.get(k, -1)


def test_migrate_redis_keys(monkeypatch):
    """강등 시 lead_no 키들이 old→new 로 이관되고 old 는 삭제."""
    fake = _FakeRedis()
    fake.set('visit_notice_msg:L-03491', 'C053Q2X1NP8|123.456')
    fake.set('slack_list_posted:L-03491', '1')
    fake.set('visit_auto_completed:L-03491', '1')
    fake.set('lead_card_msg:L-03491', 'C0BB|789.0')

    class _Wrap:
        redis = fake
    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client', lambda: _Wrap())

    sb._migrate_lead_redis_keys('L-03491', 'ETC-429d99')

    # new 로 이관
    assert fake.get('visit_notice_msg:ETC-429d99') == 'C053Q2X1NP8|123.456'
    assert fake.get('slack_list_posted:ETC-429d99') == '1'
    assert fake.get('visit_auto_completed:ETC-429d99') == '1'
    assert fake.get('lead_card_msg:ETC-429d99') == 'C0BB|789.0'
    # old 삭제
    assert fake.get('visit_notice_msg:L-03491') is None
    assert fake.get('slack_list_posted:L-03491') is None


def test_migrate_redis_keys_absent_noop(monkeypatch):
    """없는 키는 건드리지 않음 (crash X)."""
    fake = _FakeRedis()

    class _Wrap:
        redis = fake
    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client', lambda: _Wrap())
    sb._migrate_lead_redis_keys('L-99999', 'ETC-abc')  # 아무 키도 없음
    assert fake.store == {}


def test_migrate_list_row_no_env(monkeypatch):
    """토큰/list_id 미설정이면 False (호출부 fallback)."""
    monkeypatch.setenv('SLACK_VISIT_BOT_TOKEN', '')
    monkeypatch.setenv('SLACK_BOT_TOKEN', '')
    monkeypatch.setenv('SLACK_VISIT_LIST_ID', '')
    assert sb._migrate_visit_list_row('L-1', 'ETC-1') is False


def test_vlist_option_maps_complete():
    """select 옵션 맵에 방문유형 3종·주요 플랫폼 존재 (스키마 동기화 확인)."""
    assert set(sb._VLIST_VTYPE_OPT) == {'기타', '거래처', '온라인'}
    for p in ('전화', '홈페이지', '거래처', '기타', '당근'):
        assert p in sb._VLIST_PLATFORM_OPT


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
