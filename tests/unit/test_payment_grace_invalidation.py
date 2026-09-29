# -*- coding: utf-8 -*-
"""grace 통합 그룹 무효화 회귀 테스트 (2026-09-29 G4109-YG↔G4133-YM).

카드가 회색화/삭제되면 그 카드를 참조하던 grace pending 그룹(payment_pending_group)이
역인덱스(payment_grace_ref:{ts})를 통해 함께 지워져야 한다. 안 지우면 같은 시그니처
재입금이 유령 ts 를 chat.update 해 message_not_found(삭제) 또는 취소 카드 오부활(회색).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.payment_sync import _drop_grace_group_by_ts
from dashboard.utils.redis_client import get_redis_client


@pytest.fixture
def rc():
    return get_redis_client().redis


def test_drop_removes_group_and_ref(rc):
    ts = '9999999999.000001'
    gk = 'payment_pending_group:잔금:deadbeefdeadbeef'
    rc.set(gk, '{"ts":"9999999999.000001"}', ex=600)
    rc.set(f'payment_grace_ref:{ts}', gk, ex=600)
    _drop_grace_group_by_ts(ts)
    assert rc.get(gk) is None
    assert rc.get(f'payment_grace_ref:{ts}') is None


def test_drop_noop_when_no_ref(rc):
    # ref 없으면 조용히 무시 (예외 없이)
    ts = '9999999999.000002'
    rc.delete(f'payment_grace_ref:{ts}')
    _drop_grace_group_by_ts(ts)  # 예외 안 나면 통과
    assert rc.get(f'payment_grace_ref:{ts}') is None


def test_drop_empty_ts_safe():
    _drop_grace_group_by_ts('')   # 빈 ts 무해
    _drop_grace_group_by_ts(None)


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
