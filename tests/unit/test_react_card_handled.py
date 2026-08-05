# -*- coding: utf-8 -*-
"""카드 처리완료 ✅ 리액션 공통 헬퍼 회귀 테스트 (2026-08-04).

'카드에 붙는 ✅ = 리드 처리 완료' 통일. 모달 완료·기존 lead 연결·자동 스레드
감지 3경로가 _react_card_handled 로 카드 root 에 ✅ 를 단다. 기존은 자동감지만
매니저 답글에 ✅ 를 달아 채널 스캔 시 처리 여부가 안 보였음.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.blueprints.slack_bot as sb


class _FakeClient:
    def __init__(self, raise_exc=None):
        self.calls = []
        self._raise = raise_exc

    def reactions_add(self, channel, timestamp, name):
        self.calls.append({'channel': channel, 'timestamp': timestamp, 'name': name})
        if self._raise:
            raise self._raise
        return {'ok': True}


def test_react_adds_check_to_root():
    """카드 root 에 white_check_mark 를 정확히 1회 단다."""
    c = _FakeClient()
    assert sb._react_card_handled(c, 'C0BB', '123.456') is True
    assert len(c.calls) == 1
    assert c.calls[0] == {'channel': 'C0BB', 'timestamp': '123.456',
                          'name': 'white_check_mark'}


def test_react_empty_args_noop():
    """channel/ts 빈값이면 호출 안 하고 False."""
    c = _FakeClient()
    assert sb._react_card_handled(c, '', '123.456') is False
    assert sb._react_card_handled(c, 'C0BB', '') is False
    assert c.calls == []


def test_react_swallows_error_returns_false():
    """already_reacted 등 예외는 삼키고 False (조용히 죽지 않고 로그만)."""
    c = _FakeClient(raise_exc=Exception('already_reacted'))
    assert sb._react_card_handled(c, 'C0BB', '123.456') is False
    assert len(c.calls) == 1  # 시도는 했음


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
