# -*- coding: utf-8 -*-
"""입금 메모 정정 → 카드 갱신 + 스레드 스냅샷 회귀 테스트 (2026-07-29).

메모 정정 시 기존 카드를 스레드에 스냅샷으로 남기고 최신 파싱으로 chat_update.
안전 가드: 정정된 단계 외 다른 줄이 바뀌면(예: 매출이동 은행코드 재계산) 자동갱신 skip.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.payment_sync as ps


_NOTES = ['일시 07/21, 16:53\n입금 200,000원\n계좌번호 255******31304\n적요 신세게유통',
          '', '2026-07-29\n2,490,000원\n현금']

_OLD_R = ('⠀\n:white_check_mark: *수금완료* — :id: *R3888-JSH*\n'
          '--------------------------------------------\n'
          '주소 : 인천 남동구 인하로 513-1 구월빌딩 1층\n'
          '공사내용 : LG 싱글 냉난방 4W 28평 설치 공사\n\n[입금 이력]\n'
          '계약금  07/21  하나 (R)  200,000원  신세게유통\n'
          '잔금  07/29  N  50,000원  현금 수령\n'
          '--------------------------------------------\n총액 : 2,690,000원\n⠀')

_CORR = {'project': 'R3888-JSH', 'row': 3889, 'stage': '잔금', 'ts': '123.45',
         'u': 200000, 'v': 0, 'w': 2490000, 'aa': True,
         'address': '인천 남동구 인하로 513-1 구월빌딩 1층',
         'construction': 'LG 싱글 냉난방 4W 28평 설치 공사', 'invoice': '',
         'total_r': 2690000, 'total_t': 2690000, 'unpaid': 0}


class _FakeSlack:
    def __init__(self, old):
        self.old = old; self.threads = []; self.updates = []
    def conversations_history(self, **k):
        return {'messages': [{'text': self.old}]}
    def chat_postMessage(self, **k):
        self.threads.append(k)
    def chat_update(self, **k):
        self.updates.append(k)


@pytest.fixture(autouse=True)
def _patch(monkeypatch):
    import dashboard.blueprints.slack_helpers as sh
    monkeypatch.setattr(sh, 'safe_slack_call', lambda fn, **kw: fn(**kw))
    monkeypatch.setattr(ps, '_fetch_row_notes', lambda sid, sn, row: list(_NOTES))


def test_clean_correction_updates_and_snapshots():
    """계약금 줄 일치·잔금만 stale → 카드 갱신(2,490,000) + 스레드 스냅샷(50,000)."""
    fs = _FakeSlack(_OLD_R)
    assert ps._correct_payment_card(fs, 'C', _CORR, 'sid', 'sn') is True
    assert len(fs.updates) == 1 and '2,490,000원' in fs.updates[0]['text']
    assert len(fs.threads) == 1 and '50,000원' in fs.threads[0]['text']
    assert fs.threads[0].get('thread_ts') == '123.45'


def test_other_line_changed_skips():
    """계약금이 (N)으로 표기(매출이동 수동조정) → 파서 재계산(R)과 달라 자동갱신 skip."""
    fs = _FakeSlack(_OLD_R.replace('하나 (R)', '하나 (N)'))
    assert ps._correct_payment_card(fs, 'C', _CORR, 'sid', 'sn') is False
    assert not fs.updates and not fs.threads


def test_unified_card_skipped(monkeypatch):
    """통합 카드는 v1 대상 아님 → skip."""
    fs = _FakeSlack(_OLD_R.replace('수금완료', '수금완료 (통합 입금)'))
    assert ps._correct_payment_card(fs, 'C', _CORR, 'sid', 'sn') is False


def test_incomplete_parse_skips(monkeypatch):
    """새 파싱이 미완성(거래처 없음)이면 좋은 카드 안 덮어씀."""
    monkeypatch.setattr(ps, '_fetch_row_notes',
                        lambda sid, sn, row: ['', '', '2026-07-29\n2,490,000원'])  # 거래처 없음
    fs = _FakeSlack(_OLD_R)
    assert ps._correct_payment_card(fs, 'C', _CORR, 'sid', 'sn') is False


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
