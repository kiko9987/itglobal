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


def test_build_stage_card_text_middle_stage_no_crash():
    """_build_stage_card_text 중도금/잔금-이력 경로 unpaid 인자 누락 크래시 회귀 (2026-07-29).

    _build_stage_with_history_message 는 unpaid 필수인데 _build_stage_card_text 가
    안 넘겨 중도금 정정 시 TypeError. → unpaid=c.get('unpaid',0) 전달로 수정.
    """
    payments = ps._parse_notes(
        ['일시 07/10\n입금 1,000,000원\n적요 김철수',
         '일시 07/20\n입금 2,000,000원\n적요 김철수', ''],
        stage_vals={'계약금': 1000000, '중도금': 2000000, '잔금': 0})
    c = {'address': 'A', 'construction': 'B', 'invoice': '', 'total_r': 5000000,
         'total_t': 5000000, 'unpaid': 2000000, 'u': 1000000, 'v': 2000000, 'w': 0}
    txt = ps._build_stage_card_text('중도금', 'TEST-01', payments, c,
                                    {'계약금': 1000000, '중도금': 2000000, '잔금': 0})
    assert '중도금 입금' in txt
    assert '미수금 : -2,000,000원' in txt


# ── [특이사항] 재렌더 경로 유지 (2026-10-15 R4101-JSH: 정정 후 '중고매입 상계처리' 사라짐) ──

_R4101_NOTES = ['2026/09/11\n하나,09/11, 13:00\n255******31304\n입금1,200,000원\n김영은',
                '대체전표\n2026-10-15\n400,000원\n중고매입 상계처리',
                '2026/10/12\n농협, 10/12 10:45\n352-****-1682-33\n입금 10,400,000원\n김영은']
_R4101 = {'project': 'R4101-JSH', 'row': 4102, 'stage': '중도금', 'ts': '9.9',
          'u': 1200000, 'v': 400000, 'w': 10400000,
          'address': '시흥 서울대학로278번길 19-14', 'construction': '중고 10마력 천9대 설치공사',
          'invoice': '잔금 - N입금', 'total_r': 12000000, 'total_t': 12000000, 'unpaid': 0}


def test_append_special_notes_inserts_above_last_separator():
    txt = ps._append_special_notes('⠀\n머리\n' + ps._SEP + '\n본문\n' + ps._SEP + '\n미수금 : 0원\n⠀',
                                   _R4101_NOTES)
    assert '[특이사항]\n중도금 중고매입 상계처리\n' + ps._SEP + '\n미수금 : 0원' in txt
    assert ps._append_special_notes('본문', ['', '', '']) == '본문'     # 해당 줄 없으면 그대로


def test_correction_rerender_keeps_special_notes(monkeypatch):
    # 특이사항만 빠졌던 카드 → 정정 시 특이사항 포함 본문으로 갱신 (같은 줄 비교라 안전 가드 통과)
    monkeypatch.setattr(ps, '_fetch_row_notes', lambda sid, sn, row: list(_R4101_NOTES))
    pays = ps._parse_notes(_R4101_NOTES, stage_vals={'계약금': 1200000, '중도금': 400000, '잔금': 10400000})
    bare = ps._build_stage_card_text('중도금', 'R4101-JSH', pays, _R4101,
                                     {'계약금': 1200000, '중도금': 400000, '잔금': 10400000})
    fs = _FakeSlack(bare)
    assert ps._correct_payment_card(fs, 'C', _R4101, 'sid', 'sn') is True
    txt = fs.updates[0]['text']
    assert '중도금 대체수금' in txt and '[특이사항]\n중도금 중고매입 상계처리' in txt


def test_reprice_rerender_keeps_special_notes(monkeypatch):
    monkeypatch.setattr(ps, '_fetch_row_notes', lambda sid, sn, row: list(_R4101_NOTES))
    fs = _FakeSlack('옛 카드')
    rep = dict(_R4101, stage='잔금')
    assert ps._reprice_payment_card(fs, 'C', rep, 'sid', 'sn') is True
    txt = fs.updates[0]['text']
    assert '수금완료' in txt and '중도금  10/15  대체수금  400,000원' in txt
    assert '[특이사항]\n중도금 중고매입 상계처리' in txt


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))


def test_unpaid_line_overpay_distinct_from_owed():
    """카드 맨 아래 줄 — 덜 받음 '미수금 : -X', 더 받음 '과입금 : X (반환 확인)' (2026-10-16 R4104-SD)."""
    assert ps._unpaid_line(2201320) == '미수금 : -2,201,320원'
    assert ps._unpaid_line(0) == '미수금 : 0원'
    assert ps._unpaid_line(-235000) == '과입금 : 235,000원 (반환 확인)'
    assert ps._unpaid_line(-2) == '미수금 : -2원'          # 반올림 끝전은 과입금 아님(기존 표시)
    payments = ps._parse_notes(['일시 09/14\n입금 235,000원\n적요 (주)포마그룹', '',
                                '일시 10/16\n입금 2,750,000원\n적요 박종서'],
                               stage_vals={'계약금': 235000, '중도금': 0, '잔금': 2750000})
    c = {'address': 'A', 'construction': 'B', 'invoice': '', 'total_r': 2500000,
         'total_t': 2750000, 'unpaid': -235000}
    txt = ps._build_stage_card_text('잔금', 'R4104-SD', payments, c,
                                    {'계약금': 235000, '중도금': 0, '잔금': 2750000})
    assert '과입금 : 235,000원 (반환 확인)' in txt and '미수금 :' not in txt


def test_refund_card_full_history_and_rerender_keeps_refund():
    """반환 카드 — 전체 이력(뒤 단계 입금 포함) + 재렌더해도 '반환' 카드 유지 (2026-10-16 R4104-SD)."""
    notes = ['2026/09/14\n하나,09/14, 15:07\n255******31304\n입금235,000원\n(주)포마그룹\n\n'
             '2026/10/16\n하나,10/16, 13:20\n255******31304\n출금235,000원\n(주)포마그룹',
             '', '2026/10/16\n하나,10/16, 12:08\n255******31304\n입금2,750,000원\n박종서']
    sv = {'계약금': 0, '중도금': 0, '잔금': 2750000}
    pays = ps._parse_notes(notes, stage_vals=sv)
    c = {'address': '마포구', 'construction': '컨버터블', 'invoice': '', 'total_r': 2500000,
         'total_t': 2750000, 'unpaid': 0}
    txt = ps._build_stage_card_text('계약금', 'R4104-SD', pays, c, sv)
    assert '*계약금 반환*' in txt and '반환액 : -235,000원' in txt
    assert '잔금  26/10/16  하나 (R)  2,750,000원  박종서' in txt       # 뒤 단계 입금도 이력에
    assert '계약금  26/09/14' in txt and '반환  26/10/16' in txt
    # 일반 단계 카드는 여전히 현재 단계까지만 (다음 단계 미리 적은 메모 차단)
    pays2 = ps._parse_notes([notes[0].split('\n\n')[0], '', notes[2]], stage_vals={'계약금': 235000, '중도금': 0, '잔금': 2750000})
    txt2 = ps._build_stage_card_text('계약금', 'R4104-SD', pays2, dict(c, unpaid=-235000),
                                     {'계약금': 235000, '중도금': 0, '잔금': 2750000})
    assert '*계약금 입금*' in txt2 and '2,750,000원' not in txt2


def test_complete_balance_card_when_paid_by_deposit_stage(monkeypatch):
    """계약금·중도금 입금으로 완납되면 기존 잔금 카드를 수금완료로 재렌더 (2026-10-16 R4101-JSH)."""
    called = []
    monkeypatch.setattr(ps, 'rerender_stage_card', lambda p, s, row=None: called.append((p, s, row)) or True)
    rc = type('R', (), {'get': lambda self, k: '1.1' if k == 'payment_slack:ts:R4101-JSH:잔금' else None})()
    assert ps._complete_balance_card_if_paid(rc, 'R4101-JSH', '계약금', 0, 4102) is True
    assert called == [('R4101-JSH', '잔금', 4102)]
    called.clear()
    assert ps._complete_balance_card_if_paid(rc, 'R4101-JSH', '계약금', 1200000, 4102) is False   # 아직 미수금
    assert ps._complete_balance_card_if_paid(rc, 'R4101-JSH', '잔금', 0, 4102) is False           # 잔금은 원래 경로
    assert ps._complete_balance_card_if_paid(rc, 'G0000-AA', '계약금', 0, 10) is False            # 잔금 카드 없음
    assert called == []
