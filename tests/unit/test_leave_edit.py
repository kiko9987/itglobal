"""휴가 고치기·변경 신청·연차 조정·관리자 등록 (2026-11-18, 플렉스 관리자 기능 맞추기)."""
from datetime import date

import pytest

from dashboard.services import leave as lv

JW, YG, KIKO, MW, SB = 'U051ENK8AAY', 'U050XNG5VBR', 'U04UL2ZLJAX', 'U0BJ9M1MRJ4', 'U0BHC2JV7U5'
TODAY = date(2026, 12, 3)


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(lv, '_db', lambda: d)
    monkeypatch.setattr(lv, '_table_ready', False)
    monkeypatch.setenv('LEAVE_APPROVER_IDS', f'{JW},{YG}')
    monkeypatch.setattr(lv, 'manager_ids', lambda: (SB, KIKO))
    lv.upsert_staff('박민우', '2022-02-03', MW, accrue_from='2026-12-03')
    lv.add_grant('박민우', 5, date(2026, 12, 2), date(2027, 2, 2), 'flex:박민우', '플렉스 잔여')
    return d


def _new(start, end=None, ltype=lv.ANNUAL):
    return lv.create('박민우', MW, MW, '박민우', ltype, '', start, end or start, '개인 사정')


def _approve(rid):
    rid = rid['id'] if isinstance(rid, dict) else rid
    lv.transition(rid, 'approve', JW, '박정우')
    return lv.transition(rid, 'approve', YG, '박용구')


def test_self_edit_pending_restarts_approval_and_reallocates(db):
    r = _new(date(2026, 12, 14), date(2026, 12, 15))            # 월·화 2일
    lv.transition(r['id'], 'approve', JW, '박정우')                 # 1차 승인 → 2차 대기
    assert lv.get(r['id'])['status'] == lv.WAIT_2
    new, errs = lv.edit_request(r['id'], lv.ANNUAL, '', date(2026, 12, 14), date(2026, 12, 16), '개인 사정',
                                MW, '박민우', today=TODAY)
    assert errs == {} and new['days'] == 3 and new['status'] == lv.WAIT_1   # 승인 처음부터
    assert lv.balance('박민우', TODAY) == 2
    line = lv.step_line(new, lv.steps(new['id']))
    assert '✅' not in line and line.startswith('⏳')               # 고치기 전 1차 승인은 지운 셈
    assert '고침' in lv.card_text(new, lv.steps(new['id']), 'requester')


def test_self_edit_checks_balance_without_double_counting_itself(db):
    r = _new(date(2026, 12, 14), date(2026, 12, 18))            # 5일 = 남은 연차 전부
    new, errs = lv.edit_request(r['id'], lv.ANNUAL, '', date(2026, 12, 21), date(2026, 12, 24), '',
                                MW, '박민우', today=TODAY)
    assert errs == {} and new['days'] == 4 and new['reason'] == '개인 사정'   # 사유 비우면 원래 사유
    _, errs = lv.edit_request(r['id'], lv.ANNUAL, '', date(2026, 12, 21), date(2026, 12, 29), '',
                              MW, '박민우', today=TODAY)
    assert '남은 연차' in errs['start']


def test_self_cannot_edit_approved_but_admin_can(db):
    r = _approve(_new(date(2026, 12, 14))['id'])
    _, errs = lv.edit_request(r['id'], lv.ANNUAL, '', date(2026, 12, 15), None, '', MW, '박민우')
    assert '변경 신청' in errs['_']
    _, errs = lv.edit_request(r['id'], lv.ANNUAL, '', date(2026, 12, 15), None, '', JW, '박정우', as_admin=True)
    assert '경영지원' in errs['_']
    new, errs = lv.edit_request(r['id'], lv.ANNUAL, '', date(2026, 12, 15), date(2026, 12, 16), '',
                                SB, '황샛별', as_admin=True, today=TODAY)
    assert errs == {} and new['status'] == lv.APPROVED and new['days'] == 2   # 상태 그대로
    assert lv.balance('박민우', TODAY) == 3
    assert '경영지원 고침' in lv.card_text(new, lv.steps(new['id']), 'requester')


def test_change_request_approved_cancels_original(db):
    orig = _approve(_new(date(2026, 12, 14)))
    assert lv.own_edit_mode(orig, MW, TODAY) == 'change'
    new, errs = lv.change_request(orig['id'], lv.ANNUAL, '', date(2026, 12, 21), date(2026, 12, 22), '날짜 바뀜',
                                  MW, '박민우', today=TODAY)
    assert errs == {} and new['replaces'] == orig['id'] and new['status'] == lv.WAIT_1
    assert lv.own_edit_mode(lv.get(orig['id']), MW, TODAY) == ''        # 변경 신청 대기 중엔 또 못 함
    _, errs = lv.change_request(orig['id'], lv.ANNUAL, '', date(2026, 12, 28), None, '', MW, '박민우', today=TODAY)
    assert '이미 변경 신청' in errs['_']
    assert '바꾸기 전' in lv.card_text(new, lv.steps(new['id']), 'approver')
    done = _approve(new['id'])
    assert done['status'] == lv.APPROVED
    o = lv.get(orig['id'])
    assert o['status'] == lv.CANCELLED and lv.steps(o['id'])[-1]['result'] == 'replaced'
    assert lv.balance('박민우', TODAY) == 3                              # 원래 1일은 돌아오고 새 2일만


def test_change_request_rejected_keeps_original(db):
    orig = _approve(_new(date(2026, 12, 14)))
    new, _ = lv.change_request(orig['id'], lv.ANNUAL, '', date(2026, 12, 21), None, '', MW, '박민우', today=TODAY)
    lv.transition(new['id'], 'reject', JW, '박정우', '안 됨')
    assert lv.get(orig['id'])['status'] == lv.APPROVED
    assert lv.balance('박민우', TODAY) == 4


def test_change_request_can_reuse_its_own_days(db):
    orig = _approve(_new(date(2026, 12, 14), date(2026, 12, 18)))   # 5일 전부 승인
    new, errs = lv.change_request(orig['id'], lv.ANNUAL, '', date(2026, 12, 21), date(2026, 12, 24), '',
                                  MW, '박민우', today=TODAY)
    assert errs == {} and new is not None                              # 원래 휴가 몫을 다시 쓸 수 있음
    _approve(new['id'])
    assert lv.balance('박민우', TODAY) == 1


def test_change_request_by_final_approver_is_immediate(db, monkeypatch):
    lv.upsert_staff('박용구', '2007-01-01', YG, accrue_from='2026-12-03')
    lv.add_grant('박용구', 10, date(2026, 12, 2), date(2026, 12, 31), 'flex:박용구', '')
    orig = lv.create('박용구', YG, YG, '박용구', lv.ANNUAL, '', date(2026, 12, 14), date(2026, 12, 14), 'x')
    assert orig['status'] == lv.APPROVED
    new, _ = lv.change_request(orig['id'], lv.ANNUAL, '', date(2026, 12, 15), None, '', YG, '박용구', today=TODAY)
    assert new['status'] == lv.APPROVED and lv.get(orig['id'])['status'] == lv.CANCELLED
    assert lv.balance('박용구', TODAY) == 9


def test_adjust_plus_minus(db):
    g, err = lv.adjust('박민우', 2, '보상 연차', SB, '황샛별', today=TODAY)
    assert err == '' and g['days'] == 2 and g['valid_to'] == '2027-02-02'
    assert lv.balance('박민우', TODAY) == 7
    g, err = lv.adjust('박민우', -1.5, '결근 차감', SB, '황샛별', today=TODAY)
    assert err == '' and lv.balance('박민우', TODAY) == 5.5
    assert [a['days'] for a in lv.adjustments('박민우')] == [-1.5, 2]
    assert '−1.5일 조정: 결근 차감 (황샛별)' in lv.adjustments('박민우')[0]['note']
    assert lv.adjust('박민우', -9, '많이', SB, '황샛별', today=TODAY)[1].startswith('남은 연차')
    assert lv.adjust('박민우', 0.3, 'x', SB, '황샛별', today=TODAY)[1].startswith('일수는')
    assert lv.adjust('박민우', 1, '', SB, '황샛별', today=TODAY)[1] == '사유를 적어 주세요.'
    assert lv.adjust('박민우', 1, 'x', MW, '박민우', today=TODAY)[1].startswith('경영지원만')
    # 뺀 뒤에도 신청은 남은 만큼만
    r = _new(date(2026, 12, 14), date(2026, 12, 18))
    assert r is not None and lv.balance('박민우', TODAY) == 0.5


def test_admin_register_is_approved_and_does_not_deduct(db):
    row, errs = lv.admin_register('박민우', 'spouse_birth', '', date(2026, 12, 14), date(2026, 12, 18), '출생',
                                  SB, '황샛별')
    assert errs == {} and row['status'] == lv.APPROVED and row['days'] == 5
    assert lv.balance('박민우', TODAY) == 5
    assert '경영지원 등록' in lv.card_text(row, lv.steps(row['id']), 'info')
    _, errs = lv.admin_register('박민우', lv.ANNUAL, '', date(2026, 12, 21), None, '', SB, '황샛별')
    assert 'type' in errs
    _, errs = lv.admin_register('박민우', 'reward', '', date(2026, 12, 16), None, '', SB, '황샛별')
    assert '겹쳐요' in errs['start']
    _, errs = lv.admin_register('박민우', 'reward', '', date(2026, 12, 21), None, '', MW, '박민우')
    assert '경영지원만' in errs['_']
    # 관리자 등록 휴가는 본인이 변경 신청하지 않는다(경영지원이 고침)
    assert lv.own_edit_mode(row, MW, TODAY) == ''


def test_requester_card_has_edit_button(db):
    r = _new(date(2026, 12, 14))
    acts = [b['action_id'] for blk in lv.card_blocks(r, lv.steps(r['id']), 'requester') if blk['type'] == 'actions'
            for b in blk['elements']]
    assert acts == [lv.A_EDIT, lv.A_CANCEL]
