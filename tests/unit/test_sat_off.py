# -*- coding: utf-8 -*-
"""토요일 휴무(기술영업 격주휴무, 2026-11-20 부장님).

- 기술영업은 달력 토요일에 자기 이름을 넣고 뺀다. 취합 담당(빈 대리)은 남의 이름도.
- 취합 담당이 결재를 올리면 박정우·박용구 둘 중 한 분 승인으로 확정. 돌려보내면 다시 작성 중.
- 확정된 달은 변경 신청 → 둘 중 한 분 승인. 승인권자는 바로 고칠 수 있다.
- 확정된 달은 휴가 일수에 쓰인다(쉬는 토요일은 빼고 근무 토요일은 셈).
슬랙은 가짜 클라이언트로 대신한다(운영 슬랙으로 아무것도 안 나감).
"""
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest
from cryptography.fernet import Fernet

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import sat_off as so

ADMIN = 'kiko@itg-aircon.com'
JW, YG, CO, KIKO = 'UJW', 'UYG', 'UCO', 'UKIKO'
A, B, C = 'UA0001', 'UB0001', 'UC0001'
TODAY = date(2026, 11, 20)
DEC = '2026-12'
S5, S12, S19, S26 = (date(2026, 12, d) for d in (5, 12, 19, 26))


class FakeSlack:
    def __init__(self):
        self.posts = []

    def conversations_open(self, users):
        return {'channel': {'id': f'D-{users}'}}

    def chat_postMessage(self, channel, text='', blocks=None):
        self.posts.append({'ch': channel, 'text': text})
        return {'ts': '1.1'}

    def to(self, uid):
        return [p['text'] for p in self.posts if p['ch'] == f'D-{uid}']


@pytest.fixture
def slack(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, lv):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('LEAVE_APPROVER_IDS', f'{JW},{YG}')
    monkeypatch.setenv('SAT_OFF_COORDINATOR_IDS', CO)
    monkeypatch.setenv('LEAVE_ANNOUNCE_CHANNEL', 'C-ANNOUNCE')
    monkeypatch.setattr(lv, 'manager_ids', lambda: (KIKO,))
    gw._invalidate()
    people = (('영업가', A, '기술영업'), ('영업나', B, '기술영업'), ('영업다', C, '기술영업'),
              ('빈대리', CO, '기술영업'), ('박정우', JW, 'CEO'), ('박용구', YG, 'CEO'),
              ('박민우', 'UMW', '영업'), ('경영', KIKO, '경영지원'))
    for name, uid, dept in people:
        hr.create_person({'name': name, 'status': '재직', 'hire_date': '2022-05-23', 'department': dept,
                          'slack_id': uid, 'work_email': f'{uid.lower()}@itg-aircon.com'}, ADMIN)
        lv.upsert_staff(name, '2022-05-23', uid, accrue_from='2026-11-11')
        lv.add_grant(name, 15, date(2026, 11, 10), date(2027, 11, 9), f'flex:{name}', '플렉스 잔여')
    fake = FakeSlack()
    monkeypatch.setattr(so, '_client', lambda: fake)
    monkeypatch.setattr(so, '_run_bg', lambda fn: fn())
    yield fake
    gw._invalidate()


def _confirm(month=DEC):
    so.submit(month, CO, '빈대리', TODAY)
    so.decide(month, True, JW, '박정우')


# ── 대상·고르기 ─────────────────────────────────────────────────────────
def test_members_are_tech_sales_only(slack):
    assert [m['name'] for m in so.members()] == ['빈대리', '영업가', '영업나', '영업다']   # '영업'(박민우)은 아님
    assert so.is_member('영업가') and not so.is_member('박민우')
    assert so.role(A, '영업가') == {'member': True, 'coordinator': False, 'approver': False, 'manager': False}
    assert so.role(CO, '빈대리')['coordinator'] and so.role(CO, '빈대리')['manager']
    assert so.role(YG, '박용구')['approver'] and not so.role(YG, '박용구')['member']


def test_member_toggles_own_name_only(slack):
    r = so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    assert r['names'] == ['영업가'] and r['count'] == 1 and not r['warn']
    assert so.picks(DEC)['2026-12-05'] == ['영업가']
    with pytest.raises(PermissionError):
        so.toggle(S5, '영업나', True, A, '영업가', TODAY)                  # 남의 이름은 안 됨
    so.toggle(S5, '영업가', False, A, '영업가', TODAY)
    assert so.picks(DEC)['2026-12-05'] == []
    with pytest.raises(so.SatOffError, match='대상'):
        so.toggle(S5, '박민우', True, 'UMW', '박민우', TODAY)              # 기술영업이 아님


def test_day_rules(slack):
    with pytest.raises(so.SatOffError, match='토요일만'):
        so.toggle(date(2026, 12, 4), '영업가', True, A, '영업가', TODAY)
    with pytest.raises(so.SatOffError, match='지난 토요일'):
        so.toggle(date(2026, 11, 14), '영업가', True, A, '영업가', TODAY)
    with pytest.raises(so.SatOffError, match='개천절'):
        so.toggle(date(2026, 10, 3), '영업가', True, A, '영업가', date(2026, 9, 1))
    with pytest.raises(so.SatOffError, match='6달 뒤'):
        so.toggle(date(2027, 6, 5), '영업가', True, A, '영업가', TODAY)


def test_coordinator_picks_for_others_without_monthly_limit(slack):
    """한 달 몇 번 기준은 없다(11-24 부장님 — 급한 주말 공사로 자주 바뀜). 3번째도 경고 없이."""
    for d in (S5, S12):
        so.toggle(d, '영업나', True, CO, '빈대리', TODAY)
    r = so.toggle(S19, '영업나', True, CO, '빈대리', TODAY)
    assert r['count'] == 3 and r['warn'] == ''
    v = so.view(DEC, CO, '빈대리', TODAY)
    assert v['pick_others'] and v['can_submit'] and not v['can_decide'] and 'limit' not in v
    assert {m['name']: m['count'] for m in v['members']}['영업나'] == 3


# ── 휴무 출근(11-24 부장님) ─────────────────────────────────────────────
def test_worked_after_verbal_approval_by_self(slack):
    """쉬기로 한 토요일에 급히 나옴 — 구두 승인 뒤 본인이 기록. 쉬는 사람에서 빠지고 휴무 출근으로 남는다."""
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    so.toggle(S12, '영업가', True, A, '영업가', TODAY)
    with pytest.raises(so.SatOffError, match='작성 중'):
        so.mark_worked(S5, '영업가', '박정우', '현장', A, '영업가', TODAY)       # 작성 중엔 이름을 빼면 됨
    _confirm()
    v = so.view(DEC, A, '영업가', TODAY)
    assert v['can_mark_worked'] and v['approvers'] == ['박정우', '박용구']
    with pytest.raises(so.SatOffError, match='구두로 승인'):
        so.mark_worked(S5, '영업가', '영업나', '현장', A, '영업가', TODAY)       # 승인권자가 아님
    with pytest.raises(so.SatOffError, match='어떤 일'):
        so.mark_worked(S5, '영업가', '박정우', ' ', A, '영업가', TODAY)
    with pytest.raises(so.SatOffError, match='쉬는 날이 아니'):
        so.mark_worked(S19, '영업가', '박정우', '현장', A, '영업가', TODAY)
    with pytest.raises(PermissionError):
        so.mark_worked(S5, '영업가', '박정우', '현장', B, '영업나', TODAY)       # 남의 것은 취합 담당·승인권자만
    r = so.mark_worked(S5, '영업가', '박정우', '강남 현장 공사', A, '영업가', TODAY)
    assert r['verbal_by'] == '박정우' and r['label'] == '12/5(토)'
    assert so.picks(DEC)['2026-12-05'] == [] and so.counts(DEC)['영업가'] == 1     # 쉬는 사람에서 빠짐
    v = so.view(DEC, A, '영업가', TODAY)
    assert [s['worked'] for s in v['saturdays']][:2] == [['영업가'], []]
    assert {m['name']: (m['count'], m['worked']) for m in v['members']}['영업가'] == (1, 1)
    assert v['worked'][0]['reason'] == '강남 현장 공사' and v['worked'][0]['can_undo']
    assert not so.view(DEC, B, '영업나', TODAY)['worked'][0]['can_undo']        # 남은 되돌리지 못함
    assert any('휴무 출근 (구두 승인 박정우)' in x['detail'] for x in so.recent_log(DEC))
    # 근무로 바뀐 날 — 휴가 일수·차량 배정은 근무로 본다
    assert so.work_saturdays('영업가', S5, S12) == {S5: True, S12: False}
    assert lv.count_days(lv.ANNUAL, S5, S5, '영업가') == 1
    assert not slack.to(JW) or not any('휴무 출근' in t for t in slack.to(JW))     # 승인권자 DM 없음(구두 승인)


def test_worked_undo_and_re_add_clear_record(slack):
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    _confirm()
    so.mark_worked(S5, '영업가', '박용구', '현장', CO, '빈대리', TODAY)          # 취합 담당이 대신 기록
    with pytest.raises(PermissionError):
        so.undo_worked(S5, '영업가', B, '영업나')
    so.undo_worked(S5, '영업가', A, '영업가')
    assert so.picks(DEC)['2026-12-05'] == ['영업가'] and so.worked(DEC) == []
    with pytest.raises(so.SatOffError, match='기록이 없'):
        so.undo_worked(S5, '영업가', A, '영업가')
    # 출근 기록 뒤 승인권자가 다시 쉬는 날로 넣으면 출근 기록은 지운다(둘 다 남지 않게)
    so.mark_worked(S5, '영업가', '박정우', '현장', A, '영업가', TODAY)
    so.toggle(S5, '영업가', True, JW, '박정우', TODAY)
    assert so.picks(DEC)['2026-12-05'] == ['영업가'] and so.worked(DEC) == []


def test_worked_past_day_within_a_month(slack):
    """다음 주에 적어도 된다 — 한 달 안 지난 토요일까지. 그보다 오래된 건 승인권자만."""
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    _confirm()
    so.mark_worked(S5, '영업가', '박정우', '현장', A, '영업가', date(2026, 12, 9))
    so.undo_worked(S5, '영업가', A, '영업가')
    with pytest.raises(so.SatOffError, match='한 달보다'):
        so.mark_worked(S5, '영업가', '박정우', '현장', A, '영업가', date(2027, 1, 20))
    so.mark_worked(S5, '영업가', '박정우', '현장', JW, '박정우', date(2027, 1, 20))


# ── 결재 ───────────────────────────────────────────────────────────────
def test_submit_and_one_of_two_approves(slack):
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    with pytest.raises(PermissionError):
        so.submit(DEC, A, '영업가', TODAY)                               # 직원은 못 올림
    so.submit(DEC, CO, '빈대리', TODAY)
    assert so.month_row(DEC)['status'] == so.WAIT
    with pytest.raises(so.SatOffError, match='결재 대기'):
        so.toggle(S12, '영업가', True, A, '영업가', TODAY)                # 결재 중엔 잠김
    with pytest.raises(PermissionError):
        so.decide(DEC, True, CO, '빈대리')                               # 취합 담당은 승인 못 함
    so.decide(DEC, True, YG, '박용구')                                  # 대표님 한 분으로 확정
    assert so.month_row(DEC)['status'] == so.CONFIRMED
    with pytest.raises(so.SatOffError, match='결재 대기 중인 달이 아니'):
        so.decide(DEC, True, JW, '박정우')                               # 나머지 한 분은 할 일 없음


def test_return_needs_reason_and_reopens(slack):
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    so.submit(DEC, CO, '빈대리', TODAY)
    with pytest.raises(so.SatOffError, match='이유'):
        so.decide(DEC, False, JW, '박정우')
    so.decide(DEC, False, JW, '박정우', '12일 현장 겹침')
    row = so.month_row(DEC)
    assert row['status'] == so.OPEN and row['note'] == '12일 현장 겹침'
    so.toggle(S12, '영업가', True, A, '영업가', TODAY)                    # 다시 고칠 수 있음


def test_notify_messages(slack):
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    so.submit(DEC, CO, '빈대리', TODAY)
    so.notify('submit', DEC, '빈대리')
    assert any('2026년 12월 스케줄 근무표' in t and '12/5(토) 영업가' in t for t in slack.to(JW))
    assert slack.to(YG)
    so.decide(DEC, True, JW, '박정우')
    so.notify('approve', DEC, '박정우')
    assert any('승인했어요' in t for t in slack.to(CO))
    assert [p['text'] for p in slack.posts if p['ch'] == 'C-ANNOUNCE'] == \
        ['📅 2026년 12월 스케줄 근무표가 확정됐어요. 토요일에 쉬는 사람이에요.\n• 12/5(토) 영업가']
    assert any('토요일에 쉬는 사람\n• 12/5(토) 영업가' in t for t in slack.to(JW))   # 이름이 쉬는 사람임을 밝힘


# ── 확정 뒤 바꾸기 ──────────────────────────────────────────────────────
def test_change_request_flow(slack):
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    _confirm()
    with pytest.raises(so.SatOffError, match='변경 신청'):
        so.toggle(S12, '영업가', True, A, '영업가', TODAY)
    with pytest.raises(so.SatOffError, match='이유'):
        so.request_change(DEC, '영업가', ['2026-12-19'], ['2026-12-05'], '', A, '영업가', TODAY)
    c = so.request_change(DEC, '영업가', ['2026-12-19'], ['2026-12-05'], '5일 현장 시공', A, '영업가', TODAY)
    assert c['status'] == so.CH_WAIT and c['what'] == '12/5(토) → 12/19(토)'
    assert so.picks(DEC)['2026-12-05'] == ['영업가']                     # 승인 전엔 그대로
    assert not so.view(DEC, A, '영업가', TODAY)['can_request_change']   # 대기 중엔 새로 고르지 않게
    with pytest.raises(so.SatOffError, match='이미 결재 대기'):
        so.request_change(DEC, '영업가', ['2026-12-26'], [], '또', A, '영업가', TODAY)
    so.notify('change_request', DEC, '영업가', change=c)
    assert any('영업가 님 스케줄 근무 변경 신청' in t for t in slack.to(YG))
    c = so.decide_change(c['id'], True, JW, '박정우', today=TODAY)
    assert c['status'] == so.CH_OK
    assert so.picks(DEC)['2026-12-05'] == [] and so.picks(DEC)['2026-12-19'] == ['영업가']
    so.notify('change_approve', DEC, '박정우', change=c)
    assert any('승인했어요' in t for t in slack.to(A))
    assert any(p['ch'] == 'C-ANNOUNCE' and '영업가 쉬는 토요일 12/5(토) → 12/19(토)' in p['text'] for p in slack.posts)


def test_change_reject_and_cancel(slack):
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    _confirm()
    c = so.request_change(DEC, '영업가', [], ['2026-12-05'], '근무', A, '영업가', TODAY)
    with pytest.raises(so.SatOffError, match='이유'):
        so.decide_change(c['id'], False, YG, '박용구', today=TODAY)
    assert so.decide_change(c['id'], False, YG, '박용구', '그날 인원 부족', TODAY)['status'] == so.CH_NO
    assert so.picks(DEC)['2026-12-05'] == ['영업가']
    c = so.request_change(DEC, '영업가', ['2026-12-12'], [], '가족 행사', A, '영업가', TODAY)
    with pytest.raises(PermissionError):
        so.cancel_change(c['id'], B, '영업나')
    assert so.cancel_change(c['id'], A, '영업가')['status'] == so.CH_CANCEL


def test_approver_edits_confirmed_month_directly(slack):
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    _confirm()
    so.toggle(S12, '영업나', True, JW, '박정우', TODAY)                  # 승인권자는 바로
    assert so.picks(DEC)['2026-12-12'] == ['영업나']
    assert any('확정 뒤 직접 수정' in l['detail'] for l in so.recent_log(DEC))
    c = so.request_change(DEC, '영업가', ['2026-12-26'], ['2026-12-05'], '조정', YG, '박용구', TODAY)
    assert c['status'] == so.CH_OK and so.picks(DEC)['2026-12-26'] == ['영업가']


# ── 휴가 일수 ──────────────────────────────────────────────────────────
def test_leave_days_follow_confirmed_schedule(slack):
    a, b = date(2026, 12, 9), date(2026, 12, 14)                         # 수~월, 12(토) 낌
    assert lv.count_days(lv.ANNUAL, a, b, '영업가') == 4                 # 확정 전: 가운데 토요일은 안 셈(예전 규칙)
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    so.toggle(S19, '영업나', True, B, '영업나', TODAY)
    _confirm()
    assert lv.count_days(lv.ANNUAL, a, b, '영업가') == 5                 # 12일은 근무 토요일 → 셈
    assert lv.count_days(lv.ANNUAL, S5, S5, '영업가') == 0               # 쉬는 토요일 → 0
    assert '쉬는 날' in lv.validate('영업가', lv.ANNUAL, '', S5, S5)['start']
    assert lv.count_days(lv.ANNUAL, S12, S12, '영업가') == 1
    assert lv.count_days(lv.ANNUAL, date(2026, 12, 17), date(2026, 12, 21), '영업나') == 3   # 19 쉬는 토 빼고 17·18·21
    assert lv.count_days(lv.ANNUAL, a, b, '박민우') == 4                 # '영업'(대상 아님)은 예전 규칙 그대로
    assert lv.count_days(lv.ANNUAL, date(2027, 1, 6), date(2027, 1, 11), '영업가') == 4   # 확정 안 된 달은 예전 규칙


def test_affected_leaves_listed_after_confirm(slack):
    a, b = date(2026, 12, 9), date(2026, 12, 14)
    r = lv.create('영업가', A, A, '영업가', lv.ANNUAL, '', a, b, '가족 여행')
    assert r['days'] == 4
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    _confirm()
    bad = so.affected_leaves(DEC)
    assert [(x['name'], x['saved'], x['now']) for x in bad] == [('영업가', 4, 5)]
    so.notify('approve', DEC, '박정우')
    assert any('휴가 일수가 달라지는' in t and '영업가' in t for t in slack.to(KIKO))


# ── 화면 데이터 ────────────────────────────────────────────────────────
def test_view_flags_by_role(slack):
    v = so.view(DEC, A, '영업가', TODAY)
    assert v['status_label'] == '작성 중' and v['can_toggle'] and not v['pick_others'] and not v['can_submit']
    assert [s['label'] for s in v['saturdays']] == ['12/5(토)', '12/12(토)', '12/19(토)', '12/26(토)']
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    _confirm()
    v = so.view(DEC, A, '영업가', TODAY)
    assert v['can_request_change'] and not v['can_toggle']
    v = so.view(DEC, JW, '박정우', TODAY)
    assert v['can_toggle'] and not v['can_request_change'] and v['log']
    assert so.calendar_month(2026, 12) == [{'day': '2026-12-05', 'names': ['영업가'], 'status': so.CONFIRMED,
                                            'status_label': '확정'}]


def test_menu_visibility(slack, monkeypatch):
    from dashboard.blueprints.gw_satoff import can_view_email
    monkeypatch.setattr(hr, 'can_view_all_leave', lambda e: e == 'boss@itg-aircon.com')
    assert can_view_email(f'{A.lower()}@itg-aircon.com')                 # 기술영업
    assert can_view_email(f'{JW.lower()}@itg-aircon.com')                # 승인권자
    assert can_view_email('umw@itg-aircon.com')                           # '영업'도 보기는 됨(11-25 시공·미화만 뺌)
    assert can_view_email('boss@itg-aircon.com')
    nav = gw.build_nav('/leave/saturday', 'editor', sat_off=True)
    items = [it for a in nav['areas'] for it in a['children']]
    assert any(it['href'] == '/leave/saturday' and it['active'] for it in items)
    nav = gw.build_nav('/leave', 'editor')
    assert not any(it['href'] == '/leave/saturday' for a in nav['areas'] for it in a['children'])


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))


def test_dm_some_alerts_when_nobody_received(caplog):
    """결재 요청 DM 이 한 명에게도 안 가면 ERROR(알림 종에도 없어 아무도 모르던 것, 11-24 조용한 실패 점검 5번)."""
    import logging

    class Broken:
        def conversations_open(self, users):
            raise RuntimeError('slack down')
    with caplog.at_level(logging.WARNING):
        so._dm_some(Broken(), ['U1', 'U2'], 'x', DEC, '결재 요청')
    assert any(r.levelno == logging.ERROR and '아무에게도 못 보냄' in r.getMessage() for r in caplog.records)
    caplog.clear()

    class Half:
        def conversations_open(self, users):
            if users == 'U1':
                raise RuntimeError('no')
            return {'channel': {'id': 'D' + users}}

        def chat_postMessage(self, **k):
            return {'ok': True}
    with caplog.at_level(logging.WARNING):
        so._dm_some(Half(), ['U1', 'U2'], 'x', DEC, '결재 요청')
    assert not any(r.levelno == logging.ERROR for r in caplog.records)


def test_notify_failure_shown_on_screen(slack, monkeypatch):
    """알림을 못 보내면 기록에 남겨 화면(view 의 notify_fail)이 '직접 알려 주세요'를 띄운다 —
    ERROR 알림은 부장님께만 가서 버튼을 누른 사람은 몰랐음(11-24 조용한 실패 점검 추가분)."""
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    so.submit(DEC, CO, '빈대리', TODAY)
    real = slack.chat_postMessage

    def down(channel, text='', blocks=None):
        raise RuntimeError('channel_not_found')
    monkeypatch.setattr(slack, 'chat_postMessage', down)
    so.notify('submit', DEC, '빈대리')
    msg = so.view(DEC, CO, '빈대리', TODAY)['notify_fail']
    assert '결재 요청' in msg and '직접 알려 주세요' in msg
    monkeypatch.setattr(slack, 'chat_postMessage', real)
    so.decide(DEC, True, YG, '박용구')                                   # 다음 처리가 있으면 지난 일
    so.notify('approve', DEC, '박용구')
    assert so.view(DEC, CO, '빈대리', TODAY)['notify_fail'] == ''
    monkeypatch.setattr(so, '_client', lambda: None)
    so.notify('approve', DEC, '박용구')
    assert '휴가 봇 연결 없음' in so.view(DEC, CO, '빈대리', TODAY)['notify_fail']


def test_worked_api_route(slack, monkeypatch):
    """POST /api/satoff/worked — on=true 기록, on=false 되돌리기. 휴가 일수 영향이 없으면 알림 없음."""
    import inspect
    from flask import Flask
    from dashboard.blueprints import gw_satoff as so_bp
    so.toggle(S5, '영업가', True, A, '영업가', TODAY)
    _confirm()
    monkeypatch.setattr(so_bp, '_actor', lambda: ('영업가', A))
    view = inspect.unwrap(so_bp.api_worked)                      # 로그인·개발 중 문은 다른 시험이 봄
    app = Flask(__name__)
    before = len(slack.posts)
    with app.test_request_context(json={'day': '2026-12-05', 'verbal_by': '박정우', 'reason': '현장'}):
        j = view().get_json()
    assert j['success'] and j['name'] == '영업가' and so.picks(DEC)['2026-12-05'] == []
    assert len(slack.posts) == before                              # 구두 승인 — 알림 없음
    with app.test_request_context(json={'day': '2026-12-05', 'verbal_by': '', 'reason': '현장'}):
        r = view()
    assert r[1] == 400                                             # 이미 출근 기록 → 쉬는 날이 아님 등 업무 오류
    with app.test_request_context(json={'day': '2026-12-05', 'on': False}):
        assert view().get_json()['success']
    assert so.picks(DEC)['2026-12-05'] == ['영업가']


# ── 누가 보나(11-25 부장님: 시공·미화 빼고 모두에게 엶) ─────────────────────
def test_view_open_to_all_but_construction_and_cleaning(slack, monkeypatch):
    from dashboard.blueprints import gw_satoff as so_bp
    for name, uid, dept in (('시공가', 'USG', '기술시공(내부)'), ('미화가', 'UMH', '미화'), ('무부서', 'UNO', '')):
        hr.create_person({'name': name, 'status': '재직', 'hire_date': '2022-05-23', 'department': dept,
                          'slack_id': uid, 'work_email': f'{uid.lower()}@itg-aircon.com'}, ADMIN)
    monkeypatch.setattr(hr, 'can_view_all_leave', lambda e: False)
    ok = {e: so_bp.can_view_email(e) for e in ('ua0001@itg-aircon.com', 'uco@itg-aircon.com', 'ujw@itg-aircon.com',
                                                'umw@itg-aircon.com', 'ukiko@itg-aircon.com', 'uno@itg-aircon.com',
                                                'usg@itg-aircon.com', 'umh@itg-aircon.com', 'nobody@itg-aircon.com', '')}
    assert ok == {'ua0001@itg-aircon.com': True, 'uco@itg-aircon.com': True, 'ujw@itg-aircon.com': True,   # 기술영업·취합·CEO
                  'umw@itg-aircon.com': True, 'ukiko@itg-aircon.com': True, 'uno@itg-aircon.com': True,    # 영업·경영지원·부서 없음
                  'usg@itg-aircon.com': False, 'umh@itg-aircon.com': False,                               # 시공·미화
                  'nobody@itg-aircon.com': False, '': False}
    # 보기만 하는 직원(영업) — 고르기·휴무 출근은 서비스가 막는다
    v = so.view(DEC, 'UMW', '박민우', TODAY)
    assert not v['can_toggle'] and not v['can_submit'] and not v['can_mark_worked'] and not v['me']['member']
    with pytest.raises(so.SatOffError, match='대상'):
        so.toggle(S5, '박민우', True, 'UMW', '박민우', TODAY)
    assert not hasattr(so_bp.gw_satoff_bp, 'before_request_funcs') or not so_bp.gw_satoff_bp.before_request_funcs.get(None)
