"""휴가 (2026-11-05) — 장부·자동 발생·신청·승인 순서·#휴가_일정 문구."""
from datetime import date

import pytest

from dashboard.services import leave as lv

JW, YG, KIKO, MW = 'U051ENK8AAY', 'U050XNG5VBR', 'U04UL2ZLJAX', 'U0BJ9M1MRJ4'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(lv, '_db', lambda: d)
    monkeypatch.setattr(lv, '_table_ready', False)
    monkeypatch.delenv('LEAVE_APPROVER_IDS', raising=False)
    return d


def _seed(name='박민우', uid=MW, hire='2022-02-03', left=8.0, until='2027-02-02'):
    lv.upsert_staff(name, hire, uid, accrue_from='2026-11-06')
    lv.add_grant(name, left, date(2026, 11, 5), date.fromisoformat(until), f'flex:{name}', '플렉스 잔여')


def test_count_days_skips_weekend_and_holiday():
    assert lv.count_days(lv.ANNUAL, date(2026, 11, 6), date(2026, 11, 9)) == 2      # 금·월
    assert lv.count_days(lv.ANNUAL, date(2026, 12, 24), date(2026, 12, 25)) == 1    # 성탄절
    assert lv.count_days(lv.HALF_AM, date(2026, 11, 6), None) == 0.5


def test_annual_days_rule():
    assert [lv.annual_days(n) for n in (1, 2, 3, 4, 5, 21, 30)] == [15, 15, 16, 16, 17, 25, 25]


def test_due_grants_monthly_then_annual():
    p = {'name': '권진현', 'hire_date': '2026-11-02', 'accrue_from': '2026-11-02'}
    g = lv.due_grants(p, date(2027, 3, 5))
    assert [x['valid_from'] for x in g] == [date(2026, 12, 2), date(2027, 1, 2), date(2027, 2, 2), date(2027, 3, 2)]
    assert all(x['valid_to'] == date(2027, 11, 1) for x in g)
    # 1년 이상 근속 기념일은 FISCAL_FROM 전까지만(그 뒤는 회계연도 정산표로)
    old = {'name': '주영민', 'hire_date': '2024-04-08', 'accrue_from': '2026-11-06'}
    assert lv.due_grants(old, date(2027, 6, 1)) == []
    mid = {'name': 'X', 'hire_date': '2025-12-01', 'accrue_from': '2026-11-06'}
    a = [x for x in lv.due_grants(mid, date(2026, 12, 31)) if x['source'].startswith('annual')]
    assert a and a[0]['days'] == 15 and a[0]['valid_to'] == date(2027, 11, 30)


def test_accrue_is_idempotent(db):
    lv.upsert_staff('권진현', '2026-11-02', 'U0CETGE1480', accrue_from='2026-11-02')
    assert len(lv.accrue(date(2026, 12, 2))) == 1
    assert lv.accrue(date(2026, 12, 2)) == []
    assert lv.balance('권진현', date(2026, 12, 3)) == 1


def test_validate_balance_and_overlap(db):
    _seed(left=1.5)
    assert lv.validate('박민우', lv.ANNUAL, '', date(2026, 11, 9), date(2026, 11, 9)) == {}
    e = lv.validate('박민우', lv.ANNUAL, '', date(2026, 11, 9), date(2026, 11, 10))
    assert '남은 연차가 1.5일' in e['start']
    r = lv.create('박민우', MW, MW, '박민우', lv.HALF_AM, '', date(2026, 11, 9), None)
    assert r['status'] == lv.WAIT_1 and lv.balance('박민우', date(2026, 11, 9)) == 1.0
    assert lv.validate('박민우', lv.HALF_PM, '', date(2026, 11, 9), None) == {}            # 오전+오후는 됨
    assert '겹쳐요' in lv.validate('박민우', lv.ANNUAL, '', date(2026, 11, 9), date(2026, 11, 9))['start']
    # 관리자 부여 휴가는 경영지원에게도 안 보인다(11-05)
    assert 'type' in lv.validate('박민우', 'reward', '', date(2026, 11, 9), date(2026, 11, 9))
    assert 'type' in lv.validate('박민우', 'reward', '', date(2026, 11, 12), date(2026, 11, 12), for_manager=True)


def test_event_and_sick_limits(db):
    _seed()
    e = lv.validate('박민우', lv.EVENT, 'mourn2', date(2026, 11, 9), date(2026, 11, 12))
    assert '3일까지' in e['end']
    assert lv.validate('박민우', lv.EVENT, 'mourn2', date(2026, 11, 9), date(2026, 11, 11)) == {}
    assert 'sub' in lv.validate('박민우', lv.EVENT, '', date(2026, 11, 9), date(2026, 11, 9))


def test_approval_chain_and_release(db):
    _seed(left=3)
    r = lv.create('박민우', MW, MW, '박민우', lv.ANNUAL, '', date(2026, 11, 9), date(2026, 11, 10))
    assert lv.current_approver(r) == JW
    assert not lv.can_act('approve', r, YG)
    r = lv.transition(r['id'], 'approve', JW, '박정우')
    assert r['status'] == lv.WAIT_2 and lv.current_approver(r) == YG
    assert not lv.can_act('approve', r, JW)          # JW 는 이미 했다 — 이제 YG 차례
    r = lv.transition(r['id'], 'approve', YG, '박용구')
    assert r['status'] == lv.APPROVED and lv.balance('박민우', date(2026, 11, 9)) == 1
    assert lv.transition(r['id'], 'approve', YG, '박용구') is None   # 두 번 눌러도 한 번만


def test_double_click_guard_and_reject_returns_days(db):
    _seed(left=3)
    r = lv.create('박민우', MW, MW, '박민우', lv.ANNUAL, '', date(2026, 11, 9), date(2026, 11, 10))
    assert lv.balance('박민우', date(2026, 11, 9)) == 1
    r = lv.transition(r['id'], 'reject', JW, '박정우', '현장 일정')
    assert r['status'] == lv.REJECTED and lv.balance('박민우', date(2026, 11, 9)) == 3
    assert lv.transition(r['id'], 'reject', JW, '박정우') is None
    assert '현장 일정' in lv.card_text(r, lv.steps(r['id']), 'requester', 3)


def test_requester_is_approver(db):
    _seed('박정우', JW, '2019-08-07', 18, '2027-08-06')
    _seed('박용구', YG, '2007-01-01', 24, '2026-12-31')
    r = lv.create('박정우', JW, JW, '박정우', lv.ANNUAL, '', date(2026, 11, 9), date(2026, 11, 9))
    assert r['status'] == lv.WAIT_2 and lv.current_approver(r) == YG
    r = lv.create('박용구', YG, YG, '박용구', lv.ANNUAL, '', date(2026, 11, 9), date(2026, 11, 9))
    assert r['status'] == lv.APPROVED
    assert '✅ 박정우' in lv.step_line(r, lv.steps(r['id'])) and '✅ 박용구' in lv.step_line(r, lv.steps(r['id']))


def test_cancel_rules(db):
    _seed(left=3)
    r = lv.create('박민우', MW, MW, '박민우', lv.ANNUAL, '', date(2026, 11, 9), date(2026, 11, 9))
    r = lv.transition(r['id'], 'approve', JW, '박정우')
    r = lv.transition(r['id'], 'approve', YG, '박용구')
    assert r['status'] == lv.APPROVED
    assert lv.can_act('cancel', r, MW, today=date(2026, 11, 8))
    assert not lv.can_act('cancel', r, MW, today=date(2026, 11, 9))      # 당일·지난 휴가는 본인 취소 불가
    assert lv.can_act('cancel', r, KIKO, today=date(2026, 11, 20))        # 경영지원은 됨
    r = lv.transition(r['id'], 'cancel', MW, '박민우')
    assert r['status'] == lv.CANCELLED and lv.balance('박민우', date(2026, 11, 9)) == 3


def test_proxy_request_for_staff_without_slack(db):
    lv.upsert_staff('권태숙', '2023-07-24', '', accrue_from='2026-11-06')
    lv.add_grant('권태숙', 12, date(2026, 11, 5), date(2027, 7, 23), 'flex:권태숙')
    r = lv.create('권태숙', '', KIKO, '고광일', lv.ANNUAL, '', date(2026, 11, 16), date(2026, 11, 16))
    assert r['status'] == lv.WAIT_1 and r['submitter_id'] == KIKO
    assert lv.can_act('cancel', r, KIKO)
    assert '대신 신청: 고광일' in lv.card_text(r, lv.steps(r['id']), 'requester', 11)


def test_announce_text_hides_type_and_reason():
    base = {'name': '빈승정', 'start': '2026-11-09', 'end': '2026-11-09', 'reason': '병원'}
    assert lv.announce_text({**base, 'ltype': lv.SICK}) == '🌴 [빈승정] - 11월 9일 하루종일 휴가입니다.'
    assert lv.announce_text({**base, 'ltype': lv.HALF_PM}) == '🌴 [빈승정] - 11월 9일 오후 반차입니다.'
    assert lv.announce_text({**base, 'ltype': lv.ANNUAL, 'end': '2026-11-11'}) == \
        '🌴 [빈승정] - 11월 9일 ~ 11월 11일 휴가입니다.'


def test_allocation_uses_soonest_expiring_first(db):
    lv.upsert_staff('강정권', '2026-04-01', 'U0BHH8UD99P', accrue_from='2026-11-06')
    lv.add_grant('강정권', 1, date(2026, 11, 1), date(2027, 3, 31), 'a')
    lv.add_grant('강정권', 2, date(2026, 11, 1), date(2027, 1, 31), 'b')
    r = lv.create('강정권', 'U0BHH8UD99P', 'U0BHH8UD99P', '강정권', lv.ANNUAL, '', date(2026, 11, 9), date(2026, 11, 9))
    g = {x['source']: x['used'] for x in lv.grants('강정권')}
    assert g == {'b': 1, 'a': 0} and r


# ── 슬랙 화면 ────────────────────────────────────────────────────────
class _FakeClient:
    def __init__(self):
        self.posts, self.updates = [], []

    def conversations_open(self, users):
        return {'channel': {'id': f'D-{users}'}}

    def chat_postMessage(self, channel, text, blocks=None):
        self.posts.append((channel, text, blocks))
        return {'ts': f'{len(self.posts)}.0'}

    def chat_update(self, channel, ts, text, blocks=None):
        self.updates.append((channel, ts, text, blocks))


@pytest.fixture
def slack(db, monkeypatch):
    from dashboard.blueprints import slack_leave as sl
    from dashboard.blueprints import slack_eapproval as se
    c = _FakeClient()
    monkeypatch.setattr(sl, 'leave_client', lambda: c)
    monkeypatch.setattr(se, '_name', lambda uid: {JW: '박정우', YG: '박용구', KIKO: '고광일', MW: '박민우'}.get(uid, uid))
    monkeypatch.delenv('LEAVE_ANNOUNCE_CHANNEL', raising=False)
    return sl, c


def _buttons(blocks):
    return [e['action_id'] for b in blocks or [] if b['type'] == 'actions' for e in b['elements']]


def test_flow_dm_only_and_announce_dates_only(slack):
    sl, c = slack
    _seed(left=3)
    r = lv.create('박민우', MW, MW, '박민우', lv.SICK, '', date(2026, 11, 9), date(2026, 11, 9), '병원 진료')
    sl.start(r)
    assert [p[0] for p in c.posts] == [f'D-{MW}', f'D-{JW}']     # 채널엔 안 올라감
    assert _buttons(c.posts[1][2]) == [lv.A_APPROVE, lv.A_REJECT]
    assert sl.apply_action('approve', r['id'], YG) == '지금 승인할 차례가 아니에요.'
    assert sl.apply_action('approve', r['id'], JW) is None
    assert c.posts[-1][0] == f'D-{YG}'
    assert sl.apply_action('approve', r['id'], YG) is None
    ann = [p for p in c.posts if p[0] == lv.ANNOUNCE_CHANNEL_DEFAULT]
    assert len(ann) == 1 and ann[0][1] == '🌴 [박민우] - 11월 9일 하루종일 휴가입니다.'
    # JW 카드는 버튼이 사라진다
    jw_updates = [u for u in c.updates if u[0] == f'D-{JW}']
    assert jw_updates and _buttons(jw_updates[-1][3]) == []


def test_cancel_after_announce_strikes_message(slack):
    sl, c = slack
    _seed('박용구', YG, '2007-01-01', 24, '2026-12-31')
    r = lv.create('박용구', YG, YG, '박용구', lv.HALF_AM, '', date(2026, 12, 7), None)
    sl.start(r)
    assert any(p[0] == lv.ANNOUNCE_CHANNEL_DEFAULT for p in c.posts)
    assert sl.apply_action('cancel', r['id'], YG) is None
    ann = [u for u in c.updates if u[0] == lv.ANNOUNCE_CHANNEL_DEFAULT]
    assert ann and ann[-1][2].startswith('~🌴 [박용구]') and ann[-1][2].endswith('(취소)')


def test_leave_view_fields(db):
    from dashboard.blueprints import slack_leave as sl
    _seed()
    lv.upsert_staff('권태숙', '2023-07-24', '')
    v = sl.leave_view({'name': '박민우', 'mgr': False})
    ids = [b.get('block_id') for b in v['blocks'] if b.get('block_id')]
    assert ids == ['lv_type', 'lv_start', 'lv_end', 'lv_reason']
    types = [o['value'] for o in v['blocks'][1]['element']['options']]
    assert 'reward' not in types and lv.HALF_AM in types
    v = sl.leave_view({'name': '권태숙', 'mgr': True}, {'type': lv.EVENT})
    ids = [b.get('block_id') for b in v['blocks'] if b.get('block_id')]
    assert ids[:3] == ['lv_who', 'lv_type', 'lv_sub'] and 'reward' not in str(v['blocks'])
    assert '권태숙 (슬랙 없음)' in str(v['blocks'][0])
    v = sl.leave_view({'name': '박민우'}, {'type': lv.HALF_PM})
    assert [b.get('block_id') for b in v['blocks'] if b.get('block_id')] == ['lv_type', 'lv_start', 'lv_reason']
    st = {'state': {'values': {'lv_type': {'lv_type': {'selected_option': {'value': 'sick'}}},
                               'lv_start': {'lv_start': {'selected_date': '2026-11-09'}},
                               'lv_reason': {'lv_reason': {'value': ' 진료 '}}}}}
    assert sl.read_view(st) == {'who': '', 'type': 'sick', 'sub': '', 'start': '2026-11-09', 'end': '',
                                'reason': '진료'}


def test_start_view_gating_and_manifest(db, monkeypatch):
    from pathlib import Path
    from dashboard.blueprints import slack_leave as sl
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: uid)               # 슬랙 이름 조회(네트워크) 안 함
    monkeypatch.delenv('LEAVE_ENABLED', raising=False)
    monkeypatch.delenv('LEAVE_PILOT_IDS', raising=False)
    monkeypatch.delenv('EAPPROVAL_APPROVER_IDS', raising=False)
    _seed()
    assert '준비 중' in str(sl.start_view(MW))                   # 열기 전엔 플렉스로
    monkeypatch.setenv('LEAVE_ENABLED', 'true')
    assert sl.start_view(MW)['callback_id'] == lv.CB_SUBMIT
    assert '명단에 아직 없어요' in str(sl.start_view('U-NOBODY'))
    docs = Path(__file__).resolve().parents[2] / 'docs'
    manifest = (docs / 'slack_leave_manifest.yml').read_text(encoding='utf-8')
    assert f'  {lv.FN_OPEN_LEAVE}:' in manifest and '/slack/leave-events' in manifest
    # 휴가는 전자결재 봇과 따로 — 전자결재 메뉴·manifest 엔 휴가가 없다
    from dashboard.services import e_approval as ea
    assert lv.A_NEW not in str(ea.hub_blocks()) and 'leave' not in (docs / 'slack_eapproval_manifest.yml').read_text(
        encoding='utf-8')


def test_home_tab_pilot_only(db, monkeypatch):
    from dashboard.blueprints import slack_leave as sl
    monkeypatch.delenv('LEAVE_ENABLED', raising=False)
    monkeypatch.delenv('LEAVE_PILOT_IDS', raising=False)
    _seed()
    lv.upsert_staff('고광일', '2025-01-01', KIKO)
    assert lv.A_NEW not in str(sl.home_view(MW)) and '박민우' in str(sl.home_view(MW))
    assert lv.A_NEW in str(sl.home_view(KIKO)) and lv.A_MINE in str(sl.home_view(KIKO))
    monkeypatch.setenv('LEAVE_ENABLED', 'true')
    assert lv.A_NEW in str(sl.home_view(MW))


def test_leave_app_handlers_registered():
    """휴가 봇 Bolt 앱에 단계·버튼·창 핸들러가 다 붙는다(이름 충돌 없이)."""
    from slack_bolt import App
    from dashboard.blueprints import slack_leave as sl
    app = App(token='xoxb-test', signing_secret='s', token_verification_enabled=False)
    sl.register_leave_handlers(app)
    names = str([l.ack_function for l in app._listeners])
    for fn in ('handle_open_leave_step', 'handle_leave_submit', 'handle_leave_reject_open', 'handle_leave_mine',
               'handle_leave_home', 'handle_leave_view_closed'):
        assert fn in names
