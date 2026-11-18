"""플렉스 휴가 관리 맞추기(2026-11-18) — 연차 조정 적용일·취소·일괄, 월별 상세, 맞춤 휴가, 촉진."""
from datetime import date

import pytest

from dashboard.services import leave as lv

JW, YG, KIKO, MW, SB, TH = 'U051ENK8AAY', 'U050XNG5VBR', 'U04UL2ZLJAX', 'U0BJ9M1MRJ4', 'U0BHC2JV7U5', 'UTH'
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
    lv.upsert_staff('김태현', '2023-05-10', TH, accrue_from='2026-12-03')
    lv.add_grant('김태현', 3, date(2026, 12, 2), date(2027, 5, 9), 'flex:김태현', '플렉스 잔여')
    return d


def _new(start, end=None, ltype=lv.ANNUAL, name='박민우', uid=MW):
    return lv.create(name, uid, uid, name, ltype, '', start, end or start, '개인 사정')


# ── 1. 연차 조정 보완 ────────────────────────────────────────────────
def test_adjust_applies_on_future_date(db):
    g, err = lv.adjust('박민우', 2, '보상', SB, '황샛별', today=TODAY, applies_on=date(2026, 12, 20))
    assert err == '' and g['valid_from'] == '2026-12-20' and g['valid_to'] == '2027-02-02'
    assert lv.balance('박민우', TODAY) == 5                     # 적용일 전엔 그대로
    assert lv.balance('박민우', date(2026, 12, 21)) == 7


def test_adjust_minus_checked_at_applies_on_and_window(db):
    _, err = lv.adjust('박민우', -6, '결근', SB, '황샛별', today=TODAY)
    assert '남은 연차(5일)' in err
    _, err = lv.adjust('박민우', 1, '보상', SB, '황샛별', today=TODAY, applies_on=date(2028, 1, 5))
    assert '1년 안' in err
    _, err = lv.adjust('박민우', 1, '보상', MW, '박민우', today=TODAY)
    assert '경영지원' in err


def test_cancel_adjust_keeps_record_and_restores_balance(db):
    g, _ = lv.adjust('박민우', -1, '병원', SB, '황샛별', today=TODAY)
    assert lv.balance('박민우', TODAY) == 4
    c, err = lv.cancel_adjust(g['id'], KIKO, '고광일')
    assert err == '' and c['days'] == 0 and c['orig_days'] == -1 and '취소(고광일' in c['note']
    assert lv.balance('박민우', TODAY) == 5
    rows = lv.adjustments('박민우')
    assert len(rows) == 1 and rows[0]['cancelled']
    _, err = lv.cancel_adjust(g['id'], KIKO, '고광일')            # 두 번은 안 됨
    assert '취소할 수 있는' in err


def test_cancel_plus_adjust_refused_when_used(db):
    g, _ = lv.adjust('박민우', 3, '보상', SB, '황샛별', today=TODAY)
    # 소멸이 빠른 플렉스 잔여(2/2)부터 떼므로 6일 신청이면 조정분(같은 소멸일, id 뒤)까지 쓴다
    r = _new(date(2026, 12, 7), date(2026, 12, 14))            # 월~월 6일
    assert r and lv.balance('박민우', TODAY) == 2
    _, err = lv.cancel_adjust(g['id'], SB, '황샛별')
    assert '이미 휴가에 썼어요' in err


def test_cancel_plus_adjust_refused_when_minus_would_go_negative(db):
    plus, _ = lv.adjust('박민우', 3, '보상', SB, '황샛별', today=TODAY)
    lv.adjust('박민우', -7, '정정', SB, '황샛별', today=TODAY)
    assert lv.balance('박민우', TODAY) == 1
    _, err = lv.cancel_adjust(plus['id'], SB, '황샛별')
    assert '0보다 적어져요' in err


def test_adjust_bulk_all_or_nothing(db):
    made, errs = lv.adjust_bulk([{'name': '박민우', 'days': 1}, {'name': '김태현', 'days': -9}], '연말 정정',
                                SB, '황샛별', today=TODAY)
    assert made == [] and '김태현' in errs and '박민우' not in errs
    assert lv.balance('박민우', TODAY) == 5                     # 하나라도 틀리면 아무것도 안 들어감
    made, errs = lv.adjust_bulk([{'name': '박민우', 'days': 1}, {'name': '김태현', 'days': -0.5, 'reason': '병원'},
                                 {'name': '박민우', 'days': ''}], '연말 정정', SB, '황샛별', today=TODAY)
    assert errs == {} and len(made) == 2
    assert lv.balance('박민우', TODAY) == 6 and lv.balance('김태현', TODAY) == 2.5
    assert '병원' in lv.adjustments('김태현')[0]['note'] and '연말 정정' in lv.adjustments('박민우')[0]['note']
    _, errs = lv.adjust_bulk([{'name': '박민우', 'days': 1}], '', SB, '황샛별', today=TODAY)
    assert errs == {'박민우': '사유를 적어 주세요.'}
    _, errs = lv.adjust_bulk([], '사유', SB, '황샛별', today=TODAY)
    assert '한 명 이상' in errs['_']


# ── 3. 월별 연차 상세 ────────────────────────────────────────────────
def test_monthly_detail_rows_add_up(db):
    lv.adjust('박민우', -1, '병원', SB, '황샛별', today=TODAY)                 # 12월 조정
    _new(date(2026, 12, 14))                                                   # 12월 사용 1
    _new(date(2027, 1, 11), date(2027, 1, 12))                                 # 1월 사용 2 (승인 대기도 뺌)
    # 2/3 입사 기념일 연차는 아직 안 생김(자동 발생 전) — 2/2 플렉스 잔여 소멸만
    d = lv.monthly_detail('박민우', 2026, today=TODAY)
    assert [m['month'] for m in d['months']] == [12] and d['opening'] == 0
    dec = d['months'][0]
    assert (dec['grant'], dec['adjust'], dec['used'], dec['left']) == (5, -1, 1, 3)
    d27 = lv.monthly_detail('박민우', 2027, today=TODAY)
    jan, feb = d27['months'][0], d27['months'][1]
    assert d27['opening'] == 3 and jan['used'] == 2 and jan['left'] == 1 and jan['projected']
    assert feb['expired'] == 1 and feb['left'] == 0                    # 플렉스 잔여 남은 2일 + 조정 −1 → 순 1일 소멸
    prev = d27['opening']
    for m in d27['months']:
        assert round(prev + m['grant'] + m['adjust'] - m['used'] - m['expired'], 2) == m['left']
        prev = m['left']


def test_monthly_detail_empty_for_unknown(db):
    assert lv.monthly_detail('없는사람', 2026) == {'year': 2026, 'months': [], 'opening': 0.0}


# ── 4. 맞춤 휴가 부여 ────────────────────────────────────────────────
def test_custom_grant_then_employee_applies(db):
    assert 'health' not in dict(lv.type_options_for('박민우', TODAY))
    errs = lv.validate('박민우', 'health', '', date(2026, 12, 7), date(2026, 12, 7))
    assert errs.get('type')                                             # 부여 전엔 못 고름
    g, err = lv.custom_grant('박민우', 'health', 2, None, date(2026, 12, 31), '보건 휴가', SB, '황샛별', today=TODAY)
    assert err == '' and g['valid_from'] == '2026-12-03'
    opts = dict(lv.type_options_for('박민우', TODAY))
    assert opts['health'] == '보건 (남은 2일)'
    assert lv.validate('박민우', 'health', '', date(2026, 12, 7), date(2026, 12, 7)) == {}
    errs = lv.validate('박민우', 'health', '', date(2026, 12, 7), date(2026, 12, 9))
    assert '2일만' in errs['start']
    errs = lv.validate('박민우', 'health', '', date(2026, 12, 31), date(2027, 1, 4))   # 12/31·1/4 = 2일, 기간 넘음
    assert '기간을 넘어요' in errs['end']
    r = lv.create('박민우', MW, MW, '박민우', 'health', '', date(2026, 12, 7), date(2026, 12, 7), '보건')
    assert r['status'] == lv.WAIT_1 and lv.custom_left('박민우', 'health', TODAY) == 1
    assert lv.balance('박민우', TODAY) == 5                             # 연차는 안 빠짐
    # 본인이 신청한 맞춤 휴가는 본인이 취소 가능(경영지원 등록분만 경영지원 전용)
    assert lv.can_act('cancel', r, MW, TODAY)
    lv.transition(r['id'], 'cancel', MW, '박민우')
    assert lv.custom_left('박민우', 'health', TODAY) == 2
    # 남은 게 있으면 부여 취소 가능, 신청이 걸려 있으면 안 됨
    r2 = lv.create('박민우', MW, MW, '박민우', 'health', '', date(2026, 12, 8), date(2026, 12, 8), '보건')
    _, err = lv.cancel_custom_grant(g['id'], SB, '황샛별')
    assert '먼저 취소' in err
    lv.transition(r2['id'], 'cancel', MW, '박민우')
    c, err = lv.cancel_custom_grant(g['id'], SB, '황샛별')
    assert err == '' and lv.custom_left('박민우', 'health', TODAY) == 0
    assert lv.custom_grants('박민우')[0]['cancelled_at']


def test_custom_grant_checks(db):
    assert lv.custom_grant('박민우', 'annual', 1, None, None, '', SB, 'x', today=TODAY)[1] == '부여할 휴가 종류를 골라 주세요.'
    assert '0.5일 단위' in lv.custom_grant('박민우', 'health', 0.3, None, None, '', SB, 'x', today=TODAY)[1]
    assert '경영지원만' in lv.custom_grant('박민우', 'health', 1, None, None, '', MW, 'x', today=TODAY)[1]
    assert '지난 기간' in lv.custom_grant('박민우', 'health', 1, date(2026, 1, 1), date(2026, 2, 1), '', SB, 'x',
                                         today=TODAY)[1]
    g, err = lv.custom_grant('박민우', 'infertility', 3, None, None, '', SB, 'x', today=TODAY)
    assert err == '' and g['valid_to'] == lv.ANYTIME
    assert lv.custom_balances('박민우', TODAY) == [{'ltype': 'infertility', 'label': '난임 치료', 'left': 3.0,
                                                   'valid_to': lv.ANYTIME}]


def test_manager_registered_admin_leave_still_manager_only(db):
    r, errs = lv.admin_register('박민우', 'reward', '', date(2026, 12, 21), None, '포상', SB, '황샛별')
    assert errs == {} and lv.manager_registered(r)
    assert not lv.can_act('cancel', r, MW, TODAY) and lv.own_edit_mode(r, MW, TODAY) == ''
    assert lv.can_act('cancel', r, SB, TODAY)


# ── 5. 연차 사용 촉진 ────────────────────────────────────────────────
@pytest.fixture
def promo(db, monkeypatch):
    from dashboard.services import leave_promo as lp
    monkeypatch.setattr(lp, '_ready', False)
    monkeypatch.setenv('LEAVE_ENABLED', 'true')
    return lp


def test_promo_schedule_matches_law_and_flex(promo):
    s = promo.schedule(promo.ANNUAL, date(2027, 1, 31))           # 김영숙: 플렉스 촉구 8/10, 지정 기한 12/1
    assert (s['notice_from'], s['notice_to'], s['designate_due']) == (date(2026, 8, 1), date(2026, 8, 10), date(2026, 12, 1))
    s = promo.schedule(promo.FIRST, date(2027, 5, 9))               # 1년 미만: 3개월 전·1개월 전
    assert (s['notice_from'], s['designate_due']) == (date(2027, 2, 10), date(2027, 4, 10))
    s = promo.schedule(promo.FIRST_LATE, date(2027, 5, 9))          # 촉구 뒤 생긴 월차: 1개월 전 5일 안, 10일 전
    assert (s['notice_from'], s['notice_to'], s['designate_due']) == (date(2027, 4, 10), date(2027, 4, 14), date(2027, 4, 30))


def test_promo_daily_sends_in_window_then_plan(promo):
    # 박민우 플렉스 잔여 5일, 2027-02-02 소멸 → 촉구 기간 2026-08-03 ~ 08-12 (이미 지남)
    assert promo.candidates(TODAY) == []
    missed = promo.candidates(TODAY, include_missed=True)
    assert [(c['name'], c['missed']) for c in missed if c['name'] == '박민우'] == [('박민우', True)]
    # 김태현 2027-05-09 소멸 → 2026-11-10 ~ 11-19 (오늘 12/3 이면 지남), 11/15 에 돌리면 보냄
    sent = []
    out = promo.daily(date(2026, 11, 16), notify=lambda k, p: sent.append(k))
    names = [r['name'] for r in out['sent']]
    assert names == ['김태현'] and sent[0] == 'notice' and 'managers' in sent
    rec = out['sent'][0]
    assert rec['days'] == 3 and rec['plan_due'] == '2026-11-26' and rec['designate_due'] == '2027-03-10'
    assert '근로기준법 제61조' in promo.notice_text(rec) and '11월 26일까지' in promo.notice_text(rec)
    assert promo.daily(date(2026, 11, 17))['sent'] == []            # 한 번만
    # 본인 확인·계획: 이어진 날은 한 신청으로, 반차 따로
    assert promo.ack(rec['id'], MW) != '' and promo.ack(rec['id'], TH) == ''
    _, _, err = promo.submit_plan(rec['id'], TH, [{'date': '2026-12-11'}, {'date': '2026-12-14'},
                                                  {'date': '2026-12-15'}, {'date': '2026-12-16', 'half': 'am'}],
                                  today=date(2026, 11, 18))
    assert '3일이라 3.5일' in err
    r, made, err = promo.submit_plan(rec['id'], TH, [{'date': '2026-12-11'}, {'date': '2026-12-14'},
                                                     {'date': '2026-12-16', 'half': 'am'}], today=date(2026, 11, 18))
    assert err == '' and len(made) == 2 and made[0]['days'] == 2 and made[1]['ltype'] == lv.HALF_AM   # 금·월은 한 신청
    assert all(m['status'] == lv.WAIT_1 and m['reason'] == promo.PLAN_REASON for m in made)
    assert promo.get(rec['id'])['ack_at'] and promo.state(r, date(2026, 11, 18))['key'] == 'planned'
    _, _, err = promo.submit_plan(rec['id'], TH, [{'date': '2026-12-21'}], today=date(2026, 11, 18))
    assert '0.5일이라 1일' in err


def test_promo_designate_after_plan_due(promo):
    rec, err = promo.send_notice('박민우', promo.ANNUAL, '2027-02-02', today=date(2026, 11, 20))
    assert err == '' and rec['days'] == 5
    _, _, err = promo.designate(rec['id'], SB, '황샛별', [{'date': '2026-12-07'}], today=date(2026, 11, 25))
    assert '계획 기한' in err
    assert promo.state(rec, date(2026, 12, 3))['key'] == 'need_designate'
    _, _, err = promo.designate(rec['id'], MW, '박민우', [{'date': '2026-12-07'}], today=TODAY)
    assert '경영지원만' in err
    _, _, err = promo.designate(rec['id'], SB, '황샛별', [{'date': '2026-12-12'}], today=TODAY)
    assert '쉬는 날' in err
    r, made, err = promo.designate(rec['id'], SB, '황샛별', [{'date': d} for d in
                                   ('2026-12-07', '2026-12-08', '2026-12-09', '2026-12-10', '2026-12-11')], today=TODAY)
    assert err == '' and len(made) == 1 and made[0]['status'] == lv.APPROVED and made[0]['days'] == 5
    assert r['status'] == promo.DESIGNATED and promo.period_left('박민우', '2027-02-02') == 0
    assert '12월 7일 (월)' in promo.designate_text(r) and '출근하지 않으셔도' in promo.designate_text(r)
    assert promo.state(r, TODAY)['key'] == 'designated'


def test_promo_import_flex(promo, tmp_path):
    import json
    p = tmp_path / 'flex.json'
    p.write_text(json.dumps({'users': {'U1': '박민우', 'U2': '김태현'}, 'docs': [
        {'u': 'U1', 'task': 't1', 'to': '2027-02-02', 'boostedAt': '2026-08-11T23:00:00', 'days': 8,
         'status': 'SHOULD_MANAGER_WRITE', 'userBy': '2026-08-22', 'mgrBy': '2026-12-03', 'planned': [], 'canceled': False},
        {'u': 'U2', 'task': 't2', 'to': '2027-05-09', 'boostedAt': '2026-11-09T23:00:00', 'days': 3,
         'status': 'PENDING_APPROVAL', 'approval': 'ONGOING', 'userBy': '2026-11-20', 'mgrBy': '2027-03-09',
         'planned': ['2027-01-05', '2026-12-30'], 'writtenAt': 1793404036000, 'canceled': False},
        {'u': 'U2', 'task': 't0', 'to': '2027-05-09', 'boostedAt': '2026-11-01T23:00:00', 'days': 3,
         'status': 'CANCELED', 'userBy': '2026-11-12', 'mgrBy': '2027-03-09', 'planned': [], 'canceled': True},
    ]}, ensure_ascii=False), encoding='utf-8')
    rows = promo.import_flex(str(p))
    assert [(r['name'], r['status'], r['source']) for r in rows] == [('김태현', 'planned', 'flex:t2'),
                                                                     ('박민우', 'noticed', 'flex:t1')]
    assert rows[1]['noticed_at'] == '2026-08-12 08:00:00'          # UTC 23시 → 한국 다음 날 8시
    promo.import_flex(str(p), apply=True)
    promo.import_flex(str(p), apply=True)                            # 두 번 넣어도 한 번
    recs = promo.records()
    assert len(recs) == 2
    mw = next(r for r in recs if r['name'] == '박민우')
    assert promo.state(mw, TODAY)['key'] == 'need_designate'         # 12/3 까지 지정
    kt = next(r for r in recs if r['name'] == '김태현')
    assert promo.state(kt, TODAY)['key'] == 'planned' and [d['date'] for d in kt['plan_dates']] == ['2026-12-30', '2027-01-05']
    # 플렉스 기록이 있으면 같은 기간 촉구를 또 보내지 않음
    assert all(c['name'] != '김태현' for c in promo.candidates(date(2026, 11, 16)))
