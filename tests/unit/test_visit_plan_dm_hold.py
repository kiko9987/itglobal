"""방문 일정 DM 을 차량 배정까지 기다렸다 한 통으로(2026-11-25 부장님).
- 방문 일정 확정: 시트·List·캔버스는 바로, 담당자·당번 DM 은 기다림(차량 배정을 이미 확정한 날이면 바로 차와 함께)
- 차량 배정 확정: 방문 일정 + 차를 사람마다 한 통, 그 사람은 차량 DM 을 따로 안 받음(방문 없이 차만 탄 사람은 차량 DM)
- 16:30 자동 발송·[일정만 먼저 보내기]: 일정만 먼저, 그 뒤 차는 차량 DM 으로 따로
이름·번호는 모두 가짜."""
from datetime import date, timedelta

import dashboard.services.visit_assignment_sync as vas
import dashboard.utils.redis_client as rcmod
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import visit_plan as vp
from test_visit_plan import DAY, LEADS, STAFF, _Client, _lead, _NowThread, _RC, _Redis, db  # noqa: F401  (db = 픽스처)

RAY, SORENTO = '237허5358', '03모8021'
PLAN = {'courses': [{'members': ['TH'], 'leads': ['L-3'], 'memo': ''},
                    {'members': ['YM', 'LSH'], 'leads': ['L-1', 'L-2'], 'memo': ''}],
        'online': [{'code': 'JK', 'shift': ''}], 'online_shared': False}


def _env(monkeypatch, seen=None):
    """확정 본체·캔버스 같은 바깥은 가짜로. seen 에 commit_plan 에 넘긴 send_dm·car_by_mgr 를 모은다."""
    seen = [] if seen is None else seen
    monkeypatch.setattr(vas, 'commit_plan', lambda parsed, target_date=None, **k: seen.append(k) or {'ok': True, 'dm': {}})
    monkeypatch.setattr(vp.threading, 'Thread', _NowThread)
    monkeypatch.setattr(vp, '_audit', lambda res, by: [])
    monkeypatch.setattr(va, 'request_assignment', lambda day=None, by='': 0)
    monkeypatch.setattr(vp, '_staff', lambda: [dict(p) for p in STAFF])
    monkeypatch.setattr(vp, '_dm_leads', lambda: LEADS)
    monkeypatch.setattr(va, 'CAR_DM_ON', True)
    car_dm = []
    monkeypatch.setattr(va, '_dm', lambda sid, text: car_dm.append((sid, text)) or True)
    plan_dm = []

    def fake_send(parsed, target_date, car_by_mgr=None):
        plan_dm.append(car_by_mgr)
        codes = {m for a in parsed['assignments'] for m in a['assign']}
        return {'visit_mgr_sent': len(codes), 'visit_mgr_failed': 0, 'online_duty_sent': 1, 'mgr_done': sorted(codes)}
    monkeypatch.setattr(vas, 'send_plan_dms', fake_send)
    return seen, car_dm, plan_dm


def _confirm(plan=PLAN):
    vp.start_confirm(DAY, plan, 'p@example.com', LEADS, lambda: LEADS)
    assert vp.run_state(DAY)['state'] == 'done'


def test_confirm_holds_dm_until_cars(db, monkeypatch):
    seen, car_dm, plan_dm = _env(monkeypatch)
    _confirm()
    assert seen[-1] == {'send_dm': False, 'car_by_mgr': None}            # 시트·List·캔버스만, DM 은 기다림
    assert vp.dm_mode(DAY) == vp.DM_HELD and vp.run_state(DAY)['result']['dm_mode'] == vp.DM_HELD
    assert vp.load(DAY, LEADS, 'p@example.com')['dm_mode'] == vp.DM_HELD
    assert vp.confirmed_info(DAY)['dm_mode'] == vp.DM_HELD
    _confirm()                                                            # 기다리는 중에 다시 확정해도 그대로 기다림
    assert seen[-1]['send_dm'] is False and not plan_dm

    # 차량 배정 확정 — 방문 담당자는 일정+차 한 통, 방문 없이 차만 탄 JK(온라인 당번)는 차량 DM
    res = va.save_day(DAY, [{'plate': RAY, 'people': ['TH', 'JK']}, {'plate': SORENTO, 'people': ['YM', 'LSH']}],
                      'p@example.com', confirm=True)
    cars = plan_dm[-1]
    assert set(cars) == {'TH', 'YM', 'LSH'}
    assert '레이 5358' in cars['TH'] and '강해' in cars['TH'] and '쏘렌토' in cars['YM'] and '이나래' in cars['YM']
    assert [s for s, _ in car_dm] == ['UJK'] and res['notified'] == 1
    assert res['plan_dm']['ok'] and res['plan_dm']['dm'] == 3 and 'covered' not in res['plan_dm']
    assert vp.dm_mode(DAY) == vp.DM_COMBINED

    # 확정 뒤 차를 바꿈 — 방문 담당자는 한 통으로 다시(바뀐 사람만은 DM 본체가 가림), 차만 탄 사람은 차량 DM
    car_dm.clear()
    va.save_day(DAY, [{'plate': RAY, 'people': ['TH']}, {'plate': SORENTO, 'people': ['YM', 'LSH', 'JK']}], 'p@example.com')
    assert len(plan_dm) == 2 and '함께' not in plan_dm[-1]['TH']
    assert [s for s, _ in car_dm] == ['UJK'] and '바뀌었어요' in car_dm[0][1]

    # 방문 일정을 다시 확정 — 이미 차와 함께 보낸 날이라 바로 차와 함께(바뀐 사람만)
    _confirm()
    assert seen[-1]['send_dm'] is True and set(seen[-1]['car_by_mgr']) == {'TH', 'YM', 'LSH'}


def test_confirm_after_cars_sends_together_now(db, monkeypatch):
    """차량 배정을 먼저 확정해 둔 날이면 방문 일정 확정 때 바로 차와 함께."""
    seen, car_dm, plan_dm = _env(monkeypatch)
    va.save_day(DAY, [{'plate': RAY, 'people': ['TH']}], 'p@example.com', confirm=True)
    assert [s for s, _ in car_dm] == ['UTH']                            # 방문 일정 확정 전이라 차량 DM 만
    _confirm()
    assert seen[-1]['send_dm'] is True and '레이 5358' in seen[-1]['car_by_mgr']['TH']
    assert seen[-1]['car_by_mgr']['YM'] == ':car: 차량 : 배정 없음'
    assert vp.dm_mode(DAY) == vp.DM_COMBINED


def test_plan_dm_failure_falls_back_to_car_dm(db, monkeypatch):
    """리드 시트를 못 읽어 한 통으로 못 보내면 — 차량 DM 은 모두에게 예전처럼, 방문 일정 DM 은 계속 기다림."""
    _, car_dm, plan_dm = _env(monkeypatch)
    _confirm()

    def down():
        raise ValueError('리드 시트를 읽지 못해 방문 일정 DM 을 보내지 못했어요.')
    monkeypatch.setattr(vp, '_dm_leads', down)
    res = va.save_day(DAY, [{'plate': RAY, 'people': ['TH', 'JK']}], 'p@example.com', confirm=True)
    assert res['plan_dm'] == {'ok': False, 'message': '리드 시트를 읽지 못해 방문 일정 DM 을 보내지 못했어요.'}
    assert sorted(s for s, _ in car_dm) == ['UJK', 'UTH'] and not plan_dm
    assert vp.dm_mode(DAY) == vp.DM_HELD


def test_auto_send_and_send_now(db, monkeypatch):
    """16:30 — 차량 배정을 확정하지 않은 다음 근무일은 일정만 먼저 + 관리자에게 알림. 그 뒤 차는 차량 DM 으로 따로."""
    _, car_dm, plan_dm = _env(monkeypatch)
    monkeypatch.setattr(va, 'next_workday', lambda today=None: DAY)
    monkeypatch.setattr(va, 'assigners', lambda: ['jw@example.com', 'sj@example.com'])
    monkeypatch.setattr(hr, 'is_admin_email', lambda e: e in ('jw@example.com', 'p@example.com'))
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'slack_id': 'U' + e.split('@')[0].upper()})
    assert vp.auto_send_held() is None                                    # 확정 전 — 할 일 없음
    _confirm()
    res = vp.auto_send_held()
    assert res['ok'] and res['dm'] == 3 and plan_dm == [None]           # 차 없이 일정만
    assert sorted(s for s, _ in car_dm) == ['UJW', 'UP'] and '먼저 보냈어요' in car_dm[0][1]   # 관리자(배정 담당·확정한 사람)에게
    assert vp.dm_mode(DAY) == vp.DM_PLAIN and res['admin_notified'] == 2
    assert vp.auto_send_held() is None                                    # 두 번은 안 보냄
    car_dm.clear()
    va.save_day(DAY, [{'plate': RAY, 'people': ['TH']}], 'p@example.com', confirm=True)
    assert [s for s, _ in car_dm] == ['UTH'] and plan_dm == [None]       # 차는 차량 DM 으로 따로


def test_send_now_route(db, monkeypatch):
    import dashboard.auth as auth
    from flask import Flask
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.services import groupware as gw
    _, _, plan_dm = _env(monkeypatch)
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'name': '사무', 'department': '영업', 'status': '재직'})
    _confirm()
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        assert cl.post('/api/vehicles/plan/send-now', json={'day': DAY}).status_code == 403
        with cl.session_transaction() as s:
            s['user'] = {'email': 'p@example.com', 'permission_level': 'admin'}
        j = cl.post('/api/vehicles/plan/send-now', json={'day': DAY}).get_json()
        assert j['success'] and j['dm'] == 3 and 'covered' not in j and plan_dm == [None]
        assert cl.post('/api/vehicles/plan/send-now', json={'day': DAY}).status_code == 409    # 이미 보냄
        assert cl.get(f'/api/vehicles/assign?day={DAY}').get_json()['plan']['dm_mode'] == vp.DM_PLAIN


def test_engine_car_line_and_resend_only_on_change(monkeypatch):
    """DM 본체 — 차량 줄을 메모 아래에 붙이고, 차만 바뀌어도 그 사람에게 '차량 배정 변경'으로 다시. None 이면 비교 안 함."""
    fc, fr = _Client(), _Redis()
    monkeypatch.setattr(vas, '_get_visit_client', lambda: fc)
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: _RC(fr))
    monkeypatch.setattr(vas, '_email_from_initial', lambda ini, itn: f'{ini.lower()}@example.com')
    monkeypatch.setattr(vas, '_get_visit_card_permalink', lambda c, n: '')
    monkeypatch.setattr(vas, '_load_addr_candidates', lambda: (_ for _ in ()).throw(AssertionError('시트를 다시 읽으면 안 됨')))
    itn = {p['code']: p['name'] for p in STAFF}
    target = date.today() + timedelta(days=2)
    a = [{'lead': _lead('L-1'), 'assign': ['YM', 'LSH'], 'memo': ''}]
    parsed = {'assignments': a, 'online_duty': [], 'off_duty': [], 'online_shared': False}

    def run(cars):
        return vas.send_plan_dms(parsed, target, car_by_mgr=cars)
    r = run({'YM': ':car: 차량: *레이 5358*\n함께 타는 사람: 이나래', 'LSH': ':car: 차량: *레이 5358*'})
    ym = [p['text'] for p in fc.posted if p['channel'] == 'U_ym']
    assert r['visit_mgr_sent'] == 2 and sorted(r['mgr_done']) == ['LSH', 'YM']
    assert '>:car: 차량: *레이 5358*\n>함께 타는 사람: 이나래' in ym[0] and ym[0].index(':car:') < ym[0].index('L-1')
    run({'YM': ':car: 차량: *레이 5358*\n함께 타는 사람: 이나래', 'LSH': ':car: 차량: *레이 5358*'})
    assert len(fc.posted) == 2                                            # 같으면 다시 안 보냄
    r = run({'YM': ':car: 차량: *쏘렌토 8021*', 'LSH': ':car: 차량: *레이 5358*'})
    ym = [p['text'] for p in fc.posted if p['channel'] == 'U_ym']
    assert len(fc.posted) == 3 and '차량 배정 변경' in ym[1] and '쏘렌토' in ym[1] and len(fc.deleted) == 1
    assert r['mgr_done'] == ['LSH', 'YM'] or sorted(r['mgr_done']) == ['LSH', 'YM']
    run(None)                                                             # 일정만 다시(차 비교 안 함) — 안 보냄
    assert len(fc.posted) == 3
