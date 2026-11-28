"""조용한 실패 점검 3차 — 차량 쪽(11-24 점검 17·18·23·24번). 이름은 모두 가짜.

17. 차량 배정 동시 편집 — 화면이 불러올 때 받은 판(base)과 다르면 409, base 없는 옛 화면은 그대로 허용
18. 배정 화면 불러오기 — 리드 시트·휴가 등을 못 읽으면 partial 로 알림(성공은 그대로, 배정은 막지 않음)
23. 운행일지 배정·운전자 읽기 — 실패는 ERROR + assign_error, 전역 대신 요청마다 따로
24. 차량 관리 결재 비용 모으기 — 표 없음만 DEBUG, 한 줄 실패가 나머지를 막지 않음, 0원 줄은 넣지 않음
"""
import json
import sqlite3

import pytest
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_driver as vd
from dashboard.services import vehicle_log as vl
from dashboard.services import vehicle_manage as vm
from dashboard.services import visit_plan as vp
from test_vehicle_assign import ASSIGNER, STAFF
from test_vehicle_log import LEADS, T1, _lead, _xlsx

DAY = '2026-11-25'
RAY = '237허5358'
SORENTO = '16소4041'


class _Log:
    """모듈 logger 대신 — 어느 단계로 무엇을 남겼는지 모은다."""

    def __init__(self):
        self.calls = []

    def _add(self, level):
        def f(msg, *a, **k):
            self.calls.append((level, str(msg)))
        return f

    def __getattr__(self, name):
        if name in ('debug', 'info', 'warning', 'error', 'critical', 'exception'):
            return self._add(name)
        raise AttributeError(name)

    def at(self, level):
        return [m for lv, m in self.calls if lv == level]


@pytest.fixture
def vdb(tmp_path, monkeypatch):
    """차량 배정·운행일지·운전자 답을 한 임시 DB 로(test_vehicle_assign 의 db 와 같은 격리 + 운전자 답 표)."""
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, vl, va, hr, vd):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(vp, '_table_ready', False)      # 방문 일정 표도 같은 임시 DB 에(va._db 를 씀)
    monkeypatch.setattr(va, 'staff', lambda: [dict(p) for p in STAFF])
    monkeypatch.setattr(va, 'assigners', lambda: [ASSIGNER])
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ())
    monkeypatch.setattr(va, 'CAR_DM_ON', True)
    sent = []
    monkeypatch.setattr(va, '_dm', lambda sid, text: sent.append((sid, text)) or True)
    d.create_user('배정', ASSIGNER, 'pw-Test-1234!', 'admin')    # 고치기는 관리자만(11-25)
    gw._invalidate()
    yield sent
    gw._invalidate()


def _people(day=DAY):
    return {c['plate']: c['people'] for c in va.day_view(day)['cars'] if c['people']}


# ── 17. 차량 배정 동시 편집 ────────────────────────────────────────────────
def test_stale_base_rejected_and_nothing_overwritten(vdb):
    base = va.day_view(DAY)['base']
    # SJ 가 먼저 저장 — 새 판을 돌려받는다
    r1 = va.save_day(DAY, [{'plate': RAY, 'people': ['TH']}], ASSIGNER, base=base)
    assert r1['changed'] == [RAY] and r1['base'] and r1['base'] != base
    # 같은 화면을 열어 둔 JW 가 옛 판으로 저장 → 거절, SJ 가 넣은 TH 는 그대로
    with pytest.raises(va.StaleSave) as e:
        va.save_day(DAY, [{'plate': SORENTO, 'people': ['YM']}], 'jw@example.com', base=base)
    assert str(e.value) == va.STALE_MSG == '다른 분이 먼저 고쳤어요. 새로 불러온 뒤 다시 해 주세요.'
    assert _people() == {RAY: ['TH']}
    # SJ 가 돌려받은 판으로 연달아 저장 — 막히지 않음
    r2 = va.save_day(DAY, [{'plate': RAY, 'people': ['TH', 'YM']}], ASSIGNER, base=r1['base'])
    assert r2['changed'] == [RAY] and r2['base'] != r1['base']
    # 바뀐 게 없는 저장은 판도 그대로(다음 저장이 막히지 않게)
    r3 = va.save_day(DAY, [{'plate': RAY, 'people': ['TH', 'YM']}], ASSIGNER, base=r2['base'])
    assert r3['changed'] == [] and r3['base'] == r2['base'] == va.day_version(DAY)
    # base 를 안 보내는 옛 화면(재시작 직후 열려 있던 것)은 예전처럼 저장
    va.save_day(DAY, [{'plate': SORENTO, 'people': ['JK']}], ASSIGNER)
    assert _people() == {SORENTO: ['JK']}


def test_clearing_a_car_and_confirming_change_the_base(vdb):
    """차를 비우면 그 줄이 지워져 updated_at 최댓값이 그대로일 수 있다 — 판은 로그 id 라 바뀐다. 확정도 판을 바꾼다."""
    r = va.save_day(DAY, [{'plate': RAY, 'people': ['TH']}, {'plate': SORENTO, 'people': ['YM']}], ASSIGNER)
    b1 = r['base']
    b2 = va.save_day(DAY, [{'plate': RAY, 'people': ['TH']}], ASSIGNER, base=b1)['base']    # 쏘렌토 비움
    assert b2 != b1
    with pytest.raises(va.StaleSave):
        va.save_day(DAY, [{'plate': RAY, 'people': ['TH']}, {'plate': SORENTO, 'people': ['YM']}], 'jw@example.com', base=b1)
    # 다른 분이 확정만 해도(바뀐 줄 없음) 판이 바뀐다 — 옛 화면이 '확정 전'인 줄 알고 저장하지 않게
    b3 = va.save_day(DAY, [{'plate': RAY, 'people': ['TH']}], ASSIGNER, confirm=True, base=b2)['base']
    assert b3 != b2


def test_stale_save_after_confirm_sends_no_dm(vdb):
    """확정 뒤 옛 화면의 저장이 앞 사람 변경을 덮어쓰면 '빠졌어요' DM 까지 나가던 것 — 거절이면 DM 도 없다."""
    sent = vdb
    base = va.day_view(DAY)['base']
    b = va.save_day(DAY, [{'plate': RAY, 'people': ['TH', 'YM']}], ASSIGNER, confirm=True, base=base)['base']
    sent.clear()
    va.save_day(DAY, [{'plate': RAY, 'people': ['TH', 'YM', 'JK']}], ASSIGNER, base=b)   # SJ 가 JK 추가
    sent.clear()
    with pytest.raises(va.StaleSave):     # JW 는 JK 추가 전 화면 그대로 저장 시도
        va.save_day(DAY, [{'plate': RAY, 'people': ['TH', 'YM']}], 'jw@example.com', base=b)
    assert sent == [] and _people() == {RAY: ['TH', 'YM', 'JK']}


def _assign_app(monkeypatch, leads=None):
    import dashboard.auth as auth
    from dashboard.blueprints import gw_vehicles as bp
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(hr, 'hr_admins', lambda: [])
    monkeypatch.setattr(bp, 'render_template', lambda name, **k: name)
    monkeypatch.setattr(bp, '_leads', leads or (lambda: [_lead('L-1', DAY, '부천 길주로 1', '서준')]))
    monkeypatch.setattr(va, 'day_off', lambda d: {'off': [], 'half': []})
    monkeypatch.setattr(bp, '_resolve_later', lambda s, e: None)
    log = _Log()
    monkeypatch.setattr(bp, 'logger', log)
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    return app, bp, log


def test_api_save_409_on_stale_base(vdb, monkeypatch):
    app, bp, _ = _assign_app(monkeypatch)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ASSIGNER, 'permission_level': 'user'}
        j = cl.get(f'/api/vehicles/assign?day={DAY}').get_json()
        base = j['base']
        assert j['success'] and base and j['partial'] == []
        r = cl.post('/api/vehicles/assign', json={'day': DAY, 'base': base, 'rows': [{'plate': RAY, 'people': ['TH']}]})
        assert r.status_code == 200 and r.get_json()['base'] != base
        r = cl.post('/api/vehicles/assign', json={'day': DAY, 'base': base, 'rows': [{'plate': RAY, 'people': ['YM']}]})
        assert r.status_code == 409
        assert r.get_json() == {'success': False, 'conflict': True,
                                'message': '다른 분이 먼저 고쳤어요. 새로 불러온 뒤 다시 해 주세요.'}
        assert _people() == {RAY: ['TH']}
        # base 없는 옛 화면 요청은 그대로
        r = cl.post('/api/vehicles/assign', json={'day': DAY, 'rows': [{'plate': RAY, 'people': ['YM']}]})
        assert r.status_code == 200 and _people() == {RAY: ['YM']}
        # 숫자처럼 보이는 base 도 글자로 비교(화면은 받은 그대로 돌려보냄)
        new = cl.get(f'/api/vehicles/assign?day={DAY}').get_json()['base']
        assert cl.post('/api/vehicles/assign', json={'day': DAY, 'base': new, 'rows': []}).status_code == 200


# ── 18. 배정 화면 불러오기 — 못 읽은 칸을 partial 로 ────────────────────────
def test_api_get_partial_when_leads_or_leave_fail(vdb, monkeypatch):
    def down():
        raise RuntimeError('시트 503')
    app, bp, log = _assign_app(monkeypatch, leads=down)

    def leave_down(d):
        raise sqlite3.OperationalError('database is locked')
    monkeypatch.setattr(va, 'day_off', leave_down)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ASSIGNER, 'permission_level': 'user'}
        j = cl.get(f'/api/vehicles/assign?day={DAY}').get_json()
    # 배정은 계속 할 수 있게 성공으로 열되, 못 읽은 칸을 알린다
    assert j['success'] and j['can_edit'] and len(j['cars']) == 12
    assert 'leads' in j['partial'] and 'off' in j['partial']
    assert j['visitors'] == [] and j['off'] == []
    errs = log.at('error')
    assert any('리드 시트' in m for m in errs) and any('휴가' in m for m in errs)


def test_api_get_empty_lead_sheet_is_failure(vdb, monkeypatch):
    """_leads() 는 시트 오류를 [] 로 돌려준다(load_leads_data → None) — 수천 줄 시트가 비었으면 못 읽은 것."""
    app, bp, log = _assign_app(monkeypatch, leads=lambda: [])
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ASSIGNER, 'permission_level': 'user'}
        j = cl.get(f'/api/vehicles/assign?day={DAY}').get_json()
    assert j['success'] and j['partial'] == ['leads']
    assert any('리드 시트' in m for m in log.at('error'))


def test_api_get_other_parts_fail(vdb, monkeypatch):
    app, bp, log = _assign_app(monkeypatch)

    def boom(*a, **k):
        raise ValueError('깨짐')
    monkeypatch.setattr(va, 'day_schedule', boom)
    monkeypatch.setattr(va, 'online_duty', boom)
    monkeypatch.setattr(vp, 'confirmed_work', boom)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ASSIGNER, 'permission_level': 'user'}
        j = cl.get(f'/api/vehicles/assign?day={DAY}').get_json()
    # 확정한 팀(vp.confirmed_teams, 11-22)도 confirmed_work 를 읽어 같이 실패 → 'teams' 도 알림
    assert j['success'] and set(j['partial']) == {'schedule', 'online', 'plan', 'teams'}
    assert j['visitors'] and j['visitors'][0]['code'] == 'TH'      # 리드는 읽혔으니 방문 사람은 그대로
    assert len(log.at('error')) == 4 and not log.at('warning')


# ── 23. 운행일지 배정·운전자 읽기 ─────────────────────────────────────────
def test_no_module_globals_for_shifts():
    """오전·오후 운전자·운전자 답은 요청마다 따로 — 모듈 전역이면 동시 요청 때 서로 덮어씀."""
    assert not hasattr(vl, '_SHIFTS') and not hasattr(vl, '_CONFIRMED')


def test_assignments_separate_per_call(vdb):
    va.save_day('2026-11-23', [{'plate': SORENTO, 'people': ['YG'], 'pm': ['SJ']}], ASSIGNER)
    va.save_day('2026-12-01', [{'plate': SORENTO, 'people': ['JK']}], ASSIGNER)
    nov = vl._assignments('2026-11-01', '2026-11-30')
    dec = vl._assignments('2026-12-01', '2026-12-31')     # 나중 호출이 앞 결과를 덮어쓰지 않는다
    assert nov.error == '' and dec.error == ''
    assert vl.trip_driver('2026-11-23', SORENTO, '14:00', nov) == '오늘'
    assert vl.trip_driver('2026-11-23', SORENTO, '09:00', nov) == '임산'
    assert vl.trip_driver('2026-11-23', SORENTO, '14:00', dec) == ''
    assert vl.trip_driver('2026-12-01', SORENTO, '09:00', dec) == '강해'


def test_assignments_read_fail_is_error_and_flagged(vdb, monkeypatch):
    log = _Log()
    monkeypatch.setattr(vl, 'logger', log)
    vl.import_upload('a.xlsx', _xlsx(trips=T1), 'x')

    def down(*a):
        raise sqlite3.OperationalError('database is locked')
    monkeypatch.setattr(va, 'assigned_people', down)
    a = vl._assignments('2026-11-01', '2026-11-30')
    assert a.error == 'assign' and a.rows == {} and a.fixed == {}
    assert any('차량 배정을 못 읽어' in m for m in log.at('error'))
    s = vl.month_summary('2026-11', LEADS)
    assert s['assign_error'] == 'assign' and s['vehicles'][0]['source'] == 'guess'    # 추정만 — 화면이 한 줄로 알림
    assert vl.vehicle_days('2026-11', '12가3456', LEADS)['assign_error'] == 'assign'


def test_driver_answers_read_fail_keeps_assignment(vdb, monkeypatch):
    log = _Log()
    monkeypatch.setattr(vl, 'logger', log)
    va.save_day('2026-11-23', [{'plate': SORENTO, 'people': ['YG']}], ASSIGNER)

    def down(*a):
        raise sqlite3.OperationalError('database is locked')
    monkeypatch.setattr(vd, 'actual_drivers', down)
    a = vl._assignments('2026-11-01', '2026-11-30')
    assert a.error == 'driver' and a.rows[('2026-11-23', SORENTO)] == ['임산'] and not a.confirmed
    assert any('운전자 답' in m for m in log.at('error'))


def test_assignments_ok_has_no_error(vdb, monkeypatch):
    log = _Log()
    monkeypatch.setattr(vl, 'logger', log)
    vl.import_upload('a.xlsx', _xlsx(trips=T1), 'x')
    assert vl.month_summary('2026-11', LEADS)['assign_error'] == ''
    assert vl.vehicle_days('2026-11', '12가3456', LEADS)['assign_error'] == ''
    assert not log.at('error')


# ── 24. 차량 관리 결재 비용 모으기 ────────────────────────────────────────
@pytest.fixture
def mdb(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, va, vm, hr):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(vm, '_last_sync', 0.0)
    monkeypatch.setattr(vm, '_zero_warned', set())
    log = _Log()
    monkeypatch.setattr(vm, 'logger', log)
    gw._invalidate()
    yield d, log
    gw._invalidate()


def _flex_table(conn):
    conn.execute('CREATE TABLE flex_approval_archive (doc_no TEXT, title TEXT, form TEXT, status TEXT, writer TEXT, '
                 'created_at TEXT, updated_at TEXT, attachments TEXT, memo TEXT, fields TEXT, amount INTEGER, '
                 'project_code TEXT, partner TEXT, imported_at TEXT)')


def _approvals_table(conn):
    conn.execute('CREATE TABLE approvals (id INTEGER PRIMARY KEY, form TEXT, status TEXT, content TEXT, '
                 'amount INTEGER, partner TEXT, created_at TEXT)')


def test_missing_tables_only_debug(mdb):
    d, log = mdb
    assert vm.collect_from_approvals() == []
    assert not log.at('error') and len(log.at('debug')) == 2


def test_other_db_errors_are_error(mdb):
    d, log = mdb
    with d._get_connection() as conn:
        _flex_table(conn)
        conn.execute('CREATE TABLE approvals (id INTEGER PRIMARY KEY, form TEXT, status TEXT)')   # 컬럼이 바뀐 표
        conn.commit()
    assert vm.collect_from_approvals() == []
    errs = log.at('error')
    assert len(errs) == 1 and '전자결재' in errs[0]


def test_one_bad_row_does_not_skip_the_rest(mdb):
    d, log = mdb
    with d._get_connection() as conn:
        _flex_table(conn)
        _approvals_table(conn)
        conn.executemany('INSERT INTO flex_approval_archive (doc_no, form, status, created_at, memo, fields, amount) '
                         'VALUES (?,?,?,?,?,?,?)', [
                             ('2026-A', '개인경비청구서', '완료', '2026-11-05 02:06', '', '[]', 0),     # fields 가 목록 — 읽기 실패
                             ('2026-B', '개인경비청구서', '완료', '2026-11-06 02:06', '',
                              json.dumps({'사용일자': '2026-11-06', '청구 금액': '40000', '사유': '6873 주차비'}), 40000)])
        conn.execute("INSERT INTO approvals VALUES (30, 'fund', 'done', '카니발 엔진오일', 210000, '가나정비', '2026-11-11 16:43')")
        conn.commit()
    rows = vm.collect_from_approvals()
    assert sorted(r['key'] for r in rows) == ['appr:30', 'flex:2026-B']
    errs = log.at('error')
    assert len(errs) == 1 and '1줄' in errs[0] and '2026-A' in errs[0]


def test_zero_amount_not_inserted_and_comes_back_after_fix(mdb):
    d, log = mdb
    with d._get_connection() as conn:
        _flex_table(conn)
        _approvals_table(conn)
        conn.execute("INSERT INTO approvals VALUES (31, 'fund', 'done', '5358 유리', 0, '', '2026-11-12 10:00')")
        conn.execute("INSERT INTO approvals VALUES (32, 'fund', 'done', '4041 타이어', 100000, '', '2026-11-12 11:00')")
        conn.commit()
    assert vm.sync_costs(force=True) == 1
    assert not vm.costs(plate=RAY) and vm.costs(plate=SORENTO)[0]['amount'] == 100000
    warns = log.at('warning')
    assert len(warns) == 1 and 'appr:31' in warns[0]
    vm.sync_costs(force=True)
    assert len(log.at('warning')) == 1              # 같은 줄은 한 번만 알림
    with d._get_connection() as conn:               # 결재 금액을 고치면 다음 모으기 때 들어간다
        conn.execute('UPDATE approvals SET amount = 55000 WHERE id = 31')
        conn.commit()
    assert vm.sync_costs(force=True) == 1 and vm.costs(plate=RAY)[0]['amount'] == 55000
