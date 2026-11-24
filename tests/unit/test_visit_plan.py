"""그룹웨어 방문 일정(2026-11-22) — 캔버스1 현장에서 시작해 코스로 묶고 사람 넣기, 일정 다듬기·점검, 확정(뒤에서 돎)과
/일정확정 본체 공용(리드 직접 연결·팀 메모 DM·DM 대상일 고정). 이름·번호는 모두 가짜."""
from datetime import date, timedelta

import pytest

import dashboard.services.visit_assignment_sync as vas
import dashboard.utils.redis_client as rcmod
from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import visit_plan as vp

STAFF = [{'code': c, 'name': n, 'slack_id': 'U' + c, 'dept': ''} for c, n in [
    ('YM', '김가람'), ('LSH', '이나래'), ('TH', '서준'), ('JK', '강해'), ('SJ', '오늘')]]
DAY = (date.today() + timedelta(days=2)).isoformat()


def _lead(no, who='', owner='', mark=''):
    return {'리드 No': no, '방문 예정일': DAY, '상태': '방문 예약', '영업 담당자': who, '고객명': no + '상호',
            '고객 연락처': '010-0000-0000', '방문 주소': f'가상시 {no}로 1', '플랫폼': '홈페이지',
            '상담 내용': '천장형 실측', '온라인 상담자': owner, '본인 방문 여부': mark}


LEADS = [_lead('L-1', '김가람,이나래'), _lead('L-2', '김가람,이나래'), _lead('L-3', '서준'),
         _lead('L-4'), _lead('L-5', '', '강해', '본인 방문 필수')]


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, va, hr):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(vp, '_table_ready', False)
    d.create_user('박기획', 'p@example.com', 'pw-Test-1234!', 'admin')   # 고치기·확정은 관리자만(11-25) — 짜는 사람은 관리자
    monkeypatch.setattr(va, 'staff', lambda: [dict(p) for p in STAFF])
    monkeypatch.setattr(va, 'day_off', lambda day: {'off': [{'code': 'SJ', 'name': '오늘', 'why': '연차'}], 'half': []})
    monkeypatch.setattr(vas, '_load_initial_maps',
                        lambda: ({p['code']: p['name'] for p in STAFF}, {p['name']: p['code'] for p in STAFF}))
    monkeypatch.setattr(vp, '_name_of_email', lambda e: {'p@example.com': '박기획'}.get(e, e))
    monkeypatch.setattr(vp, 'write_canvas2', lambda day, plan, leads, parsed, by: {'ok': True, 'diff': [], 'read': 0})   # 진짜 캔버스2 금지
    gw._invalidate()
    yield d
    gw._invalidate()


def test_starts_empty_from_canvas1(db):
    """시트 영업 담당자(=캔버스2 결과)로 팀을 묶지 않는다 — 캔버스1 현장이 모두 '배정 안 된 방문'에서 시작(11-22 부장님)."""
    extra = [{**_lead('L-8'), '상태': '방문 완료'}, {**_lead('L-9'), '상태': '유선 상담'},
             {**_lead('L-10'), '상담 내용': '실측 ─── [11.20 09:00 YM · 방문 완료]'}]
    data = vp.load(DAY, LEADS + extra, 'p@example.com')
    assert data['source'] == 'empty' and data['plan'] == {'courses': [], 'online': [], 'online_shared': False}
    assert [v['lead_no'] for v in data['visits']] == ['L-1', 'L-2', 'L-3', 'L-4', 'L-5']   # 캔버스1 기준만
    assert all('who' not in v for v in data['visits'])
    assert next(v for v in data['visits'] if v['lead_no'] == 'L-5')['owner'] == 'JK'
    assert [w['kind'] for w in data['warnings']] == ['left']
    l1 = data['visits'][0]
    assert l1['place'] == '가상 L-1로' or l1['region'] == '가상'     # 지역 꼬리표(가짜 주소라 시 이름만)
    assert l1['region'] == '가상'


def test_who_tag(db):
    """누구 현장인지(11-23 부장님) — 거래처·기타는 늘 상담자 이름, 온라인은 본인 방문 필수일 때만."""
    extra = [{**_lead('L-20', '', '서준'), '플랫폼': '거래처'},                       # 거래처 · 필수 아님
             {**_lead('L-21', '', '김가람', '본인 방문 필수'), '플랫폼': '거래처'},    # 거래처 · 필수
             _lead('L-22', '', '이나래')]                                               # 온라인 · 필수 아님
    vs = {v['lead_no']: v for v in vp.load(DAY, LEADS + extra, 'p@example.com')['visits']}
    assert (vs['L-20']['contact'], vs['L-20']['must'], vs['L-20']['owner']) == ('서준', False, '')
    assert (vs['L-21']['contact'], vs['L-21']['must'], vs['L-21']['owner']) == ('김가람', True, 'YM')
    assert (vs['L-22']['contact'], vs['L-22']['must']) == ('', False)
    assert (vs['L-5']['contact'], vs['L-5']['must']) == ('강해', True)                 # 온라인 · 필수
    assert all('본인 방문 필수' not in v['marks'] for v in vs.values())                 # 이름과 한 꼬리표로


def test_cars_by_person(db, monkeypatch):
    """코스 제목 줄 차 꼬리표(11-23) — 차량 배정에 저장된 차를 사람별로, 오후 사람도. 못 읽으면 빈 값."""
    monkeypatch.setattr(va, '_day_rows', lambda day: {'11가1111': va._row(['YM', 'TH'], '', ['YM', 'JK']),
                                                      '22나2222': va._row([])})
    monkeypatch.setattr(va, 'fleet', lambda include_inactive=False: [{'plate': '11가1111', 'label': '스타렉스'}])
    cars = vp.cars_by_person(DAY)
    assert cars == {'YM': [{'plate': '11가1111', 'label': '스타렉스'}], 'TH': [{'plate': '11가1111', 'label': '스타렉스'}],
                    'JK': [{'plate': '11가1111', 'label': '스타렉스'}]}
    monkeypatch.setattr(va, '_day_rows', lambda day: (_ for _ in ()).throw(RuntimeError('db')))
    assert vp.cars_by_person(DAY) == {}


def test_region_rank():
    """현장 목록 지역 순서 서울 → 인천 → 경기 → 그 밖(11-23 부장님). 축약 주소엔 '경기'가 없어 시 이름으로."""
    from dashboard.services import vehicle_log as vl
    ranks = {a: vp._where(a, vl)['region_rank'] for a in
             ('마포구 양화로 45', '인천 서구 가정로 1', '용인 기흥구 동백로 1', '광주 오포읍 1', '광주 동구 금남로 1', '천안 서북구 1')}
    assert list(ranks.values()) == [0, 1, 2, 2, 3, 3]


def test_clean_plan_rules():
    codes, nos = ['YM', 'TH'], ['L-1', 'L-2']
    raw = {'courses': [{'name': '  평택  라인 ', 'members': ['ym', 'YM'], 'leads': ['L-2', 'L-1', 'L-1'], 'memo': '  평택\n끝나고  '},
                       {'members': ['TH'], 'leads': ['L-1']}, {'members': [], 'leads': []}],
           'online': [{'code': 'TH', 'shift': '오후'}, {'code': 'TH'}, {'code': 'YM', 'shift': '밤'}],
           'online_shared': True}
    p = vp.clean_plan(raw, nos, codes)
    assert p['courses'] == [{'name': '평택 라인', 'members': ['YM'], 'leads': ['L-2', 'L-1'], 'memo': '평택 끝나고'},
                            {'name': '', 'members': ['TH'], 'leads': [], 'memo': ''}]   # 한 현장은 한 코스에만, 가는 순서 그대로
    assert p['online'] == [{'code': 'TH', 'shift': '오후'}, {'code': 'YM', 'shift': ''}]
    assert p['online_shared'] is False                                            # 당번이 있으면 공용 아님
    with pytest.raises(ValueError, match='바뀐 건'):
        vp.clean_plan({'courses': [{'members': ['YM'], 'leads': ['L-9']}]}, nos, codes)
    with pytest.raises(ValueError, match='명단에 없는'):
        vp.clean_plan({'courses': [{'members': ['ZZ'], 'leads': []}]}, nos, codes)
    assert vp.clean_plan({'courses': [{'members': ['ZZ'], 'leads': ['L-9']}]}, nos, codes, strict=False)['courses'] == []
    # 앞 판 저장분(팀 칸·현장별 사람)도 코스로 읽는다
    assert vp.clean_plan({'teams': [{'members': ['TH'], 'leads': ['L-2'], 'memo': 'm'}]}, nos, codes)['courses'] == \
        [{'name': '', 'members': ['TH'], 'leads': ['L-2'], 'memo': 'm'}]
    old = vp.clean_plan({'sites': {'L-1': {'members': ['YM']}, 'L-2': {'members': ['YM'], 'memo': 'x'}}}, nos, codes)
    assert old['courses'] == [{'name': '', 'members': ['YM'], 'leads': ['L-1', 'L-2'], 'memo': 'x'}]


def test_warnings(db):
    vs = vp.visits(DAY, LEADS)
    names = {p['code']: p['name'] for p in STAFF}
    plan = {'courses': [{'name': '', 'members': [], 'leads': ['L-1'], 'memo': ''},
                        {'name': '고덕', 'members': ['SJ', 'YM'], 'leads': ['L-5'], 'memo': ''},
                        {'name': '', 'members': ['YM'], 'leads': ['L-2', 'L-3'], 'memo': ''}],
            'online': [{'code': 'SJ', 'shift': ''}], 'online_shared': False}
    ws = vp.warnings(plan, vs, va.day_off(DAY), names)
    by = {w['kind']: w for w in ws}
    assert '2건' in by['left']['text']                              # L-1(사람 없는 코스)·L-4(코스 없음)
    assert by['nobody']['text'] == '코스 1에 사람을 넣지 않았어요(현장 1건).'
    assert by['owner']['text'].startswith('고덕의 본인 방문 필수 현장')
    offs = [w for w in ws if w['kind'] == 'off']
    assert len(offs) == 2                                            # 코스에도, 온라인 당번에도 연차인 오늘
    assert all(w['level'] == 'error' and '빼야 확정돼요' in w['text'] for w in offs)   # 휴가자는 확정을 막음(11-23)
    assert 'twice' in by and all(w['level'] == 'warn' for w in ws if w['kind'] != 'off')
    # 확정 전 토요 휴무는 막지 않고 경고만
    soft = {'off': [{'code': 'SJ', 'name': '오늘', 'why': '스케줄 근무 쉬는 날(확정 전)', 'soft': True}], 'half': []}
    assert {w['level'] for w in vp.warnings(plan, vs, soft, names) if w['kind'] == 'off'} == {'warn'}


def test_work_course(db, monkeypatch):
    """사내 업무(11-23 부장님) — 현장 없이 사람 + 할 일. 이름은 종류마다 따로 세고, 할 일이 바뀐 사람에게만 DM."""
    raw = {'courses': [{'members': ['YM'], 'leads': ['L-1'], 'memo': ''},
                       {'kind': 'work', 'members': ['TH', 'LSH'], 'leads': ['L-2'], 'memo': '창고 정비'},
                       {'kind': 'work', 'members': ['JK'], 'leads': [], 'memo': ''}],
           'online': [], 'online_shared': False}
    plan = vp.clean_plan(raw, ['L-1', 'L-2'], ['YM', 'TH', 'LSH', 'JK'])
    assert plan['courses'][1] == {'name': '', 'members': ['TH', 'LSH'], 'leads': [], 'memo': '창고 정비', 'kind': 'work'}
    assert 'kind' not in plan['courses'][0]                                   # 방문 코스는 예전 모양 그대로
    assert vp.course_labels(plan) == ['코스 1', '사내 업무 1', '사내 업무 2']
    names = {p['code']: p['name'] for p in STAFF}
    ws = vp.warnings(plan, vp.visits(DAY, LEADS), va.day_off(DAY), names)
    texts = [w['text'] for w in ws]
    assert '사내 업무 2에 할 일을 적어 주세요.' in texts
    assert not any(w['kind'] == 'idle' for w in ws)                           # 남는 현장 직원 없음(오늘=연차)
    plan['courses'].pop()                                                     # 강해가 남음
    idle = [w for w in vp.warnings(plan, vp.visits(DAY, LEADS), va.day_off(DAY), names) if w['kind'] == 'idle']
    assert idle and '강해' in idle[0]['text'] and idle[0]['level'] == 'warn'
    # DM — 처음은 모두, 다시 확정하면 바뀐 사람만
    sent = []
    monkeypatch.setattr(va, '_dm', lambda sid, text: sent.append((sid, text)) or True)
    assert vp._send_work_dms(DAY, plan, None) == {'work_dm': 2, 'work_dm_failed': 0}
    assert sent[0][0] == 'UTH' and '할 일 : 창고 정비' in sent[0][1] and '함께 : 이나래' in sent[0][1]
    sent.clear()
    assert vp._send_work_dms(DAY, plan, plan) == {'work_dm': 0, 'work_dm_failed': 0}
    new = {**plan, 'courses': [plan['courses'][0], {**plan['courses'][1], 'members': ['TH']}]}
    assert vp._send_work_dms(DAY, new, plan)['work_dm'] == 1 and sent[0][0] == 'UTH'   # 같이 하는 사람이 바뀜
    # 차량 배정은 확정한 사내 업무 사람을 '배정 없음'에서 뺀다
    saved = vp.save(DAY, plan, 'p@example.com', LEADS)
    vp._finish(DAY, 'done', {}, 'p@example.com', saved)
    assert vp.confirmed_work(DAY) == [{'code': 'TH', 'task': '창고 정비'}, {'code': 'LSH', 'task': '창고 정비'}]
    assert vp.confirmed_info(DAY)['courses'] == 1                            # 방문 코스만 셈


def test_past_day_from_sheet(db, monkeypatch):
    """오늘·지난 날(슬랙 /일정확정으로 정한 날) — 방문 완료까지, 시트 영업 담당자 조합대로 코스(보기만, 11-23 부장님)."""
    past = (date.today() - timedelta(days=1)).isoformat()
    leads = [{**l, '방문 예정일': past} for l in LEADS] + [{**_lead('L-9', '서준'), '방문 예정일': past, '상태': '방문 완료'}]
    monkeypatch.setattr(va, 'online_duty', lambda d: [{'code': 'JK', 'name': '강해', 'shift': '오후'}])
    d = vp.load(past, leads, 'p@example.com')
    assert d['source'] == 'sheet' and not d['editable']
    assert 'L-9' in [v['lead_no'] for v in d['visits']]                      # 방문 완료도 보임
    assert [(c['members'], c['leads']) for c in d['plan']['courses']] == [(['YM', 'LSH'], ['L-1', 'L-2']), (['TH'], ['L-3', 'L-9'])]
    assert d['plan']['online'] == [{'code': 'JK', 'shift': '오후'}]
    with vp._conn() as conn:                                                  # 보기만 — 저장하지 않음
        assert conn.execute('SELECT COUNT(*) FROM visit_plan_day WHERE day = ?', (past,)).fetchone()[0] == 0
    # 앞으로 짤 날은 그대로 캔버스1 기준(시트 담당자로 묶지 않음)
    assert vp.load(DAY, LEADS, 'p@example.com')['source'] == 'empty'


def test_urgent_assign(db, monkeypatch):
    """당일 배정(11-23 부장님) — 오늘만, 관리자만(11-25: 전엔 오늘 온라인 당번도). 시트·List·긴급 DM, 두 번은 안 됨."""
    today = date.today().isoformat()
    leads = [{**_lead('L-1', '김가람'), '방문 예정일': today},
             {**_lead('L-7', '', '서준'), '방문 예정일': today, '플랫폼': '기타', '상담 시간': today.replace('-', '.') + '. 08:12'}]
    monkeypatch.setattr(vp, 'can_plan', lambda e: e == 'p@example.com')
    monkeypatch.setattr(va, 'online_duty', lambda d: [{'code': 'JK', 'name': '강해', 'shift': ''}])
    assert vp.can_urgent('p@example.com', today) and not vp.can_urgent('jk@example.com', today)   # 관리자만 — 당번은 이제 못 함
    assert not vp.can_urgent('ym@example.com', today) and not vp.can_urgent('p@example.com', DAY)  # 그 밖·오늘 아님
    wrote, sent, lists = {}, [], []
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'update_lead', lambda no, vals: wrote.update({no: vals}) or {'success': True})
    monkeypatch.setattr(vas, '_update_slack_list_managers', lambda a, pm, sheet_leads=None: lists.append(a[0]['assign']) or (1, 0, []))
    monkeypatch.setattr(vas, '_get_visit_client', lambda: None)
    monkeypatch.setattr(va, '_dm', lambda sid, text: sent.append((sid, text)) or True)
    import dashboard.services.visit_canvas_sync as vcs
    monkeypatch.setattr(vcs, 'rebuild_canvas_async', lambda: None)
    for who in ('ym@example.com', 'jk@example.com'):
        with pytest.raises(PermissionError):
            vp.urgent_assign(today, 'L-7', ['TH'], who, leads, background=False)
    with pytest.raises(ValueError, match='이미 담당자'):
        vp.urgent_assign(today, 'L-1', ['TH'], 'p@example.com', leads, background=False)
    with pytest.raises(ValueError, match='연차'):
        vp.urgent_assign(today, 'L-7', ['SJ'], 'p@example.com', leads, background=False)
    res = vp.urgent_assign(today, 'L-7', ['TH'], 'p@example.com', leads, background=False)
    assert res == {'names': '서준', 'dm': 1, 'dm_failed': 0}
    assert wrote == {'L-7': {'영업 담당자': '서준'}} and lists == [['TH']]
    assert sent[0][0] == 'UTH' and ':rotating_light:' in sent[0][1] and '당일 배정' in sent[0][1] and 'L-7' in sent[0][1]
    with pytest.raises(ValueError, match='방금 다른 사람'):
        vp.urgent_assign(today, 'L-7', ['YM'], 'p@example.com', leads, background=False)
    d = vp.load(today, leads, 'p@example.com')                               # 시트가 늦게 바뀌어도 코스에 붙어 보임
    assert any('L-7' in c['leads'] and c['members'] == ['TH'] for c in d['plan']['courses'])
    assert d['urgent']['L-7']['members'] == ['TH'] and d['urgent']['L-7']['state'] == 'done'
    assert d['urgent']['L-7']['result'] == {'dm': 1, 'dm_failed': 0}
    assert next(v for v in d['visits'] if v['lead_no'] == 'L-7')['reg_at'] == today + ' 08:12'


def test_editable_only_future():
    today = date(2026, 11, 22)
    assert not vp.editable('2026-11-22', today) and not vp.editable('2026-11-21', today)
    assert vp.editable('2026-11-23', today) and not vp.editable('0001-11-23x', today)


def test_build_parsed_order_memo_off(db):
    plan = {'courses': [{'name': '', 'members': ['TH'], 'leads': ['L-3'], 'memo': '먼저'},
                        {'name': '', 'members': [], 'leads': ['L-4'], 'memo': ''},
                        {'name': '', 'members': ['YM', 'LSH'], 'leads': ['L-2', 'L-1'], 'memo': ''}],
            'online': [{'code': 'JK', 'shift': '오전'}, {'code': 'TH', 'shift': ''}], 'online_shared': False}
    p = vp.build_parsed(DAY, plan, LEADS)
    assert [a['lead_no'] for a in p['assignments']] == ['L-3', 'L-2', 'L-1']   # 코스 순서 → 가는 순서, 사람 없는 코스 빠짐
    assert p['assignments'][0]['memo'] == '먼저' and p['assignments'][1]['assign'] == ['YM', 'LSH']
    assert p['assignments'][0]['lead']['리드 No'] == 'L-3'
    assert p['online_duty'] == ['JK', 'TH'] and p['online_duty_shifts'] == {'JK': '오전'}
    assert p['off_duty'] == ['SJ']
    with pytest.raises(ValueError, match='찾지 못했'):
        vp.build_parsed(DAY, plan, LEADS[:2])


class _NowThread:
    def __init__(self, target, args=(), **k):
        self.target, self.args = target, args

    def start(self):
        self.target(*self.args)


def test_confirm_runs_and_locks_slack(db, monkeypatch):
    seen = {}

    def fake_commit(parsed, target_date=None):
        seen['parsed'], seen['target'] = parsed, target_date
        return {'ok': True, 'updated_count': 2, 'failed': [], 'list_updated': 2, 'list_failed': 0,
                'list_missing': [], 'dm': {'target_date': DAY, 'visit_mgr_sent': 3, 'online_duty_sent': 1}}
    monkeypatch.setattr(vas, 'commit_plan', fake_commit)
    monkeypatch.setattr(vp.threading, 'Thread', _NowThread)
    monkeypatch.setattr(vp, '_audit', lambda res, by: [])
    monkeypatch.setattr(va, 'request_assignment', lambda day=None, by='': seen.setdefault('req', day))
    monkeypatch.setattr(va, 'next_workday', lambda today=None: DAY)
    plan = {'courses': [{'name': '', 'members': ['YM', 'LSH'], 'leads': ['L-1', 'L-2'], 'memo': '오전 먼저'},
                        {'name': '', 'members': ['TH'], 'leads': ['L-3'], 'memo': ''}],
            'online': [{'code': 'JK', 'shift': '오후'}], 'online_shared': False}
    assert vp.confirmed_upcoming_day() is None
    with pytest.raises(ValueError, match='코스가 없어요'):
        vp.start_confirm(DAY, {'courses': [{'members': [], 'leads': ['L-1']}]}, 'p@example.com', LEADS, lambda: LEADS)
    with pytest.raises(ValueError, match='빼야 확정돼요'):       # 연차인 오늘(SJ)을 넣으면 확정 안 됨(11-23)
        vp.start_confirm(DAY, {'courses': [{'members': ['SJ'], 'leads': ['L-1']}]}, 'p@example.com', LEADS, lambda: LEADS)
    assert vp.start_confirm(DAY, plan, 'p@example.com', LEADS, lambda: LEADS) == {'state': 'running'}
    assert seen['target'] == date.fromisoformat(DAY) and seen['req'] == DAY
    st = vp.run_state(DAY)
    assert st['state'] == 'done' and st['result']['dm'] == 3 and st['result']['sheet'] == 2
    assert va.online_duty(DAY) == [{'code': 'JK', 'name': '강해', 'shift': '오후'}]
    assert vp.confirmed_teams(DAY) == [
        {'members': ['YM', 'LSH'], 'visits': 2, 'big': False, 'label': '코스 1', 'memo': '오전 먼저', 'leads': ['L-1', 'L-2'], 'also': {}},
        {'members': ['TH'], 'visits': 1, 'big': False, 'label': '코스 2', 'memo': '', 'leads': ['L-3'], 'also': {}}]
    assert vp.confirmed_upcoming_day() == {'day': DAY, 'label': va.day_label(DAY), 'by': '박기획'}
    data = vp.load(DAY, LEADS, 'p@example.com')
    assert data['source'] == 'saved' and data['confirmed_name'] == '박기획' and not data['changed_since_confirm']
    vp.save(DAY, {**plan, 'courses': plan['courses'][:1]}, 'p@example.com', LEADS)
    assert vp.load(DAY, LEADS, 'p@example.com')['changed_since_confirm'] is True


def test_confirm_failure_kept(db, monkeypatch):
    monkeypatch.setattr(vas, 'commit_plan', lambda parsed, target_date=None: {'ok': False, 'reason': '⏳ 처리 중입니다.'})
    monkeypatch.setattr(vp.threading, 'Thread', _NowThread)
    plan = {'courses': [{'members': ['TH'], 'leads': ['L-3']}], 'online': [], 'online_shared': False}
    vp.start_confirm(DAY, plan, 'p@example.com', LEADS, lambda: LEADS)
    st = vp.run_state(DAY)
    assert st['state'] == 'failed' and st['result']['message'] == '처리 중입니다.'
    assert vp.confirmed_teams(DAY) is None


def test_save_rejects_today(db):
    with pytest.raises(ValueError, match='오늘과 지난'):
        vp.save(date.today().isoformat(), {'courses': []}, 'p@example.com', LEADS)


def test_planner_permission(monkeypatch):
    """고치기는 관리자 계정만(11-25) — 일정 담당 명단(PLANNER_EMAILS)이어도 관리자가 아니면 못 고침."""
    monkeypatch.setattr(hr, 'is_admin_email', lambda e: (e or '').strip().lower() == 'boss@example.com')
    monkeypatch.setattr('dashboard.services.visit_assignment_sync.PLANNER_EMAILS', ('p2@example.com',))
    assert vp.can_plan('boss@example.com') and not vp.can_plan('p2@example.com') and not vp.can_plan('')


def test_can_view_excludes_crew_and_cleaning(monkeypatch):
    """보기는 시공·미화를 뺀 재직자 모두(11-25 부장님) — 관리자는 부서와 상관없이, 인사카드 없거나 퇴사면 안 보임."""
    people = {'sales@example.com': {'department': '기술영업', 'status': '재직'},
              'office@example.com': {'department': '경영지원', 'status': '재직'},
              'crew@example.com': {'department': '기술시공(내부)', 'status': '재직'},
              'clean@example.com': {'department': '미화', 'status': '재직'},
              'gone@example.com': {'department': '영업', 'status': '퇴사'},
              'bossc@example.com': {'department': '기술시공(내부)', 'status': '재직'}}
    monkeypatch.setattr(hr, 'find_by_email', lambda e: people.get(e))
    monkeypatch.setattr(hr, 'is_admin_email', lambda e: e == 'bossc@example.com')
    assert vp.can_view('sales@example.com') and vp.can_view('office@example.com') and vp.can_view('bossc@example.com')
    assert not any(vp.can_view(e) for e in ('crew@example.com', 'clean@example.com', 'gone@example.com', 'nobody@example.com', ''))


# ── /일정확정 본체 공용 ──────────────────────────────────────────────────
def test_resolve_uses_attached_lead():
    lead = {'리드 No': 'L-1'}
    assert vas._resolve_lead_for_assignment({'lead': lead, 'phone_digits': '01000000000'}, {}, []) is lead


def test_dm_safe_escapes_mentions():
    assert vas._dm_safe(' <!channel> a&b\n끝 ') == '&lt;!channel&gt; a&amp;b 끝'


class _Client:
    def __init__(self):
        self.posted, self.deleted, self.n = [], [], 0

    def users_lookupByEmail(self, email):
        return {'user': {'id': 'U_' + email.split('@')[0]}}

    def chat_postMessage(self, channel, text, **k):
        self.n += 1
        self.posted.append({'channel': channel, 'text': text})
        return {'ok': True, 'channel': channel, 'ts': str(self.n)}

    def chat_delete(self, channel, ts, **k):
        self.deleted.append(ts)
        return {'ok': True}


class _Redis:
    def __init__(self):
        self.d = {}

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, ex=None, nx=False):
        self.d[k] = v
        return True

    def delete(self, *ks):
        for k in ks:
            self.d.pop(k, None)

    def scan_iter(self, match='*', count=100):
        import fnmatch
        return [k for k in list(self.d) if fnmatch.fnmatch(k, match)]


class _RC:
    def __init__(self, r):
        self.redis = r


def test_memo_in_dm_and_resend_on_memo_change(monkeypatch):
    fc, fr = _Client(), _Redis()
    monkeypatch.setattr(vas, '_get_visit_client', lambda: fc)
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: _RC(fr))
    monkeypatch.setattr(vas, '_email_from_initial', lambda ini, itn: f'{ini.lower()}@example.com')
    monkeypatch.setattr(vas, '_get_visit_card_permalink', lambda c, n: '')
    itn = {p['code']: p['name'] for p in STAFF}
    target = date.today() + timedelta(days=3)          # 여러 날 걸친 방문만 있어도 고른 날로

    def run(memo):
        a = [{'lead': _lead('L-1'), 'assign': ['YM'], 'memo': memo}]
        return vas._send_dms_for_next_visit(a, {}, itn, ['JK'], [], addr_candidates=[], target_date=target)
    r = run('평택 끝나고 <!here>')
    assert r['target_date'] == target.isoformat()
    ym = [p for p in fc.posted if p['channel'] == 'U_ym']
    assert len(ym) == 1 and '>:memo: 메모 : 평택 끝나고 &lt;!here&gt;' in ym[0]['text']
    assert ym[0]['text'].index(':memo:') < ym[0]['text'].index('L-1')             # 코스 메모는 위쪽에 한 번
    run('평택 끝나고 <!here>')
    assert len([p for p in fc.posted if p['channel'] == 'U_ym']) == 1           # 같으면 다시 안 보냄
    run('오산 먼저')
    ym = [p for p in fc.posted if p['channel'] == 'U_ym']
    assert len(ym) == 2 and '메모 : 오산 먼저' in ym[1]['text'] and '일정 업데이트' in ym[1]['text']


def test_canvas_commit_still_matches_by_phone(monkeypatch):
    """캔버스(/일정확정) 길은 그대로 — 리드가 안 붙은 배정은 전화번호로 시트에서 찾는다."""
    calls = {}
    monkeypatch.setattr(vas, '_match_leads_by_phone', lambda a: calls.setdefault('match', {'01011112222': [_lead('L-7')]}))
    monkeypatch.setattr(vas, '_load_addr_candidates', lambda: [])
    monkeypatch.setattr(vas, '_load_initial_maps', lambda: ({'YM': '김가람'}, {'김가람': 'YM'}))
    monkeypatch.setattr('dashboard.services.lead_service.update_lead',
                        lambda no, vals: calls.setdefault('upd', []).append((no, vals)) or {'success': True})
    monkeypatch.setattr(vas, '_update_slack_list_managers', lambda a, pm, sheet_leads=None: (0, 0, []))
    monkeypatch.setattr(vas, '_send_dms_for_next_visit', lambda *a, **k: calls.setdefault('dm_target', k.get('target_date')) or {})
    monkeypatch.setattr('dashboard.services.visit_canvas_sync.rebuild_canvas_async', lambda: None)
    monkeypatch.setattr('dashboard.services.visit_canvas_sync._fetch_visit_leads', lambda: [])
    parsed = {'assignments': [{'phone': '010-1111-2222', 'phone_digits': '01011112222', 'assign': ['YM'], 'raw': ''}],
              'online_duty': [], 'off_duty': []}
    res = vas._apply_assignments(parsed)
    assert res['ok'] and calls['upd'] == [('L-7', {'영업 담당자': '김가람'})] and calls['dm_target'] is None


def test_plan_routes_permission(db, monkeypatch):
    """보기는 시공·미화 뺀 직원, 고치기·확정은 관리자만(11-25). 현장 세션은 길이 막혀 있다(문지기 목록에 없음)."""
    import dashboard.auth as auth
    from flask import Flask
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.blueprints.groupware import field_path_allowed
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ())
    monkeypatch.setattr(bp, '_leads', lambda: LEADS)
    monkeypatch.setattr(bp, '_plan_leads', lambda: LEADS)
    monkeypatch.setattr(bp, '_fresh_leads', lambda: LEADS)
    monkeypatch.setattr(vp.threading, 'Thread', _NowThread)
    monkeypatch.setattr(vas, 'commit_plan', lambda parsed, target_date=None: {'ok': True, 'dm': {}})
    monkeypatch.setattr(vp, '_audit', lambda res, by: [])
    monkeypatch.setattr(va, 'request_assignment', lambda day=None, by='': 0)
    monkeypatch.setattr(bp, 'render_template', lambda name, **k: name)
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'name': '사무', 'department': '영업', 'status': '재직'}
                        if e == 'staff@example.com' else None)
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    plan = {'courses': [{'members': ['TH'], 'leads': ['L-3'], 'memo': ''}], 'online': [], 'online_shared': False}
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        assert cl.get('/visits/plan').get_data(as_text=True) == 'gw/visit_plan.html'    # 보기는 열림(11-25)
        j = cl.get(f'/api/vehicles/plan?day={DAY}').get_json()
        assert j['success'] and j['can_edit'] is False and len(j['visits']) == 5
        assert cl.post('/api/vehicles/plan', json={'day': DAY, 'plan': plan}).status_code == 403
        assert cl.post('/api/vehicles/plan/confirm', json={'day': DAY, 'plan': plan}).status_code == 403
        with cl.session_transaction() as s:
            s['user'] = {'email': 'p@example.com', 'permission_level': 'admin'}
        assert cl.get(f'/api/vehicles/plan?day={DAY}').get_json()['can_edit'] is True
        r = cl.post('/api/vehicles/plan', json={'day': DAY, 'plan': plan}).get_json()
        assert r['success'] and r['source'] == 'saved' and r['saved_name'] == '박기획'
        bad = cl.post('/api/vehicles/plan', json={'day': DAY, 'plan': {'courses': [{'members': ['TH'], 'leads': ['L-99']}]}})
        assert bad.status_code == 400
        assert cl.post('/api/vehicles/plan/confirm', json={'day': DAY, 'plan': plan}).get_json()['run'] == {'state': 'running'}
        assert cl.get(f'/api/vehicles/plan/state?day={DAY}').get_json()['run']['state'] == 'done'
        assert cl.get('/visits/plan').status_code == 200                     # 영업 > 방문 일정 화면
        info = vp.confirmed_info(DAY)
        assert info['courses'] == 1 and info['confirmed_at']
    assert not field_path_allowed('GET', '/api/vehicles/plan') and not field_path_allowed('POST', '/api/vehicles/plan/confirm')


def test_menu_planner_only():
    """영업 > 방문 일정 메뉴는 일정 담당에게만(11-22)."""
    labels = lambda nav: [c['label'] for a in nav['areas'] if a['key'] == 'sales' for c in a['children']]
    assert '방문 일정' not in labels(gw.build_nav('/leads', 'user'))
    on = labels(gw.build_nav('/visits/plan', 'user', planner=True))
    assert on.index('방문 일정') == on.index('고객 리드') + 1
    assert '방문 일정' not in [c['label'] for a in gw.build_nav('/home', 'field', field=True)['areas'] for c in a['children']]


def test_view_all_edit_admin(db, monkeypatch):
    """방문 일정·차량 배정 — 개발 중 명단과 상관없이 시공·미화 뺀 모두 보기, 고치기는 관리자 계정만(11-25 부장님)."""
    import dashboard.auth as auth
    from flask import Flask
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.blueprints import groupware as gw_bp
    from dashboard.blueprints import hr as hr_bp
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ('boss@example.com',))
    monkeypatch.setattr(bp, '_leads', lambda: LEADS)
    monkeypatch.setattr(bp, '_plan_leads', lambda: LEADS)
    monkeypatch.setattr(bp, 'render_template', lambda name, **k: name)
    monkeypatch.setattr(hr_bp, 'render_template', lambda name, **k: name)
    people = {'staff@example.com': {'name': '사무', 'department': '기술영업', 'status': '재직'},
              'crew@example.com': {'name': '시공', 'department': '기술시공(내부)', 'status': '재직'},
              'clean@example.com': {'name': '미화', 'department': '미화', 'status': '재직'}}
    monkeypatch.setattr(hr, 'find_by_email', lambda e: people.get(e))
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    plan = {'courses': [{'members': ['TH'], 'leads': ['L-3'], 'memo': ''}], 'online': [], 'online_shared': False}
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        assert cl.get('/visits/plan').get_data(as_text=True) == 'gw/visit_plan.html'
        assert cl.get('/vehicles/assign').get_data(as_text=True) == 'gw/vehicle_assign.html'
        assert cl.get(f'/api/vehicles/plan?day={DAY}').get_json()['can_edit'] is False
        assert cl.post('/api/vehicles/plan', json={'day': DAY, 'plan': plan}).status_code == 403
        assert cl.post('/api/vehicles/plan/confirm', json={'day': DAY, 'plan': plan}).status_code == 403
        with cl.session_transaction() as s:
            s['user'] = {'email': 'p@example.com', 'permission_level': 'admin'}
        assert cl.get(f'/api/vehicles/plan?day={DAY}').get_json()['can_edit'] is True
        for who in ('crew@example.com', 'clean@example.com'):
            with cl.session_transaction() as s:
                s['user'] = {'email': who, 'permission_level': 'user'}
            assert cl.get('/visits/plan').status_code == 302 and cl.get('/vehicles/assign').status_code == 302
            assert cl.get(f'/api/vehicles/plan?day={DAY}').status_code == 403
            assert cl.get(f'/api/vehicles/assign?day={DAY}').status_code == 403
    assert gw_bp._planner_menu('a@example.com', 'admin') and gw_bp._planner_menu('a@example.com', 'super_admin')
    assert gw_bp._planner_menu('staff@example.com', 'user')
    assert not gw_bp._planner_menu('crew@example.com', 'user') and not gw_bp._planner_menu('clean@example.com', 'user')


def _confirm(monkeypatch, plan, notices=None):
    monkeypatch.setattr(vas, 'commit_plan', lambda parsed, target_date=None: {'ok': True, 'dm': {}})
    monkeypatch.setattr(vp.threading, 'Thread', _NowThread)
    monkeypatch.setattr(vp, '_audit', lambda res, by: [])
    monkeypatch.setattr(va, 'request_assignment', lambda day=None, by='': 0)
    monkeypatch.setattr(vp, '_staff', lambda: [dict(p) for p in STAFF])
    if notices is not None:
        monkeypatch.setattr(va, 'notify_plan_changed', lambda day, text, by='': notices.append((day, text)) or 1)
    vp.start_confirm(DAY, plan, 'p@example.com', LEADS, lambda: LEADS)
    assert vp.run_state(DAY)['state'] == 'done'


def test_confirmed_teams_follow_courses(db, monkeypatch):
    """차량 배정 팀 = 방문 코스 그대로(11-23) — 코스 순서·이름·메모, 두 코스에 든 사람은 앞 코스에 두고 also, 같은 사람 코스는 이름만 이어 붙임."""
    plan = {'courses': [{'members': ['TH'], 'leads': ['L-3'], 'memo': '자재 싣고'},
                        {'members': ['YM', 'LSH'], 'leads': ['L-1'], 'memo': ''},
                        {'members': ['TH', 'JK'], 'leads': ['L-4'], 'memo': ''},
                        {'kind': 'work', 'members': ['SJ'], 'leads': [], 'memo': '창고 정리'},
                        {'members': ['YM', 'LSH'], 'leads': ['L-2'], 'memo': ''}],
            'online': [], 'online_shared': False}
    monkeypatch.setattr(va, 'day_off', lambda day: {'off': [], 'half': []})
    _confirm(monkeypatch, plan)
    t = vp.confirmed_teams(DAY)
    assert [(x['label'], x['members'], x['visits']) for x in t] == [
        ('코스 1', ['TH'], 1), ('코스 2 · 코스 4', ['YM', 'LSH'], 2), ('코스 3', ['JK'], 1)]
    assert t[0]['memo'] == '자재 싣고' and t[0]['also'] == {'TH': ['코스 3']} and t[1]['leads'] == ['L-1', 'L-2']


def test_confirmed_courses_panel(db, monkeypatch):
    """차량 배정 위 칸을 코스로(11-23) — 코스마다 사람·메모·현장(지역 꼬리표), 사내 업무는 할 일, 코스에 없는 현장·사라진 리드."""
    plan = {'courses': [{'members': ['YM', 'LSH'], 'leads': ['L-1', 'L-2'], 'memo': '오전 먼저'},
                        {'kind': 'work', 'members': ['JK'], 'leads': [], 'memo': '창고 정리'}],
            'online': [], 'online_shared': False}
    _confirm(monkeypatch, plan)
    assert vp.confirmed_courses((date.today() + timedelta(days=5)).isoformat(), LEADS) is None   # 확정 안 한 날
    c = vp.confirmed_courses(DAY, LEADS)
    v1, w1 = c['courses']
    assert (v1['label'], v1['kind'], v1['members'], v1['memo']) == ('코스 1', 'visit', ['YM', 'LSH'], '오전 먼저')
    assert [s['lead_no'] for s in v1['sites']] == ['L-1', 'L-2'] and v1['sites'][0]['addr'] == '가상시 L-1로 1'
    assert 'place' in v1['sites'][0] and 'who' not in v1['sites'][0]
    assert (w1['label'], w1['kind'], w1['task'], w1['memo'], w1['sites']) == ('사내 업무 1', 'work', '창고 정리', '', [])
    assert [s['lead_no'] for s in c['left']] == ['L-3', 'L-4', 'L-5']
    gone = vp.confirmed_courses(DAY, [l for l in LEADS if l['리드 No'] != 'L-2'])     # 날짜를 옮긴 리드
    assert gone['courses'][0]['sites'][1] == {'lead_no': 'L-2', 'missing': True}


def test_reconfirm_tells_car_assigner(db, monkeypatch):
    """다시 확정해 코스 사람이 바뀌면 차량 배정 쪽에 알림(11-23) — 처음 확정·사람 그대로면 안 보냄."""
    notices = []
    base = {'courses': [{'members': ['YM', 'LSH'], 'leads': ['L-1']}, {'members': ['TH'], 'leads': ['L-3']}],
            'online': [], 'online_shared': False}
    _confirm(monkeypatch, base, notices)
    assert notices == []                                                       # 처음 확정
    _confirm(monkeypatch, {**base, 'courses': [{**base['courses'][0], 'memo': '메모만'}, base['courses'][1]]}, notices)
    assert notices == []                                                       # 사람 그대로
    _confirm(monkeypatch, {**base, 'courses': [{'members': ['YM'], 'leads': ['L-1']}, {'members': ['TH', 'JK'], 'leads': ['L-3']}]}, notices)
    assert notices == [(DAY, '추가 강해 · 빠짐 이나래')]
    names = {p['code']: p['name'] for p in STAFF}
    swap = {'courses': [{'members': ['YM', 'TH'], 'leads': ['L-1']}, {'members': ['JK'], 'leads': ['L-3']}]}
    assert vp.plan_change_text({'courses': [{'members': ['YM'], 'leads': ['L-1']}, {'members': ['TH', 'JK'], 'leads': ['L-3']}]},
                               swap, names) == '같이 가는 사람이 바뀜'


def test_cars_see_sheet_courses_on_past_day(db, monkeypatch):
    """차량 배정도 슬랙으로 정한 오늘·지난 날은 시트 담당자 코스로(11-23) — 앞으로 짤 날은 그룹웨어 확정 전엔 없음."""
    past = (date.today() - timedelta(days=1)).isoformat()
    leads = [{**l, '방문 예정일': past} for l in LEADS] + [{**_lead('L-9', '서준'), '방문 예정일': past, '상태': '방문 완료'}]
    monkeypatch.setattr(va, 'online_duty', lambda d: [])
    monkeypatch.setattr(vp, '_staff', lambda: [dict(p) for p in STAFF])
    c = vp.confirmed_courses(past, leads)
    assert c['source'] == 'sheet'
    assert [(x['label'], x['members'], [s['lead_no'] for s in x['sites']]) for x in c['courses']] == [
        ('코스 1', ['YM', 'LSH'], ['L-1', 'L-2']), ('코스 2', ['TH'], ['L-3', 'L-9'])]
    assert [s['lead_no'] for s in c['left']] == ['L-4', 'L-5']
    assert [(t['label'], t['members']) for t in vp.confirmed_teams(past, leads)] == [('코스 1', ['YM', 'LSH']), ('코스 2', ['TH'])]
    assert vp.confirmed_teams(past) is None                                   # 리드 없이 부르면(옛 호출) 그대로
    assert vp.confirmed_courses(DAY, LEADS) is None and vp.confirmed_teams(DAY, LEADS) is None


def test_canvas2_round_trip(db, monkeypatch):
    """캔버스2를 JW 모양으로 써서(11-23) /일정확정 해석기로 다시 읽으면 그룹웨어 코스와 같아야 한다 — 다르면 diff."""
    import re as _re
    from dashboard.services import visit_canvas_sync as vcs
    leads = [{**_lead(n, '', owner, mark), '고객 연락처': f'010-1111-{i:04d}', '플랫폼': plat}
             for i, (n, owner, mark, plat) in enumerate([('L-1', '', '', '홈페이지'), ('L-2', '', '', '당근'),
                                                         ('ETC-3', '서준', '본인 방문 필수', '기타'), ('L-4', '', '', '거래처')])]
    monkeypatch.setattr(vcs, '_get_initial_map', lambda: {p['name']: p['code'] for p in STAFF})
    monkeypatch.setattr(vas, '_match_leads_by_phone', lambda parsed: {vas._normalize_phone(l['고객 연락처']): [l] for l in leads})
    monkeypatch.setattr(vas, '_load_addr_candidates', lambda: [])
    plan = {'courses': [{'members': ['YM', 'LSH'], 'leads': ['L-1', 'L-2'], 'memo': '메모는 안 씀'},
                        {'members': ['TH'], 'leads': ['ETC-3'], 'memo': ''},
                        {'kind': 'work', 'members': ['JK'], 'leads': [], 'memo': '창고 정리'}],
            'online': [{'code': 'SJ', 'shift': '오후'}], 'online_shared': False}
    md = vp.canvas2_markdown(DAY, plan, leads, ['JK'], '박기획', at='11/23 16:38')
    assert md.startswith('# 방문 일정 스케줄 (그룹웨어 확정 11/23 16:38 · 박기획)')
    assert '### YM+LSH' in md and '### TH' in md and '메모는 안 씀' not in md and '창고 정리' not in md
    assert '### 온라인\nSJ(오후)' in md and '### 휴무\nJK' in md
    # 캔버스 HTML 흉내 — 줄마다 <p>
    html = ''.join('<p>' + _re.sub(r'^(#+ |- |> )', '', ln) + '</p>' for ln in md.split('\n') if ln.strip())
    chk = vp.canvas2_check(html, plan)
    assert chk == {'ok': True, 'diff': [], 'read': 3}
    other = {**plan, 'courses': [{'members': ['YM'], 'leads': ['L-1', 'L-2', 'L-4'], 'memo': ''}] + plan['courses'][1:]}
    bad = vp.canvas2_check(html, other)
    assert not bad['ok'] and any('L-1' in d for d in bad['diff']) and any('L-4' in d for d in bad['diff'])


def test_same_site(db):
    """같은 주소의 다른 리드(11-23 L-04642 ↔ L-04650) — 문의 리드끼리, 앞뒤 14일, 드랍·취소 빼고. 기타(ETC-)는 안 봄."""
    addr = '인천 영종구 하늘별빛로65번길 8 비씨디몰'
    past = (date.fromisoformat(DAY) - timedelta(days=1)).isoformat()
    leads = LEADS + [
        {**_lead('L-40'), '방문 주소': addr, '방문 예정일': f'{past}~{DAY[-2:]}'},              # 원래 등록(이틀짜리)
        {**_lead('L-41'), '방문 주소': addr + ' 1층', '방문 예정일': past, '상태': '방문 완료', '영업 담당자': '박민재'},   # 다시 등록
        {**_lead('L-42'), '방문 주소': addr, '방문 예정일': past, '상태': '문의 드랍'},           # 드랍은 안 봄
        {**_lead('ETC-43'), '방문 주소': addr, '방문 예정일': past, '상태': '방문 완료'},         # 기타 방문은 안 봄
        {**_lead('L-44'), '방문 주소': '부산 영종구 하늘별빛로65번길 8', '방문 예정일': past}]     # 다른 시
    vs = vp.visits(DAY, leads)
    r = vp.same_site(DAY, leads, vs)
    assert list(r) == ['L-40'] and [x['lead_no'] for x in r['L-40']] == ['L-41']
    assert r['L-40'][0]['status'] == '방문 완료' and r['L-40'][0]['who'] == '박민재'
    assert 'same_site' in vp.load(DAY, leads, 'p@example.com')


def test_visit_tabs_only_when_both_open(db, monkeypatch):
    """방문 일정 ↔ 차량 배정 탭(11-24) — 두 화면은 같은 사람에게 열려(11-25) 볼 수 있으면 둘 다."""
    from flask import Flask
    from dashboard.blueprints import gw_vehicles as bp
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'department': '기술영업', 'status': '재직'} if e == 'staff@example.com' else None)
    app = Flask(__name__)
    app.secret_key = 'x'
    cases = [('p@example.com', 'admin', {'plan': True, 'cars': True}),
             ('staff@example.com', 'user', {'plan': True, 'cars': True}),
             ('nobody@example.com', 'user', {'plan': False, 'cars': False})]    # 인사카드 없음
    for email, level, want in cases:
        with app.test_request_context('/visits/plan'):
            from flask import session
            session['user'] = {'email': email, 'permission_level': level}
            assert bp._vtabs() == want, email


def test_plan_save_stops_when_leads_unreadable(db, monkeypatch):
    """리드 시트를 못 읽으면 저장·불러오기·확정을 멈춘다 — 빈 리드로 저장하면 짜 둔 코스 현장이 모두 지워졌음
    (11-24 조용한 실패 점검 1번)."""
    import dashboard.auth as auth
    from flask import Flask
    from dashboard.blueprints import gw_vehicles as bp
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ())
    monkeypatch.setattr(bp, '_leads', lambda: LEADS)
    monkeypatch.setattr(bp, '_plan_leads', lambda: LEADS)
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    plan = {'courses': [{'members': ['TH'], 'leads': ['L-3'], 'memo': ''}], 'online': [], 'online_shared': False}
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'p@example.com', 'permission_level': 'admin'}
        assert cl.post('/api/vehicles/plan', json={'day': DAY, 'plan': plan}).get_json()['success']

        def down():
            raise bp._LeadsUnavailable()
        monkeypatch.setattr(bp, '_plan_leads', down)
        monkeypatch.setattr(bp, '_leads', lambda: [])
        for r in (cl.post('/api/vehicles/plan', json={'day': DAY, 'plan': plan}),
                  cl.get(f'/api/vehicles/plan?day={DAY}'),
                  cl.post('/api/vehicles/plan/confirm', json={'day': DAY, 'plan': plan})):
            assert r.status_code == 503 and '리드 시트' in r.get_json()['message']
    saved = vp.load(DAY, LEADS)['plan']
    assert [c['leads'] for c in saved['courses']] == [['L-3']]          # 짜 둔 코스 그대로
