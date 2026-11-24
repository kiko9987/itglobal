"""차량 운행(2026-11-20) — U+커넥트 운행일지 읽기·중복 건너뛰기·지역 맞추기·추정 운전자·화면 API."""
import io
import zipfile

import openpyxl
import pytest
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import vehicle_log as vl

ADMIN = 'hr-admin@example.com'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    from dashboard.services import vehicle_assign as va
    for mod in (gw, vl, va, hr):      # 운전자 표시가 차량 배정(va)·인사(hr)를 읽으므로 같이 격리
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(va, 'staff', lambda: [])
    monkeypatch.setattr(gw, 'DEV_PREVIEW_USERS', ())   # 개발 중 문은 test_vehicle_assign 에서 따로 본다
    gw._invalidate()
    yield d
    gw._invalidate()


def _xlsx(plate='12가3456', model='레이 (TAM)', trips=()):
    """U+커넥트 운행일지(국세청) 양식 그대로 — 머리 9줄 + 운행 + 합계 2줄."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['과세기간', '2026-10-21 ~ 2026-11-20', '업무승용차 운행기록부', '상호명', '아이티글로벌'])
    ws.append([None, None, None, '사업자등록번호', '8848101425'])
    ws.append([])
    ws.append(['차종', None, '자동차등록번호'])
    ws.append([model, None, plate])
    ws.append([])
    ws.append(['사용일자(요일)', '사용자', None, '운행내역'])
    ws.append([None, '부서', '성명', '주행전 계기판의 거리(km)', '주행후 계기판의 거리(km)', '주행거리(km)', '출발지', '도착지',
               '출발시간', '도착시간', '업무용 사용거리(km)', None, '비 고'])
    ws.append([None] * 10 + ['출퇴근용(km)', '일반 업무용(km)', None])
    for day, start, frm, to, dist in trips:
        ws.append([day, '기본 사업장', '-', '1,000.0', '1,010.0', str(dist), frm, to, start, '23:59:00', '0', str(dist), ''])
    ws.append(['', '', '', '과세기간 중 주행거리(km)', None, None, None, None, '', None, '과세기간 업무용 사용 거리(km)', None, '업무용 사용 비율(%)'])
    ws.append(['', '', '', '2,261.3', None, None, None, None, '', None, '2,261.3', None, '100.0'])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _zip(files):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as z:
        for name, data in files:
            info = zipfile.ZipInfo(name.encode('cp949').decode('cp437'))   # 윈도 한글 이름(UTF-8 표시 없음)
            z.writestr(info, data)
    return buf.getvalue()


T1 = [('2026-11-20(금)', '09:10:00', '경기도 김포시 걸포동', '경기도 용인시 수지구 풍덕천동', 40.0),
      ('2026-11-20(금)', '13:00:00', '경기도 용인시 수지구 풍덕천동', '경기도 김포시 걸포동', 41.5),
      ('2026-11-19(목)', '18:00:00', '경기도 김포시 걸포동', '경기도 김포시 고촌읍 향산리', 8.0)]


def test_parse_sheet_reads_trips_and_skips_head_and_totals():
    s = vl.parse_upload('a.xlsx', _xlsx(trips=T1))[0]
    assert (s['plate'], s['model'], s['period']) == ('12가3456', '레이 (TAM)', '2026-10-21 ~ 2026-11-20')
    assert [(t['day'], t['start_time'], t['km']) for t in s['trips']] == [
        ('2026-11-20', '09:10:00', 40.0), ('2026-11-20', '13:00:00', 41.5), ('2026-11-19', '18:00:00', 8.0)]
    assert s['trips'][0]['odo_start'] == 1000.0 and s['trips'][0]['to_addr'] == '경기도 용인시 수지구 풍덕천동'


def test_not_a_log_is_refused():
    wb = openpyxl.Workbook()
    wb.active.append(['이름', '금액'])
    buf = io.BytesIO()
    wb.save(buf)
    with pytest.raises(ValueError, match='운행일지'):
        vl.parse_upload('x.xlsx', buf.getvalue())
    with pytest.raises(ValueError, match='운행일지'):
        vl.parse_upload('x.csv', b'a,b')
    with pytest.raises(ValueError, match='xlsx'):
        vl.parse_upload('x.zip', _zip([('메모.txt', b'x')]))


def test_zip_upload_and_duplicates_skipped(db):
    data = _zip([('LGUPlus_FMS_System_운행일지_상세_12가3456.xlsx', _xlsx(trips=T1)),
                 ('LGUPlus_FMS_System_운행일지_상세_34나5678.xlsx', _xlsx('34나5678', '포터2 (HR)', T1[:1]))])
    assert vl.import_upload('운행일지.zip', data, ADMIN) == {'vehicles': 2, 'added': 4, 'skipped': 0,
                                                             'period': '2026-10-21 ~ 2026-11-20'}
    again = vl.import_upload('운행일지.zip', data, ADMIN)
    assert (again['added'], again['skipped']) == (0, 4)          # 같은 운행(차량·날짜·출발 시각)은 한 번만
    assert vl.months() == ['2026-11'] and vl.trip_range() == ('2026-11-19', '2026-11-20')
    assert vl.last_upload()['skipped'] == 4


@pytest.mark.parametrize('addr,want', [
    ('경기도 용인시 수지구 풍덕천동', ('용인', '수지구', '풍덕천동')),
    ('용인 수지구 심곡로 11', ('용인', '수지구', '')),
    ('마포구 양화로 45', ('서울', '마포구', '')),                    # 축약 주소는 '서울'이 빠져 있다
    ('서울특별시 마포구 서교동', ('서울', '마포구', '서교동')),
    ('인천광역시 검단구 원당동', ('인천', '검단구', '원당동')),
    ('인천 부평구 부평대로63번길 24-1', ('인천', '부평구', '')),
    ('경기도 김포시 고촌읍 향산리', ('김포', '', '고촌읍')),
    ('구리 경춘로 1', ('구리', '', '')),
    ('걸포로 134-84', ('', '', '')),
    ('10펼', ('', '', '')),
])
def test_area(addr, want):
    assert vl.area(addr) == want


def test_same_area_rules():
    trip = vl.area('경기도 부천시 오정구 고강동')
    assert vl.same_area(trip, vl.area('부천 길주로 1'))                 # 방문지에 구가 없으면 시만
    assert not vl.same_area(trip, vl.area('부천 원미구 길주로 1'))
    assert vl.same_area(trip, ('부천', '오정구', '고강동'))
    assert not vl.same_area(trip, ('부천', '오정구', '원종동'))          # 방문지 동을 알면 동까지
    # 회사가 있는 김포는 방문지 동을 모르면 맞추지 않는다(출퇴근과 늘 겹침)
    assert not vl.same_area(vl.area('경기도 김포시 고촌읍 향산리'), vl.area('김포 김포한강4로 1'))
    assert vl.same_area(vl.area('경기도 김포시 양촌읍 양곡리'), ('김포', '', '양촌읍'))
    assert vl.is_home(vl.area('경기도 김포시 걸포동'))


def _lead(no, day, addr, who, status='방문 완료'):
    return {'리드 No': no, '방문 예정일': day, '방문 주소': addr, '영업 담당자': who, '상태': status}


LEADS = [_lead('L-1', '2026-11-20', '용인 수지구 심곡로 11', '김가람,이나래'),
         _lead('L-2', '2026-11-20', '부천 길주로 1', '박다솜'),
         _lead('L-3', '2026-11-19', '김포 김포한강4로 1', '최마루'),
         _lead('L-4', '2026-11-20', '용인 수지구 포은대로 1', '정바다', status='유선 상담'),   # 가지 않은 리드
         _lead('L-5', '2026-11-20', '용인 수지구 포은대로 2', '')]                           # 담당자 없음


def test_summary_guesses_driver_from_same_area_visits(db, monkeypatch):
    vl.import_upload('a.xlsx', _xlsx(trips=T1), ADMIN)
    s = vl.month_summary('2026-11', LEADS)
    car = s['vehicles'][0]
    assert (car['plate'], car['km'], car['trips'], car['days']) == ('12가3456', 89.5, 3, 2)
    assert car['drivers'] == ['김가람', '이나래'] and car['visits'] == 1
    assert s['pending_dongs'] == 3                 # 동을 아직 찾지 않은 방문 주소(L-1·L-2·L-3)
    d = vl.vehicle_days('2026-11', '12가3456', LEADS)
    first = d['days'][0]
    assert first['day'] == '2026-11-20' and first['drivers'] == ['김가람', '이나래']
    assert [t['home'] for t in first['trips']] == [False, True]
    assert first['trips'][0]['visits'][0]['lead_no'] == 'L-1'
    assert d['days'][1]['drivers'] == []           # 김포 고촌읍 — 방문지 동을 모르면 맞추지 않음

    # 행안부로 동을 찾으면 동까지 맞춘다 — L-3(김포)이 고촌읍이면 19일이 맞춰지고, L-1 이 다른 동이면 20일은 빠진다
    found = {'김포 김포한강4로 1': '고촌읍', '용인 수지구 심곡로 11': '동천동'}
    monkeypatch.setattr(vl, '_lookup_dong', lambda a: found.get(a, ''))
    assert vl.resolve_dongs(vl.unresolved(LEADS, '2026-11-01', '2026-11-30'), pause=0) == 3
    assert vl.unresolved(LEADS, '2026-11-01', '2026-11-30') == []
    d = vl.vehicle_days('2026-11', '12가3456', LEADS)
    assert d['days'][0]['drivers'] == [] and d['days'][1]['drivers'] == ['최마루']


def test_placeholder_manager_is_not_a_driver():
    """영업 담당자 칸의 '-'·'—' 는 사람이 아니다(운전자 칸에 "-, 박민재"로 나오던 것, 11-21)."""
    leads = [_lead('L-9', '2026-11-20', '용인 수지구 심곡로 11', '-'),
             _lead('L-10', '2026-11-20', '용인 수지구 심곡로 12', '— , 김가람'),
             _lead('L-11', '2026-11-20', '용인 수지구 심곡로 13', '김가람/-')]
    v = vl.visits_by_day(leads, '2026-11-01', '2026-11-30')['2026-11-20']
    assert [x['who'] for x in v] == [['김가람'], ['김가람']]     # '-' 만 있는 L-9 는 방문에서 빠짐


def test_range_visit_counts_every_day():
    """'2026-11-24~26' 같은 기간 방문은 걸친 날마다 센다(11-24 — 첫날에만 세어 25일 코스에서 빠지던 것)."""
    leads = [_lead('ETC-1', '2026-11-24~26', '영등포구 영중로8길 14', '한바람', status='방문 예정'),
             _lead('ETC-2', '2026-11-29~12-02', '부천 길주로 1', '오가온', status='방문 예정'),   # 달을 넘는 기간
             _lead('ETC-3', '2026-11-24~20', '부천 길주로 2', '유하늘', status='방문 예정'),     # 끝이 앞 → 첫날만
             _lead('ETC-4', '2026-01-02~12-30', '부천 길주로 3', '진새봄', status='방문 예정')]  # 너무 긴 기간 → 첫날만
    v = vl.visits_by_day(leads, '2026-11-01', '2026-11-30')
    assert [d for d in sorted(v) if any(x['lead_no'] == 'ETC-1' for x in v[d])] == ['2026-11-24', '2026-11-25', '2026-11-26']
    assert [d for d in sorted(v) if any(x['lead_no'] == 'ETC-2' for x in v[d])] == ['2026-11-29', '2026-11-30']   # 기간 밖은 자름
    assert [d for d in sorted(v) if any(x['lead_no'] == 'ETC-3' for x in v[d])] == ['2026-11-24']
    assert vl._visit_days('2026-01-02~12-30') == ['2026-01-02']
    assert vl._visit_days('2026-11-25') == ['2026-11-25'] and vl._visit_days('') == []


def test_lookup_dong_uses_juso_jibun(monkeypatch):
    from dashboard.services import address_resolver as ar
    seen = []
    def fake(q):
        seen.append(q)
        return (('경기도 용인시 수지구 심곡로 11 (동천동)', '경기도 용인시 수지구 동천동 1-1', ''),)
    monkeypatch.setattr(ar, '_juso_search_cached', fake)
    assert vl._lookup_dong('용인 수지구 심곡로 11 2층 OO학원') == '동천동'
    assert seen == ['용인 수지구 심곡로 11']        # 건물·층은 떼고 도로명+번호로
    def down(q):
        raise ar._JusoTransientError()
    monkeypatch.setattr(ar, '_juso_search_cached', down)
    assert vl._lookup_dong('용인 수지구 심곡로 11') is None   # 잠시 안 되면 저장하지 않고 다음에


def test_routes_hr_admin_only(db, monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.blueprints.groupware import field_path_allowed
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(hr, 'hr_admins', lambda: [ADMIN])
    monkeypatch.setattr(bp, '_leads', lambda: LEADS)
    monkeypatch.setattr(bp, '_resolve_later', lambda s, e: None)
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'admin'}
        assert cl.get('/api/vehicles').status_code == 403
        assert cl.post('/api/vehicles/upload', data={'file': (io.BytesIO(_xlsx(trips=T1)), 'a.xlsx')}).status_code == 403
        assert cl.get('/vehicles').status_code == 302
        with cl.session_transaction() as s:
            s['user'] = {'email': ADMIN, 'permission_level': 'user'}
        r = cl.post('/api/vehicles/upload', data={'file': (io.BytesIO(_xlsx(trips=T1)), 'a.xlsx')})
        assert r.get_json()['added'] == 3
        bad = cl.post('/api/vehicles/upload', data={'file': (io.BytesIO(b'x'), 'a.txt')})
        assert bad.status_code == 400 and '운행일지' in bad.get_json()['message']
        j = cl.get('/api/vehicles?month=2026-11').get_json()
        assert j['success'] and j['vehicles'][0]['drivers'] == ['김가람', '이나래'] and j['months'] == ['2026-11']
        assert cl.get('/api/vehicles?month=bogus').get_json()['month']           # 이상한 월은 이번 달로
        car = cl.get('/api/vehicles/car?month=2026-11&plate=12가3456').get_json()
        assert car['success'] and len(car['days']) == 2
        assert cl.get('/api/vehicles/car?month=2026-11').status_code == 400
    assert not field_path_allowed('GET', '/api/vehicles') and not field_path_allowed('POST', '/api/vehicles/upload')


def test_menu_only_for_hr_admin():   # 11-20 '차량' 영역으로 옮김(배정과 함께)
    items = [it for a in gw.AREAS for it in a['items'] if it['href'] == '/vehicles']
    assert items and items[0].get('hr_admin')
