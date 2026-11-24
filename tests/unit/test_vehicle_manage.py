"""차량 관리(2026-11-22) — 결재에서 차량 비용 모으기(법인카드 메모·개인경비·자금집행·비품·전자결재),
종류 나누기, 여러 대 나눔, 중복 없이 넣기, 만기 D-day, 직접 입력, 계기판 km, 화면 권한. 이름은 모두 가짜."""
import json
from datetime import date

import pytest
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_manage as vm

ADMIN = 'hr-admin@example.com'

CARD_MEMO = ('작성요령\n법인카드 지출 결의서 사용기간2026월 10월 작성자/직급김가람 / 매니저소속부서영업부사용카드1234   '
             '내            역   no.사용일자가맹점명사용금액사용내역(상세기재)'
             '12026.10.02써브웨이 김포점 ₩           11,400YM 점심'
             '22026.10.06SK쉴더스_주차료 ₩            3,0001783 주차비'
             '32026.10.08(주)빅쓰리 ₩            8,0001783 주차비'
             '426.10.09걸포주유소 ₩           50,00060거 5484 모닝 주유비'
             '합  계')


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, va, vm, hr):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(vm, '_last_sync', 0.0)
    gw._invalidate()
    with d._get_connection() as conn:
        conn.execute('CREATE TABLE IF NOT EXISTS flex_approval_archive (doc_no TEXT, title TEXT, form TEXT, status TEXT, '
                     'writer TEXT, created_at TEXT, updated_at TEXT, attachments TEXT, memo TEXT, fields TEXT, amount INTEGER, '
                     'project_code TEXT, partner TEXT, imported_at TEXT)')
        conn.execute('CREATE TABLE IF NOT EXISTS approvals (id INTEGER PRIMARY KEY, form TEXT, status TEXT, content TEXT, '
                     'amount INTEGER, partner TEXT, created_at TEXT)')
        rows = [
            ('2026-1', '법인카드 지출결의', '완료', '2026-11-01 09:00', CARD_MEMO, json.dumps({'이름': '김가람'}), 0),
            ('2026-2', '개인경비청구서', '완료', '2026-11-05 02:06', '',
             json.dumps({'사용일자': '2026-11-04', '청구 금액': '40000', '사유': '트럭 , 6873 주차비'}), 40000),
            ('2026-3', '자금집행요청', '완료', '2026-10-14 04:52', '',
             json.dumps({'내용': '8021 콤프레샤 교체 및 4904 전조등 교체비용', '상호명': '가나정비', '총 금액(VAT포함)': '419000'}), 419000),
            ('2026-4', '비품 신청', '완료', '2026-06-18 08:53', '',
             json.dumps({'품명': '워셔액 2L', '금액(VAT포함)': '22480', '사유': '8988 법인차량 워셔액 구매'}), 22480),
            ('2026-5', '자금집행요청', '반려', '2026-10-15 04:52', '',
             json.dumps({'내용': '4041 타이어', '총 금액(VAT포함)': '100000'}), 100000),            # 반려 — 안 넣음
            ('2026-6', '자금집행요청', '완료', '2026-10-16 04:52', '',
             json.dumps({'내용': '사무실 에어컨 18,021원 부품', '총 금액(VAT포함)': '18021'}), 18021),  # 차 번호 아님
        ]
        conn.executemany('INSERT INTO flex_approval_archive (doc_no, form, status, created_at, memo, fields, amount) '
                         'VALUES (?,?,?,?,?,?,?)', rows)
        conn.executemany('INSERT INTO approvals (id, form, status, content, amount, partner, created_at) VALUES (?,?,?,?,?,?,?)', [
            (20, 'fund', 'done', '카니발 엔진오일, 와이퍼 교체', 210000, '가나정비', '2026-11-11 16:43:25'),
            (21, 'fund', 'cancelled', '5358 유리', 55000, '', '2026-11-12 10:00:00'),
        ])
        conn.commit()
    yield d
    gw._invalidate()


def test_find_plates_and_kinds(db):
    fleet = va.fleet(True)
    assert vm.find_plates('817루 4904 타이어', fleet) == ['817루4904']
    assert vm.find_plates('카니발 와이퍼', fleet) == ['143고9687']
    assert vm.find_plates('에어컨 18,021원 · 카드 4966', fleet) == []
    assert vm.kind_of('1783 주차비') == '주차'
    assert vm.kind_of('60거5484 현장미팅중 주정차단속') == '과태료'
    assert vm.kind_of('9131 전기트럭 주유구 구매') == '정비'          # '주유'가 들어 있어도 부품
    assert vm.kind_of('5484 주유비') == '유류'
    assert vm.kind_of('노후 교체 카니발 와이퍼') == '소모품'
    assert vm.kind_of('포터 817루4904', vendor='가나모터스') == '정비'


def test_parse_card_memo(db):
    rows = vm.parse_card_memo(CARD_MEMO, va.fleet(True))
    assert [(r['day'], r['amount'], r['plates']) for r in rows] == [
        ('2026-10-06', 3000, ['53우1783']), ('2026-10-08', 8000, ['53우1783']), ('2026-10-09', 50000, ['60거5484'])]
    assert rows[0]['shop'] == 'SK쉴더스_주차료'


def test_sync_costs_once_with_split(db):
    assert vm.sync_costs(force=True) == 8   # 카드 3 + 개인경비 1 + 자금집행 2대 + 비품 1 + 전자결재 1
    assert vm.sync_costs(force=True) == 0   # 같은 결재 줄은 한 번만
    c = vm.costs(plate='03모8021')
    assert c[0]['amount'] == 209500 and c[0]['shared'] == 2 and c[0]['kind'] == '정비' and c[0]['source'] == 'approval'
    assert vm.costs(plate='204두8988')[0]['kind'] == '소모품'
    assert vm.costs(plate='143고9687')[0]['amount'] == 210000
    assert not vm.costs(plate='16소4041') and not vm.costs(plate='237허5358')   # 반려·취소 결재는 안 넣음
    with pytest.raises(ValueError):
        vm.delete_cost(c[0]['id'])                                               # 결재에서 온 줄은 못 지움


def test_overview_due_and_totals(db):
    vm.sync_costs(force=True)
    vm.update_info('16소4041', {'inspection_end': '2026-12-10'}, ADMIN)
    ov = vm.overview(today=date(2026, 11, 22))
    cars = {c['plate']: c for c in ov['cars']}
    assert cars['231허6854']['insurance_d'] == -31 and cars['231허6854']['warn']      # 렌트 공제 증명서 지남
    assert cars['16소4041']['inspection_d'] == 18 and cars['16소4041']['warn']
    assert not cars['03모8021']['warn'] and cars['03모8021']['insurer'] == '현대해상'
    assert ov['cars'][0]['warn']                                                       # 만기 임박이 맨 위
    assert ov['month_total'] == 40000 + 210000 and ov['alerts'] == 3
    assert cars['53우1783']['cost_year'] == 11000 and cars['817루4904']['last_fix_day'] == '2026-10-14'


def test_manual_cost_and_odometer(db):
    r = vm.add_cost('204두8988', '2026-11-20', '세차', '손세차', '30,000', '', ADMIN)
    assert vm.costs(plate='204두8988')[0]['amount'] == 30000
    vm.delete_cost(r['id'])
    assert not vm.costs(plate='204두8988')
    with pytest.raises(ValueError):
        vm.add_cost('204두8988', '2026-11-20', '없는종류', 'x', 1, '', ADMIN)
    vm.add_odometer('204두8988', '2026-10-31', '12,000', ADMIN)
    res = vm.add_odometer('204두8988', '2026-11-30', 12850.5, ADMIN)
    assert res['readings'][0]['driven'] == 850.5
    with pytest.raises(ValueError):
        vm.add_odometer('204두8988', '2026-12-31', 11000, ADMIN)                     # 앞 기록보다 작음


def test_routes_hr_admin_only(db, monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import gw_vehicles as bp
    from dashboard.blueprints import hr as hr_bp
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(gw, 'dev_preview_ok', lambda e: True)
    monkeypatch.setattr(hr, 'hr_admins', lambda: [ADMIN])
    monkeypatch.setattr(bp, 'render_template', lambda name: name)
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(bp.gw_vehicles_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'admin'}
        assert cl.get('/api/vehicles/manage').status_code == 403
        assert cl.get('/vehicles/manage').status_code == 302
        with cl.session_transaction() as s:
            s['user'] = {'email': ADMIN, 'permission_level': 'user'}
        j = cl.get('/api/vehicles/manage').get_json()
        assert j['success'] and len(j['cars']) == 14 and '정비' in j['kinds']
        car = cl.get('/api/vehicles/manage/car?plate=817루4904').get_json()
        assert car['info']['insurer'] == 'KB손해보험' and car['costs']
        assert cl.post('/api/vehicles/manage/info/817루4904', json={'inspection_end': '2027-04-13'}).get_json()['info']['inspection_end'] == '2027-04-13'
        assert cl.post('/api/vehicles/manage/info/817루4904', json={'inspection_end': '내일'}).status_code == 400
        add = cl.post('/api/vehicles/manage/costs', json={'plate': '817루4904', 'day': '2026-11-21', 'kind': '세차',
                                                           'item': '세차', 'amount': 10000}).get_json()
        assert cl.post(f"/api/vehicles/manage/costs/{add['id']}/delete", json={}).get_json()['success']
        assert cl.post('/api/vehicles/manage/odometer', json={'plate': '204두8988', 'day': '2026-11-21', 'km': 100}).get_json()['success']


def test_menu_three_items():
    area = next(a for a in gw.AREAS if a['key'] == 'vehicle')
    assert [i['label'] for i in area['items']] == ['실시간 위치', '운행일지', '차량 관리']   # 배정은 방문 일정 탭으로만, 위치는 전 직원(11-24)
    assert area['items'][0]['field'] and area['items'][2]['hr_admin']
