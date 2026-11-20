"""증명서 바로 발급 (2026-11-20) — 작은 PDF(글꼴 줄이기·PNG)·증명서 칸·발급 기록·슬랙 [📄 바로 발급]·그룹웨어 API."""
import os
import struct
import zlib

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import certificates as ct
from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import pdf_mini as pm

FONT = r'C:\Windows\Fonts\malgun.ttf'
need_font = pytest.mark.skipif(not os.path.isfile(FONT), reason='맑은 고딕 없음')
ADMIN = 'hr-admin@example.com'


def _png(w=4, h=3, rgba=True) -> bytes:
    rows = b''
    for y in range(h):
        rows += b'\x00' + bytes((255, 0, 0, 128) if rgba else (0, 255, 0)) * w
    def chunk(t, d):
        return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xFFFFFFFF)
    ihdr = struct.pack('>IIBBBBB', w, h, 8, 6 if rgba else 2, 0, 0, 0)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', ihdr) + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b'')


def test_png_decode_rgba_and_rgb():
    w, h, rgb, a = pm.decode_png(_png())
    assert (w, h) == (4, 3) and rgb[:3] == b'\xff\x00\x00' and len(rgb) == 36 and set(a) == {128}
    w, h, rgb, a = pm.decode_png(_png(rgba=False))
    assert a is None and rgb[:3] == b'\x00\xff\x00'
    with pytest.raises(ValueError):
        pm.decode_png(b'not png')


@need_font
def test_font_subset_keeps_only_used_glyphs():
    f = pm.TTF.load(FONT)
    g = f.gid('재')
    assert g and f.width('재직', 10) > 0
    sub = pm.TTF(f.subset({g}))
    assert sub.num_glyphs == f.num_glyphs                      # 번호는 그대로
    assert sub.loca[g + 1] - sub.loca[g] > 0                    # 쓴 글자는 남음
    other = f.gid('뷁')
    assert sub.loca[other + 1] - sub.loca[other] == 0           # 안 쓴 글자는 비움
    assert len(f.subset({g})) < len(f.data) / 5


@need_font
def test_doc_output_is_pdf_with_text_and_image():
    d = pm.Doc(title='시험')
    d.font('R', pm.TTF.load(FONT))
    d.text(40, 40, '가나 ABC', 12, 'R', align='center', spacing=3)
    d.text(40, 60, '다음 줄', 12, 'R')
    d.image(_png(), 10, 10, 20, 20)
    out = d.output()
    assert out.startswith(b'%PDF-1.7') and out.rstrip().endswith(b'%%EOF')
    assert b'/Identity-H' in out and b'/SMask' in out and b'/ToUnicode' in out
    assert ' 3.00 Tc' in d.ops[0] and ' 0.00 Tc' in d.ops[1]     # 간격은 그 줄에만(그래픽 상태가 번지지 않게)


def test_cert_rows_employed_vs_career(monkeypatch):
    monkeypatch.delenv('CERT_GROUP_NAME', raising=False)
    p = {'name': '박민우', 'birth_date': '1990-03-05', 'address': '김포', 'hire_date': '2023-07-24', 'status': '재직',
         'department': '기술영업', 'job_rank': '매니저', 'job_role': '영업', 'group_hire_date': '2022-01-01'}
    rows = dict(ct.cert_rows(p, '재직증명서', '금융기관 제출'))
    assert dict(rows['기본 정보'])['생년월일'] == '1990년 3월 5일' and '주민등록번호' not in dict(rows['기본 정보'])
    assert dict(rows['재직 정보'])['재직 기간'] == '2023년 7월 24일 ~ 현재'
    assert dict(rows['회사 정보'])['회사명'] == '(주)아이티글로벌그룹'          # 기본=그룹(플렉스 발급본)
    itg = dict(ct.cert_rows(p, '재직증명서', 'x', company_key='itg'))
    assert dict(itg['회사 정보'])['회사명'] == '(주)아이티글로벌' and dict(itg['회사 정보'])['사업자등록번호'] == '884-81-01425'
    g = dict(ct.cert_rows(p, '재직증명서', 'x', group_date=True))
    assert dict(g['재직 정보'])['재직 기간'].startswith('2022년 1월 1일')
    c = dict(ct.cert_rows(dict(p, status='퇴사', resign_date='2026-10-26'), '경력증명서', 'x', ssn_full='900305-1234567'))
    assert dict(c['경력 정보'])['근무 기간'] == '2023년 7월 24일 ~ 2026년 10월 26일'
    assert dict(c['기본 정보'])['주민등록번호'] == '900305-1234567'
    monkeypatch.setenv('CERT_GROUP_NAME', '(주)새이름')
    assert ct.company('group')['name'] == '(주)새이름' and ct.company_key('(주)새이름') == 'group'
    with pytest.raises(ValueError):
        ct.company('')


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, ct):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('GW_HR_DOCS_DIR', str(tmp_path / 'docs'))
    (tmp_path / 'seal.png').write_bytes(_png(8, 8))            # 그룹 직인(플렉스에서 받은 처음 이름)
    monkeypatch.setenv('CERT_SEAL_DIR', str(tmp_path))
    monkeypatch.setenv('EAPPROVAL_EXECUTOR_DM_IDS', 'UHANDLER')
    monkeypatch.setattr(hr, 'hr_admins', lambda: [ADMIN])
    gw._invalidate()
    yield d
    gw._invalidate()


def _person(**kw):
    rec = {'flex_id': 'F1', 'name': '박민우', 'status': '재직', 'hire_date': '2022-02-03', 'department': '영업',
           'work_email': 'staff@example.com', 'login_email': 'staff@example.com', 'access': 'staff',
           'ssn': '900101-1234567', 'slack_id': 'UREQ'}
    rec.update(kw)
    hr.upsert_from_flex(rec, [])
    return hr.find_by_email('staff@example.com')


@need_font
def test_issue_pdf_records_and_saves_document(db):
    p = _person()
    res = ct.issue_pdf(p['id'], '재직증명서', '금융기관 제출', ADMIN, ssn=True, company_key='group')
    assert res['pdf'].startswith(b'%PDF') and res['doc_no'].endswith('-0001')
    docs = hr.documents(p['id'])
    assert [d['doc_type'] for d in docs] == ['발급 증명서'] and docs[0]['file_name'] == res['file_name']
    assert ct.issues(p['id'])[0]['ssn_included'] == 1
    assert any(l['action'] == 'cert_issue' for l in hr.access_log(p['id']))
    assert ct.issues(p['id'])[0]['company'] == 'group' and '아이티글로벌그룹' in res['file_name']
    assert ct.issue_pdf(p['id'], '경력증명서', 'x', ADMIN, company_key='group')['doc_no'].endswith('-0002')


@need_font
def test_issue_pdf_guards(db, monkeypatch):
    p = _person()
    with pytest.raises(ValueError):
        ct.issue_pdf(p['id'], '이상한 증명서', 'x', ADMIN, company_key='group')
    with hr._conn() as conn:
        conn.execute("UPDATE hr_people SET status='퇴사', resign_date='2026-10-01', ssn_enc=NULL WHERE id=?", (p['id'],))
    with pytest.raises(ValueError, match='경력증명서'):
        ct.issue_pdf(p['id'], '재직증명서', 'x', ADMIN, company_key='group')
    with pytest.raises(ValueError, match='주민등록번호'):
        ct.issue_pdf(p['id'], '경력증명서', 'x', ADMIN, ssn=True, company_key='group')
    with pytest.raises(ValueError, match='발급 회사'):
        ct.issue_pdf(p['id'], '경력증명서', 'x', ADMIN)
    with pytest.raises(RuntimeError, match='아이티글로벌 직인'):          # 아이티글로벌 직인 그림은 아직 없음
        ct.issue_pdf(p['id'], '경력증명서', 'x', ADMIN, company_key='itg')


class _Client:
    def __init__(self):
        self.calls = []

    def conversations_open(self, users):
        return {'channel': {'id': 'D' + users}}

    def files_upload_v2(self, **kw):
        self.calls.append(('upload', kw['channel'], [u['filename'] for u in kw['file_uploads']]))

    def chat_update(self, **kw):
        self.calls.append(('update', kw['channel']))

    def chat_postMessage(self, **kw):
        self.calls.append(('post', kw['channel'], kw.get('text')))
        return {'ts': '1'}


@need_font
def test_slack_auto_issue(db, monkeypatch):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '황샛별')
    monkeypatch.setattr(se, '_notify', lambda c, uid, text: c.calls.append(('notify', uid, text)))
    p = _person()
    vals = {'ctype': '재직증명서', 'company': 'group', 'purpose': '금융기관 제출', 'dest': '○○은행', 'ssn': ct.SSN_OPTS[0],
            'receive': ct.RECEIVE[1], 'copies': '2', 'need_by': '', 'memo': ''}
    row = ct.create('UREQ', '박민우', vals)
    c = _Client()
    assert sc.auto_issue(c, row['id'], 'UHANDLER') == ''
    ups = [x for x in c.calls if x[0] == 'upload']
    assert ups[0][1] == 'DUREQ' and ups[1][1] == 'DUHANDLER'          # 원본 신청 → 경영지원에게도
    assert ct.get(row['id'])['status'] == ct.ISSUED
    assert ct.issues(p['id'])[0]['purpose'] == '금융기관 제출 (○○은행)' and ct.issues(p['id'])[0]['request_id'] == row['id']
    assert sc.auto_issue(c, row['id'], 'UHANDLER') == '이미 처리된 신청이에요.'
    old = ct.create('UREQ', '박민우', dict(vals, company=''))              # 회사 고르기 전 신청
    assert '발급 회사' in sc.auto_issue(c, old['id'], 'UHANDLER') and ct.get(old['id'])['status'] == ct.PENDING
    # 인사카드에 슬랙 연결이 없으면 바로 발급 대신 안내
    row2 = ct.create('UNOCARD', '누구', dict(vals, receive=ct.RECEIVE[0]))
    assert '슬랙 연결' in sc.auto_issue(c, row2['id'], 'UHANDLER')
    assert ct.get(row2['id'])['status'] == ct.PENDING


@need_font
def test_gw_certificate_api(db, monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import hr as hr_bp_mod
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    p = _person()
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(hr_bp_mod.hr_bp)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ADMIN, 'permission_level': 'admin'}
        r = cl.post(f"/api/hr/people/{p['id']}/certificate", json={'company': 'group', 'ctype': '재직증명서', 'purpose': '관공서 제출'})
        assert r.status_code == 200 and r.data.startswith(b'%PDF') and r.headers['X-Doc-No'].endswith('-0001')
        r = cl.post(f"/api/hr/people/{p['id']}/certificate", json={'company': 'group', 'ctype': '재직증명서', 'purpose': 'x', 'ssn': True,
                                                                     'group_date': True})
        assert r.status_code == 200
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        assert cl.post(f"/api/hr/people/{p['id']}/certificate", json={'ctype': '재직증명서'}).status_code == 403


def test_transparent_pixels_take_ink_color():
    """투명 칸이 검정이면 보기 프로그램이 줄일 때 도장 가장자리가 어두워진다(11-20) — 진한 칸 평균색으로 바꿔 넣는다."""
    rows = b''.join(b'\x00' + bytes((200, 30, 40, 255, 0, 0, 0, 0)) for _ in range(2))
    def chunk(t, d):
        return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xFFFFFFFF)
    png = (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 6, 0, 0, 0))
           + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b''))
    w, h, rgb, a = pm._unbleed(pm.decode_png(png))
    assert rgb[3:6] == bytes((200, 30, 40)) and a[1] == 0
