# -*- coding: utf-8 -*-
"""그룹웨어 공지사항·개발 노트(11-25). 슬랙은 가짜 클라이언트(운영 슬랙으로 아무것도 안 나감)."""
import sys

sys.path.insert(0, '.')

import pytest
from cryptography.fernet import Fernet

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import notices as ns

ADMIN = 'kiko@itg-aircon.com'


@pytest.fixture
def env(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, ns):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('GW_NOTICE_FILES_DIR', str(tmp_path / 'files'))
    gw._invalidate()
    people = {}
    for name, dept, access in (('사무가', '기술영업', 'staff'), ('사무나', '경영지원', 'staff'),
                               ('현장가', '기술시공', 'field'), ('미화가', '미화', 'field')):
        people[name] = hr.create_person({'name': name, 'status': '재직', 'hire_date': '2022-05-23', 'department': dept,
                                         'access': access, 'work_email': f'{len(people)}x@itg-aircon.com'}, ADMIN)
    yield {n: hr.get_person(pid if isinstance(pid, int) else pid['id']) for n, pid in people.items()}
    gw._invalidate()


def V(p, admin=False):
    return ns.viewer(p, admin=admin)


def _cat(name):
    return next(c for c in ns.categories() if c['name'] == name)


AUTHOR = {'name': '고광일', 'email': ADMIN, 'key': 'eadmin'}


def _post(cat, title, kind=ns.NOTICE, **kw):
    data = dict({'title': title, 'body': '본문', 'category_id': _cat(cat)['id'] if cat else None}, **kw)
    return ns.save(kind, data, AUTHOR, True)['id']


def test_default_categories_and_audience(env):
    # 전사 / 영업 / 시공(11-25 부장님) — 경영지원만 보는 공지는 없음
    assert [c['name'] for c in ns.categories()] == ['전사 공지', '영업 공지', '시공 공지']
    a, b, f, m = env['사무가'], env['사무나'], env['현장가'], env['미화가']
    _post('전사 공지', '급여일 변경')
    _post('영업 공지', '매출계약서 양식')
    _post('시공 공지', '현장 안전')
    _post(None, '그룹웨어 업데이트', kind=ns.DEVNOTE)
    titles = lambda v, k=ns.NOTICE: sorted(n['title'] for n in ns.list_notices(k, v))
    assert titles(V(a)) == ['급여일 변경', '매출계약서 양식']          # 기술영업
    assert titles(V(b)) == ['급여일 변경']                           # 경영지원(관리자가 아니면 전사만)
    assert titles(V(f)) == ['급여일 변경', '현장 안전']               # 기술시공
    assert titles(V(m)) == ['급여일 변경']                           # 미화
    assert titles(V(b, admin=True)) == ['급여일 변경', '매출계약서 양식', '현장 안전']
    assert titles(V(a), ns.DEVNOTE) == ['그룹웨어 업데이트']
    assert titles(V(f), ns.DEVNOTE) == []                       # 개발 노트는 현장·미화 제외
    field_only = next(n['id'] for n in ns.list_notices(ns.NOTICE, V(a, True)) if n['title'] == '현장 안전')
    with pytest.raises(PermissionError):
        ns.get(field_only, V(a))                                   # 영업은 시공 공지를 못 봄


def test_only_admin_writes(env):
    with pytest.raises(PermissionError):
        ns.save(ns.NOTICE, {'title': 't', 'category_id': _cat('전사 공지')['id']}, AUTHOR, False)
    with pytest.raises(ns.NoticeError):
        ns.save(ns.NOTICE, {'title': '', 'category_id': _cat('전사 공지')['id']}, AUTHOR, True)
    with pytest.raises(ns.NoticeError):
        ns.save(ns.NOTICE, {'title': '카테고리 없음'}, AUTHOR, True)


def test_pinned_first_then_newest_and_unread(env):
    a = env['사무가']
    first = _post('전사 공지', '옛 공지')
    _post('전사 공지', '새 공지')
    _post('전사 공지', '고정 공지', pinned=True)
    items = ns.list_notices(ns.NOTICE, V(a))
    assert [n['title'] for n in items] == ['고정 공지', '새 공지', '옛 공지']
    assert ns.unread_count(ns.NOTICE, V(a)) == 3
    ns.get(first, V(a))
    assert ns.unread_count(ns.NOTICE, V(a)) == 2
    assert ns.get(first, V(a), mark_read=False)['read'] is True


def test_read_status_counts_only_audience(env):
    a, b, f = env['사무가'], env['사무나'], env['현장가']
    c = ns.save_category({'name': '내근 공지', 'audience': ns.AUD_OFFICE}, True)
    nid = ns.save(ns.NOTICE, {'title': '결재 방법', 'category_id': c['id']}, AUTHOR, True)['id']
    ns.get(nid, V(a))
    r = ns.read_status(nid)
    assert r['total'] == 2                                        # 현장·미화 제외 2명
    assert [p['name'] for p in r['read']] == ['사무가'] and [p['name'] for p in r['unread']] == ['사무나']
    all_id = _post('전사 공지', '차량 보험')
    ns.get(all_id, V(f))
    r = ns.read_status(all_id)
    assert r['total'] == 4 and [p['name'] for p in r['read']] == ['현장가']
    assert 'reads' in ns.get(all_id, V(b, admin=True)) and 'reads' not in ns.get(all_id, V(b))


def test_dept_category(env):
    a, b = env['사무가'], env['사무나']
    c = ns.save_category({'name': '영업 회의 공지', 'audience': ns.AUD_DEPT, 'depts': ['기술영업']}, True)
    ns.save(ns.NOTICE, {'title': '영업 회의', 'category_id': c['id']}, AUTHOR, True)
    assert [n['title'] for n in ns.list_notices(ns.NOTICE, V(a))] == ['영업 회의']
    assert ns.list_notices(ns.NOTICE, V(b)) == []
    with pytest.raises(ns.NoticeError):
        ns.save_category({'name': '빈 부서', 'audience': ns.AUD_DEPT, 'depts': []}, True)
    with pytest.raises(ns.NoticeError):
        ns.save_category({'name': '영업 회의 공지', 'audience': ns.AUD_ALL}, True)      # 같은 이름
    with pytest.raises(ns.NoticeError):
        ns.delete_category(c['id'], True)                                           # 글이 있으면 못 지움
    with pytest.raises(PermissionError):
        ns.save_category({'name': 'x'}, False)


def test_files_follow_notice_audience(env):
    a, f = env['사무가'], env['현장가']
    nid = _post('영업 공지', '양식')
    got = ns.add_file(nid, '계약서.pdf', b'%PDF-1', True)
    assert ns.file_for(got['id'], V(a))['name'] == '계약서.pdf'
    with pytest.raises(PermissionError):
        ns.file_for(got['id'], V(f))
    with pytest.raises(ns.NoticeError):
        ns.add_file(nid, 'run.exe', b'MZ', True)
    with pytest.raises(PermissionError):
        ns.add_file(nid, 'x.pdf', b'1', False)
    ns.remove_file(got['id'], True)
    assert ns.get(nid, V(a))['files'] == []


def test_delete_hides(env):
    a = env['사무가']
    nid = _post('전사 공지', '지울 글')
    ns.delete(nid, True, '고광일')
    assert ns.list_notices(ns.NOTICE, V(a)) == []
    with pytest.raises(LookupError):
        ns.get(nid, V(a))


class FakeSlack:
    def __init__(self):
        self.posts = []

    def chat_postMessage(self, channel, text, **kw):
        self.posts.append((channel, text))
        return {'ts': '1.2'}


def test_slack_notify_needs_channel(env):
    nid = _post('전사 공지', '급여일')
    fake = FakeSlack()
    with pytest.raises(ns.NoticeError):
        ns.notify(nid, True, client=fake)                          # 카테고리에 채널 없음
    c = _cat('전사 공지')
    ns.save_category({'name': c['name'], 'audience': c['audience'], 'slack_channel': 'C0ABCDEF12'}, True, c['id'])
    ns.notify(nid, True, client=fake)
    assert fake.posts[0][0] == 'C0ABCDEF12' and '[전사 공지] 급여일' in fake.posts[0][1] and f'/notices/{nid}' in fake.posts[0][1]
    assert ns.list_notices(ns.NOTICE, V(env['사무가']))[0]['slack_sent'] is True
    with pytest.raises(ns.NoticeError):
        ns.save_category({'name': '잘못', 'audience': ns.AUD_ALL, 'slack_channel': '#전체'}, True)


def test_body_kept_as_written(env):
    nid = ns.save(ns.NOTICE, {'title': '  제목  ', 'body': '첫 줄\r\n둘째 줄', 'category_id': _cat('전사 공지')['id']},
                  AUTHOR, True)['id']
    n = ns.get(nid, V(env['사무가']))
    assert n['title'] == '제목' and n['body'] == '첫 줄\n둘째 줄'
