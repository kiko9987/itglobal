"""입사·퇴사 할 일 (2026-11-20) — 인사카드 맞추기·기한·완료·슬랙 DM·그룹웨어 할 일·API."""
from datetime import date

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import hr_tasks as ht

ADMIN = 'kiko@itg-aircon.com'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, ht):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('EAPPROVAL_EXECUTOR_DM_IDS', 'UHANDLER')
    monkeypatch.delenv('HR_TASKS_FROM', raising=False)
    gw._invalidate()
    yield d
    gw._invalidate()


def _person(fid, name, **kw):
    rec = {'flex_id': fid, 'name': name, 'status': '재직', 'hire_date': '2026-12-01', 'department': '영업',
           'access': 'staff'}
    rec.update(kw)
    hr.upsert_from_flex(rec, [])
    with hr._conn() as conn:
        return dict(conn.execute('SELECT * FROM hr_people WHERE flex_id=?', (fid,)).fetchone())


def test_sync_makes_tasks_only_from_start_date(db):
    _person('A', '새사람', hire_date='2026-12-01')
    _person('B', '옛사람', hire_date='2024-01-01')                                  # 플렉스 때 일 — 안 만듦
    _person('C', '나간사람', hire_date='2024-01-01', status='퇴사', resign_date='2026-11-25')
    rows = ht.sync()
    got = sorted((r['name'], r['kind'], r['due_date']) for r in rows)
    assert got == [('나간사람', 'ins_loss', '2026-12-10'), ('나간사람', 'resign_docs', '2026-12-09'),
                   ('나간사람', 'resign_settle', '2026-12-09'),
                   ('새사람', 'hire_docs', '2026-12-08'), ('새사람', 'ins_acquire', '2026-12-15')]
    assert ht.sync() == []                                                           # 두 번 돌려도 그대로


def test_date_change_cancels_old_and_complete_once(db):
    p = _person('A', '새사람', hire_date='2026-12-01')
    ht.sync()
    with hr._conn() as conn:
        conn.execute("UPDATE hr_people SET hire_date='2026-12-03' WHERE id=?", (p['id'],))
    new = ht.sync()
    assert {r['event_date'] for r in new} == {'2026-12-03'}
    assert [t['event_date'] for t in ht.open_tasks(p['id'])] == ['2026-12-03', '2026-12-03']
    t = ht.open_tasks(p['id'])[0]
    assert ht.complete(t['id'], '황샛별')['status'] == ht.DONE
    assert ht.complete(t['id'], '황샛별') is None
    assert len(ht.for_person(p['id'])) == 2                                         # 취소분은 안 보임


def test_due_label_and_reminder(db):
    _person('A', '새사람', hire_date='2026-12-01')
    ht.sync()
    t = [r for r in ht.open_tasks() if r['kind'] == 'hire_docs'][0]
    assert ht.due_label(t, date(2026, 12, 6)) == '기한 12/8 · 2일 남음'
    assert ht.due_label(t, date(2026, 12, 8)) == '기한 12/8 · 오늘까지'
    assert ht.due_label(t, date(2026, 12, 10)) == '기한 12/8 · 2일 지남'
    assert ht.reminder_rows(date(2026, 12, 1)) == []
    assert [r['kind'] for r in ht.reminder_rows(date(2026, 12, 5))] == ['hire_docs']
    assert '새사람 — 입사 서류 받기' in ht.reminder_text(ht.reminder_rows(date(2026, 12, 5)), date(2026, 12, 5))
    blocks = ht.card_blocks(t, date(2026, 12, 5))
    assert blocks[1]['elements'][0]['action_id'] == ht.A_DONE and '/hr/people#' in blocks[1]['elements'][1]['url']
    done = ht.complete(t['id'], '황샛별')
    assert len(ht.card_blocks(done)) == 1 and '완료' in ht.card_blocks(done)[0]['text']['text']


class _C:
    def __init__(self):
        self.calls = []

    def conversations_open(self, users):
        return {'channel': {'id': 'D' + users}}

    def chat_postMessage(self, **kw):
        self.calls.append(('post', kw['channel'], kw.get('text')))
        return {'ts': str(len(self.calls))}

    def chat_update(self, **kw):
        self.calls.append(('update', kw['channel']))


def test_run_daily_dm_and_reminder(db, monkeypatch):
    from dashboard.blueprints import slack_hr_tasks as sh
    _person('A', '새사람', hire_date='2026-12-01')
    c = _C()
    sh.run_daily(date(2026, 12, 7), client=c)                       # 월요일, 서류 기한 12/8
    posts = [x for x in c.calls if x[0] == 'post']
    assert len(posts) == 3                                           # 새 할 일 2장 + 기한 알림 1
    assert all(p[1] == 'DUHANDLER' for p in posts) and '입사·퇴사 할 일 1건' in posts[-1][2]
    c2 = _C()
    sh.run_daily(date(2026, 12, 12), client=c2)                     # 토요일 — 알림 없음, 새 일도 없음
    assert c2.calls == []


def test_gw_home_todo_and_api(db, monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import hr as hr_bp_mod
    from dashboard.blueprints import slack_hr_tasks as sh
    from dashboard.services import gw_home
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(sh, 'sync_async', lambda: None)
    p = _person('A', '새사람', hire_date=date.today().isoformat())
    ht.sync()
    items = [i for i in gw_home.todos(None, '', False, True) if i['id'] == 'hr-tasks']
    assert items and items[0]['href'] == f"/hr/people#{p['id']}" and '1건' in items[0]['title']   # 7일 안 기한만
    assert not [i for i in gw_home.todos(None, '', False, False) if i['id'] == 'hr-tasks']
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(hr_bp_mod.hr_bp)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    tid = ht.open_tasks(p['id'])[0]['id']
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ADMIN, 'permission_level': 'admin'}
        card = cl.get(f"/api/hr/people/{p['id']}").get_json()['person']
        assert [t['title'] for t in card['tasks']] == ['입사 서류 받기', '4대보험 취득 신고']
        r = cl.post(f'/api/hr/tasks/{tid}/done').get_json()
        assert r['success'] and [t['status'] for t in r['tasks']].count('done') == 1
        with cl.session_transaction() as s:
            s['user'] = {'email': 'x@itg-aircon.com', 'permission_level': 'user'}
        assert cl.post(f'/api/hr/tasks/{tid}/done').status_code == 403
