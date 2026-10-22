# -*- coding: utf-8 -*-
"""거래처명 열 규칙 (2026-10-22 사용자 확정) — 자동 채움·같은 곳 소급·표시·경고·안전장치.

- 거래처명: 거래처 유입이거나 2건 이상인 곳 = 이름(법인 표기 뺀), 아니면 유입 묶음(온라인·전화·소개·중고판매·기타).
- 같은 곳의 두 번째 공사가 생기면(등록 때 / 사업자명이 나중에 채워질 때) 앞 공사의 유입 묶음도 그 이름으로.
- 사업자명은 등록증 상호 그대로(계산서 기준).
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.services.partner_names as pn
from dashboard.utils.google_sheets import main_field_to_index, main_field_to_letter


def _rec(code, inflow, biz, partner):
    return {'프로젝트 코드': code, '유입 구분': inflow, '사업자명': biz, '거래처명': partner}


# ── 이름 규칙 ─────────────────────────────────────────────

@pytest.mark.parametrize('inflow,group', [
    ('당근', '온라인'), ('홈페이지', '온라인'), ('숨고', '온라인'), ('', '온라인'),
    ('전화', '전화'), ('소개', '소개'), ('중고판매', '중고판매'), ('방문', '기타'),
])
def test_inflow_group(inflow, group):
    assert pn.inflow_group(inflow) == group


def test_short_name_and_is_partner_name():
    assert pn.short_name('주식회사 메이드바이디자인') == '메이드바이디자인'
    assert pn.short_name('(주)에이치  디자인 [폐업]') == '에이치 디자인'
    for v in ('온라인', '전화', '거래처', '-', '', '당근', '개인', '김', 'nan', None):
        assert not pn.is_partner_name(v), v
    assert pn.is_partner_name('메이드바이')
    assert pn.place_key('메이드바이 디자인') == pn.place_key('(주)메이드바이디자인')


def test_display_name_prefers_partner():
    assert pn.display_name({'거래처명': '메이드바이', '사업자명': '홍길동'}) == '메이드바이'
    assert pn.display_name({'거래처명': '온라인', '사업자명': '(주)한빛'}) == '(주)한빛'
    assert pn.display_name({'거래처명': '온라인', '사업자명': '-'}) == '-'


# ── 새 공사 거래처명 + 소급 대상 ───────────────────────────

def test_new_online_first_project_gets_group():
    name, promote, place = pn.resolve_new_partner('당근', '-', '', [])
    assert (name, promote, place) == ('온라인', [], '')


def test_new_partner_inflow_uses_short_biz_or_dash():
    assert pn.resolve_new_partner('거래처', '주식회사 한빛인테리어', '', [])[0] == '한빛인테리어'
    assert pn.resolve_new_partner('거래처', '-', '', [])[0] == '-'          # 확인 필요


def test_given_name_wins():
    name, _, _ = pn.resolve_new_partner('거래처', '홍길동', '메이드바이', [])
    assert name == '메이드바이'


def test_second_project_same_place_promotes_first():
    recs = [_rec('G1', '당근', '(주)한빛', '온라인'), _rec('G2', '거래처', '다른곳', '다른곳')]
    name, promote, place = pn.resolve_new_partner('전화', '주식회사 한빛', '', recs, code='G9')
    assert (name, promote, place) == ('한빛', ['G1'], '한빛')


def test_promote_skips_names_and_dash():
    recs = [_rec('G1', '당근', '한빛', '온라인'), _rec('G2', '거래처', '한빛', '메이드바이'),
            _rec('G3', '거래처', '한빛', '-')]
    _, promote, _ = pn.resolve_new_partner('전화', '한빛', '', recs, code='G9')
    assert promote == ['G1']


def test_place_name_ignores_interior_of_customer_paid_job():
    """고객이 돈 낸 인테리어 경유 공사(거래처명=인테리어)는 그 고객 곳 이름 후보가 아님."""
    recs = [_rec('G1', '거래처', '한빛상사', '메이드바이')]
    assert pn.place_name('한빛상사', recs) == '한빛상사'
    # 사람이 부르는 이름으로 고쳐 둔 직접 공사 이름은 이어 씀
    recs.append(_rec('G2', '당근', '한빛상사', '한빛'))
    assert pn.place_name('한빛상사', recs) == '한빛'
    # 같은 곳을 가리키는 이름(띄어쓰기만 다름)이 있으면 그것
    recs.append(_rec('G3', '거래처', '한빛상사', '한빛 상사'))
    assert pn.place_name('한빛상사', recs) == '한빛 상사'


def test_known_partner_names_excludes_groups():
    recs = [_rec('G1', '', '', '온라인'), _rec('G2', '', '', '메이드바이'), _rec('G3', '', '', '메이드바이'),
            _rec('G4', '', '', '한빛')]
    assert pn.known_partner_names(recs) == ['메이드바이', '한빛']


# ── 소급 기록 (쓰기 큐·버전) ───────────────────────────────

class _Mgr:
    def __init__(self, rows):
        self.rows = rows            # code -> (row_number, {field: value})

    def get_field_to_letter(self): return main_field_to_letter()
    def get_field_to_index(self): return main_field_to_index()

    def find_row_by_project_code(self, sid, code, rng):
        return self.rows.get(code, (None, None))[0]

    def get_row_values(self, sid, sn, row):
        f2i = main_field_to_index()
        vals = [''] * (max(f2i.values()) + 1)
        for code, (r, data) in self.rows.items():
            if r == row:
                for f, v in data.items():
                    vals[f2i[f]] = v
        return vals


@pytest.fixture
def promote_env(monkeypatch):
    import dashboard.services.lead_service as ls
    import dashboard.services.project_service as ps
    import dashboard.services.sheet_write_queue as q
    import dashboard.utils.redis_client as rcm
    import dashboard.utils.user_database as udb
    import dashboard.services.project_slack_notifier as psn
    out = {'enq': [], 'cache': [], 'notify': [], 'mgr': None, 'recs': []}
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: out['mgr'])
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: out['recs'])
    monkeypatch.setattr(ps, 'update_project_in_cache', lambda c, f: out['cache'].append((c, f)) or True)
    monkeypatch.setattr(q, 'enqueue', lambda op, payload, meta=None: out['enq'].append((op, payload)))
    monkeypatch.setattr(rcm, 'get_redis_client',
                        lambda: type('C', (), {'redis': type('R', (), {'setex': lambda *a: None})()})())
    monkeypatch.setattr(udb, 'get_audit_repository',
                        lambda: type('A', (), {'log_action': lambda *a, **k: True})())
    monkeypatch.setattr(psn, 'notify_project_field_changes', lambda *a, **k: out['notify'].append((a, k)) or True)
    return out


def test_promote_writes_partner_and_bumps_version(promote_env):
    promote_env['mgr'] = _Mgr({
        'G1': (10, {'프로젝트 코드': 'G1', '거래처명': '온라인', '_version': '3'}),
        'G2': (11, {'프로젝트 코드': 'G2', '거래처명': '메이드바이', '_version': '1'}),   # 이미 이름 → 그대로
        'G3': (12, {'프로젝트 코드': 'G3', '거래처명': '-', '_version': '0'}),            # 확인 필요 → 그대로
    })
    promote_env['recs'] = [{'프로젝트 코드': 'G1', '거래처명': '온라인', '_version': '5'}]   # 큐에 남은 PM 저장
    n = pn.promote_previous(['G1', 'G2', 'G3'], '한빛', why='같은 곳 공사 G9 등록')
    assert n == 1
    (op, payload), = promote_env['enq']
    f2l = main_field_to_letter()
    assert op == 'project_update_sheet' and payload['row_number'] == 10
    assert payload['cells'] == [{'col': f2l['거래처명'], 'value': '한빛'}, {'col': f2l['_version'], 'value': '6'}]
    assert promote_env['cache'] == [('G1', {'거래처명': '한빛', '_version': '6'})]
    (args, kw), = promote_env['notify']
    assert args[1] == [{'field_name': '거래처명', 'old_value': '온라인', 'new_value': '한빛'}]
    assert kw['editor'] == '시스템 자동기록 · 같은 곳 공사 G9 등록'


def test_promote_ignores_non_name():
    assert pn.promote_previous(['G1'], '온라인') == 0
    assert pn.promote_previous([], '한빛') == 0


def test_biz_change_binds_same_place(monkeypatch):
    import dashboard.services.project_service as ps
    recs = [_rec('G1', '당근', '한빛', '온라인'), _rec('G5', '숨고', '-', '온라인')]
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: recs)
    got = []
    monkeypatch.setattr(pn, 'promote_previous', lambda codes, name, why='', editor='': got.append((codes, name)) or 2)
    assert pn.apply_place_on_biz_change('G5', '(주)한빛') == 2
    assert got == [(['G5', 'G1'], '한빛')]
    # 같은 곳이 없으면 아무것도 안 함
    got.clear()
    assert pn.apply_place_on_biz_change('G5', '다른상호') == 0 and got == []
    assert pn.apply_place_on_biz_change('G5', '-') == 0


# ── 등록 경로 (PM·슬랙 공용 _prepare_project_defaults) ──────────

def test_prepare_defaults_fills_partner_and_returns_promote(monkeypatch):
    import dashboard.services.project_service as ps
    from dashboard.blueprints.projects import _prepare_project_defaults
    monkeypatch.setattr(ps, 'get_project_records',
                        lambda force_refresh=False: [_rec('G1', '당근', '한빛', '온라인')])

    class _M:
        def get_field_to_letter(self): return main_field_to_letter()

    data = {'프로젝트 코드': 'G9', '유입 구분': '전화', '사업자명': '(주)한빛'}
    assert _prepare_project_defaults(data, 100, _M()) == (['G1'], '한빛')
    assert data['거래처명'] == '한빛'
    data = {'프로젝트 코드': 'G10', '유입 구분': '당근', '사업자명': ''}
    assert _prepare_project_defaults(data, 101, _M()) == ([], '')
    assert data['거래처명'] == '온라인' and data['사업자명'] == '-'


# ── 리컨사일러: 새 감시 필드는 조용히 기준값 ─────────────────

class _FakeRedis:
    def __init__(self):
        self.kv, self.h = {}, {}

    def get(self, k): return self.kv.get(k)
    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.kv:
            return None
        self.kv[k] = v
        return True
    def delete(self, k): self.kv.pop(k, None); self.h.pop(k, None)
    def hset(self, k, mapping): self.h.setdefault(k, {}).update(mapping)
    def hgetall(self, k): return dict(self.h.get(k, {}))
    def expire(self, k, t): pass
    def scan_iter(self, match='', count=0):
        return [k for k in list(self.kv) if k.startswith(match.rstrip('*'))]


def test_reconciler_new_watch_field_is_silent_baseline(monkeypatch):
    import dashboard.services.project_reconcile as pr
    import dashboard.services.project_slack_notifier as psn
    import dashboard.services.project_service as ps
    import dashboard.utils.redis_client as rcm
    fr = _FakeRedis()
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    code = 'G4100-JW'
    fr.set(f'project_card_msg:{code}', 'C1|1.1')
    watch = psn._NOTIFY_FIELDS - pr._RECONCILE_EXCLUDE
    assert '거래처명' in watch
    before = {f: '' for f in watch if f != '거래처명'}          # 열 추가 전 스냅샷
    before['프로젝트 코드'] = code
    fr.hset(f'card_field_snap:{code}', mapping=before)
    rec = dict(before, 거래처명='메이드바이')
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: [rec])
    sent = []
    monkeypatch.setattr(psn, 'notify_project_field_changes', lambda *a, **k: sent.append(a) or True)
    monkeypatch.setattr(psn, 'notify_amount_edit_to_settlement', lambda *a, **k: None)
    pr.reconcile_project_cards()
    assert sent == []
    assert fr.h[f'card_field_snap:{code}']['거래처명'] == '메이드바이'
    # 그다음 실제로 바뀌면 감지
    rec['거래처명'] = '한빛'
    pr.reconcile_project_cards()
    assert len(sent) == 1 and sent[0][1][0]['field_name'] == '거래처명'


# ── 입금 경고: 거래처명(인테리어)이 낸 입금은 정상 ─────────────

def test_payment_warning_accepts_partner_depositor(monkeypatch):
    import dashboard.services.project_service as ps
    from dashboard.services.payment_sync import _partner_misplacement_warning
    recs = [_rec('G1', '거래처', '홍길동', '메이드바이디자인'), _rec('G2', '거래처', '메이드바이디자인', '메이드바이디자인')]
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: recs)
    assert _partner_misplacement_warning('G1', '메이드바이디자인') == ''
    recs[0]['거래처명'] = '온라인'
    assert '다른 프로젝트' in _partner_misplacement_warning('G1', '메이드바이디자인')


# ── 수금 폴러 헤더 확인 ─────────────────────────────────────

class _Svc:
    def __init__(self, header):
        self.header = header

    def spreadsheets(self): return self
    def values(self): return self
    def get(self, **k): return self
    def execute(self): return {'values': [self.header]}


def test_poll_header_ok(monkeypatch):
    from dashboard.services import payment_sync as ps_
    idx = main_field_to_index()
    header = [''] * (max(idx.values()) + 1)
    for f, i in idx.items():
        header[i] = f
    assert ps_._poll_header_ok(_Svc(['엉뚱']), 'sid', 'sn', idx)          # 테스트(헤더 지도 꺼짐)에선 통과
    monkeypatch.setenv('SHEET_HEADER_MAP_FORCE', '1')
    assert ps_._poll_header_ok(_Svc(header), 'sid', 'sn', idx)
    shifted = header[:4] + ['거래처명'] + header[4:]                       # 열을 막 넣은 상황
    assert not ps_._poll_header_ok(_Svc(shifted), 'sid', 'sn', idx)


# ── 슬랙 /공사확정 거래처명 검색 ─────────────────────────────

def test_slack_partner_search(monkeypatch):
    import dashboard.services.project_service as ps
    from dashboard.blueprints.slack_bot import _search_partner_names
    recs = [_rec('G1', '', '', '메이드바이'), _rec('G2', '', '', '온라인'), _rec('G3', '', '', '메이드바이 2호')]
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: recs)
    vals = [o['value'] for o in _search_partner_names('메이드')]
    assert vals[0] == '메이드' and set(vals[1:]) == {'메이드바이', '메이드바이 2호'}   # 새 거래처 먼저
    vals = [o['value'] for o in _search_partner_names('메이드바이')]
    assert '온라인' not in vals and vals[0] in ('메이드바이', '메이드바이 2호')        # 똑같은 이름 있으면 새 항목 없음
    assert _search_partner_names('온라인') == []                                       # 분류값은 새 거래처로 안 만듦
