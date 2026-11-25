"""전화 문의 보정(10초) — 리드 No·플랫폼 두 칸만 먼저 읽고, 번호 매길 행이 있을 때만 시트 전체를 읽는다(2026-11-25).
리드 캐시는 만료가 가까울 때만 뒤에서 새로 채운다. 이름·번호는 모두 가짜."""
from dashboard.services import lead_sync as ls
from dashboard.services import sync_scheduler as sch


class _Req:
    def __init__(self, values):
        self.values = values


class _Mgr:
    def __init__(self, values=None, boom=False):
        self.values, self.boom, self.ranges = values, boom, []
        mgr = self

        class _V:
            def get(self, spreadsheetId, range, **k):
                mgr.ranges.append(range)
                return _Req(mgr.values)

        class _S:
            def values(self):
                return _V()

        class _Svc:
            def spreadsheets(self):
                return _S()
        self.service = _Svc()

    def _execute_with_retry(self, factory, name=''):
        if self.boom:
            raise RuntimeError('시트 오류')
        return {'values': factory().values}


def _setup(monkeypatch, mgr):
    monkeypatch.setattr(ls, '_get_sheet_config', lambda: {'sheet_id': 'S', 'sheet_name': '리드'})
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: mgr)


HEAD = ['리드 No', '상담 시간', '플랫폼']


def test_light_check(monkeypatch):
    m = _Mgr([HEAD, ['L-1', 't', '전화'], ['', 't', '홈페이지'], [], ['L-2']])
    _setup(monkeypatch, m)
    assert ls._wf_light_has_candidates() is False and m.ranges == ['리드!A:C']       # 두 칸만
    m.values = [HEAD, ['L-1', 't', '전화'], ['', 't', ' 거래처 ']]
    assert ls._wf_light_has_candidates() is True
    m.values = []
    assert ls._wf_light_has_candidates() is None                                     # 모르면 전체로
    _setup(monkeypatch, _Mgr(boom=True))
    assert ls._wf_light_has_candidates() is None


def test_sync_skips_full_read_when_nothing(monkeypatch):
    full = []
    monkeypatch.setattr(ls, 'load_leads_data', lambda *a, **k: full.append(k) or None)
    monkeypatch.setattr(ls, '_wf_light_has_candidates', lambda: False)
    r = ls.sync_workflow_phone_leads()
    assert r.get('light') and not full
    monkeypatch.setattr(ls, '_wf_light_has_candidates', lambda: None)               # 두 칸 읽기 실패 → 예전처럼 전체
    ls.sync_workflow_phone_leads()
    assert full == [{'force_refresh': True}]
    monkeypatch.setattr(ls, '_wf_light_has_candidates', lambda: True)
    ls.sync_workflow_phone_leads()
    assert len(full) == 2


def test_cache_warm_only_near_expiry(monkeypatch):
    import dashboard.services.lead_service as lsvc
    import dashboard.utils.smart_cache_manager as scm
    loads, ttl = [], {'ms': 50000}

    class _R:
        def pttl(self, k):
            return ttl['ms']

    class _W:
        redis = _R()

    class _C:
        redis = _W()
    monkeypatch.setattr(sch, '_redis_healthy', lambda: True)
    monkeypatch.setattr(scm, 'get_smart_cache', lambda: _C())
    monkeypatch.setattr(lsvc, 'load_leads_data', lambda *a, **k: loads.append(k))
    sch._safe_leads_cache_warm()
    assert loads == []                     # 아직 넉넉함
    ttl['ms'] = 20000
    sch._safe_leads_cache_warm()
    ttl['ms'] = -2                         # 캐시 없음
    sch._safe_leads_cache_warm()
    assert loads == [{'force_refresh': True}] * 2
