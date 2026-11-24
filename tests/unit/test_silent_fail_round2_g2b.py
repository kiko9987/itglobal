"""입찰 공고 모으기 조용한 실패 막기(11-24 조용한 실패 점검 15·16).

15. 검색어·누리장터·공동주택 실패가 WARNING 으로만 묻히고 화면엔 '공고 없음' → ERROR(전부 실패면 force_alert),
    지난번 모으기 결과(못 모은 곳 수)를 저장해 목록 API 로 돌려줌.
16. 지역·면허 조회가 실패한 공고를 버리던 것 → '다시 볼 공고'로 남겨 다음 모으기 때 다시 조회.
    _get 은 429·5xx 를 몇 번 다시 시도.
나라장터 API 는 부르지 않는다(가짜 requests·가짜 search/fetch_limits).
"""
from datetime import datetime, timedelta

import pytest
import requests

from dashboard.services import g2b_bids as g
from test_g2b_bids import MECH, _fake_api, _item, db  # noqa: F401  (db 픽스처 재사용)


class _Log:
    """g.logger 대신 — 수준·문구·extra 를 모아 둔다."""

    def __init__(self):
        self.records = []

    def _add(self, level, msg, **kw):
        self.records.append((level, msg, kw.get('extra')))

    def info(self, msg, **kw):
        self._add('INFO', msg, **kw)

    def warning(self, msg, **kw):
        self._add('WARNING', msg, **kw)

    def error(self, msg, **kw):
        self._add('ERROR', msg, **kw)

    def errors(self):
        return [r for r in self.records if r[0] == 'ERROR']


@pytest.fixture
def log(monkeypatch):
    lg = _Log()
    monkeypatch.setattr(g, 'logger', lg)
    return lg


# ── 16(a) _get 다시 시도 ─────────────────────────────────────────────
class _Resp:
    def __init__(self, status=200, body=None, text=''):
        self.status_code = status
        self._body = body
        self.text = text

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(f'{self.status_code} Error for url: https://x/?serviceKey=test-key&type=json')

    def json(self):
        if self._body is None:
            raise ValueError('Expecting value')
        return self._body


_OK = {'response': {'header': {'resultCode': '00'}, 'body': {'items': [{'bidNtceNo': 'A1'}], 'totalCount': 1}}}


def _fake_get(monkeypatch, seq):
    calls = []

    def get(url, params=None, timeout=None):
        calls.append(url)
        r = seq.pop(0)
        if isinstance(r, Exception):
            raise r
        return r
    monkeypatch.setattr(g.requests, 'get', get)
    monkeypatch.setattr(g, 'RETRY_WAITS', (0, 0))
    return calls


@pytest.mark.parametrize('status', [429, 500, 502, 503, 504])
def test_get_retries_busy_status_then_succeeds(db, monkeypatch, status):
    calls = _fake_get(monkeypatch, [_Resp(status), _Resp(status), _Resp(200, _OK)])
    body = g._get(g.OP_LIST, {'pageNo': 1})
    assert g._items(body) == [{'bidNtceNo': 'A1'}] and len(calls) == 3


def test_get_gives_up_after_retries(db, monkeypatch):
    calls = _fake_get(monkeypatch, [_Resp(502), _Resp(502), _Resp(502), _Resp(200, _OK)])
    with pytest.raises(requests.HTTPError):
        g._get(g.OP_LIST, {})
    assert len(calls) == 3   # 처음 + 다시 두 번


def test_get_retries_connection_error(db, monkeypatch):
    calls = _fake_get(monkeypatch, [requests.ConnectionError('끊김'), requests.Timeout('느림'), _Resp(200, _OK)])
    assert g._items(g._get(g.OP_LIST, {})) and len(calls) == 3


def test_get_does_not_retry_client_error(db, monkeypatch):
    calls = _fake_get(monkeypatch, [_Resp(404), _Resp(200, _OK)])
    with pytest.raises(requests.HTTPError):
        g._get(g.OP_LIST, {})
    assert len(calls) == 1


def test_get_xml_quota_error_is_readable(db, monkeypatch):
    xml = ('<OpenAPI_ServiceResponse><cmmMsgHeader><errMsg>SERVICE ERROR</errMsg>'
           '<returnAuthMsg>LIMITED_NUMBER_OF_SERVICE_REQUESTS_EXCEEDS_ERROR</returnAuthMsg>'
           '<returnReasonCode>22</returnReasonCode></cmmMsgHeader></OpenAPI_ServiceResponse>')
    _fake_get(monkeypatch, [_Resp(200, None, xml)])
    with pytest.raises(RuntimeError) as e:
        g._get(g.OP_LIST, {})
    assert 'LIMITED_NUMBER_OF_SERVICE_REQUESTS_EXCEEDS_ERROR(22)' in str(e.value)


# ── 15 목록 실패 → ERROR·상태 저장 ────────────────────────────────────
def test_partial_failure_logs_error_and_saves_status(db, monkeypatch, log):
    _fake_api(monkeypatch, [_item('A1')], {'A1': (['경기도'], [[MECH]])})
    real = g.search

    def search(kw, start, end, op=g.OP_LIST):
        if kw == '공조':
            raise RuntimeError('HTTP 502 Bad Gateway serviceKey=test-key')
        return real(kw, start, end, op)
    monkeypatch.setattr(g, 'search', search)
    r = g.collect()
    assert r['errors'] == 2 and r['new'] == 1          # 공사·물품 '공조' 두 곳
    errs = log.errors()
    assert len(errs) == 1 and '일부 실패 2/' in errs[0][1] and errs[0][2] is None   # 일부 실패는 보통 ERROR
    assert 'test-key' not in errs[0][1]
    data = g.list_bids(g.ELIGIBLE)
    st = data['collect_status']
    assert st['errors'] == 2 and not st['all_failed'] and st['sources'] == len(g.KEYWORDS) * 2 + 1 + len(g.NURI_KINDS)
    assert data['last_collect'] and [x['bid_no'] for x in data['items']] == ['A1']


def test_all_failed_force_alert_and_keeps_last_collect(db, monkeypatch, log):
    _fake_api(monkeypatch, [_item('A1')], {'A1': (['경기도'], [[MECH]])})
    g.collect()
    before = g.list_bids()['last_collect']
    assert before

    def boom(*a, **k):
        raise RuntimeError('SERVICE_KEY_IS_NOT_REGISTERED_ERROR serviceKey=test-key')
    monkeypatch.setattr(g, 'search', boom)
    monkeypatch.setattr(g, 'nuri_list', boom)
    monkeypatch.setattr(g, 'apt_list', boom)
    monkeypatch.setattr(g, 'mark_collected', lambda: pytest.fail('전부 실패면 마지막으로 모은 때를 바꾸지 않음'))
    r = g.collect()
    assert r['errors'] == len(g.KEYWORDS) * 2 + 1 + len(g.NURI_KINDS)
    errs = log.errors()
    assert len(errs) == 1 and '전부 실패' in errs[0][1] and errs[0][2] == {'force_alert': True}
    assert 'test-key' not in errs[0][1] and 'serviceKey=***' in errs[0][1]
    data = g.list_bids()
    assert data['collect_status']['all_failed'] and data['last_collect'] == before


def test_success_has_no_error_and_clean_status(db, monkeypatch, log):
    _fake_api(monkeypatch, [_item('A1')], {'A1': (['경기도'], [[MECH]])})
    g.collect()
    assert not log.errors()
    st = g.list_bids()['collect_status']
    assert st['errors'] == 0 and st['pending'] == 0 and not st['all_failed']


def test_status_missing_is_empty(db):
    assert g.list_bids()['collect_status'] == {}


# ── 16(b) 지역·면허 조회 실패 공고는 '다시 볼 공고'로 ─────────────────
def test_limit_failure_not_saved_then_retried_next_collect(db, monkeypatch, log):
    calls = _fake_api(monkeypatch, [_item('A1'), _item('B2')], {'A1': (['경기도'], [[MECH]]), 'B2': (['경기도'], [[MECH]])})
    real = g.fetch_limits

    def flaky(no, ord_, source='나라장터'):
        if no == 'B2':
            raise RuntimeError('HTTP 503 Service Unavailable')
        return real(no, ord_, source)
    monkeypatch.setattr(g, 'fetch_limits', flaky)
    r = g.collect()
    assert r['new'] == 1 and r['limit_errors'] == 1 and r['pending'] == 1
    # 지역·면허를 모르는 공고가 '입찰 가능'(전국·면허 제한 없음)으로 보이면 안 됨 — 어느 탭에도 없다
    seen = [x['bid_no'] for t in (g.ELIGIBLE, g.BLOCKED) for x in g.list_bids(t)['items']]
    assert seen == ['A1']
    assert g.list_bids()['collect_status']['pending'] == 1
    assert '_pending_since' in g._load_pending()['B2']
    # 다음 모으기 — 이틀 조회 기간을 지나 목록에 안 나와도 남겨 둔 항목으로 다시 조회해 저장
    monkeypatch.setattr(g, 'fetch_limits', real)
    monkeypatch.setattr(g, 'search', lambda *a, **k: [])
    r = g.collect()
    assert r['new'] == 1 and r['pending'] == 0 and g._load_pending() == {}
    b2 = [x for x in g.list_bids(g.ELIGIBLE)['items'] if x['bid_no'] == 'B2'][0]
    assert b2['regions'] == ['경기도']
    assert calls['limits'] == 2   # A1 한 번 + B2 다시 한 번(A1 은 다시 안 부름)


def test_pending_keeps_newer_order_from_list(db, monkeypatch, log):
    _fake_api(monkeypatch, [_item('A1')], {'A1': (['경기도'], [[MECH]])})
    monkeypatch.setattr(g, 'fetch_limits', lambda *a, **k: (_ for _ in ()).throw(RuntimeError('x')))
    g.collect()
    assert g._load_pending()['A1']['bidNtceOrd'] == '000'
    _fake_api(monkeypatch, [_item('A1', ord_='001', title='OO초 냉난방개선 부대공사(정정)')], {'A1': (['경기도'], [[MECH]])})
    g.collect()
    rows = g.list_bids(g.ELIGIBLE)['items']
    assert rows[0]['bid_ord'] == '001' and g._load_pending() == {}


def test_limit_calls_stop_after_streak(db, monkeypatch, log):
    items = [_item(f'A{i}') for i in range(8)]
    calls = {'n': 0}
    _fake_api(monkeypatch, items, {})

    def down(*a, **k):
        calls['n'] += 1
        raise RuntimeError('HTTP 502')
    monkeypatch.setattr(g, 'fetch_limits', down)
    r = g.collect()
    assert calls['n'] == g.LIMIT_FAIL_STOP and r['pending'] == 8 and r['new'] == 0


def test_closed_pending_dropped(db, monkeypatch, log):
    past = (datetime.now() - timedelta(hours=1)).strftime('%Y-%m-%d %H:%M:%S')
    now_s = g._now()
    keep = g._save_pending({'OLD': _item('OLD', bidClseDt=past, opengDt=past), 'NEW': _item('NEW')}, now_s)
    assert list(keep) == ['NEW'] and list(g._load_pending()) == ['NEW']
    # 마감 일시를 모르는 공고는 오래되면 버림
    ancient = (datetime.now() - timedelta(days=g.PENDING_DAYS + 1)).strftime('%Y-%m-%d %H:%M:%S')
    keep = g._save_pending({'X': {'bidNtceNo': 'X', '_pending_since': ancient}, 'Y': {'bidNtceNo': 'Y'}}, now_s)
    assert list(keep) == ['Y']


def test_stale_pending_alerts(db, monkeypatch, log):
    _fake_api(monkeypatch, [_item('A1')], {})
    monkeypatch.setattr(g, 'fetch_limits', lambda *a, **k: (_ for _ in ()).throw(RuntimeError('x')))
    g.collect()
    assert not log.errors()   # 처음 실패는 다음 모으기로 넘김(WARNING)
    p = g._load_pending()
    p['A1']['_pending_since'] = (datetime.now() - timedelta(hours=25)).strftime('%Y-%m-%d %H:%M:%S')
    g._kv_set('pending_limits', p)
    g.collect()
    errs = log.errors()
    assert len(errs) == 1 and '하루 넘게' in errs[0][1] and errs[0][2] == {'force_alert': True}
    assert g.list_bids()['collect_status']['stale'] == 1


def test_reclassify_never_sees_pending(db, monkeypatch, log):
    """다시 분류(회사 정보 저장·재시작)는 저장된 줄만 — 지역·면허 없는 공고가 '가능'으로 바뀌지 않음."""
    _fake_api(monkeypatch, [_item('A1')], {})
    monkeypatch.setattr(g, 'fetch_limits', lambda *a, **k: (_ for _ in ()).throw(RuntimeError('x')))
    g.collect()
    g.reclassify()
    assert g.list_bids(g.ELIGIBLE)['items'] == [] and g.home_summary()[0] == 0
