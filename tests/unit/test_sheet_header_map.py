# -*- coding: utf-8 -*-
"""공사 현황 열 지도를 헤더에서 읽기 (2026-10-22 거래처명 열 추가 준비).

① 헤더 읽기: 지금 헤더면 예전 고정 배치와 똑같은 지도, 중복·필수 헤더 누락이면 쓰지 않음,
   읽기 실패면 마지막 지도(Redis) → 기본 배치.
② 열이 끼어든 배치(D 뒤 '거래처명')를 흉내 내 — 새 공사 수식·행 폭·버전 칸·입금 메모 열·수금 폴러 기준값·
   반환 후보·정합성 문구가 모두 새 위치를 따라가는지.
"""
import json
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.utils.google_sheets as gs

CURRENT = dict(gs._STATIC_COLUMN_MAPPING)


def _shifted():
    """D(유입 구분) 뒤에 '거래처명' — E 부터 한 칸씩 밀린 배치."""
    fields = [CURRENT[l] for l in sorted(CURRENT, key=gs.column_index)]
    fields.insert(4, '거래처명')
    return {gs.column_letter(i): f for i, f in enumerate(fields)}


class _Svc:
    def __init__(self, header=None, fail=False):
        self.header, self.fail, self.calls = header, fail, 0

    def spreadsheets(self):
        svc = self

        class _S:
            def values(self):
                class _V:
                    def get(self, spreadsheetId, range):
                        svc.calls += 1
                        if svc.fail:
                            raise RuntimeError('HttpError 503')
                        return type('E', (), {'execute': lambda _s: {'values': [svc.header]}})()
                return _V()
        return _S()


class _R:
    def __init__(self):
        self.d = {}

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, **kw):
        self.d[k] = v


@pytest.fixture
def live(monkeypatch):
    """pytest 중에도 헤더 읽기 경로를 켬 + 캐시 초기화 + 가짜 Redis."""
    monkeypatch.setenv('SHEET_HEADER_MAP_FORCE', '1')
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'SID')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    gs._header_map_cache.update(at=0.0, map=None)
    r = _R()
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': r})())
    yield r
    gs._header_map_cache.update(at=0.0, map=None)


def _header(mapping):
    return [mapping[l] for l in sorted(mapping, key=gs.column_index)]


def test_current_header_equals_old_fixed_layout(live):
    m = gs.main_column_mapping(_Svc(_header(CURRENT)))
    assert m == CURRENT
    assert json.loads(live.d['sheet_header_map:SID:공사 현황']) == CURRENT   # 마지막 지도 저장


def test_cache_one_read_per_minute(live):
    svc = _Svc(_header(CURRENT))
    for _ in range(5):
        gs.main_column_mapping(svc)
    assert svc.calls == 1


def test_refresh_in_progress_returns_last_map_without_waiting(live):
    gs.main_column_mapping(_Svc(_header(CURRENT)))
    gs._header_map_cache['at'] = 0.0                 # 다시 읽을 때가 됐는데
    svc = _Svc(_header(_shifted()))
    gs._header_map_lock.acquire()                    # 다른 스레드가 읽는 중
    try:
        assert gs.main_column_mapping(svc) == CURRENT and svc.calls == 0
    finally:
        gs._header_map_lock.release()
    assert gs.main_column_mapping(svc) == _shifted() and svc.calls == 1


def test_bad_headers_rejected(live):
    dup = _header(CURRENT) + ['계약금']
    assert gs._read_header_mapping(_Svc(dup)) is None
    missing = [h for h in _header(CURRENT) if h != '_version']
    assert gs._read_header_mapping(_Svc(missing)) is None


def test_read_failure_uses_last_good_then_static(live):
    live.d['sheet_header_map:SID:공사 현황'] = json.dumps(_shifted(), ensure_ascii=False)
    assert gs.main_column_mapping(_Svc(fail=True)) == _shifted()        # 마지막 지도
    gs._header_map_cache.update(at=0.0, map=None)
    live.d.clear()
    assert gs.main_column_mapping(_Svc(fail=True)) == CURRENT           # 기본 배치


def test_pytest_default_is_static(monkeypatch):
    monkeypatch.delenv('SHEET_HEADER_MAP_FORCE', raising=False)
    assert gs.main_column_mapping() == CURRENT


# ── 열이 끼어든 배치 흉내 ──
@pytest.fixture
def shifted(monkeypatch):
    m = _shifted()
    monkeypatch.setattr(gs, 'main_column_mapping', lambda service=None: dict(m))
    return m


def test_shifted_letters(shifted):
    f2l = gs.main_field_to_letter()
    assert (f2l['거래처명'], f2l['사업자명'], f2l['총액 1'], f2l['계약금'], f2l['잔금 부담자'], f2l['_version']) \
        == ('E', 'F', 'S', 'V', 'AO', 'AT')
    from dashboard.constants import payment_field_column
    assert [payment_field_column(f) for f in ('계약금', '중도금', '잔금', '계산서')] == ['V', 'W', 'X', 'Z']


def test_shifted_new_row(shifted):
    from dashboard.blueprints import projects as P

    class M:
        def get_column_mapping(self):
            return dict(shifted)

        def get_field_to_letter(self):
            return {f: l for l, f in shifted.items()}

    f = P._new_row_formulas(M().get_field_to_letter(), 10)
    assert f['총액 2'].startswith('=IF(T10=TRUE, S10+FLOOR(S10*0.1,1)')
    assert f['미수금'] == '=IF(ABS($U10-$V10-$W10-$X10)<2, 0, $U10-$V10-$W10-$X10)'
    assert f['순익'] == '=S10-(AF10+AG10+AH10+AI10)'
    assert f['마진율'] == '=IF(OR(S10=0, AJ10=0), 0, AJ10/S10)'
    vals = P._build_row_values({'거래처명': '디자인성실한', '사업자명': '코나솔루션', 'Lead No': 'L-1'}, M(), 10)
    assert len(vals) == 46
    assert (vals[4], vals[5], vals[44], vals[45]) == ('디자인성실한', '코나솔루션', 'L-1', '0')


def test_current_new_row_formulas_unchanged():
    # 지금 배치에서 만든 수식은 예전에 박아 둔 수식과 글자 하나 다르지 않다
    from dashboard.blueprints import projects as P
    f2l = {f: l for l, f in CURRENT.items()}
    r = 4300
    old = {
        '총액 2': f'=IF(S{r}=TRUE, R{r}+FLOOR(R{r}*0.1,1) + IF(MOD(R{r}+FLOOR(R{r}*0.1,1), 10)=1, -1, IF(MOD(R{r}+FLOOR(R{r}*0.1,1), 10)=9, 1, 0)), R{r})',
        '미수금': f'=IF(ABS($T{r}-$U{r}-$V{r}-$W{r})<2, 0, $T{r}-$U{r}-$V{r}-$W{r})',
        '순익': f'=R{r}-(AE{r}+AF{r}+AG{r}+AH{r})',
        '마진율': f'=IF(OR(R{r}=0, AI{r}=0), 0, AI{r}/R{r})',
    }
    assert P._new_row_formulas(f2l, r) == old


def test_shifted_payment_baseline_and_card(shifted):
    from dashboard.services import payment_sync as ps
    I = gs.main_field_to_index()
    vals = [''] * 46
    vals[I['계약금']], vals[I['중도금']], vals[I['잔금']] = 1000, 2000, 3000
    vals[I['미수금']], vals[I['수금 확인']] = 500, True
    vals[I['현장 주소']], vals[I['총액 2']], vals[I['계산서']] = '서울 어딘가', 6500, '미발행'
    b = ps.row_baseline_mapping(vals, ['', '', ''])
    assert (b['u'], b['v'], b['w'], b['x'], b['aa']) == (1000, 2000, 3000, 500, 'true')
    c = ps._card_fields(vals, I)
    assert (c['address'], c['total_t'], c['invoice'], c['unpaid']) == ('서울 어딘가', 6500, '미발행', 500)
    assert ps._sheet_col('계약금') == 'V' and ps._poll_width(I) == I['수금 확인'] + 1 == 31


def test_shifted_refund_candidates_read_new_positions(shifted):
    from dashboard.services import refund_match as rm
    I = gs.main_field_to_index()
    vals = [''] * 46
    vals[0] = 'G4100-XX'
    vals[I['총액 2']], vals[I['계약금']], vals[I['잔금']] = 1000000, 600000, 500000
    payers = rm._row_payers(vals, 600000, 0, 500000)
    assert payers == {}
    vals[I['계약금 부담자']] = '임대인 · 홍길동 · 600,000 · 계산서'
    vals[I['잔금 부담자']] = '임차인 · 김철수 · 400,000 · 계산서'
    assert set(rm._row_payers(vals, 600000, 0, 500000)) == {'계약금', '잔금'}


def test_shifted_invariant_message_letter(shifted):
    from dashboard.services import invariant_checks as ic
    out = ic.sheet_cell_anomalies([(10, 'G1', {'순익': '#REF!'})])
    assert '순익(AJ10)' in out[0][2]


def test_shifted_manager_ranges(shifted, monkeypatch):
    m = object.__new__(gs.GoogleSheetsManager)
    assert m.get_last_column_letter() == 'AT'
    assert m.get_field_to_index()['_version'] == 45
    assert m.get_field_column_mapping()['거래처명'] == 'E'
