# -*- coding: utf-8 -*-
"""계산서 요청 금액 사전 점검 (2026-10-15 G4146-JW·G4157-JW).

부가세 포함 금액을 'VAT 별도'로 요청해 부가세가 두 번 붙은 사고 → 요청 화면에서 경고, 같은 금액으로
다시 제출하면 통과. 과거 'VAT 별도' 요청 251건(7/24~10/15) 백테스트: ①은 이 두 건만, ②는 4건.
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb


def _rec(code, t1, t2):
    return [{'프로젝트 코드': code, '총액 1': t1, '총액 2': t2}]


@pytest.fixture
def proj(monkeypatch):
    import dashboard.services.project_service as ps
    state = {'recs': [], 'issued': 0}
    monkeypatch.setattr(ps, 'get_project_records', lambda: state['recs'])
    monkeypatch.setattr(sb, '_project_issued_invoice', lambda code: {'issued_gross': state['issued']})
    return state


@pytest.mark.parametrize('amt, supply', [(6_887_100, 6_261_000), (10_016_600, 9_106_000)])
def test_vat_included_amount_sent_as_separate_warns(proj, amt, supply):
    w = sb._invoice_amount_warning('G4157-JW', str(amt), 'sep')
    assert f'{amt:,}원은 부가세 포함 금액으로 보입니다' in w and f'공급가 {supply:,}원' in w


@pytest.mark.parametrize('amt, vat', [
    (6_701_000, 'sep'),       # 천 원 단위 공급가(정상) — G4146 제이케이디 요청
    (10_016_600, 'incl'),     # VAT 포함으로 고르면 정상
    (5_911_818, 'sep'),       # ÷1.1 이 천 원 단위 아님
])
def test_normal_amounts_pass(proj, amt, vat):
    proj['recs'] = _rec('X1', 50_000_000, 55_000_000)
    assert sb._invoice_amount_warning('X1', str(amt), vat) == ''


def test_over_total_warns(proj):
    # R3862-YG 형: 총액(부가세 포함) 6,503,000원을 'VAT 별도'로 요청 → 7,153,300원
    proj['recs'] = _rec('R3862-YG', 5_911_818, 6_503_000)
    w = sb._invoice_amount_warning('R3862-YG', '6503000', 'sep')
    assert '총액 6,503,000원을 넘습니다' in w and '7,153,300원' in w


def test_over_total_counts_already_issued(proj):
    proj['recs'] = _rec('G4146-JW', 12_962_000, 14_258_200)
    proj['issued'] = 12_880_780
    w = sb._invoice_amount_warning('G4146-JW', '6701000', 'sep')
    assert w.startswith('⚠️ 이미 발행 12,880,780원 + 이번 7,371,100원')
    # 남은 금액 안쪽이면 통과, 반올림 수준(1,000원 이내) 초과도 통과
    assert sb._invoice_amount_warning('G4146-JW', '1252200', 'sep') == ''
    proj['issued'] = 0
    proj['recs'] = _rec('G3959-YG', 3_895_455, 4_285_000)
    assert sb._invoice_amount_warning('G3959-YG', '4285500', 'incl') == ''


def test_replace_skips_over_total(proj):
    proj['recs'] = _rec('R3862-YG', 5_911_818, 6_503_000)
    assert sb._invoice_amount_warning('R3862-YG', '6503000', 'sep', has_replace=True) == ''


class _FakeRedis:
    def __init__(self):
        self.s = {}

    def get(self, k):
        return self.s.get(k)

    def set(self, k, v, ex=None):
        self.s[k] = v


def _view(amt, vat='sep'):
    return {'id': 'V1', 'private_metadata': '{"code": "G4157-JW"}',
            'state': {'values': {'amt': {'value': {'value': amt}},
                                 'vat': {'value': {'selected_option': {'value': vat}}}}}}


def test_modal_check_warns_once_then_passes(proj, monkeypatch):
    fr = _FakeRedis()
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    assert '부가세 포함 금액' in sb._invoice_submit_amount_check(_view('10,016,600'))
    assert sb._invoice_submit_amount_check(_view('10,016,600')) == ''      # 재제출 = 확인
    assert '부가세 포함 금액' in sb._invoice_submit_amount_check(_view('6,887,100'))   # 금액 바꾸면 다시 점검
    assert sb._invoice_submit_amount_check(_view('9,106,000')) == ''


def test_modal_check_skips_without_redis(proj, monkeypatch):
    import dashboard.utils.redis_client as rcm

    def _boom():
        raise RuntimeError('redis down')
    monkeypatch.setattr(rcm, 'get_redis_client', _boom)
    assert sb._invoice_submit_amount_check(_view('10,016,600')) == ''      # 막히지 않게 점검 생략
