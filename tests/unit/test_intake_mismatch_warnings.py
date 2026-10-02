# -*- coding: utf-8 -*-
"""입금 지정 이상 징후 점검 (2026-10-02 SB 건의).

고객에게 통장을 잘못 안내해 입금 계좌의 사업자와 프로젝트 코드가 다르거나(과입금 포함) 해도
경영지원이 짚어주지 않으면 매니저가 모른다 → 지정 시 경고, 매니저가 [지정]을 다시 누르면(확인)
카드 ⚠️ 줄 + 경영지원 DM.
- 계좌·사업자: G(글로벌)↔R(글로벌그룹)·P 불일치만. N통장(농협)·현금 제외, 계좌 미확인은 판정 안 함.
- 과입금: 입금액 > 미수금 (총액 2 > 0 일 때만).
"""
import sys
import time

sys.path.insert(0, '.')

import dashboard.blueprints.slack_bot as sb

G_MEMO = '기업, 10/02 10:14\n452***38801011\n입금 650,000원\n김옥현'
R_HANA = '하나, 10/02 10:14\n255-****-431304\n입금 650,000원\n김옥현'
R_GIUP = '기업, 10/02 10:14\n452***86504039\n입금 650,000원\n관공서'      # 하도급지킴이 = R
N_MEMO = '농협, 10/02 10:14\n352-****-1682-33\n입금 650,000원\n김옥현'


def _rec(code, total2=1_000_000, unpaid=1_000_000):
    return [{'프로젝트 코드': code, '총액 2': total2, '미수금': unpaid}]


def _kinds(warns):
    return [w['kind'] for w in warns]


def test_hana_R_account_to_G_project_warns():
    w = sb._intake_warnings('G4139-MJ', 650000, {}, R_HANA, records=_rec('G4139-MJ'))
    assert _kinds(w) == ['account']
    assert w[0]['field'] == 'project'
    assert '하나은행 (글로벌그룹)' in w[0]['line'] and '글로벌(G)' in w[0]['line']


def test_G_account_to_R_project_warns():
    w = sb._intake_warnings('R4100-MS', 650000, {}, G_MEMO, records=_rec('R4100-MS'))
    assert _kinds(w) == ['account']
    assert '글로벌그룹(R)' in w[0]['line']


def test_subcontract_guard_giup_account_is_R_not_G():
    # 기업은행이지만 하도급지킴이 계좌 = 글로벌그룹 → R 프로젝트면 정상, G 면 불일치
    assert sb._intake_warnings('R4100-MS', 650000, {}, R_GIUP, records=_rec('R4100-MS')) == []
    assert _kinds(sb._intake_warnings('G4139-MJ', 650000, {}, R_GIUP, records=_rec('G4139-MJ'))) == ['account']


def test_matching_account_no_warning():
    assert sb._intake_warnings('G4139-MJ', 650000, {}, G_MEMO, records=_rec('G4139-MJ')) == []
    assert sb._intake_warnings('R4100-MS', 650000, {}, R_HANA, records=_rec('R4100-MS')) == []


def test_N_account_and_cash_excluded():
    assert sb._intake_warnings('G4139-MJ', 650000, {}, N_MEMO, records=_rec('G4139-MJ')) == []
    cash = '현금, 10/02\n입금 650,000원\n현금 수령 (MJ)'
    assert sb._intake_warnings('R4100-MS', 650000, {}, cash, records=_rec('R4100-MS')) == []


def test_unknown_account_not_judged_by_bank_name():
    # 계좌 마스킹으로 판별 불가 → 은행명(기업=G/R 모호)으로 추측하지 않음
    memo = '기업, 10/02 10:14\n452***12\n입금 650,000원\n김옥현'
    assert sb._intake_warnings('R4100-MS', 650000, {}, memo, records=_rec('R4100-MS')) == []


def test_preview_acct_code_fallback():
    w = sb._intake_warnings('G4139-MJ', 650000, {'acct_code': 'R'}, '', records=_rec('G4139-MJ'))
    assert _kinds(w) == ['account']


def test_plant_project_with_G_account_warns():
    assert _kinds(sb._intake_warnings('P0012-SH', 650000, {}, G_MEMO, records=_rec('P0012-SH'))) == ['account']


def test_overpay():
    w = sb._intake_warnings('G4139-MJ', 500000, {}, G_MEMO, records=_rec('G4139-MJ', 1_000_000, 300_000))
    assert _kinds(w) == ['overpay'] and w[0]['field'] == 'stage'
    assert '입금 500,000원이 미수금 300,000원보다 많음' in w[0]['line']
    assert '>' not in w[0]['line']      # 슬랙 인용 기호 금지


def test_exact_remaining_not_overpay():
    assert sb._intake_warnings('G4139-MJ', 300000, {}, G_MEMO, records=_rec('G4139-MJ', 1_000_000, 300_000)) == []


def test_fully_paid_project_any_deposit_is_overpay():
    assert _kinds(sb._intake_warnings('G4139-MJ', 100000, {}, G_MEMO,
                                      records=_rec('G4139-MJ', 650_000, 0))) == ['overpay']


def test_no_total_amount_no_overpay():
    # 공사 금액 미입력(총액 2 = 0) → 미수금 0 이라도 과입금 판정 안 함
    assert sb._intake_warnings('G4139-MJ', 650000, {}, G_MEMO, records=_rec('G4139-MJ', 0, 0)) == []


def test_both_warnings():
    w = sb._intake_warnings('G4139-MJ', 500000, {}, R_HANA, records=_rec('G4139-MJ', 1_000_000, 300_000))
    assert _kinds(w) == ['account', 'overpay']


def test_precheck_bounded_returns_warnings_when_dup_check_slow(monkeypatch):
    monkeypatch.setattr(sb, '_intake_duplicate_check', lambda *a, **k: time.sleep(5) or None)
    monkeypatch.setattr(sb, '_intake_warnings', lambda *a, **k: [{'kind': 'account'}])
    t = time.time()
    dup, warns = sb._intake_precheck_bounded('G4139-MJ', '잔금', 650000, '10/02', {}, R_HANA, timeout=0.5)
    assert time.time() - t < 1.5          # ack 예산 보호
    assert dup is None and warns == [{'kind': 'account'}]


class _FakeRedis:
    def __init__(self):
        self.s = {}

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.s:
            return None
        self.s[k] = v
        return True


class _FakeDM:
    def __init__(self):
        self.sent = []

    def chat_postMessage(self, **kw):
        self.sent.append(kw)


class _FakeClient:
    def chat_getPermalink(self, **kw):
        return {'permalink': 'https://slack/x'}


def test_notify_sends_once_per_intake(monkeypatch):
    fr, dm = _FakeRedis(), _FakeDM()
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    monkeypatch.setattr(sb, '_dm_client', lambda: dm)
    monkeypatch.setattr(sb, '_resolve_manager_initial', lambda u: 'MJ')
    items = [('G4139-MJ', '잔금', 650000, ['계좌·사업자 불일치 — 입금 하나은행 (글로벌그룹) / 프로젝트 글로벌(G)'])]
    sb._notify_intake_warns(_FakeClient(), 'iid1', 'C1', '1.1', items, 'U1', True)
    sb._notify_intake_warns(_FakeClient(), 'iid1', 'C1', '1.1', items, 'U1', True)   # 재지정 등 재호출
    assert len(dm.sent) == 1
    txt = dm.sent[0]['text']
    assert dm.sent[0]['channel'] == sb._SETTLEMENT_CHECKER_ID
    assert '`G4139-MJ` · 잔금 · 650,000원 — 지정 MJ' in txt and '경고를 확인하고' in txt
    assert 'https://slack/x' in txt


def test_notify_skips_when_no_warnings(monkeypatch):
    dm = _FakeDM()
    monkeypatch.setattr(sb, '_dm_client', lambda: dm)
    sb._notify_intake_warns(_FakeClient(), 'iid2', 'C1', '1.1', [('G1', '잔금', 1, [])], 'U1', True)
    assert dm.sent == []


def test_pending_and_done_cards_show_warning_lines():
    warns = ['과입금 의심 — 입금 500,000원 > 미수금 300,000원']
    pend = sb._build_intake_pending_blocks('iid', 'G4139-MJ', '잔금', 500000, G_MEMO, '', warns=warns)
    done = sb._build_intake_done_blocks('G4139-MJ', '잔금', 500000, G_MEMO, '', '', warns=warns)
    assert warns[0] in pend[0]['text']['text'] and warns[0] in done[0]['text']['text']
    split = sb._build_intake_split_pending_blocks(
        'iid', [{'project_code': 'G1', 'stage': '잔금', 'amount': 1, 'warns': warns},
                {'project_code': 'R2', 'stage': '잔금', 'amount': 2}], 3, '', G_MEMO)
    assert split[0]['text']['text'].count('과입금 의심') == 1


def test_acked_account_mismatch_becomes_fund_move_request(monkeypatch):
    """매니저가 모달 경고 후 [지정] 재클릭 = 자금 이동 요청 (카드·DM 문구, 2026-10-02 사용자 문구)."""
    w = sb._intake_warnings('G4139-MJ', 650000, {}, R_HANA, records=_rec('G4139-MJ'))[0]
    assert '자금 이동을 요청하려면 [지정]을 한 번 더' in w['modal']
    assert sb._warn_line(w, True).startswith('자금 이동 요청 — 입금: 하나은행 (글로벌그룹)')
    assert sb._warn_line(w, False).startswith('계좌·사업자 불일치')     # 경고 미표시면 중립
    ov = {'kind': 'overpay', 'line': '과입금 의심 — x'}
    assert sb._warn_line(ov, True) == '과입금 의심 — x'
    fr, dm = _FakeRedis(), _FakeDM()
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    monkeypatch.setattr(sb, '_dm_client', lambda: dm)
    monkeypatch.setattr(sb, '_resolve_manager_initial', lambda u: 'MJ')
    sb._notify_intake_warns(_FakeClient(), 'iid9', 'C1', '1.1',
                            [('G4139-MJ', '잔금', 650000, [sb._warn_line(w, True)])], 'U1', True)
    assert dm.sent[0]['text'].startswith(':warning: *자금 이동 요청* — 매니저가 계좌·사업자 불일치를 확인하고')
