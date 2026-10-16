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
    assert w[0]['line'] == '과입금 의심 — 미수금보다 200,000원 더 들어옴 (총액 1,000,000 · 이미 받은 돈 700,000 · 미수금 300,000)'
    assert '미수금보다 200,000원 더 들어온 입금입니다' in w[0]['modal'] and '이번 입금 500,000' in w[0]['modal']
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
    w = sb._intake_warnings('G4139-MJ', 650000, {}, R_HANA, records=_rec('G4139-MJ'))
    items = [('G4139-MJ', '잔금', 650000, w)]
    dep = {'partner': '대연이엔지주식회', 'date_md': '09/30'}
    sb._notify_intake_warns(_FakeClient(), 'iid1', 'C1', '1.1', items, 'U1', True, deposit=dep)
    sb._notify_intake_warns(_FakeClient(), 'iid1', 'C1', '1.1', items, 'U1', True, deposit=dep)  # 재호출
    assert len(dm.sent) == 1
    txt = dm.sent[0]['text']
    assert dm.sent[0]['channel'] == sb._SETTLEMENT_CHECKER_ID
    assert '*자금 이동 요청*  `G4139-MJ`' in txt
    assert 'MJ 님이 경고를 확인하고 입금을 지정했습니다. 법인 간 이체로 자금을 옮겨 주세요.' in txt
    assert '프로젝트 : G4139-MJ (글로벌) · 잔금' in txt
    assert '입금 : 650,000원 · 대연이엔지주식회 · 09/30' in txt
    assert '입금된 계좌 : 하나은행 (글로벌그룹)' in txt
    assert '옮길 곳 : 글로벌 계좌 (기업은행)' in txt
    assert 'https://slack/x' in txt


def test_dm_overpay_and_not_acked_wording(monkeypatch):
    monkeypatch.setattr(sb, '_resolve_manager_initial', lambda u: 'YM')
    ov = sb._intake_warnings('G4125-YM', 500000, {}, G_MEMO, records=_rec('G4125-YM', 3_014_000, 300_000))
    t = sb._build_intake_warn_dm([('G4125-YM', '잔금', 500000, ov)], 'U', True)
    assert '*과입금 확인*' in t and '고객 확인 후 반환 여부를 판단해 주세요' in t
    assert '미수금보다 200,000원 더 들어옴 (총액 3,014,000 · 이미 받은 돈 2,714,000 · 미수금 300,000)' in t
    acc = sb._intake_warnings('G4125-YM', 3014000, {}, R_HANA, records=_rec('G4125-YM', 3_014_000, 3_014_000))
    t2 = sb._build_intake_warn_dm([('G4125-YM', '잔금', 3014000, acc)], 'U', False)
    assert '*계좌·사업자 불일치 확인*' in t2 and '경고가 표시되지 않았습니다' in t2
    assert '입금 계좌가 맞는지 확인해 주세요' in t2


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
                            [('G4139-MJ', '잔금', 650000, [w])], 'U1', True)
    assert ':warning: *자금 이동 요청*  `G4139-MJ`' in dm.sent[0]['text']


# ── 시트 메모 이력 (2026-10-02 '잘못 들어왔다는 히스토리는 메모에 다 남기기') ──

from dashboard.services.payment_sync import _parse_notes, _hash_payments  # noqa: E402

_ANN = ['⚠️ 계좌·사업자 불일치: 하나은행 (글로벌그룹) 입금 / 프로젝트 글로벌(G) — 자금 이동 요청 (지정 YM · 확인 SB)',
        '⚠️ 과입금: 당시 미수금 300,000원 / 초과 200,000원 (지정 YM · 확인 SB)',
        '⚠️ 반환 200,000원 입금 500,000원 2026-10-02 R>G']   # 위험 단어 섞여도 무시돼야 함


def test_parser_ignores_warning_annotation_lines():
    base = '2026/09/30\n하나,09/30, 17:02\n255******31304\n입금3,014,000원\n대연이엔지주식회'
    withann = base + '\n' + '\n'.join(_ANN)
    a = _parse_notes(['', '', base], stage_vals={'잔금': 3014000})
    b = _parse_notes(['', '', withann], stage_vals={'잔금': 3014000})
    assert a == b and len(b) == 1 and not b[0]['is_refund'] and b[0]['transfer_to'] == ''
    assert _hash_payments(a) == _hash_payments(b)     # 폴러 phash 불변 → 카드 정정 안 생김


def test_annotation_never_becomes_partner():
    base = '2026/10/02 15:01\n입금 500,000원\n452***38801011\n기업'      # 입금자 줄 없는 문자
    b = _parse_notes(['', '', base + '\n' + _ANN[1]], stage_vals={'잔금': 500000})
    a = _parse_notes(['', '', base], stage_vals={'잔금': 500000})
    assert a == b


def test_warn_note_and_compose(monkeypatch):
    monkeypatch.setattr(sb, '_resolve_manager_initial', lambda u: {'U1': 'YM', 'U2': 'SB'}.get(u, '-'))
    acc = sb._intake_warnings('G4125-YM', 3014000, {}, R_HANA, records=_rec('G4125-YM', 3014000, 3014000))[0]
    ov = sb._intake_warnings('G4125-YM', 500000, {}, G_MEMO, records=_rec('G4125-YM', 3014000, 300000))[0]
    assert sb._warn_note(acc, True) == '⚠️ 계좌·사업자 불일치: 하나은행 (글로벌그룹) 입금 / 프로젝트 글로벌(G) — 자금 이동 요청'
    assert sb._warn_note(acc, False).endswith('— 지정 시 경고 미표시')
    assert sb._warn_note(ov, True) == '⚠️ 과입금: 당시 미수금 300,000원 / 초과 200,000원'
    assert sb._compose_warn_notes([sb._warn_note(ov, True)], 'U1', 'U2') == [
        '⚠️ 과입금: 당시 미수금 300,000원 / 초과 200,000원 (지정 YM · 확인 SB)']
    assert sb._compose_warn_notes([], 'U1', 'U2') == []


class _CommitMgr:
    """_commit_intake_to_sheet 용 가짜 시트 (W 잔금 셀만)."""
    def __init__(self):
        self.vals = {'W10': 0, 'T10': 3014000, 'U10': 0, 'V10': 0}
        self.notes = {}

    def find_row_by_project_code(self, *a, **k): return 10
    def get_cell_value(self, sid, sn, cell): return self.vals.get(cell, '')
    def update_cell_value(self, sid, sn, cell, v): self.vals[cell] = v; return True
    def get_cell_note(self, sid, sn, cell): return self.notes.get(cell, '')
    def update_cell_note(self, sid, sn, cell, n): self.notes[cell] = n; return True
    def get_field_to_letter(self): return {}      # 계산서/Y 재계산 단계는 건너뜀


def test_commit_appends_history_right_below_deposit(monkeypatch):
    fm = _CommitMgr()
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'SID'); monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: fm)
    import dashboard.utils.user_database as udb     # ⚠️ 운영 users.db 감사 로그 차단
    monkeypatch.setattr(udb, 'get_audit_repository',
                        lambda: type('A', (), {'log_action': lambda self, **k: None})())
    memo = '[Web발신]\n하나,09/30, 17:02\n255******31304\n입금 3,014,000원\n대연이엔지주식회'
    ok, old, new, err = sb._commit_intake_to_sheet('G4125-YM', '잔금', 3014000, memo, 'U2',
                                                   note_extra=[_ANN[0]])
    assert ok and new == 3014000 and fm.vals['W10'] == 3014000
    note = fm.notes['W10']
    assert note.splitlines()[-1] == _ANN[0] and '대연이엔지주식회' in note.splitlines()[-2]
    p = _parse_notes(['', '', note], stage_vals={'잔금': 3014000})
    assert len(p) == 1 and p[0]['partner'] == '대연이엔지주식회' and p[0]['amount'] == 3014000


# ── 잔금 후 미수금 (2026-10-15 R4101-JSH: 잔금 10,400,000원 지정 후에도 미수금 1,600,000원) ──

def test_final_payment_leaving_balance_warns():
    w = sb._intake_warnings('R4101-JSH', 10_400_000, {}, R_HANA,
                            records=_rec('R4101-JSH', 13_200_000, 12_000_000), stage='잔금')
    assert _kinds(w) == ['short_final'] and w[0]['field'] == 'stage'
    assert w[0]['left'] == 1_600_000
    assert '미수금 1,600,000원이 남습니다' in w[0]['modal'] and "'중도금'" in w[0]['modal']
    assert w[0]['line'] == '잔금 후 미수금 남음 — 입금 후에도 1,600,000원 남음'
    assert w[0]['note'].startswith('⚠')            # 수금 파서가 건너뛰는 이력 줄
    assert '>' not in w[0]['line']


def test_final_payment_clearing_balance_no_warning():
    assert sb._intake_warnings('R4101-JSH', 12_000_000, {}, R_HANA,
                               records=_rec('R4101-JSH', 13_200_000, 12_000_000), stage='잔금') == []


def test_non_final_stage_leaving_balance_no_warning():
    for st in ('계약금', '중도금', None):
        assert sb._intake_warnings('R4101-JSH', 10_400_000, {}, R_HANA,
                                   records=_rec('R4101-JSH', 13_200_000, 12_000_000), stage=st) == []


def test_final_overpay_is_overpay_not_short():
    w = sb._intake_warnings('R4101-JSH', 13_000_000, {}, R_HANA,
                            records=_rec('R4101-JSH', 13_200_000, 12_000_000), stage='잔금')
    assert _kinds(w) == ['overpay']


def test_final_no_total_amount_no_warning():
    assert sb._intake_warnings('R4101-JSH', 100_000, {}, R_HANA,
                               records=_rec('R4101-JSH', 0, 0), stage='잔금') == []


def test_split_same_project_other_stage_counts_toward_balance(monkeypatch):
    # 한 입금을 같은 프로젝트 중도금+잔금으로 나눠 완납 → 경고 없음
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'get_project_records', lambda: _rec('R4101-JSH', 13_200_000, 12_000_000))
    splits = [{'project_code': 'R4101-JSH', 'stage': '중도금', 'amount': 2_000_000},
              {'project_code': 'R4101-JSH', 'stage': '잔금', 'amount': 10_000_000}]
    assert sb._split_warnings_bounded(splits, {}, R_HANA, timeout=2.0) == [[], []]
    splits[0]['amount'] = 1_000_000                 # 합쳐도 100만 원 남음 → 잔금 행만 경고
    res = sb._split_warnings_bounded(splits, {}, R_HANA, timeout=2.0)
    assert res[0] == [] and _kinds(res[1]) == ['short_final'] and res[1][0]['left'] == 1_000_000


def test_dm_short_final_wording(monkeypatch):
    monkeypatch.setattr(sb, '_resolve_manager_initial', lambda u: 'JSH')
    w = sb._intake_warnings('R4101-JSH', 10_400_000, {}, R_HANA,
                            records=_rec('R4101-JSH', 13_200_000, 12_000_000), stage='잔금')
    t = sb._build_intake_warn_dm([('R4101-JSH', '잔금', 10_400_000, w)], 'U', True)
    assert '*잔금 후 미수금 확인*' in t
    assert '총액을 고쳐야 하는지(할인 등), 남은 돈을 더 받는지 확인해 주세요' in t
    assert '입금 후 미수금 : 1,600,000원 (현재 미수금 12,000,000원)' in t


def test_overpay_r4104_wording():
    """R4104-SD: 총액 2,750,000 · 계약금 235,000 받음 → 전액 2,750,000 입금 = 미수금보다 235,000 더 (2026-10-16)."""
    w = sb._intake_warnings('R4104-SD', 2750000, {}, R_HANA, records=_rec('R4104-SD', 2_750_000, 2_515_000),
                            stage='잔금')
    assert [x for x in w if x['kind'] == 'overpay'][0]['line'] ==         '과입금 의심 — 미수금보다 235,000원 더 들어옴 (총액 2,750,000 · 이미 받은 돈 235,000 · 미수금 2,515,000)'
