# -*- coding: utf-8 -*-
"""확인 때 금액 경고 다시 계산 (2026-11-23 R4272-YM).

YM 이 잔금 지정(부가세 별도, 미수금 3,080,000) → '입금 후에도 280,000원 남음' 경고 → 그 사이 부가세 없음으로
바뀌어 미수금 2,800,000 → SB 확인 때도 옛 경고가 카드·시트 메모에 그대로 남았다.
"""
import sys

sys.path.insert(0, '.')

import dashboard.blueprints.slack_bot as sb

N_MEMO = '농협, 11/23 14:52\n352-****-1682-33\n입금 2,800,000원\n김연주'
G_HANA_TO_R = '하나, 10/02 10:14\n255-****-431304\n입금 2,800,000원\n김연주'   # R 통장


def _rec(code, total2, unpaid):
    return [{'프로젝트 코드': code, '총액 2': total2, '미수금': unpaid}]


def _designation(code, amount, memo, total2, unpaid, stage='잔금', acked=True):
    ws = sb._intake_warnings(code, amount, {}, memo, records=_rec(code, total2, unpaid), stage=stage)
    return ([sb._warn_line(w, acked) for w in ws],
            [n for n in (sb._warn_note(w, acked) for w in ws) if n])


def test_r4272_short_final_dropped_after_vat_off():
    warns, notes = _designation('R4272-YM', 2_800_000, N_MEMO, 3_080_000, 3_080_000)
    assert warns == ['잔금 후 미수금 남음 — 입금 후에도 280,000원 남음']
    assert notes == ['⚠️ 잔금 후 미수금: 280,000원 남음 (당시 미수금 3,080,000원)']
    w2, n2 = sb._refresh_amount_warns('R4272-YM', '잔금', 2_800_000, {}, N_MEMO, warns, notes,
                                      records=_rec('R4272-YM', 2_800_000, 2_800_000))
    assert w2 == [] and n2 == []


def test_short_final_number_updated_when_still_short():
    warns, notes = _designation('R4272-YM', 2_800_000, N_MEMO, 3_080_000, 3_080_000)
    w2, n2 = sb._refresh_amount_warns('R4272-YM', '잔금', 2_800_000, {}, N_MEMO, warns, notes,
                                      records=_rec('R4272-YM', 2_900_000, 2_900_000))
    assert w2 == ['잔금 후 미수금 남음 — 입금 후에도 100,000원 남음']
    assert n2 == ['⚠️ 잔금 후 미수금: 100,000원 남음 (당시 미수금 2,900,000원)']


def test_note_tail_kept():
    # 지정 때 모달 경고를 못 봤으면 메모 꼬리 '— 지정 시 경고 미표시' 유지
    warns, notes = _designation('R4272-YM', 2_800_000, N_MEMO, 3_080_000, 3_080_000, acked=False)
    assert notes[0].endswith(' — 지정 시 경고 미표시')
    _, n2 = sb._refresh_amount_warns('R4272-YM', '잔금', 2_800_000, {}, N_MEMO, warns, notes,
                                     records=_rec('R4272-YM', 2_900_000, 2_900_000))
    assert n2 == ['⚠️ 잔금 후 미수금: 100,000원 남음 (당시 미수금 2,900,000원) — 지정 시 경고 미표시']


def test_overpay_dropped_when_total_raised():
    warns, notes = _designation('G4139-MJ', 500_000, N_MEMO, 1_000_000, 300_000, stage='중도금')
    assert sb._amount_warn_kind(warns[0], 1) == 'overpay'
    w2, n2 = sb._refresh_amount_warns('G4139-MJ', '중도금', 500_000, {}, N_MEMO, warns, notes,
                                      records=_rec('G4139-MJ', 1_200_000, 500_000))
    assert w2 == [] and n2 == []


def test_account_warning_untouched():
    # 계좌·사업자 불일치는 금액과 무관 → 그대로, 금액 경고만 빠짐
    warns, notes = _designation('G4139-MJ', 2_800_000, G_HANA_TO_R, 3_080_000, 3_080_000)
    assert {sb._amount_warn_kind(x, 1) for x in warns} == {None, 'short_final'}
    w2, n2 = sb._refresh_amount_warns('G4139-MJ', '잔금', 2_800_000, {}, G_HANA_TO_R, warns, notes,
                                      records=_rec('G4139-MJ', 2_800_000, 2_800_000))
    assert len(w2) == 1 and w2[0].startswith('자금 이동 요청')   # acked → 카드 줄은 '자금 이동 요청'
    assert len(n2) == 1 and n2[0].startswith('⚠️ 계좌·사업자 불일치')


def test_no_new_warning_added():
    # 지정 때 없던 금액 경고는 확인 때 새로 붙이지 않음
    w2, n2 = sb._refresh_amount_warns('R4272-YM', '잔금', 2_800_000, {}, N_MEMO, [], [],
                                      records=_rec('R4272-YM', 3_080_000, 3_080_000))
    assert w2 == [] and n2 == []


def test_project_missing_keeps_original():
    warns, notes = _designation('R4272-YM', 2_800_000, N_MEMO, 3_080_000, 3_080_000)
    w2, n2 = sb._refresh_amount_warns('R4272-YM', '잔금', 2_800_000, {}, N_MEMO, warns, notes,
                                      records=[])
    assert (w2, n2) == (warns, notes)


def test_records_error_keeps_original(monkeypatch):
    warns, notes = _designation('R4272-YM', 2_800_000, N_MEMO, 3_080_000, 3_080_000)
    import dashboard.services.project_service as ps

    def boom(*a, **k):
        raise RuntimeError('sheet down')
    monkeypatch.setattr(ps, 'get_project_records', boom)
    assert sb._refresh_amount_warns('R4272-YM', '잔금', 2_800_000, {}, N_MEMO, warns, notes) == (warns, notes)
