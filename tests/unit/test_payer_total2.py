# -*- coding: utf-8 -*-
"""나눠 내는 공사 — '현금 · 부가세 없음' 몫을 뺀 총액 2 (2026-10-21, 305호 G4148).

총액 2 = 부가세 없는 몫 + (총액 1 − 그 몫)에 부가세(절사 + 끝자리 1/9원 보정).
시트 계산식(total2_formula)·서버 계산(_calculate_total2)·부담자 저장(sync_total2_formula)이 같은 규칙.
"""
from unittest.mock import patch, MagicMock

import pandas as pd

from dashboard.services import payer_split as ps
from dashboard.services import project_service as prs

G4148 = {'프로젝트 코드': 'G4148-JW', '총액 1': 16042000, '총액 2': 17646200, '부가세': True,
         '계약금': 0, '중도금': 0, '잔금': 0,
         '계약금 부담자': '임대인 · 유미숙 · 6,000,000 · 현금 · 부가세 없음',
         '중도금 부담자': '-', '잔금 부담자': '임차인 · 주식회사 제이케이디학원 · 11,046,200 · 계산서',
         '수금 관련 특이사항': '', '제품대': 0, '도급비': 0, '자재비': 0, '기타비': 0}
F2L = {'총액 1': 'R', '부가세': 'S', '총액 2': 'T', '계약금': 'U', '중도금': 'V', '잔금': 'W'}
STD_FX = ('=IF(S4149=TRUE, R4149+FLOOR(R4149*0.1,1) + IF(MOD(R4149+FLOOR(R4149*0.1,1), 10)=1, -1, '
          'IF(MOD(R4149+FLOOR(R4149*0.1,1), 10)=9, 1, 0)), R4149)')


def test_calculate_total2_vat_free():
    assert prs._calculate_total2(16042000, True, 6000000) == 17046200
    assert prs._calculate_total2(16042000, True) == 17646200          # 기존 계산 그대로
    assert prs._calculate_total2(16042000, True, 0) == 17646200
    assert prs._calculate_total2(16042000, False, 6000000) == 16042000  # 부가세 없는 공사는 영향 없음
    assert prs._calculate_total2(5000000, True, 6000000) == 5000000     # 몫이 총액보다 크면 전부 부가세 없음
    assert prs._calculate_total2(0, True, 6000000) == 0                 # 공사 취소(총액 0)


def test_row_vat_free_dict_and_series():
    assert prs._row_vat_free(G4148) == 6000000
    assert prs._row_vat_free(pd.Series(G4148)) == 6000000
    plain = dict(G4148, **{'계약금 부담자': '임대인 · 유미숙 · 6,000,000 · 현금'})
    assert prs._row_vat_free(plain) == 0
    assert prs._row_vat_free(dict(G4148, 부가세=False)) == 0
    assert prs._row_vat_free({'총액 1': 100}) == 0


def test_backend_calculations_use_vat_free():
    df = prs._apply_backend_calculations(pd.DataFrame([G4148, dict(G4148, **{'프로젝트 코드': 'X1',
                                                                            '계약금 부담자': '-', '잔금 부담자': '-'})]))
    assert df.loc[0, '총액 2'] == 17046200 and df.loc[0, '미수금'] == 17046200
    assert df.loc[1, '총액 2'] == 17646200      # 일반 공사는 그대로


def test_formula_standard_matches_new_project_formula():
    assert ps.total2_formula(4149, 'R', 'S', 0) == STD_FX
    fx = ps.total2_formula(4149, 'R', 'S', 6000000)
    assert fx.startswith('=IF(S4149=TRUE, MIN(R4149,6000000)+MAX(R4149-6000000,0)') and fx.endswith(', R4149)')


def test_formula_semantics_match_python():
    """계산식 문자열을 그대로 계산해 서버 계산과 같은지 (FLOOR·MOD·MIN·MAX·IF 만 쓰는 식)."""
    def evaluate(fx, r, s):
        expr = fx[1:].replace('R4149', str(r)).replace('S4149', 'True' if s else 'False')
        expr = expr.replace('=TRUE', '==True').replace('MOD(', 'MOD(')
        env = {'IF': lambda c, a, b: a if c else b, 'FLOOR': lambda x, n: (x // n) * n,
               'MOD': lambda a, b: a % b, 'MIN': min, 'MAX': max, 'True': True, 'False': False}
        return eval(expr.replace(')=1', ')==1').replace(')=9', ')==9'), env)
    for r in (16042000, 16042001, 12345678, 6000000, 3000000, 0):
        for vf in (0, 6000000):
            assert evaluate(ps.total2_formula(4149, 'R', 'S', vf), r, True) == prs._calculate_total2(r, True, vf), (r, vf)
        assert evaluate(ps.total2_formula(4149, 'R', 'S', 6000000), r, False) == r


class _Mgr:
    def __init__(self, formula, row_map):
        self.formula = formula
        self.row_map = row_map
        self.writes = []

    def get_cell_formula(self, sid, sn, a):
        return self.formula

    def update_cell_value(self, sid, sn, a, v):
        self.writes.append((a, v))
        self.formula = v
        return True

    def get_row_map(self, sid, sn, row):
        return self.row_map


def _plan(vat_free):
    return {'계약금': {'role': '임대인', 'name': '유미숙', 'amount': 6000000, 'method': '현금', 'vat_free': vat_free},
            '잔금': {'role': '임차인', 'name': 'JKD', 'amount': 11046200, 'method': '계산서', 'vat_free': False}}


def test_sync_formula_writes_custom_then_restores_standard():
    rm = {'R': 16042000, 'S': True, 'U': 1000000, 'V': 0, 'W': 0}
    m = _Mgr(STD_FX, rm)
    out = ps.sync_total2_formula(m, 'sid', 'sn', 4149, F2L, _plan(True))
    assert m.writes and 'MIN(R4149,6000000)' in m.writes[0][1]
    assert out == {'총액 2': 17046200.0, '미수금': 16046200.0}
    assert ps.sync_total2_formula(m, 'sid', 'sn', 4149, F2L, _plan(True)) == {}    # 이미 맞음 → 안 씀
    out2 = ps.sync_total2_formula(m, 'sid', 'sn', 4149, F2L, _plan(False))         # 표시 뺌 → 표준식
    assert m.writes[-1][1] == STD_FX and out2['총액 2'] == 17646200.0


def test_sync_formula_leaves_ordinary_rows_alone():
    m = _Mgr(STD_FX, {})
    assert ps.sync_total2_formula(m, 'sid', 'sn', 4149, F2L, _plan(False)) == {} and not m.writes
    m2 = _Mgr(17646200, {})            # 계산식 없이 값만 있는 옛 행도 부가세 없는 몫이 없으면 그대로
    assert ps.sync_total2_formula(m2, 'sid', 'sn', 4149, F2L, _plan(False)) == {} and not m2.writes


def test_share_check_no_warning_once_total2_follows():
    rec = dict(G4148, **{'총액 2': 17046200})
    assert ps.share_check(list(ps.stage_payers(rec).values()), rec) == ''
    assert ps.vat_ratio(rec) == 1.1     # 총액 비율(1.0626…)이 아니라 부가세 10%


def test_split_issued_summary_tenant_supply_not_distorted():
    from dashboard.blueprints import slack_bot as sb
    rec = dict(G4148, **{'총액 2': 17046200, '계산서_메모': ''})
    with patch('dashboard.services.project_service.get_project_records', return_value=[rec]):
        one = sb._project_issued_invoice('G4148-JW', stage='잔금')
    assert one['remaining_supply'] == 10042000 and one['total2'] == 11046200


def test_invoice_warning_uses_vat_10pct_for_vat_free_projects():
    from dashboard.blueprints import slack_bot as sb
    rec = dict(G4148, **{'총액 2': 17046200, '계산서_메모': ''})
    with patch('dashboard.services.project_service.get_project_records', return_value=[rec]):
        # 임차인 공급가 10,042,000 → 발행 11,046,200 = 부담액 (총액 비율이면 10,670,… 로 틀어짐)
        assert sb._invoice_amount_warning('G4148-JW', '10042000', 'sep', stage='잔금') == ''
        assert '부담액' in sb._invoice_amount_warning('G4148-JW', '10500000', 'sep', stage='잔금')


def test_notify_total2_change_dm_and_skip_self():
    from dashboard.services import project_slack_notifier as pn
    fake = MagicMock()
    fake.conversations_open.return_value = {'channel': {'id': 'D1'}}
    with patch.dict('os.environ', {'SLACK_BOT_TOKEN': 'x'}), patch('slack_sdk.WebClient', return_value=fake):
        assert pn.notify_total2_payer_change('G4148-JW', 17646200, 17046200, _plan(True),
                                             editor_email='jw@itg-aircon.com', latest_data=G4148) is True
        text = fake.chat_postMessage.call_args.kwargs['text']
        assert '17,646,200원 → 17,046,200원' in text and "유미숙 6,000,000원을 '현금 · 부가세 없음'" in text
        fake.conversations_open.reset_mock()
        assert pn.notify_total2_payer_change('G4148-JW', 17646200, 17046200, _plan(True),
                                             editor_email='sb@itg-aircon.com') is True       # 승인자 본인 편집
        opened = [c.kwargs['users'] for c in fake.conversations_open.call_args_list]
        assert opened and 'U0BHC2JV7U5' not in opened                                      # → 본인 DM 만 빠짐
        assert pn.notify_total2_payer_change('G4148-JW', 17046200, 17046200, _plan(True)) is False  # 그대로
