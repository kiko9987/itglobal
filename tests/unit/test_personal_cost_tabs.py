"""개인 원가 탭 ↔ 메인 동기화 계획 (순수 함수) 테스트."""
from dashboard.services import personal_cost_tabs as pct
from dashboard.services.personal_cost_tabs import (
    COST_IDX, HEADERS, LAYOUT, COST, CALC, decide, plan_tab, cost_num, same_value,
    profit_formula, margin_formula,
)

MAIN_HDR = ['프로젝트 코드', '사업자', '담당자', '유입 구분', '거래처명', '사업자명', '현장 주소', '공사 구분',
            '기계 분류', '브랜드', '공사 시작', '공사 종료', '공사 내용', '도급 구분', '시공자', '발주처 담당자',
            '발주처 연락처', '발주처 이메일', '총액 1', '부가세', '총액 2', '계약금', '중도금', '잔금', '미수금',
            '계산서', '계약금 계산서', '중도금 계산서', '잔금 계산서', '수금 날짜', '수금 확인', '제품대', '도급비',
            '자재비', '기타비', '순익', '마진율', '수금 관련 특이사항', '계약금 부담자', '중도금 부담자',
            '잔금 부담자', '견적서 및 계약서 폴더 경로', '공사 확정', 'Airtable Record ID', 'Lead No', '_version']
IDX = {h: i for i, h in enumerate(MAIN_HDR)}


def mrow(code, owner='박용구', costs=(0, 0, 0, 0), **kw):
    r = [''] * len(MAIN_HDR)
    r[IDX['프로젝트 코드']], r[IDX['담당자']] = code, owner
    r[IDX['총액 1']] = kw.pop('total', 1000000)
    r[IDX['부가세']] = True
    for c, v in zip(('제품대', '도급비', '자재비', '기타비'), costs):
        r[IDX[c]] = v
    for k, v in kw.items():
        r[IDX[k]] = v
    return r


def trow(main_row, costs=None):
    vals = pct.main_row_values(main_row, IDX)
    if costs is not None:
        for i, ci in enumerate(COST_IDX):
            vals[ci] = costs[i]
    return vals


def fr(n):
    return [[profit_formula(r), margin_formula(r)] for r in range(2, n + 2)]


def test_layout_is_38_columns_with_costs_at_AB_AE():
    assert len(HEADERS) == 38
    assert [pct.col_letter(i) for i in COST_IDX] == ['AB', 'AC', 'AD', 'AE']
    assert pct.col_letter(pct.PROFIT_IDX) == 'AF' and pct.col_letter(pct.MARGIN_IDX) == 'AG'
    assert all(src in IDX for _, src in LAYOUT if src not in (COST, CALC))


def test_cost_num():
    assert cost_num('1,200,000원') == 1200000
    assert cost_num(3566955.0) == 3566955
    assert cost_num('') == 0 and cost_num(None) == 0 and cost_num(True) == 0
    assert cost_num('-200000') == -200000
    assert cost_num('메모') == 0


def test_same_value_does_not_mix_bool_and_number():
    assert not same_value(True, 1)
    assert same_value(3.0, 3)
    assert same_value(None, '')


def test_decide():
    z, a, b = [0, 0, 0, 0], [1, 2, 3, 4], [5, 6, 7, 8]
    assert decide(a, a, None) == 'same'
    assert decide(b, a, a) == 'to_main'          # 탭만 바뀜
    assert decide(a, b, a) == 'to_tab'           # PM 만 바뀜
    assert decide(b, [9, 9, 9, 9], a) == 'conflict_tab'
    assert decide(z, a, None) == 'to_tab'        # 기록 없고 탭 비었음 → 메인 원가 살림
    assert decide(a, z, None) == 'to_main'       # 기록 없고 탭에만 원가 → 메인으로


def test_plan_no_change_when_in_sync():
    m = {'G1-YG': mrow('G1-YG', costs=(1, 2, 3, 4))}
    tab = [HEADERS, trow(m['G1-YG'])]
    p = plan_tab('박용구', tab, fr(1), m, ['G1-YG'], IDX, {'G1-YG': [1, 2, 3, 4]}, {})
    assert not any(p[k] for k in ('cells', 'formulas', 'appends', 'pushes', 'deletes'))


def test_plan_tab_cost_edit_goes_to_main():
    m = {'G1-YG': mrow('G1-YG')}
    tab = [HEADERS, trow(m['G1-YG'], costs=[100, 0, 50, 0])]
    p = plan_tab('박용구', tab, fr(1), m, ['G1-YG'], IDX, {'G1-YG': [0, 0, 0, 0]}, {})
    assert p['pushes'] == [('G1-YG', [100, 0, 50, 0])]
    assert not p['cells']


def test_plan_pm_cost_edit_goes_to_tab():
    m = {'G1-YG': mrow('G1-YG', costs=(7, 0, 0, 0))}
    tab = [HEADERS, trow(m['G1-YG'], costs=[0, 0, 0, 0])]
    p = plan_tab('박용구', tab, fr(1), m, ['G1-YG'], IDX, {'G1-YG': [0, 0, 0, 0]}, {})
    assert not p['pushes']
    assert (2, COST_IDX[0], 7) in p['cells']


def test_plan_pending_row_is_skipped():
    m = {'G1-YG': mrow('G1-YG')}
    tab = [HEADERS, trow(m['G1-YG'], costs=[100, 0, 0, 0])]
    p = plan_tab('박용구', tab, fr(1), m, ['G1-YG'], IDX, {'G1-YG': [0, 0, 0, 0]}, {'G1-YG': [100, 0, 0, 0]})
    assert not p['pushes'] and not p['cells']


def test_plan_display_cell_follows_main_and_new_project_appended():
    m = {'G1-YG': mrow('G1-YG', **{'공사 내용': '천장형 2대'}), 'G2-YG': mrow('G2-YG', costs=(5, 0, 0, 0))}
    old = trow(mrow('G1-YG', **{'공사 내용': '천장형 1대'}))
    tab = [HEADERS, old]
    p = plan_tab('박용구', tab, fr(1), m, ['G1-YG', 'G2-YG'], IDX, {'G1-YG': [0, 0, 0, 0]}, {})
    assert p['cells'] == [(2, HEADERS.index('공사 내용'), '천장형 2대')]
    assert len(p['appends']) == 1 and p['appends'][0][0] == 'G2-YG'
    assert p['appends'][0][COST_IDX[0]] == 5
    assert ('G2-YG', [5, 0, 0, 0]) in p['last']


def test_plan_moved_code_deleted_only_when_costs_synced():
    m = {'G1-JW': mrow('G1-JW', owner='박정우', costs=(1, 0, 0, 0))}
    synced = trow(m['G1-JW'])
    p = plan_tab('박용구', [HEADERS, synced], fr(1), m, [], IDX, {'G1-JW': [1, 0, 0, 0]}, {})
    assert p['deletes'] == [2]
    assert p['last'] == []          # 박정우 탭 기록은 그대로
    unsynced = trow(m['G1-JW'], costs=[9, 9, 9, 9])
    p = plan_tab('박용구', [HEADERS, unsynced], fr(1), m, [], IDX, {'G1-JW': [1, 0, 0, 0]}, {})
    assert p['deletes'] == [] and p['notes']


def test_plan_missing_formula_is_restored():
    m = {'G1-YG': mrow('G1-YG')}
    tab = [HEADERS, trow(m['G1-YG'])]
    p = plan_tab('박용구', tab, [['', '']], m, ['G1-YG'], IDX, {'G1-YG': [0, 0, 0, 0]}, {})
    assert (2, pct.PROFIT_IDX, profit_formula(2)) in p['formulas']
    assert (2, pct.MARGIN_IDX, margin_formula(2)) in p['formulas']


def test_plan_duplicate_code_row_left_alone():
    m = {'G1-YG': mrow('G1-YG')}
    tab = [HEADERS, trow(m['G1-YG']), trow(m['G1-YG'], costs=[9, 0, 0, 0])]
    p = plan_tab('박용구', tab, fr(2), m, ['G1-YG'], IDX, {'G1-YG': [0, 0, 0, 0]}, {})
    assert not p['pushes'] and not p['deletes'] and p['notes']
