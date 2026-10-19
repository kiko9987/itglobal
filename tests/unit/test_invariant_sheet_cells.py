# -*- coding: utf-8 -*-
"""정합성 점검 — 시트 원본 칸 이상(수식 오류·칸 이동 의심), 새로 생긴 것만 알림 (2026-10-19 칸 이동 사고)."""
from dashboard.services import invariant_checks as ic

OK = {'T': 13026200, 'X': 13026200, 'AD': False, 'AE': 0, 'AF': 0, 'AG': 0, 'AH': 0, 'AI': 11842000, 'AJ': 1}


def test_detects_drag_incident_cells():
    rows = [
        (4149, 'G4148-JW', dict(OK, AH='임대인 · 유미숙 · 6,000,000 · 현금', AI='#REF! (bad)')),
        (4152, 'G4151-JW', dict(OK, AD='임대인 · 남상철 · 6,050,000 · 계산서')),
        (4153, 'G4152-JW', dict(OK)),
        (4154, 'G4153-JW', dict(OK, AH='-', AD='')),
    ]
    found = ic.sheet_cell_anomalies(rows)
    msgs = [m for _, _, m in found]
    assert len(found) == 3
    assert any('기타비(AH4149) 금액 칸에 글자' in m for m in msgs)
    assert any('순익(AI4149) 수식 오류 #REF!' in m for m in msgs)
    assert any('수금 확인(AD4152) 체크박스 칸에 값' in m for m in msgs)


def test_only_new_anomalies_alert():
    rows = [(619, 'G0618-YG', dict(OK, AJ='#DIV/0! (Function DIVIDE)')),
            (4149, 'G4148-JW', dict(OK, AI='#REF! (bad)'))]
    an = ic.sheet_cell_anomalies(rows)
    known = {an[0][0]}                                  # 옛 #DIV/0! 은 기준에 있음
    v = ic.check_sheet_cells({'cell_anomalies': an, 'known_cell_anomalies': known})
    assert len(v) == 1 and v[0]['title'] == 'G4148-JW' and v[0]['check'] == 'sheet_cell'
    assert ic.check_sheet_cells({'cell_anomalies': [], 'known_cell_anomalies': set()}) == []
