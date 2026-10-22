# -*- coding: utf-8 -*-
"""정합성 점검 — 시트 원본 칸 이상(수식 오류·칸 이동 의심), 새로 생긴 것만 알림 (2026-10-19 칸 이동 사고).

2026-10-22: 칸은 필드명으로 받고(열 위치는 헤더 지도), 기억 키도 필드명 — 열을 끼워 넣어도 기존
이상 칸이 '새 이상'으로 쏟아지지 않는다. 옛 v1 키(열 글자)는 v2 로 옮긴다.
"""
from dashboard.services import invariant_checks as ic

OK = {'총액 2': 13026200, '미수금': 13026200, '수금 확인': False, '제품대': 0, '도급비': 0, '자재비': 0,
      '기타비': 0, '순익': 11842000, '마진율': 1}


def test_detects_drag_incident_cells():
    rows = [
        (4149, 'G4148-JW', dict(OK, 기타비='임대인 · 유미숙 · 6,000,000 · 현금', 순익='#REF! (bad)')),
        (4152, 'G4151-JW', dict(OK, **{'수금 확인': '임대인 · 남상철 · 6,050,000 · 계산서'})),
        (4153, 'G4152-JW', dict(OK)),
        (4154, 'G4153-JW', dict(OK, **{'기타비': '-', '수금 확인': ''})),
    ]
    found = ic.sheet_cell_anomalies(rows)
    msgs = [m for _, _, m in found]
    assert len(found) == 3
    assert any('기타비(AH4149) 금액 칸에 글자' in m for m in msgs)
    assert any('순익(AI4149) 수식 오류 #REF!' in m for m in msgs)
    assert any('수금 확인(AD4152) 체크박스 칸에 값' in m for m in msgs)


def test_message_letter_follows_layout_key_does_not():
    # 열이 한 칸 밀린 배치 — 문구의 열 글자는 따라가고, 기억 키(필드명)는 그대로
    rows = [(4149, 'G4148-JW', dict(OK, 순익='#REF! (bad)'))]
    before = ic.sheet_cell_anomalies(rows, f2l={'순익': 'AI'})
    after = ic.sheet_cell_anomalies(rows, f2l={'순익': 'AJ'})
    assert '순익(AI4149)' in before[0][2] and '순익(AJ4149)' in after[0][2]
    assert before[0][0] == after[0][0] == 'cell:순익:4149:#REF! (bad)'


def test_only_new_anomalies_alert():
    rows = [(619, 'G0618-YG', dict(OK, 마진율='#DIV/0! (Function DIVIDE)')),
            (4149, 'G4148-JW', dict(OK, 순익='#REF! (bad)'))]
    an = ic.sheet_cell_anomalies(rows)
    known = {an[0][0]}                                  # 옛 #DIV/0! 은 기준에 있음
    v = ic.check_sheet_cells({'cell_anomalies': an, 'known_cell_anomalies': known})
    assert len(v) == 1 and v[0]['title'] == 'G4148-JW' and v[0]['check'] == 'sheet_cell'
    assert ic.check_sheet_cells({'cell_anomalies': [], 'known_cell_anomalies': set()}) == []


def test_migrate_known_v1_keys():
    migrated = ic._migrate_known_v1({'cell:AJ619:#DIV/0! (Function DIVIDE)', 'cell:AH4149:-', '__seeded__',
                                     'cell:ZZ1:x'})
    assert migrated == {'cell:마진율:619:#DIV/0! (Function DIVIDE)', 'cell:기타비:4149:-'}

# 테스트 데이터를 옛 45열 위치(계약금=U·버전=AS 등)로 짰다 — 기본 배치를 45열로 고정해 검사
#   (2026-10-22 거래처명 열 삽입 뒤에도 같은 논리를 검사. 코드는 헤더 지도를 따라감)
import pytest as _pytest  # noqa: E402
pytestmark = _pytest.mark.usefixtures('sheet_layout_45')
