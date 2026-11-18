# -*- coding: utf-8 -*-
"""PM 에서 내부 사업자(글로벌·글로벌그룹) 변경은 관리자만 (2026-11-18 KiKO, R4130-JSH 계기).

글로벌로 잔금 계산서를 발행한 뒤 고객이 글로벌그룹 통장에 입금하자 담당 매니저가 PM 에서 사업자를
글로벌그룹으로 바꿔 코드가 G4130-JSH → R4130-JSH 로 바뀌었고, 경영지원은 몰랐다.
"""
import sys

sys.path.insert(0, '.')

from dashboard.blueprints import projects as pj

F2I = {'프로젝트 코드': 0, '사업자': 1, '담당자': 2, '시공자': 3}
ROW = ['G4130-JSH', '글로벌', '조성헌', '아이티']


def test_editor_cannot_change_company():
    for perm in ('editor', 'viewer', ''):
        msg, status = pj._company_edit_gate(ROW, F2I, {'사업자': '글로벌그룹'}, perm)
        assert status == 403 and '관리자' in msg and '경영지원' in msg


def test_admin_can_change_company():
    for perm in ('admin', 'super_admin'):
        assert pj._company_edit_gate(ROW, F2I, {'사업자': '글로벌그룹'}, perm) is None


def test_unchanged_or_missing_company_passes():
    """기본정보 카드는 안 고친 칸도 같이 보낸다 — 값이 같으면 통과."""
    assert pj._company_edit_gate(ROW, F2I, {'사업자': '글로벌', '시공자': '김종연'}, 'editor') is None
    assert pj._company_edit_gate(ROW, F2I, {'사업자': ' 글로벌 '}, 'editor') is None
    assert pj._company_edit_gate(ROW, F2I, {'시공자': '김종연'}, 'editor') is None
    assert pj._company_edit_gate(ROW, F2I, {'사업자': ''}, 'editor') is None   # 빈 값은 코드도 안 바꾼다


def test_gate_is_wired_into_update_api():
    import inspect
    src = inspect.getsource(pj.update_project)
    assert '_company_edit_gate(' in src
    # 코드를 바꾸는 단계(_update_project_code_if_needed)보다 먼저 막아야 코드가 안 바뀐다
    assert src.index('_company_edit_gate(') < src.index('_update_project_code_if_needed(')
