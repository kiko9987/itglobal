# -*- coding: utf-8 -*-
"""2026-10-28 3차 — 직원 명단 한곳(staff_directory) + 담당자·사업자 변경 시 코드 번호 유지."""
import json
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.services.staff_directory as sd
from dashboard.utils import project_code_generator as pcg


@pytest.fixture
def maps(monkeypatch):
    monkeypatch.setattr(sd, 'suffix_map', lambda: {'박정우': 'JW', '주영민': 'YM', '이미영': 'MY', '김호중': 'HJ'})
    monkeypatch.setattr(sd, 'company_prefix_map', lambda: {'글로벌': 'G', '글로벌그룹': 'R', '플렌트': 'P'})


def test_regenerate_keeps_number(maps):
    # 담당자만 바꾸면 접미사만, 번호(4153)는 그대로 — 예전엔 시트 행번호로 다시 매김
    assert pcg.regenerate_project_code('G4153-JW', '', '주영민') == ('G4153-YM', '')
    # 사업자만 바꾸면 앞 글자만
    assert pcg.regenerate_project_code('G4153-JW', '글로벌그룹', '') == ('R4153-JW', '')
    # 둘 다 안 바뀜
    assert pcg.regenerate_project_code('G4153-JW', '', '') == ('G4153-JW', '')


def test_regenerate_unknown_owner_blocked(maps):
    code, err = pcg.regenerate_project_code('G4153-JW', '', '홍길동')
    assert code == 'G4153-JW' and '접미사' in err


def test_parse_uses_directory(maps):
    p = pcg.parse_project_code('R4231-HJ')
    assert p['company'] == '글로벌그룹' and p['manager'] == '김호중' and p['number'] == 4231


def test_update_code_only_when_field_really_changed(maps):
    from dashboard.blueprints.projects import _update_project_code_if_needed
    f2i = {'프로젝트 코드': 0, '사업자': 1, '담당자': 2}
    cur = ['G4153-JW', '글로벌', '박정우']
    # 안 바꾼 담당자 값이 같이 넘어와도 코드 그대로
    assert _update_project_code_if_needed({'담당자': '박정우', '사업자': '글로벌'}, list(cur), 4154, None,
                                          'G4153-JW', f2i) == ('G4153-JW', None)
    code, change = _update_project_code_if_needed({'담당자': '주영민'}, list(cur), 9999, None, 'G4153-JW', f2i)
    assert code == 'G4153-YM' and change['old_value'] == 'G4153-JW'
    with pytest.raises(ValueError):
        _update_project_code_if_needed({'담당자': '홍길동'}, list(cur), 4154, None, 'G4153-JW', f2i)


def test_active_sales_excludes_non_sales_and_inactive(monkeypatch):
    monkeypatch.setattr(sd, 'suffix_map', lambda: {'박정우': 'JW', '황샛별': 'SB', '김지훈': 'JH', '아이티': 'IT'})
    monkeypatch.setattr(sd, 'non_sales_owners', lambda: ['황샛별', '아이티'])
    monkeypatch.setattr(sd, '_users', lambda: [
        {'name': '박정우', 'is_active': True}, {'name': '황샛별', 'is_active': True},
        {'name': '김지훈', 'is_active': False, 'is_resigned': True}])
    assert sd.active_sales_managers() == ['박정우']
    assert sd.inactive_names() == ['김지훈']


def test_config_has_non_sales_and_legacy_suffix():
    cfg = json.load(open('dashboard/project_config.json', encoding='utf-8'))
    assert {'황샛별', '이미영', '아이티'} <= set(cfg['non_sales_owners'])
    assert cfg['owner_suffix_map']['김호중'] == 'HJ'
