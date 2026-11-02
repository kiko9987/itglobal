# -*- coding: utf-8 -*-
"""권진현 = JH (2026-11-02 입사, jh@ 메일 인수) — 공사 코드 접미사·역매핑·슬랙 이니셜."""
import json
import sys

sys.path.insert(0, '.')

from dashboard.utils.project_code_generator import generate_project_code, parse_project_code


def test_jh_code_roundtrip():
    code = generate_project_code(4300, '글로벌그룹', '권진현')
    assert code == 'R4299-JH'
    assert parse_project_code(code)['manager'] == '권진현'


def test_jh_in_config_and_fallback():
    cfg = json.load(open('dashboard/project_config.json', encoding='utf-8'))
    assert cfg['owner_suffix_map']['권진현'] == 'JH'
    assert '김지훈' not in cfg['owner_suffix_map']   # 같은 JH 를 두 사람이 쓰면 역매핑이 엇갈림
    from dashboard.blueprints.slack_helpers import _SALES_INITIALS_FALLBACK
    assert _SALES_INITIALS_FALLBACK['권진현'] == 'JH'
    assert '김지훈' not in _SALES_INITIALS_FALLBACK
