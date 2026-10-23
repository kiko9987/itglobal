# -*- coding: utf-8 -*-
"""/공사확정 유입 구분 선택지 = 실제 플랫폼 (2026-10-23: '온라인'으로만 고를 수 있어 홈페이지·전화가 묻혔음)."""
import sys
sys.path.insert(0, '.')

import dashboard.blueprints.slack_bot as sb


class _FakeClient:
    def __init__(self):
        self.view = None

    def views_open(self, trigger_id, view):
        self.view = view


def _source_values():
    c = _FakeClient()
    sb._open_project_modal(c, 'T', 'C', 'U')
    block = next(b for b in c.view['blocks'] if b.get('block_id') == 'source')
    return [o['value'] for o in block['element']['options']]


def test_no_online_and_real_platforms():
    vals = _source_values()
    assert '온라인' not in vals
    for v in ('거래처', '홈페이지', '전화', '당근', '카카오톡', '채널톡', '큐플레이스', '숨고', '메일', '소개', '중고판매', '기타'):
        assert v in vals
    assert len(vals) == len(set(vals))
