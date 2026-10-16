# -*- coding: utf-8 -*-
"""취소된 공사 편집 게이트 — 관리자만 수정, '공사 취소' 표기는 편집으로 못 지움 (2026-10-16)."""
import sys

sys.path.insert(0, '.')

from dashboard.blueprints.projects import _cancelled_edit_gate

F2I = {'수금 관련 특이사항': 2, '시공자': 1}


def _row(notes):
    return ['R1-X', '홍길동', notes]


def test_active_project_anyone_with_edit_right():
    assert _cancelled_edit_gate(_row(''), F2I, {'시공자': 'A'}, 'editor') is None


def test_cancelled_blocks_editor():
    msg, status = _cancelled_edit_gate(_row('공사 취소 (사유: 고객 변심)'), F2I, {'시공자': 'A'}, 'editor')
    assert status == 403 and '관리자' in msg


def test_cancelled_allows_admin_and_super_admin():
    for perm in ('admin', 'super_admin'):
        data = {'시공자': 'A', '수금 관련 특이사항': '공사 취소 (사유: 고객 변심)'}
        assert _cancelled_edit_gate(_row('공사 취소 (사유: 고객 변심)'), F2I, data, perm) is None


def test_admin_cannot_erase_cancel_mark():
    msg, status = _cancelled_edit_gate(_row('공사취소'), F2I, {'수금 관련 특이사항': ''}, 'admin')
    assert status == 400 and '공사 재개' in msg


def test_admin_edit_without_notes_field_ok():
    assert _cancelled_edit_gate(_row('공사 취소'), F2I, {'시공자': 'B'}, 'admin') is None
