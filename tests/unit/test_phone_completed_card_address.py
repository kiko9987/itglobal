# -*- coding: utf-8 -*-
"""전화 완료(회색) 카드 주소 표시 회귀 테스트 (2026-07-31 L-03476).

배경: _post_phone_lead_completed_card 가 row.to_dict()(시트 raw + _meta 미포함)를
build_inquiry_blocks 에 넘겨 addr_level='' → 카카오 성공/실패와 무관하게 항상
'주소 확인 필요' 배지 + 매니저 raw 주소를 표시. 방문 카드는 정규화본(변환 주소)을
쓰는데 온라인 완료 카드만 raw+배지라 불일치.

수정: sync_workflow_phone_leads 가 정규화 결과(addr_for_notify / _meta_address_level)를
_lead_dict 에 주입 → 온라인 리드와 동일하게 원본/변환 2줄, verified 는 배지 없음,
비verified(실패)는 배지 유지(오방문 방지). 이 테스트는 그 렌더링 계약을 고정한다.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_sync import build_inquiry_blocks

BADGE = '주소 확인 필요'


def _section_text(lead: dict) -> str:
    blocks, _ = build_inquiry_blocks(lead, 'L-09999', source='전화')
    for b in blocks:
        if b.get('type') == 'section':
            return (b.get('text') or {}).get('text', '')
    return ''


def _base(**over) -> dict:
    d = {
        '상담 시간': '2026.07.31. 14:17',
        '고객명': '고객',
        '고객 연락처': '010-3022-4068',
        '이메일': '-',
        '상담 내용': '천장형 냉난방기 견적',
    }
    d.update(over)
    return d


class TestPhoneCompletedCardAddress:
    def test_verified_diff_two_lines_no_badge(self):
        """카카오 verified 정정 성공(원본≠변환) → 원본/변환 2줄, 배지 없음 (L-03476)."""
        t = _section_text(_base(
            **{'방문 주소': '인천 부평구 부평문화로 50-1 3층',
               '_meta_address_raw': '부평구 부평문화로 50-1 3층',
               '_meta_address_level': 'verified'}))
        assert '원본 주소* : 부평구 부평문화로 50-1 3층' in t
        assert '변환 주소* : 인천 부평구 부평문화로 50-1 3층' in t
        assert BADGE not in t

    def test_verified_same_single_line_no_badge(self):
        """verified & 원본==변환 → 단일 방문 주소 라인, 배지 없음."""
        t = _section_text(_base(
            **{'방문 주소': '서울 강남구 테헤란로 1',
               '_meta_address_raw': '서울 강남구 테헤란로 1',
               '_meta_address_level': 'verified'}))
        assert '방문 주소* : 서울 강남구 테헤란로 1' in t
        assert '원본 주소' not in t
        assert BADGE not in t

    def test_non_verified_keeps_badge(self):
        """비verified(실패) → raw + '주소 확인 필요' 배지 유지(오방문 방지)."""
        t = _section_text(_base(
            **{'방문 주소': '부평구 부평문화로 오십의일',
               '_meta_address_raw': '부평구 부평문화로 오십의일',
               '_meta_address_level': ''}))
        assert BADGE in t

    def test_legacy_no_meta_still_badges(self):
        """_meta 미주입(구 동작·주입 실패 fallback) → 배지 (안전측)."""
        t = _section_text(_base(**{'방문 주소': '부평구 부평문화로 50-1 3층'}))
        assert BADGE in t


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
