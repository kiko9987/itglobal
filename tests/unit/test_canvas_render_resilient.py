# -*- coding: utf-8 -*-
"""캔버스1 렌더 회복탄력성 회귀 테스트 (2026-08-06).

_render_item 한 건이 예외를 던져도 build_canvas_markdown 전체가 실패해 캔버스가
통째로 동결(신규 방문 전부 누락)되면 안 됨. 실패 건은 조용히 빼지 말고 ⚠️
placeholder 로 노출. 캔버스1 방문 누락은 업무상 치명적이라 도입.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.visit_canvas_sync as vcs


def test_build_canvas_resilient_to_render_error(monkeypatch):
    good = {'리드 No': 'L-0001', '플랫폼': '전화', '방문 예정일': '2099-01-01',
            '고객명': '굿', '고객 연락처': '010-0000-0000', '방문 주소': '서울', '상담 내용': 'ok'}
    bad = {'리드 No': 'L-BAD', '플랫폼': '전화', '방문 예정일': '2099-01-02', '고객명': '배드'}
    monkeypatch.setattr(vcs, '_fetch_visit_leads', lambda: [good, bad])
    monkeypatch.setattr(vcs, '_get_initial_map', lambda: {})

    def fake_render(lead, im):
        if lead.get('리드 No') == 'L-BAD':
            raise ValueError('boom')
        return f"OK {lead.get('리드 No')}"
    monkeypatch.setattr(vcs, '_render_item', fake_render)

    md = vcs.build_canvas_markdown()  # 예외 없이 완주해야 함
    assert 'OK L-0001' in md                      # 정상 건 렌더됨
    assert 'L-BAD' in md and '렌더 오류' in md     # 실패 건은 placeholder 로 노출(누락 X)


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
