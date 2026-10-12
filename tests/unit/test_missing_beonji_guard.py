# -*- coding: utf-8 -*-
"""번지 없는 도로명은 검증 생략·원문 유지 + [주소 확인 필요] (L-04214).

'구로구 부일로15길 3층 pc방'은 번지가 없는데 카카오가 '3층'의 3을 번지로 넘겨짚어
'부일로15길 3'(지어낸 번지)을 verified 로 반환 → 오방문 위험. 도로명 바로 뒤가 층/호면
번지 누락으로 보고 raw 로 내보낸다. 번지가 있는 주소는 영향 없음.
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar


@pytest.mark.parametrize('text, expected', [
    ('구로구 부일로15길 3층 pc방', True),
    ('구로구 부일로15길 2~3층', True),
    ('강남구 봉은사로 101호', True),
    ('강남구 테헤란로 152 3층', False),
    ('상도로13길 4 3층', False),
    ('영등포구 국제금융로2길 36 8~9층', False),
    ('부천로431번길 16 2층', False),
    ('강남구 학동로 지하 102', False),
    ('은평구 갈현로3나길 23 102동 501호', False),
])
def test_road_missing_beonji(text, expected):
    assert ar._road_missing_beonji(text) is expected


def test_resolve_returns_raw_without_inventing_beonji(monkeypatch):
    # 카카오가 번지를 지어내도 호출 자체를 안 함 — raw 원문 유지
    monkeypatch.setattr(ar, 'verify_address',
                        lambda *a, **k: ('구로구 부일로15길 3 3층', 'verified'))
    addr, lv = ar.resolve_address('구로구 부일로15길 3층 pc방', '구로구 부일로15길 3층', 'level3')
    assert lv == 'raw'
    assert addr == '구로구 부일로15길 3층 pc방'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
