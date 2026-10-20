# -*- coding: utf-8 -*-
"""방문일 미래 시 사진 첨부 자동완료 skip 게이트 회귀 테스트 (2026-08-05).

L-03575 사고: 08-06 방문 거래처 건에 도면(참고)을 등록 시 첨부했는데, 사진 첨부
자동완료가 방문일과 무관하게 발동해 '방문 완료' 처리됨. 방문일이 미래면 첨부는
참고자료로 보고 완료 skip (사진 저장은 유지). 파싱 실패·빈값은 기존 동작(허용).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.blueprints.slack_bot as sb


def test_future_date_blocks():
    """먼 미래 방문일 → True (자동완료 skip)."""
    assert sb._visit_not_yet_reached('2999-01-01') is True


def test_future_date_range_uses_start():
    """범위 방문일도 시작일 기준 — 미래면 True."""
    assert sb._visit_not_yet_reached('2999-01-01~2999-01-03') is True


def test_past_date_allows():
    """지난 방문일 → False (자동완료 허용)."""
    assert sb._visit_not_yet_reached('2000-01-01') is False


def test_empty_or_garbage_allows():
    """빈값·파싱불가 → False (기존 동작 유지, 현장 사진 UX 보존)."""
    assert sb._visit_not_yet_reached('') is False
    assert sb._visit_not_yet_reached(None) is False
    assert sb._visit_not_yet_reached('미정') is False


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))


# ── 현장 사진 판별 게이트 (2026-10-20 L-04289) ─────────────────────────────
# 방문 당일 올린 도면(PNG 캡처·PDF·DWG)은 날짜 게이트를 통과해 오완료 → 파일 형식으로 보완.

@pytest.mark.parametrize('files,expected', [
    ([{'name': 'IMG_9761.jpg', 'filetype': 'jpg', 'mimetype': 'image/jpeg'}], True),
    ([{'name': 'KakaoTalk_20261001.jpg', 'filetype': 'jpg', 'mimetype': 'image/jpeg'}], True),
    ([{'name': 'IMG_0001.HEIC', 'filetype': 'heic', 'mimetype': 'image/heic'}], True),
    ([{'name': 'image.png', 'filetype': 'png', 'mimetype': 'image/png'}], False),   # L-04289 캡처
    ([{'name': '1013 TOTAL DWG SET.pdf', 'filetype': 'pdf', 'mimetype': 'application/pdf'}], False),
    ([{'name': '방화동.dwg', 'filetype': 'binary', 'mimetype': 'application/octet-stream'}], False),
    ([{'name': 'image.png', 'filetype': 'png'},
      {'name': 'IMG_1.jpg', 'filetype': 'jpg'}], True),                               # 섞이면 사진 있음
    ([{'name': 'x', 'mimetype': 'image/jpeg'}], True),                                 # filetype 없으면 mime
    ([{'name': 'IMG_2.JPG'}], True),                                                   # 형식 정보 없으면 확장자
    ([{'name': 'plan.png'}], False),
    ([], False),
    (None, False),
])
def test_has_site_photo(files, expected):
    assert sb._has_site_photo(files) is expected


def test_reference_only_batch_skips_auto_complete(monkeypatch):
    """도면(PNG)만 올라온 배치는 _auto_complete_visit_from_photo 를 호출하지 않는다."""
    called = []
    monkeypatch.setattr(sb, '_auto_complete_visit_from_photo',
                        lambda *a, **k: called.append(1))
    # 6단계 분기만 검증 — 실제 핸들러의 조건식과 동일
    files = [{'name': 'image.png', 'filetype': 'png', 'mimetype': 'image/png'}]
    if sb._has_site_photo(files):
        sb._auto_complete_visit_from_photo(None)
    assert called == []
