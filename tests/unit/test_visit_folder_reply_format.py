# -*- coding: utf-8 -*-
"""방문 폴더 생성 답글 = 사진 저장 답글과 같은 형식(탐색기 경로 + 폴더 ID), 드라이브 웹 링크 없음 (2026-10-13)."""
import sys
sys.path.insert(0, '.')

import dashboard.blueprints.slack_bot as sb

NAME = '(TH) 인천 미추홀구 학익소로63번길 8 청인빌딩 6층 인천지방변호사회 26.10.13'
FID = '1YYj-9jYYl5YKT7cqmbLf72KyS0wTxi12'


def test_created_text_uses_explorer_path_and_folder_id(monkeypatch):
    base = r'G:\공유 드라이브\견적서 모음'
    monkeypatch.setenv('GOOGLE_DRIVE_WINDOWS_BASE_PATH', base)
    t = sb._visit_folder_created_text(NAME, FID)
    assert '방문 폴더를 만들었습니다' in t
    assert '```' + base + '\\' + NAME + '```' in t
    assert f'```{FID}```' in t
    assert 'drive.google.com' not in t and '<http' not in t


def test_location_lines_same_block_as_photo_reply(monkeypatch):
    monkeypatch.setenv('GOOGLE_DRIVE_WINDOWS_BASE_PATH', r'G:\X')
    lines = sb._visit_folder_location_lines(NAME, FID)
    assert sb._visit_folder_created_text(NAME, FID).count(lines) == 1
    assert lines.startswith('\n💻 *탐색기 경로* :') and ':id: *폴더 ID* (새 프로젝트 등록용) :' in lines


def test_no_path_line_when_base_unset(monkeypatch):
    monkeypatch.delenv('GOOGLE_DRIVE_WINDOWS_BASE_PATH', raising=False)
    lines = sb._visit_folder_location_lines(NAME, FID)
    assert '탐색기 경로' not in lines and f'```{FID}```' in lines


class _FakeRedis:
    def __init__(self):
        self.h, self.ttl = {}, {}

    def exists(self, k):
        return int(k in self.h)

    def hset(self, k, mapping):
        self.h.setdefault(k, {}).update(mapping)

    def expire(self, k, s):
        self.ttl[k] = s


class _RC:
    def __init__(self, r):
        self.redis = r


def test_folder_notice_becomes_photo_reply_anchor(monkeypatch):
    """폴더 안내 답글을 사진 저장 답글 자리(0장)로 등록 → 사진 올리면 이 답글이 'N장 저장'으로 바뀜."""
    fake = _FakeRedis()
    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client', lambda: _RC(fake))
    assert sb._register_photo_reply_anchor('C1', '111.1', '222.2') is True
    assert fake.h['visit_photo_reply:C1:111.1'] == {'ts': '222.2', 'count': '0'}
    assert fake.ttl['visit_photo_reply:C1:111.1'] == 60 * 60 * 24 * 30


def test_anchor_does_not_override_existing_photo_reply(monkeypatch):
    fake = _FakeRedis()
    fake.h['visit_photo_reply:C1:111.1'] = {'ts': '999.9', 'count': '7'}
    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client', lambda: _RC(fake))
    assert sb._register_photo_reply_anchor('C1', '111.1', '222.2') is False
    assert fake.h['visit_photo_reply:C1:111.1'] == {'ts': '999.9', 'count': '7'}
