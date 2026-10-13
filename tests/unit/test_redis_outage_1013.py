# -*- coding: utf-8 -*-
"""2026-10-13 Redis 3.5h 장애 회귀 테스트.

WSL 자동 업데이트로 Redis 가 꺼진 동안:
- fallback 메모리 캐시가 pickle bytes 를 그대로 돌려줘 리드 조회가 전부 실패
- 조회 실패(None → {}) 상태로 방문 완료 마커가 상담 내용을 덮어씀 (12건)
- 사진 폴더가 '(미정) 주소 미상 26.10.13' 한 곳으로 섞임 (9건)
- Redis 중복 방지 표시가 전부 실패해 방문 완료가 배치마다 반복
"""
import pickle
import sys

import pandas as pd

sys.path.insert(0, '.')

import dashboard.blueprints.slack_bot as sb
from dashboard.utils.fallback_cache import FallbackMemoryCache
from dashboard.utils.smart_cache_manager import SimpleCache as SmartCacheManager


def _df():
    return pd.DataFrame([{'리드 No': 'L-04211', '상담 내용': '[10.12 10:08 YG · 방문 예약] 현재 공실',
                          '상태': '방문 예약', '방문 예정일': '2026-10-13', '플랫폼': '큐플레이스',
                          '영업 담당자': '빈승정', '방문 주소': '고양 덕양구 꽃내음3길 46', '폴더 ID': ''}])


def test_fallback_cache_unpickles_bytes_stored_by_redis_client():
    mgr = SmartCacheManager.__new__(SmartCacheManager)
    mgr._use_fallback = True
    mgr._fallback_cache = FallbackMemoryCache()
    mgr._record_hit = mgr._record_miss = lambda _k: None
    # RedisClient 자체 fallback 이 같은 메모리 캐시에 pickle bytes 로 넣어 둔 상황
    mgr._fallback_cache.set('cache:online_leads_sheet_data', pickle.dumps(_df()))
    got = mgr.get('online_leads_sheet_data')
    assert isinstance(got, pd.DataFrame) and got.iloc[0]['리드 No'] == 'L-04211'


def test_fallback_cache_bad_bytes_is_miss():
    mgr = SmartCacheManager.__new__(SmartCacheManager)
    mgr._use_fallback = True
    mgr._fallback_cache = FallbackMemoryCache()
    mgr._record_hit = mgr._record_miss = lambda _k: None
    mgr._fallback_cache.set('cache:x', b'not-a-pickle')
    assert mgr.get('x') is None


def test_find_lead_retries_sheet_when_cache_returns_bytes(monkeypatch):
    calls = []

    def fake_load(force_refresh=False):
        calls.append(force_refresh)
        return _df() if force_refresh else pickle.dumps(_df())

    monkeypatch.setattr('dashboard.services.lead_service.load_leads_data', fake_load)
    lead = sb._find_lead_by_no('L-04211')
    assert lead and lead['상담 내용'].startswith('[10.12 10:08 YG')
    assert calls == [False, True]


def test_find_lead_retries_sheet_when_cache_raises(monkeypatch):
    def fake_load(force_refresh=False):
        if not force_refresh:
            raise ConnectionError('Error 10061 connecting to localhost:6379')
        return _df()

    monkeypatch.setattr('dashboard.services.lead_service.load_leads_data', fake_load)
    assert sb._find_lead_by_no('L-04211')['상태'] == '방문 예약'


def test_find_lead_none_when_both_fail(monkeypatch):
    def fake_load(force_refresh=False):
        raise ConnectionError('down')

    monkeypatch.setattr('dashboard.services.lead_service.load_leads_data', fake_load)
    assert sb._find_lead_by_no('L-04211') is None


def test_visit_complete_marker_not_written_when_lookup_fails(monkeypatch):
    writes = []
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda _no: None)
    monkeypatch.setattr(sb, '_update_lead_dispatch', lambda no, upd: writes.append((no, upd)))
    sb._mark_visit_complete_on_sheet('L-04211', initial='SJ', dt_str='10.13 09:25')
    assert writes == []  # 예전엔 {'상담 내용': '[10.13 09:25 SJ · 방문 완료]'} 로 덮어씀


def test_visit_complete_marker_appends_to_existing(monkeypatch):
    writes = []
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda _no: _df().iloc[0].to_dict())
    monkeypatch.setattr(sb, '_update_lead_dispatch', lambda no, upd: writes.append((no, upd)))
    sb._mark_visit_complete_on_sheet('L-04211', initial='SJ', dt_str='10.13 09:25')
    assert writes[0][1]['상담 내용'] == ('[10.12 10:08 YG · 방문 예약] 현재 공실 ─── '
                                       '[10.13 09:25 SJ · 방문 완료]')


def test_visit_folder_not_created_when_lookup_fails(monkeypatch):
    created = []
    monkeypatch.setenv('GOOGLE_DRIVE_VISIT_FOLDER_ID', 'PARENT')
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda _no: None)
    monkeypatch.setattr('dashboard.utils.google_drive.find_or_create_folder',
                        lambda name, parent: created.append(name) or {'id': 'X'})

    class _DownRC:
        @property
        def redis(self):
            raise ConnectionError('down')

    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client', lambda: _DownRC())
    assert sb._ensure_visit_folder(None, 'L-04211') is None
    assert created == []  # 예전엔 '(미정) 주소 미상 26.10.13' 공용 폴더 생성·재사용


def test_local_flag_blocks_duplicates_when_redis_down(monkeypatch):
    class _DownRC:
        @property
        def redis(self):
            raise ConnectionError('down')

    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client', lambda: _DownRC())
    sb._LOCAL_FLAGS.clear()
    assert sb._try_acquire_action_lock('L-04220', 'complete') is True
    assert sb._try_acquire_action_lock('L-04220', 'complete') is False  # 4초 뒤 두 번째 클릭
    assert sb._local_flag_nx('visit_auto_completed:L-04211', 60) is True
    assert sb._local_flag_nx('visit_auto_completed:L-04211', 60) is False
    assert sb._local_flag_get('visit_auto_completed:L-04211') is True
    sb._LOCAL_FLAGS.clear()


def test_auto_complete_skips_when_lookup_fails(monkeypatch):
    class _DownRC:
        @property
        def redis(self):
            raise ConnectionError('down')

    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client', lambda: _DownRC())
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda _no: None)
    hooks = []
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: hooks.append(a))
    sb._LOCAL_FLAGS.clear()
    sb._auto_complete_visit_from_photo(None, 'C1', '1.1', 'L-04230', 'U1')
    assert hooks == []  # 방문일을 모르면 도면 첨부를 방문 완료로 오인하지 않음
