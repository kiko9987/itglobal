# -*- coding: utf-8 -*-
"""홈페이지 메일 sync — Gmail 접속 실패는 '이번 회차 건너뜀'(누락 아님). 한두 번은 WARNING,
연속 3회부터 ERROR, 성공하면 연속 횟수 초기화 (2026-10-14 '라벨 처리 실패' 오해 정리)."""
import logging
import sys

sys.path.insert(0, '.')

import dashboard.services.homepage_mail_sync as h


def _levels(caplog):
    return [r.levelno for r in caplog.records if r.name == 'dashboard.services.homepage_mail_sync']


def test_connect_fail_warning_then_error_then_reset(monkeypatch, caplog):
    monkeypatch.setattr(h, '_connect_fail_streak', 0)
    monkeypatch.setattr(h, '_get_gmail_service', lambda: object())

    def _down(service, name):
        raise TimeoutError('[WinError 10060] 연결된 구성원으로부터 응답이 없어')

    monkeypatch.setattr(h, '_get_or_create_label', _down)
    with caplog.at_level(logging.INFO, logger='dashboard.services.homepage_mail_sync'):
        for _ in range(3):
            assert 'error' in h.sync_homepage_email()
    lv = _levels(caplog)
    assert lv == [logging.WARNING, logging.WARNING, logging.ERROR]
    assert all('라벨 처리 실패' not in r.getMessage() for r in caplog.records)

    # 복구: 라벨 조회 성공 → 연속 횟수 초기화 (메일 검색은 빈 결과로)
    caplog.clear()
    monkeypatch.setattr(h, '_get_or_create_label', lambda service, name: 'LBL')
    monkeypatch.setattr(h, '_gmail_execute', lambda req, what='', tries=4: {'messages': []})

    class _Svc:
        def users(self):
            return self

        def messages(self):
            return self

        def list(self, **kw):
            return None

    monkeypatch.setattr(h, '_get_gmail_service', lambda: _Svc())
    with caplog.at_level(logging.INFO, logger='dashboard.services.homepage_mail_sync'):
        out = h.sync_homepage_email()
    assert out.get('total') == 0
    assert h._connect_fail_streak == 0
    assert any('접속 복구' in r.getMessage() for r in caplog.records)
