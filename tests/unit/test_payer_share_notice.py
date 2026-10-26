# -*- coding: utf-8 -*-
"""나눠 내는 공사 — 공사 금액·부가세를 바꿨는데 부담액 합계가 안 맞으면 안내 (2026-10-26).

308호(G4151): 공사 금액 30만 원 하향 후 부담액 그대로 / 309호(G4152): 부가세 없앤 뒤 부담액(부가세 포함) 그대로.
"""
import sys

sys.path.insert(0, '.')

from dashboard.services.payer_split import split_share_notice
import dashboard.blueprints.slack_bot as sb

G4151 = {'프로젝트 코드': 'G4151-JW', '부가세': True, '총액 1': 11542000, '총액 2': 12696200,
         '계약금': 1210000, '중도금': 0, '잔금': 0,
         '계약금 부담자': '임대인 · 남상철 · 6,050,000 · 계산서', '중도금 부담자': '-',
         '잔금 부담자': '임차인 · 권영진(지담학원) · 6,976,200 · 현금'}
G4152 = {'프로젝트 코드': 'G4152-JW', '부가세': False, '총액 1': 11242000, '총액 2': 11242000,
         '계약금': 0, '중도금': 0, '잔금': 0,
         '계약금 부담자': '임대인 · 박기영 · 6,050,000 · 현금', '중도금 부담자': '-',
         '잔금 부담자': '임차인 · 신상진(지담학원) · 6,316,200 · 현금'}


def test_notice_when_total_lowered_without_payer_update():
    msg = split_share_notice(G4151)
    assert '11,842,000' in msg and '11,542,000' in msg


def test_notice_when_vat_removed_without_payer_update():
    msg = split_share_notice(G4152)
    assert '12,366,200' in msg and '11,242,000' in msg


def test_no_notice_when_amounts_match_or_not_split():
    ok = dict(G4152, **{'계약금 부담자': '임대인 · 박기영 · 5,500,000 · 현금',
                        '잔금 부담자': '임차인 · 신상진(지담학원) · 5,742,000 · 현금'})
    assert split_share_notice(ok) == ''
    plain = {'프로젝트 코드': 'X', '부가세': True, '총액 1': 1000000, '총액 2': 1100000,
             '계약금 부담자': '-', '중도금 부담자': '-', '잔금 부담자': '-'}
    assert split_share_notice(plain) == ''


class _Cli:
    def __init__(self):
        self.posts = []

    def chat_postMessage(self, **k):
        self.posts.append(k)
        return {'ok': True}

    def conversations_open(self, users):
        return {'channel': {'id': f'D-{users}'}}


def test_amount_request_done_posts_thread_and_dm(monkeypatch):
    proj, dm = _Cli(), _Cli()
    monkeypatch.setattr(sb, '_project_client', lambda: proj)
    monkeypatch.setattr(sb, '_dm_client', lambda: dm)
    sb._notify_payer_share_after_amount_edit('C1', '111.1', 'G4151-JW', G4151, 'UJW', 'USB')
    assert len(proj.posts) == 1 and proj.posts[0]['thread_ts'] == '111.1'
    assert proj.posts[0]['text'].startswith('<@UJW> <@USB>')
    assert '부담자 설정' in proj.posts[0]['text']
    assert len(dm.posts) == 1 and dm.posts[0]['channel'] == 'D-UJW'


def test_amount_request_done_silent_when_matched(monkeypatch):
    proj, dm = _Cli(), _Cli()
    monkeypatch.setattr(sb, '_project_client', lambda: proj)
    monkeypatch.setattr(sb, '_dm_client', lambda: dm)
    ok = dict(G4151, **{'총액 1': 11842000, '총액 2': 13026200})
    sb._notify_payer_share_after_amount_edit('C1', '111.1', 'G4151-JW', ok, 'UJW', 'USB')
    assert proj.posts == [] and dm.posts == []
