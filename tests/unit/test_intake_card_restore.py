# -*- coding: utf-8 -*-
"""인입 원문 만료 대응 (2026-10-06) — 9/26 카드 2장: Redis 원문 7일 만료 → 리마인드 '(입금 내역)'
정보 없음 + [프로젝트 지정]이 '이미 기록된 입금'으로 막힘.

- 보관 기간 90일 (INTAKE_TTL)
- 카드 본문(구분선 사이 인용 줄)에서 원문 복원 → 리마인드 요약 폴백
"""
import sys
from unittest.mock import MagicMock

sys.path.insert(0, '.')

from dashboard.services.sms_intake import INTAKE_TTL, intake_text_from_card, parse_preview

# 실제 9/26 활성 카드 블록 (슬랙 반환 형태: &gt; 엔티티, 마스킹 ∗)
ACTIVE = [{"type": "section", "text": {"type": "mrkdwn", "text": (
    "⠀\n&gt;:bell: *새 입금 내역 알림 - 기업은행 (글로벌)*\n&gt;-------------------------\n"
    "&gt;[Web발신]\n&gt;2026/09/26 16:34\n&gt;입금 3,300,000원\n&gt;MJ논현동수영장완 \n"
    "&gt;452∗∗∗38801011\n&gt;기업\n&gt;-------------------------")}},
    {"type": "actions", "elements": [{"type": "button", "action_id": "payment_intake_open",
                                      "value": "82cc300eca01f480"}]}]
# 확인 대기 카드 (헤더 아래 지정 줄 하나 더)
CONFIRM = [{"type": "section", "text": {"type": "mrkdwn", "text": (
    "⠀\n&gt;:clock4: *확인 대기 — 경영지원 확인 후 기록*\n"
    "&gt;*R4030-YM*  ·  *잔금*  ·  *11,187,000원*   ·   지정 YM\n&gt;-------------------------\n"
    "&gt;[Web발신]\n&gt;하나,10/06, 10:45\n&gt;255∗∗∗∗∗∗31304\n&gt;입금 11,187,000원\n"
    "&gt;(주)벽진컴퍼니\n&gt;-------------------------")}}]


def test_ttl_outlives_unprocessed_cards():
    assert INTAKE_TTL >= 60 * 60 * 24 * 30


def test_restore_active_card_text():
    t = intake_text_from_card(ACTIVE)
    assert t == "[Web발신]\n2026/09/26 16:34\n입금 3,300,000원\nMJ논현동수영장완 \n452***38801011\n기업"
    pv = parse_preview(t)
    assert pv.get('amount') == 3300000
    assert pv.get('acct_code') == 'G'
    assert 'MJ논현동수영장완' in (pv.get('partner') or '')


def test_restore_confirm_card_text():
    t = intake_text_from_card(CONFIRM)
    assert t.startswith('[Web발신]\n하나,10/06, 10:45\n255******31304')
    assert parse_preview(t).get('amount') == 11187000


def test_no_separator_returns_empty():
    assert intake_text_from_card([{"type": "section", "text": {"text": "안내"}}]) == ''
    assert intake_text_from_card(None) == ''


def test_pin_remind_falls_back_to_card_when_record_missing(monkeypatch):
    from dashboard.services import pin_remind as pr
    import dashboard.utils.redis_client as rcmod
    import dashboard.services.project_service as psvc
    client = MagicMock()
    client.pins_list.return_value = {'items': [{'message': {'ts': '1790408102.664319', 'blocks': ACTIVE}}]}
    client.chat_getPermalink.return_value = {'permalink': 'https://x/p'}
    monkeypatch.setattr(pr, '_payment_client', lambda: client)
    monkeypatch.setattr(pr, '_payment_channel', lambda: 'C_INTAKE')
    fake_redis = MagicMock()
    fake_redis.get.return_value = None            # 원문 만료
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: MagicMock(redis=fake_redis))
    monkeypatch.setattr(psvc, 'get_project_records', lambda: [])
    out = pr.collect_intake_pending()
    assert len(out) == 1
    assert out[0]['summary'] != '(입금 내역)'
    assert '3,300,000원' in out[0]['summary'] and '09/26' in out[0]['summary']
    assert out[0]['amount'] == 3300000
