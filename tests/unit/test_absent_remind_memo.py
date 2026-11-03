# -*- coding: utf-8 -*-
"""부재중 리마인드 — 팔로업 메모 자르기·서식 문자, 부재중 첫 알림 접수일 (2026-11-03)."""
from datetime import date

from dashboard.services import absent_remind as ar

L04397 = ('이사할 매장으로 현재 사용중인 캐리어 렌텔 천40평 이전설치 문의. 및 신규 1대 설치 문의. '
          '렌탈 약정 확인 후 다시 전화 준다함.')


def test_callback_memo_keeps_tail_reason():
    assert ar._callback_memo(L04397) == L04397          # 100자 안 → 그대로(뒤의 '다시 전화 준다함' 유지)


def test_callback_memo_strips_slack_format_chars():
    assert ar._callback_memo('중고 제품 금액 및 설치비 상담 (추후 연락 준다 함)_') == '중고 제품 금액 및 설치비 상담 (추후 연락 준다 함)'
    assert ar._callback_memo('a*b~c`d') == 'abcd'
    assert ar._callback_memo('A<B & C>') == 'A&lt;B &amp; C&gt;'


def test_callback_memo_long_cut_at_space():
    s = ('가나다라 ' * 30).strip()
    out = ar._callback_memo(s)
    assert out.endswith('…') and len(out) <= ar._CALLBACK_MEMO_MAX + 1 and not out[:-1].endswith(' ')
    assert ar._callback_memo('가' * 150) == '가' * 100 + '…'    # 띄어쓰기 없으면 글자 수로


def _lead(d, status, memo=''):
    return {'리드 No': 'L-1', '상담 시간': d.strftime('%Y.%m.%d. 09:18'), '상태': status, '고객명': '권산옥',
            '플랫폼': '큐플레이스', '고객 연락처': '010-0000-0000', '상담 내용': memo}


def test_retry_first_alert_shows_received_date(monkeypatch):
    monkeypatch.setattr(ar, '_reminder_ordinal', lambda d, t: 1)
    d = date(2026, 11, 2)
    txt, _ = ar.build_remind_text([], {'x': [_lead(d, '부재중')]})
    assert '010-0000-0000 · 11.02(월) 부재중 접수' in txt and ':warning:' not in txt
    monkeypatch.setattr(ar, '_reminder_ordinal', lambda d, t: 2)
    txt, _ = ar.build_remind_text([], {'x': [_lead(d, '부재중')]})
    assert ':warning: *11.02(월) 부재중 접수 후 재시도 없음*' in txt


def test_callback_line_italic_intact(monkeypatch):
    monkeypatch.setattr(ar, '_reminder_ordinal', lambda d, t: 1)
    lead = _lead(date(2026, 10, 26), '유선 상담', '[10.26 16:02 MJ · 유선 상담] 중고 제품 상담 (추후 연락 준다 함)_')
    txt, _ = ar.build_remind_text([], {}, callback={'B': [lead]})
    assert '· _중고 제품 상담 (추후 연락 준다 함)_' in txt and '__' not in txt
