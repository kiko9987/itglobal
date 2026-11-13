# -*- coding: utf-8 -*-
"""리드 시트 수식·슬랙 표기 주입 방어 (2026-11-13 보안 점검).

- 시트: 예전엔 자유 글 6칸만 막아 상담 시간·연락처·gclid 로 수식이 들어갈 수 있었다 → 행 전체.
- 슬랙: 홈페이지·당근 리드 카드가 '<!channel>'·'<링크|글자>' 를 그대로 해석했다 → & < > 이스케이프.
"""
import pytest

from dashboard.services import lead_helpers as lh
from dashboard.services import homepage_mail_sync as hm
from dashboard.services import lead_sync as ls


@pytest.mark.parametrize('value,expected', [
    ('=IMAGE("https://x")', '\'=IMAGE("https://x")'),
    ('+821012345678', "'+821012345678"),
    ('@foo', "'@foo"),
    ('-1+1', "'-1+1"),
    ('-테스트', "'-테스트"),
    ('\t=1', "'\t=1"),
    ('\r=1', "'\r=1"),
    ('\n=1', "'\n=1"),
    ('-', '-'),                      # 빈 값 표시는 그대로
    ('', ''),
    ('홍길동', '홍길동'),
    ('010-1234-5678', '010-1234-5678'),
    ('2026.11.13. 10:00', '2026.11.13. 10:00'),
    (None, None),
    (1234, 1234),
])
def test_sheet_safe_text(value, expected):
    assert lh.sheet_safe_text(value) == expected


def test_sheet_safe_row_covers_every_column():
    row = ['L-00001', '=1+1', '홈페이지', '-', '', '=HYPERLINK("x")', '이름', '+82', 'gclid', '=IMPORTDATA("x")']
    out = lh.sheet_safe_row(row)
    assert out == ['L-00001', "'=1+1", '홈페이지', '-', '', '\'=HYPERLINK("x")', '이름', "'+82", 'gclid',
                   '\'=IMPORTDATA("x")']


@pytest.mark.parametrize('text,expected', [
    ('<!channel>', '&lt;!channel&gt;'),
    ('<https://evil.example|견적서 확인>', '&lt;https://evil.example|견적서 확인&gt;'),
    ('A & B', 'A &amp; B'),
    ('', ''),
    (None, ''),
])
def test_slack_esc(text, expected):
    assert lh.slack_esc(text) == expected


@pytest.mark.parametrize('token,expected', [
    ('Cj0KCQiA_abc-123', 'Cj0KCQiA_abc-123'),
    ('=IMAGE("https://x/?"&A1)', ''),
    ('=1+1', ''),
    ('abc<def', ''),
])
def test_gclid_only_click_id_charset(token, expected):
    body = f'새 입력폼 응답이 접수되었습니다.\ngclid\n{token}\n'
    assert hm._parse_gclid_field(body) == expected


def _card_text(lead):
    blocks, fallback = ls.build_inquiry_blocks(lead, 'L-00001', '홈페이지')
    return blocks[0]['text']['text'], fallback


def test_inquiry_card_escapes_external_text():
    text, fallback = _card_text({
        '고객명': '<!channel>',
        '고객 연락처': '<@U012345>',
        '이메일': 'a&b@example.com',
        '문의 내용': '<https://evil.example|견적서 확인> 문의',
        '방문 주소': '<!here> 서울',
        '_meta_address_level': 'verified',
        '_meta_place': '<!everyone>',
        '_meta_device': '천장형',
    })
    for raw in ('<!channel>', '<@U012345>', '<https://evil.example', '<!here>', '<!everyone>'):
        assert raw not in text, raw
        assert raw not in fallback, raw
    assert '&lt;!channel&gt;' in text
    assert 'a&amp;b@example.com' in text
    # 코드가 넣는 서식은 그대로
    assert '*이름 / 상호*' in text and '`L-00001`' in text


def test_inquiry_card_address_with_ampersand_stays_one_line():
    text, _ = _card_text({
        '고객명': '홍길동', '방문 주소': '서울 A&B빌딩', '_meta_address_raw': '서울 A&B빌딩',
        '_meta_address_level': 'verified',
    })
    assert '*원본 주소*' not in text
    assert '서울 A&amp;B빌딩' in text


def test_inquiry_card_unverified_address_keeps_marker():
    text, _ = _card_text({'고객명': '홍길동', '방문 주소': '<b>주소', '_meta_address_level': 'raw'})
    assert '&lt;b&gt;주소  ⚠️ *[주소 확인 필요]*' in text
