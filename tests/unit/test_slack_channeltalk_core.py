"""슬랙/채널톡 핵심 함수 unit test — 회귀 버그 자동 탐지.

테스트 대상:
- 스팸 감지 (_is_spam_message)
- 슬랙 카드 본문 truncate (긴 inquiry/first_message)
- lead 검색 점수 (_search_leads_for_options)
"""
import pytest


@pytest.mark.unit
class TestSpamDetection:
    """채널톡 스팸 자동 감지 — _is_spam_message"""

    def test_url_2_and_keyword_1_is_spam(self):
        from dashboard.blueprints.channeltalk import _is_spam_message
        msg = (
            '안녕하세요 사장님, https://example.com 무료 채널친구 늘리기 https://channelup.kr 사업 번창하시길.'
        )
        assert _is_spam_message(msg) is True

    def test_url_1_keyword_2_is_spam(self):
        from dashboard.blueprints.channeltalk import _is_spam_message
        msg = '대표님 광고 제안 드립니다. https://example.com 효과 보장 도움이 되시면 연락 부탁드립니다.'
        assert _is_spam_message(msg) is True

    def test_short_message_not_spam(self):
        from dashboard.blueprints.channeltalk import _is_spam_message
        msg = '무료체험 https://x.kr'
        # 50자 미만 — 스팸 판정 X (단순 마케팅 키워드만으로는 부족)
        assert _is_spam_message(msg) is False

    def test_normal_inquiry_not_spam(self):
        from dashboard.blueprints.channeltalk import _is_spam_message
        msg = '안녕하세요. 사무실 30평 에어컨 설치 견적 문의드립니다. 천장형으로 가능한가요?'
        assert _is_spam_message(msg) is False

    def test_friend_recommendation_with_url_not_spam(self):
        """정상 사용자 케이스 — URL + 일반 단어 — 마케팅 키워드 부족"""
        from dashboard.blueprints.channeltalk import _is_spam_message
        msg = '친구가 추천해줘서 문의드립니다. https://itg-aircon.com 사이트 봤습니다. 견적 가능할까요?'
        assert _is_spam_message(msg) is False

    def test_empty_message(self):
        from dashboard.blueprints.channeltalk import _is_spam_message
        assert _is_spam_message('') is False
        assert _is_spam_message(None or '') is False


@pytest.mark.unit
class TestSlackCardTruncate:
    """슬랙 카드 본문 길이 제한 (3000자 한도 안전선 2400)"""

    def test_long_inquiry_truncated(self):
        from dashboard.services.lead_sync import build_inquiry_blocks
        long_content = '에어컨 견적 문의 ' * 500  # ~3500자
        lead = {
            '상담 시간': '2026.06.30. 09:00',
            '고객명': '테스트',
            '고객 연락처': '010-1234-5678',
            '이메일': '',
            '_meta_place': '사무실',
            '_meta_device': '천장형',
            '_meta_inquiry': long_content,
            '방문 주소': '',
        }
        blocks, _ = build_inquiry_blocks(lead, 'L-99999', '당근')
        full_text = str(blocks)
        # 2400자 안전선 — '내용이 길어' 안내 문구 포함
        assert '내용이 길어' in full_text or len(full_text) < 4500

    def test_short_inquiry_not_truncated(self):
        from dashboard.services.lead_sync import build_inquiry_blocks
        lead = {
            '상담 시간': '2026.06.30. 09:00',
            '고객명': '테스트',
            '고객 연락처': '010-1234-5678',
            '이메일': '',
            '_meta_place': '사무실',
            '_meta_device': '천장형',
            '_meta_inquiry': '평수 30평 견적 문의드립니다.',
            '방문 주소': '',
        }
        blocks, _ = build_inquiry_blocks(lead, 'L-99999', '당근')
        full_text = str(blocks)
        assert '내용이 길어' not in full_text


@pytest.mark.unit
class TestLeadSearchScoring:
    """기존 lead 연결 모달 검색 점수 (_search_leads_for_options)
    실제 시트 접근 없이 mock 데이터로 검증"""

    def test_empty_query_returns_recent(self, monkeypatch):
        from dashboard.services import lead_service
        from dashboard.blueprints import slack_bot
        mock_leads = [
            {'리드 No': 'L-03100', '고객명': 'A', '고객 연락처': '010-1111-2222', '플랫폼': '당근'},
            {'리드 No': 'L-03099', '고객명': 'B', '고객 연락처': '010-3333-4444', '플랫폼': '전화'},
            {'리드 No': 'L-03001', '고객명': 'C', '고객 연락처': '010-5555-6666', '플랫폼': '홈페이지'},
        ]
        monkeypatch.setattr(lead_service, 'get_lead_records', lambda: mock_leads)
        opts = slack_bot._search_leads_for_options('', limit=5)
        # 빈 query → 최근순 (lead_no 큰 것부터)
        assert opts[0]['value'] == 'L-03100'
        assert opts[1]['value'] == 'L-03099'

    def test_lead_no_exact_match_scores_highest(self, monkeypatch):
        from dashboard.services import lead_service
        from dashboard.blueprints import slack_bot
        mock_leads = [
            {'리드 No': 'L-03100', '고객명': '홍길동', '고객 연락처': '010-9999-9999', '플랫폼': '당근'},
            {'리드 No': 'L-03099', '고객명': '김철수', '고객 연락처': '010-3100-0000', '플랫폼': '전화'},
        ]
        monkeypatch.setattr(lead_service, 'get_lead_records', lambda: mock_leads)
        opts = slack_bot._search_leads_for_options('L-03100', limit=5)
        # lead_no 매칭이 최상위 (점수 100)
        assert opts[0]['value'] == 'L-03100'

    def test_phone_digits_match(self, monkeypatch):
        from dashboard.services import lead_service
        from dashboard.blueprints import slack_bot
        mock_leads = [
            {'리드 No': 'L-03100', '고객명': 'A', '고객 연락처': '010-2768-2373', '플랫폼': '홈페이지'},
            {'리드 No': 'L-03099', '고객명': 'B', '고객 연락처': '010-9999-0000', '플랫폼': '전화'},
        ]
        monkeypatch.setattr(lead_service, 'get_lead_records', lambda: mock_leads)
        opts = slack_bot._search_leads_for_options('010-2768', limit=5)
        # phone 매칭 (점수 80)
        assert opts[0]['value'] == 'L-03100'
