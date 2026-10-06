# -*- coding: utf-8 -*-
"""홈페이지 폼 유입경로(UTM) 숨은 필드 파싱·영속화 회귀 테스트 (2026-10-02).

배경: 당근 '앱/웹사이트 전환' 캠페인(소재 A/B/C) 리드를 시트에서 구분하려고 아임웹 폼에
숨은 단답 필드 '유입경로'를 추가하고 바디 JS 가 URL utm 값을 "source/campaign/content"
로 채움. 파서는 ①필드 없는 옛 메일 하위호환 ②값이 문의 내용에 섞이지 않음
③빈 값(라벨만 남음) 안전 ④[세척] 마커와 공존 ⑤카드에선 마커 떼고 제목 괄호에 '당근 A' 표기.
"""
import sys
sys.path.insert(0, '.')

import pytest

from dashboard.services import homepage_mail_sync as h
from dashboard.services.lead_helpers import (
    normalize_inflow, split_inflow_marker, format_inflow_short, format_inflow_detail,
)
from dashboard.services.lead_sync import build_inquiry_blocks, _is_clean_lead


HEAD = (
    '상업용 냉난방기 · 시스템에어컨 전문 | (주)아이티글로벌\r\n'
    '새 입력폼 응답이 접수되었습니다.\r\n등록자 비회원\r\n등록위치 온라인 문의\r\n'
    '등록시각 2026-10-02 10:30\r\n응답\r\n\r\n개인정보 수집동의\r\n동의\r\n'
    '이름 or 상호\r\n클로드테스트\r\n연락처\r\n01000000000\r\n'
    '이메일 [견적 수신용]\r\n시공 장소\r\n상가 / 상업시설 / 의료시설\r\n'
    '필요한 기기 종류를 선택해 주세요. [복수 선택 가능]\r\n스탠드\r\n'
    '무료 방문 견적 받으실 주소를 입력해 주세요.\r\n(03971)\r\n'
    '서울 마포구 월드컵로 114 (성산동)\r\n1층\r\n'
    '문의 내용 [장문 가능]\r\n테스트 문의입니다.\r\n둘째 줄.\r\n'
)
TAIL = '\r\n입력폼 관리하기 <https://itg-aircon.com/admin/contents/form>\r\n© 끝\r\n'


def _mail(*hidden: str) -> str:
    return HEAD + ''.join(hidden) + TAIL


@pytest.fixture(autouse=True)
def _no_kakao(monkeypatch):
    monkeypatch.setattr(h, 'resolve_address', lambda text, addr, lvl: (addr or '', 'verified'))


# ── 파서 ──────────────────────────────────────────────

def test_old_mail_without_inflow_field_unchanged():
    p = h.parse_mail_body(_mail('문의유형\r\n설치\r\n'))
    assert p['inflow'] == ''
    assert p['inquiry_type'] == '설치'
    assert p['details'] == '테스트 문의입니다.\r\n둘째 줄.'


@pytest.mark.parametrize('hidden', [
    ('문의유형\r\n설치\r\n', '유입경로\r\ndaangn/web_install/A\r\n'),   # 문의유형 뒤
    ('유입경로\r\ndaangn/web_install/A\r\n', '문의유형\r\n설치\r\n'),   # 문의유형 앞
    ('유입경로\r\ndaangn/web_install/A\r\n',),                         # 문의유형 없음
])
def test_inflow_parsed_and_not_mixed_into_details(hidden):
    p = h.parse_mail_body(_mail(*hidden))
    assert p['inflow'] == 'daangn/web_install/A'
    assert '유입경로' not in p['details'] and 'daangn' not in p['details']
    assert p['details'] == '테스트 문의입니다.\r\n둘째 줄.'


@pytest.mark.parametrize('hidden', [
    ('문의유형\r\n설치\r\n', '유입경로\r\n'),                 # 빈 값, 끝
    ('유입경로\r\n', '문의유형\r\n설치\r\n'),                 # 빈 값 → 바로 다음 라벨
])
def test_empty_inflow_is_blank_and_type_still_parsed(hidden):
    p = h.parse_mail_body(_mail(*hidden))
    assert p['inflow'] == ''
    assert p['inquiry_type'] == '설치'
    assert p['details'] == '테스트 문의입니다.\r\n둘째 줄.'


def test_empty_inquiry_type_does_not_swallow_next_label():
    p = h.parse_mail_body(_mail('문의유형\r\n', '유입경로\r\ndaangn//B\r\n'))
    assert p['inquiry_type'] == ''
    assert p['inflow'] == 'daangn//B'


def test_word_inside_details_is_not_a_label():
    body = HEAD.replace('둘째 줄.', '유입경로 궁금하시면 연락주세요') + '문의유형\r\n설치\r\n' + TAIL
    p = h.parse_mail_body(body)
    assert p['inflow'] == ''
    assert '유입경로 궁금하시면' in p['details']


# ── 정규화 ────────────────────────────────────────────

@pytest.mark.parametrize('raw,expected', [
    ('daangn/web_install/A', '당근/web_install/A'),
    ('Karrot/web_install/C', '당근/web_install/C'),
    ('daangn//A', '당근/-/A'),
    ('naver//', '네이버'),
    ('newsite/x', 'newsite/x'),
    ('', ''),
    (' / / ', ''),
    ('daangn/<script>]\n/A', '당근/script/A'),
    ('a' * 80, 'a' * 30),
])
def test_normalize_inflow(raw, expected):
    assert normalize_inflow(raw) == expected


def test_short_label():
    assert format_inflow_short('당근/web_install/A') == '당근 A'
    assert format_inflow_short('당근/-/B') == '당근 B'
    assert format_inflow_short('당근/web_install') == '당근 web_install'
    assert format_inflow_short('네이버') == '네이버'
    assert format_inflow_short('') == ''


# ── to_lead 영속화 ────────────────────────────────────

def test_to_lead_marker_install():
    lead = h.to_lead(h.parse_mail_body(_mail('문의유형\r\n설치\r\n', '유입경로\r\ndaangn/web_install/A\r\n')))
    assert lead['문의 내용'].startswith('[유입:당근/web_install/A]\n테스트 문의입니다.')
    assert lead['_meta_inflow'] == '당근/web_install/A'
    assert not _is_clean_lead(lead)


def test_to_lead_marker_coexists_with_clean():
    lead = h.to_lead(h.parse_mail_body(_mail('문의유형\r\n세척\r\n', '유입경로\r\ndaangn/web_clean/B\r\n')))
    assert lead['문의 내용'].startswith('[세척]\n[유입:당근/web_clean/B]\n테스트')
    assert _is_clean_lead(lead)
    assert _is_clean_lead({'문의 내용': lead['문의 내용']})   # 시트 재렌더 경로


def test_to_lead_no_inflow_no_marker():
    lead = h.to_lead(h.parse_mail_body(_mail('문의유형\r\n설치\r\n')))
    assert '[유입:' not in lead['문의 내용']
    assert lead['_meta_inflow'] == ''


def test_split_marker_roundtrip():
    assert split_inflow_marker('[세척]\n[유입:당근/web_install/A]\n견적 문의') == (
        '당근/web_install/A', '[세척]\n견적 문의')
    assert split_inflow_marker('그냥 문의') == ('', '그냥 문의')


# ── 카드 ──────────────────────────────────────────────

def _section(lead):
    blocks, _ = build_inquiry_blocks(lead, 'L-09999', source='홈페이지')
    return next(b['text']['text'] for b in blocks if b.get('type') == 'section')


@pytest.mark.parametrize('itype', ['설치', '세척'])
def test_card_shows_inflow_line_and_hides_marker(itype):
    lead = h.to_lead(h.parse_mail_body(_mail(f'문의유형\r\n{itype}\r\n', '유입경로\r\ndaangn/web_install/A\r\n')))
    text = _section(lead)
    assert '*새 문의 접수 알림 - 온라인 (홈페이지 · 당근)*' in text   # 제목엔 채널만
    assert '>*광고 소재* : web_install A\n' in text                  # 소재는 문의시간 아래 줄
    assert text.index('*문의시간*') < text.index('*광고 소재*') < text.index('*이름 / 상호*')
    assert '[유입:' not in text and '[세척]' not in text
    assert '테스트 문의입니다.' in text
    # 시트에서 읽어 재렌더(메타 없음)해도 동일
    lead2 = {k: v for k, v in lead.items() if not k.startswith('_meta')}
    text2 = _section(lead2)
    assert '온라인 (홈페이지 · 당근)' in text2 and '*광고 소재* : web_install A' in text2 and '[유입:' not in text2


def test_card_without_inflow_has_no_line():
    lead = h.to_lead(h.parse_mail_body(_mail('문의유형\r\n설치\r\n')))
    t = _section(lead)
    assert '*새 문의 접수 알림 - 온라인 (홈페이지)*' in t
    assert '검색 키워드' not in t and '광고 소재' not in t


# ── 테스트 제출 skip (시트·슬랙 없이 라벨만) ───────────────

def test_is_test_submission_requires_both_name_and_phone():
    assert h.is_test_submission({'고객명': '클로드테스트', '고객 연락처': '010-0000-0000'})
    assert not h.is_test_submission({'고객명': '홍길동', '고객 연락처': '010-0000-0000'})
    assert not h.is_test_submission({'고객명': '클로드테스트', '고객 연락처': '010-1234-5678'})


def test_sync_skips_test_submission_without_sheet_or_slack(monkeypatch):
    import base64
    import pandas as pd
    import dashboard.services.lead_sync as ls

    body = _mail('문의유형\r\n설치\r\n', '유입경로\r\ndaangn/web_install/A\r\n')
    data = base64.urlsafe_b64encode(body.encode('utf-8')).decode().rstrip('=')
    msg = {'payload': {'mimeType': 'text/plain', 'body': {'data': data}}}

    class _Req:
        def __init__(self, v): self.v = v

    class _Msgs:
        def list(self, **k): return _Req({'messages': [{'id': 'm1'}]})
        def get(self, **k): return _Req(msg)

    class _Svc:
        def users(self): return self
        def messages(self): return _Msgs()

    calls = {'append': 0, 'slack': 0, 'marked': []}
    monkeypatch.setattr(h, '_get_gmail_service', lambda: _Svc())
    monkeypatch.setattr(h, '_get_or_create_label', lambda s, n: 'L1')
    monkeypatch.setattr(h, '_gmail_execute', lambda req, what='': req.v)
    monkeypatch.setattr(h, 'load_leads_data', lambda force_refresh=False: pd.DataFrame())
    monkeypatch.setattr(ls, '_get_existing_phone_lookup', lambda df: {})
    monkeypatch.setattr(ls, '_append_leads_to_main',
                        lambda leads: calls.__setitem__('append', calls['append'] + 1) or [])
    monkeypatch.setattr(ls, '_send_slack_notifications',
                        lambda *a, **k: calls.__setitem__('slack', calls['slack'] + 1) or set())
    monkeypatch.setattr(h, '_mark_processed',
                        lambda s, mid, lid: calls['marked'].append(mid) or True)

    h.sync_homepage_email()
    assert calls['append'] == 0 and calls['slack'] == 0
    assert calls['marked'] == ['m1']   # 재처리 루프 방지


# ── 네이버 자동추적·구글 gclid 유입 (2026-10-04 스니펫 확장) ──

@pytest.mark.parametrize('raw,marker,short', [
    ('naver/천장형에어컨/천장형 에어컨 설치업체', '네이버/천장형에어컨/천장형에어컨설치업체', '네이버 천장형에어컨설치업체'),
    ('naver//시스템에어컨 견적', '네이버/-/시스템에어컨견적', '네이버 시스템에어컨견적'),   # 확장검색(등록키워드 없음)
    ('naver/냉난방기', '네이버/냉난방기', '네이버 냉난방기'),
    ('google', '구글', '구글'),
    # 2026-10-06 소재 ID(n_ad) 4번째 칸 — 카드 표기는 그대로
    ('naver/냉난방기설치/냉난방기 설치업체/nad-a001-01-000000593496019',
     '네이버/냉난방기설치/냉난방기설치업체/nad-a001-01-000000593496019', '네이버 냉난방기설치업체'),
    ('naver/냉난방기설치//nad-a001-01-000000593496019',
     '네이버/냉난방기설치/-/nad-a001-01-000000593496019', '네이버 냉난방기설치'),
    ('google/냉난방기 설치', '구글/냉난방기설치', '구글 냉난방기설치'),
])
def test_naver_google_inflow(raw, marker, short):
    assert normalize_inflow(raw) == marker
    assert format_inflow_short(marker) == short


def test_naver_inflow_end_to_end_card():
    lead = h.to_lead(h.parse_mail_body(_mail('문의유형\r\n설치\r\n', '유입경로\r\nnaver/천장형에어컨/천장형 에어컨 설치\r\n')))
    assert lead['문의 내용'].startswith('[유입:네이버/천장형에어컨/천장형에어컨설치]\n')
    t = _section(lead)
    assert '온라인 (홈페이지 · 네이버)*' in t
    assert '>*검색 키워드* : 천장형에어컨설치\n' in t


@pytest.mark.parametrize('marker,expected', [
    ('네이버/홍대사무실스탠드냉난방기/홍대사무실스탠드냉난방기', ('검색 키워드', '홍대사무실스탠드냉난방기')),
    ('네이버/-/시스템에어컨견적', ('검색 키워드', '시스템에어컨견적')),    # 확장검색
    ('네이버/냉난방기', ('검색 키워드', '냉난방기')),                    # 검색어 없으면 등록 키워드
    ('네이버', None),
    ('당근/web_install/A', ('광고 소재', 'web_install A')),
    ('구글', None),
    ('구글/냉난방기설치', ('검색 키워드', '냉난방기설치')),
    ('네이버/냉난방기설치/냉난방기설치업체/nad-a001-01-000000593496019', ('검색 키워드', '냉난방기설치업체')),
    ('네이버/냉난방기설치/-/nad-a001-01-000000593496019', ('검색 키워드', '냉난방기설치')),
    ('', None),
])
def test_inflow_detail_line(marker, expected):
    assert format_inflow_detail(marker) == expected


def test_google_card_has_channel_only():
    lead = h.to_lead(h.parse_mail_body(_mail('문의유형\r\n설치\r\n', '유입경로\r\ngoogle\r\n')))
    t = _section(lead)
    assert '온라인 (홈페이지 · 구글)*' in t and '검색 키워드' not in t


def test_naver_ad_id_kept_in_marker_but_not_on_card():
    """소재 ID 는 시트 마커에만 남고(소재별 문의 집계용) 카드엔 노출 안 됨."""
    lead = h.to_lead(h.parse_mail_body(_mail(
        '문의유형\r\n설치\r\n', '유입경로\r\nnaver/냉난방기설치/냉난방기 설치업체/nad-a001-01-000000593496019\r\n')))
    assert lead['문의 내용'].startswith('[유입:네이버/냉난방기설치/냉난방기설치업체/nad-a001-01-000000593496019]\n')
    t = _section(lead)
    assert '>*검색 키워드* : 냉난방기설치업체\n' in t and 'nad-' not in t
