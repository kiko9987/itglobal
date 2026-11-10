# -*- coding: utf-8 -*-
"""10-29 한 달치 주소 점검에서 나온 변환 결함.

- 지번+상호에서 부번 유실 (ETC-346b74)
- 숫자뿐인 카카오 건물명('101') 부착 (L-04130)
- 괄호 속 도로명+번지 반복 (L-04315)
- POI 이름의 법인 접두어 (L-04122)
- 폭 0 문자(\\u2060) 저장 (L-04167)
- 'B1층' 표기 통일 (L-04244)
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.address_resolver as ar
from dashboard.services.lead_helpers import _normalize_road_spacing as N


def _road_doc(addr, bldg='', main='', sub=''):
    return {'address_name': addr, 'address_type': 'ROAD_ADDR',
            'address': {'address_name': addr, 'main_address_no': main, 'sub_address_no': sub},
            'road_address': {'address_name': addr, 'building_name': bldg}}


def _jibun_doc(addr, main, sub=''):
    return {'address_name': addr, 'address_type': 'REGION_ADDR',
            'address': {'address_name': addr, 'main_address_no': main, 'sub_address_no': sub},
            'road_address': None}


@pytest.fixture
def kakao(monkeypatch):
    table = {}
    monkeypatch.setattr(ar, '_kakao_key', lambda: 'x')
    monkeypatch.setattr(ar, '_kakao_search', lambda q: table.get(q))
    monkeypatch.setattr(ar, '_search_poi', lambda q: ())
    return table


def test_jibun_sub_number_kept(kakao):
    kakao['관악구 신림동 463-88 서서울공조'] = _jibun_doc('서울 관악구 신림동 463', '463')
    kakao['관악구 신림동 463-88'] = _road_doc('서울 관악구 관천로 87', main='463', sub='88')
    r = ar.verify_address('관악구 신림동 463-88 서서울공조', '관악구 신림동 463-88 서서울공조')
    assert r[0].startswith('관악구 관천로 87')
    assert '463' not in r[0]


def test_jibun_sub_mismatch_rejected_when_no_tail(kakao):
    kakao['관악구 신림동 463-88'] = _jibun_doc('서울 관악구 신림동 463', '463')
    assert ar.verify_address('관악구 신림동 463-88', '관악구 신림동 463-88') is None


def test_jibun_matching_sub_untouched(kakao):
    kakao['관악구 신림동 463-88 서서울공조'] = _road_doc(
        '서울 관악구 관천로 87', main='463', sub='88')
    r = ar.verify_address('관악구 신림동 463-88 서서울공조', '관악구 신림동 463-88 서서울공조')
    assert r[0].startswith('관악구 관천로 87')


def test_numeric_building_name_not_appended(kakao):
    kakao['경기 고양시 일산동구 장항한강3로 80'] = _road_doc(
        '경기 고양시 일산동구 장항한강3로 80', bldg='101')
    r = ar.verify_address('경기 고양시 일산동구 장항한강3로 80 (장항동, 105) 13층',
                          '경기 고양시 일산동구 장항한강3로 80')
    assert ' 101 ' not in r[0] + ' '


@pytest.mark.parametrize('raw, expected', [
    ('남양주 천마산로 113 1층 천마산로 113', '남양주 천마산로 113 1층'),
    ('남양주 천마산로 113 1층 천마산로113', '남양주 천마산로 113 1층'),
    # 다른 번지·다른 도로는 그대로
    ('강남구 테헤란로 152 강남파이낸스센터 테헤란로 153', '강남구 테헤란로 152 강남파이낸스센터 테헤란로 153'),
    ('강서구 공항대로 248 대방빌딩 2층', '강서구 공항대로 248 대방빌딩 2층'),
])
def test_repeated_road_removed(raw, expected):
    assert ar._post_normalize_display(raw) == expected


def test_corp_prefix_stripped_from_poi(monkeypatch):
    monkeypatch.setattr(ar, '_kakao_search_poi', lambda q: ())
    monkeypatch.setattr(ar, '_naver_search_local', lambda q: (
        ('사회복지법인설원복지재단', '경기도 안양시 만안구 예술공원로117번길 32'),
        ('재단법인 아름다운재단', '서울 종로구 자하문로19길 6'),
        ('사단법인', ''),
    ))
    names = [pn for pn, _ in ar._search_poi('q')]
    assert names == ['설원복지재단', '아름다운재단', '사단법인']


@pytest.mark.parametrize('raw, expected', [
    ('인천 연수구 인천타워대로 257 C동 107호⁠', '인천 연수구 인천타워대로 257 C동 107호'),
    ('송도 ⁠아트포레⁠ 107호', '송도 아트포레 107호'),
    ('서초구 서초중앙로 142 삼하빌딩 B1층 101호', '서초구 서초중앙로 142 삼하빌딩 지하1층 101호'),
    ('강남구 테헤란로 152 SB1층', '강남구 테헤란로 152 SB1층'),   # 앞이 영문이면 미대상
])
def test_input_normalization(raw, expected):
    assert N(raw) == expected


@pytest.mark.parametrize('text, expected', [
    ('인천 연수구 청명로 26 (청학동) 202', '(청학동) 202호'),
    ('금천구 시흥대로 399 시티렉스 407', '시티렉스 407호'),
    ('강남구 테헤란로 152 강남파이낸스센터 1203', '강남파이낸스센터 1203호'),
    ('남양주시 순화궁로 384, 별내대원칸타빌 헬스장', '별내대원칸타빌 헬스장'),
])
def test_building_tail_room_and_facility(text, expected):
    assert ar._extract_building_tail(text) == expected


@pytest.mark.parametrize('text', [
    '성북구 종암로 129 010 1234 5678',   # 전화번호 조각
    '성북구 종암로 129 12',               # 두 자리 숫자
    '남양주 천마산로 113 1층 천마산로 113',  # 도로명 뒤 숫자 = 번지
])
def test_building_tail_no_false_room(text):
    assert '호' not in ar._extract_building_tail(text)


@pytest.mark.parametrize('raw, expected', [
    ('시흥대로 399시티렉스 407', '시흥대로 399 시티렉스 407'),
    ('중앙로 12가길 5', '중앙로12가길 5'),          # 도로 접미는 그대로
    ('종로 5가 12', '종로 5가 12'),
    ('테헤란로 152층', '테헤란로 152층'),
])
def test_spaced_road_beonji_building_split(raw, expected):
    assert N(raw) == expected


def test_poi_fallback_si_hint_borrows_tenant_road(monkeypatch):
    def poi(q):
        if q.startswith('현대프리미어캠퍼스'):
            return (('현대프리미어캠퍼스 세마역', ''),
                    ('메가MGC커피 세마현대프리미어캠퍼스점', '경기 오산시 경기대로761번길 30'))
        if q.startswith('세마역'):
            return (('세마역 1호선', '경기 오산시 세마역1호선로 88'),)
        return ()
    monkeypatch.setattr(ar, '_kakao_search_poi', poi)
    r = ar._try_poi_fallback('오산시 세마역 현대프리미어 캠퍼스')
    assert r and '경기대로761번길 30' in r and '현대프리미어캠퍼스' in r


def test_poi_fallback_si_hint_skips_station_only(monkeypatch):
    monkeypatch.setattr(ar, '_kakao_search_poi', lambda q: (
        ('세마역 1호선', '경기 오산시 세마역1호선로 88'),))
    assert ar._try_poi_fallback('오산시 세마역') is None


# ── 석 달치 점검 결함 (2026-10-29 2차) ──

@pytest.mark.parametrize('raw, expected', [
    ('송도 문화로84번길 84, 1동 101호', '송도문화로84번길 84, 1동 101호'),
    ('서울 문화로 84', '서울 문화로 84'),          # 시/도
    ('김포 감암로 125', '김포 감암로 125'),        # 시 약칭
    ('마포구 월드컵로 114', '마포구 월드컵로 114'),  # 행정 접미
    ('강남역 테헤란로 152', '강남역 테헤란로 152'),  # 역 이름
])
def test_glue_word_to_road(raw, expected):
    assert ar._glue_word_to_road(raw) == expected


@pytest.mark.parametrize('text, expected', [
    ('관악구 난곡로 55 관악산휴먼시아아파트 관리사무실', '관악산휴먼시아아파트 관리사무실'),
    ('하이원길 500 스키하우스 2층 패트롤 사무실', '스키하우스 2층 패트롤 사무실'),
    ('동작구 보라매로5가길 24 보라매나산스위트 전기실', '보라매나산스위트 전기실'),
    ('서초구 강남대로 531 밝은눈 안과', '밝은눈 안과'),
    ('포천 가산면 마금길 151, 공장', '공장'),
    ('금천구 가산로 62 금호덴탈제약(주) 포장실', '금호덴탈제약(주) 포장실'),
    ('안양시 동안구 시민대로 383(디지털 엠파이어 B동, 908호)', '디지털 엠파이어 B동 908호'),
    ('중구 을지로 287(을지로6가) 별관', '별관'),
])
def test_building_tail_rooms_and_names(text, expected):
    # verify_address 와 같은 순서: tail 추출 → 괄호 정리
    assert ar._flatten_paren_tail(ar._extract_building_tail(text)) == expected


@pytest.mark.parametrize('text', [
    '강남구 테헤란로 152 근처',
    '강남구 테헤란로 152 공실',
    '강남구 테헤란로 152 사무실 이전 예정',
    '강남구 선릉로112길 13 단독건물(주택)',
    '안성 대덕면 건지안골길 7 산17-16',
])
def test_building_tail_noise_dropped(text):
    assert ar._extract_building_tail(text) == ''


@pytest.mark.parametrize('raw, expected', [
    ('안양 동안구 귀인로98번길 12 평촌트리지아1단지 아파트 평촌트리지아1단지아파트',
     '안양 동안구 귀인로98번길 12 평촌트리지아1단지 아파트'),
    ('송파구 위례순환로 364 송파학사 밀리토피아 송파학사', '송파구 위례순환로 364 송파학사 밀리토피아'),
    ('강남구 테헤란로 152 101동 1203호 101동', '강남구 테헤란로 152 101동 1203호 101동'),  # 유닛 반복은 유지
])
def test_repeated_name_removed(raw, expected):
    assert ar._post_normalize_display(raw) == expected


def test_line_number_not_split():
    from dashboard.services.lead_helpers import extract_korean_address as E
    r = E('경기도 남양주시 오남읍 양자리 661-3 1동 102호 중고 천장형 18평형 정도로 요청드립니다')
    assert r and '661-3' in r[0] and '중고' not in r[0]
    r2 = E('성남 분당구 판교로 393 봇들마을이지더원 단지내상가 2층')
    assert r2 and '봇들마을이지더원' in r2[0]


@pytest.mark.parametrize('tail, expected', [
    ('(목내동, (주)에이스테크) 본관', '(주)에이스테크 본관'),
    ('에스티팜 시화공장(주)', '에스티팜 시화공장(주)'),
    ('(을지로6가)', ''),
])
def test_flatten_nested_corp_paren(tail, expected):
    assert ar._flatten_paren_tail(tail) == expected


def test_compose_dedup_roman_and_cha(kakao):
    kakao['강남구 선릉로86길 31'] = _road_doc('서울 강남구 선릉로86길 31', bldg='선릉역 롯데골드로즈2')
    r = ar.verify_address('강남구 선릉로86길 31 선릉역롯데골드로즈II차 관리사무실',
                          '강남구 선릉로86길 31')
    assert r[0] == '강남구 선릉로86길 31 선릉역 롯데골드로즈2 관리사무실'


@pytest.mark.parametrize('addr, text, expected', [
    # 빠졌으면 원문 단어 그대로 붙임
    ('평택 포승읍 평택항로 294', '평택 포승읍 평택항로 294 EPS코리아', '평택 포승읍 평택항로 294 EPS코리아'),
    # 이미 있거나 정식명에 포함 → 안 붙임
    ('남양주 진관로387번길 47 더좋은공조시스템', '남양주 진관로387번길 47 더좋은공조시스템',
     '남양주 진관로387번길 47 더좋은공조시스템'),
    ('안양 일직로94번길 25 호리바코리아(주)신사옥', '안양 일직로94번길 25 호리바코리아',
     '안양 일직로94번길 25 호리바코리아(주)신사옥'),
    # 표기만 다른 같은 이름 → 안 붙임
    ('인천 서운산업로 30 케이&케이베이스볼아카데미', '인천 서운산업로 30 K&K베이스볼아카데미',
     '인천 서운산업로 30 케이&케이베이스볼아카데미'),
    # 지번·설명·절단어·사람 이름 → 안 붙임
    ('안성 건지안골길 7', '안성 건지안골길 7 산17-16', '안성 건지안골길 7'),
    ('강남구 선릉로112길 13', '강남구 선릉로112길 13 단독건물(주택)', '강남구 선릉로112길 13'),
    ('강남구 테헤란로 152', '강남구 테헤란로 152 근처', '강남구 테헤란로 152'),
    ('강남구 테헤란로 152', '강남구 테헤란로 152 천장형시스템', '강남구 테헤란로 152'),
    ('강남구 테헤란로 152', '강남구 테헤란로 152 김철수', '강남구 테헤란로 152'),
])
def test_append_missing_bare_name(addr, text, expected):
    assert ar._append_missing_bare_name(addr, text) == expected


def test_single_word_cut_by_stopword_not_tail():
    # '시스템' 절단어에 잘린 조각이 tail 로 들어가면 안 됨 (ETC-9c2ec6)
    assert ar._extract_building_tail('남양주 진관로387번길 47 더좋은공조시스템') == ''
    # 건물 신호 없는 이름이 tail 이 되어 카카오 건물명을 막으면 안 됨 (L-04237)
    assert ar._extract_building_tail('보석로 32 자이더테라스상가') == ''


def test_kakao_tenant_list_building_name_skipped(kakao):
    kakao['검단구 봉수대로 1242'] = _road_doc(
        '인천 검단구 봉수대로 1242', bldg='(주)백석개발 인천지점, (주)백석물류')
    r = ar.verify_address('검단구 봉수대로 1242, 2층 백석물류', '검단구 봉수대로 1242')
    assert r[0] == '인천 검단구 봉수대로 1242 2층 백석물류'


@pytest.mark.parametrize('name, expected', [
    ('디에트르송도시그니처뷰아파트 경로당', True),     # L-04506
    ('래미안아파트 작은도서관', True),
    ('디에트르송도시그니처뷰아파트', False),
])
def test_poi_apartment_facility(name, expected):
    assert ar._poi_has_facility(name) is expected


@pytest.mark.parametrize('raw, expected', [
    ('인천타워대로 393 상가 305동A-110호', '인천타워대로 393 상가 305동 A-110호'),
    ('인천타워대로 393 B동 1층 101호', '인천타워대로 393 B동 1층 101호'),   # 영문 동은 그대로
])
def test_dong_latin_room_spacing(raw, expected):
    assert ar._post_normalize_display(raw) == expected


def test_kakao_complex_list_building_name_kept(kakao):
    # 법인 표기 없는 쉼표 나열은 단지명 — 유지
    kakao['안양 동안구 흥안대로223번길 47'] = _road_doc(
        '경기 안양시 동안구 흥안대로223번길 47', bldg='샘마을대우,한양아파트')
    r = ar.verify_address('안양 동안구 흥안대로223번길 47 관리사무소', '안양 동안구 흥안대로223번길 47')
    assert '샘마을대우' in r[0]


@pytest.mark.parametrize('text, room', [
    ('구리 아차산로 417 교문동 108', False),               # 법정동 뒤 숫자 = 지번 (L-04454)
    ('파주시 광탄면 장지산로 31번길 112', False),           # 도로명 바로 뒤 숫자 = 번지 (L-04480)
    ('수원 영통구 태장로82번길 33 자이아파트 상가동 305', True),  # 건물 구역동 뒤는 호수
    ('인천 연수구 청명로 26 (청학동) 202', True),
])
def test_room_number_guards(text, room):
    assert ('호' in ar._extract_building_tail(text)) is room


def test_city_short_region_prefix():
    cands = ar._build_candidates('구리 아차산로 417 교문동 108', None)
    assert cands.index('구리시 아차산로 417') < cands.index('아차산로 417')


def test_label_stripped():
    assert ar._post_normalize_display('강남구 도산대로53길 39 1층 상호 : 쿤달') == '강남구 도산대로53길 39 1층 쿤달'


def test_landmark_shop_not_candidate():
    c = ar._extract_shop_candidates('성동구 성덕정3길 8 건너앞 망치버거 (현장건물뒤) 현재공실')
    assert '망치버거' not in c
    assert '망치버거' not in ar._extract_shop_candidates('성수동 망치버거 맞은편 2층')


def test_homepage_unverified_form_address_kept(monkeypatch):
    import dashboard.services.homepage_mail_sync as hm
    monkeypatch.setattr(hm, 'resolve_address', lambda *a, **k: ('', ''))
    lead = hm.to_lead({'address_block': '(10087)\n김포구래동카림에비뉴\n2층', 'inquiry': '견적 문의'})
    assert lead['방문 주소'] == '김포구래동카림에비뉴 2층'
    assert lead['_meta_address_level'] == 'raw'
    blank = hm.to_lead({'address_block': '()', 'inquiry': '견적 문의'})
    assert blank['방문 주소'] == '-'


def test_company_systems_suffix_not_cut():
    # L-04529: '파트리지시스템즈' 가 절단어 '시스템'에서 잘려 '파트리지 파트리지시스템즈'로 겹침
    from dashboard.services.lead_helpers import extract_korean_address as E
    t = '강남구 테헤란로 207 8층 파트리지시스템즈'
    assert E(t)[0] == t
    assert ar._extract_building_tail(t) == '8층 파트리지시스템즈'
    # 문의글의 '시스템에어컨'·'시스템 설치'는 기존대로 절단
    assert E('강남구 테헤란로 207 8층 사무실 시스템에어컨 견적')[0] == '강남구 테헤란로 207 8층 사무실'
    assert E('강남구 테헤란로 207 8층 한빛시스템스 시스템에어컨 설치')[0] == '강남구 테헤란로 207 8층 한빛시스템스'



# ── 2026-11-10 전체 방문 1,230건 재변환 점검 ──
def test_tail_letter_dong_and_campus_signal():
    assert ar._extract_building_tail(
        '화성 만세구 우정읍 포승향남로 2015-9 B동 티아이로지스 전기차충전소'
    ) == 'B동 티아이로지스 전기차충전소'
    assert ar._extract_building_tail('시흥 배곧2로 96 한화오션 시흥R&D캠퍼스') == '한화오션 시흥R&D캠퍼스'


def test_tail_labeled_shop_name():
    assert ar._extract_building_tail('종로구 계동길 5 상호 가가 (와인바)') == '가가 (와인바)'


def test_tail_after_trailing_dash():
    assert ar._extract_building_tail(
        '인천 연수구 해송로 143- 송도 웰카운티 1단지 관리 사무소'
    ) == '송도 웰카운티 1단지 관리 사무소'


def test_notes_time_and_demolition():
    assert ar.split_address_notes('방배로 83 연세베스트 재활의학과 3층 오전 9시 방문') == (
        '방배로 83 연세베스트 재활의학과 3층', '오전 9시 방문')
    assert ar.split_address_notes('부평구 부평대로 25 2층 한의원 예정지 (현재 철거중)') == (
        '부평구 부평대로 25 2층 한의원 예정지', '현재 철거중')
    # 시간대 낱말 없는 숫자·주소만 있는 줄은 그대로
    assert ar.split_address_notes('강남구 테헤란로 207 8층')[1] == ''


def test_planned_with_vacancy_and_question_mark():
    assert ar._mark_planned('꽃내음3길 46 1층 까페 예정지공실') == '꽃내음3길 46 1층 까페 (예정)'
    assert ar._post_normalize_display('송파구 중대로27길 1 인더크림 카페 ?') == '송파구 중대로27길 1 인더크림 카페'


def test_visit_modal_address_memo_moves_to_consultation():
    from dashboard.blueprints.slack_bot import _split_visit_address_memo as S
    assert S('부평구 부평대로 25 2층 한의원 예정지 (현재 철거중)', '실외기 분리 문의') == (
        '부평구 부평대로 25 2층 한의원 예정지', '실외기 분리 문의 / 현재 철거중', '현재 철거중')
    # 메모 없는 주소는 그대로, 이미 상담 내용에 있으면 다시 안 붙임
    assert S('강남구 테헤란로 207 8층', '10시 이후') == ('강남구 테헤란로 207 8층', '10시 이후', '')
    assert S('방배로 83 3층 오전 9시 방문', '오전 9시 방문 확정')[1] == '오전 9시 방문 확정'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
