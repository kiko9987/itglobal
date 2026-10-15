# -*- coding: utf-8 -*-
"""미지정 입금 추정 프로젝트 — 사업자명 우선 + 입금자명 속 현장주소 키워드 (2026-10-06).

계기: 9/26 엠제이디앤엠 입금자명 'MJ논현동수영장완'·'MJ다산순환로20' (사업자 약칭+현장주소)
→ 사업자명만 비교해 '추정 불가'. 실데이터 주소·금액 사용.
"""
import sys

sys.path.insert(0, '.')

from dashboard.services.pin_remind import _guess_project as g


def _r(code, name, addr, miss):
    return {'프로젝트 코드': code, '사업자명': name, '현장 주소': addr, '미수금': miss}


MJ = [
    _r('G2241-YM', '주식회사 엠제이디앤엠', '남양주시 다산순환로20 현대프리미엄캠퍼스 d동 205호 25.03.25', 1540000),
    _r('G2679-YM', '주식회사 엠제이디앤엠', '강남구 논현로 114길 15 지하1층 수영장', 3300000),
    _r('G3090-YM', '주식회사 엠제이디앤엠', '논현로 115번길14 지하1층 수영장', 440000),
]


def test_address_keywords_in_depositor_name():
    assert g('MJ다산순환로20', 1540000, MJ) == 'G2241-YM'
    # 논현 수영장 현장 2곳 → 미수금이 입금액 이상인 곳 하나
    assert g('MJ논현동수영장완', 3300000, MJ) == 'G2679-YM'


def test_address_ambiguous_returns_empty():
    recs = MJ[:2] + [_r('G3091-YM', '주식회사 엠제이디앤엠', '논현로 115번길14 지하1층 수영장', 5000000)]
    assert g('MJ논현동수영장완', 3300000, recs) == ''   # 후보 2개 → 추정 안 함


def test_name_match_unchanged_and_wins():
    recs = [_r('R4030-YM', '(주)벽진컴퍼니', '성동구 아차산로 1', 11187000),
            _r('G1111-AA', '다른회사', '벽진컴퍼니 사옥', 20000000)]
    assert g('(주)벽진컴퍼니', 11187000, recs) == 'R4030-YM'


def test_company_name_inside_address_is_not_location_evidence():
    # 옛 현장 주소 칸에 업체명 기재('해빛디자인 서초…') — 같은 업체 다른 현장 입금을 끌어가면 안 됨
    recs = [_r('G0275-YG', '(주)해빛디자인', '해빛디자인 서초 삼풍아파트 2동 1406호', 5170000),
            _r('G3769-JK', '(주)해빛디자인', '부천시 안곡로 279번길 42 역곡 동부센트레빌', 264000)]
    assert g('해빛디자인역곡현장', 264000, recs) != 'G0275-YG'
    # 사업자명 후보 여럿(같은 업체 여러 현장) → 주소로 좁힐 때 주소 칸의 업체명으로 고르면 안 됨
    recs2 = [_r('G3429-TH', '주식회사 플레이월드안산키즈카페', '안산 상록구 상록수로 34 뉴라성관광호텔 5층 플레이월드', 37400000),
             _r('G4100-YG', '플레이월드', '충남 천안 서북구 서부대로 413 컨벤션센터 2층', 18000000)]
    assert g('플레이월드', 15000000, recs2) == ''


def test_digits_only_piece_not_counted():
    recs = [_r('G5000-AA', '어떤상호', '서울 중구 2025길 205호', 1000000)]
    assert g('김20205', 500000, recs) == ''


def _rc(code, name, contact, addr, miss):
    return {**_r(code, name, addr, miss), '발주처 담당자': contact}


def test_contact_name_when_no_business_name():
    # 10/14 '이혜빈' 2,600,000원 — 개인 고객(사업자명 '-'), 입금자=발주처 담당자 (2026-10-15)
    recs = [_rc('R4173-YM', '-', '이혜빈', '마포구 성미산로 190-13 2층', 2860000),
            _rc('R4101-JSH', '-', '일단영어학원', '시흥 서울대학로278번길 19-14', 1600000)]
    assert g('이혜빈', 2600000, recs) == 'R4173-YM'
    # 입금액이 미수금보다 크면 후보 아님
    assert g('이혜빈', 3000000, recs) == ''


def test_business_name_beats_contact():
    recs = [_rc('G6000-AA', '(주)한빛상사', '김철수', '서울 중구 1', 5000000),
            _rc('G6001-AA', '다른상호', '한빛상사', '서울 중구 2', 5000000)]
    assert g('(주)한빛상사', 1000000, recs) == 'G6000-AA'


def test_contact_ambiguous_returns_empty():
    recs = [_rc('G7000-AA', '-', '이혜빈', '서울 마포구 1', 3000000),
            _rc('G7001-AA', '-', '이혜빈', '경기 고양시 2', 3000000)]
    assert g('이혜빈', 2600000, recs) == ''


def test_summary_line_keeps_emoji_shortcode():
    from dashboard.services.pin_remind import _summary_line
    s = _summary_line('매출금액 확인 필요\n:white_check_mark: *R4101-JSH* · 잔금 _10,400,000원_ 확인 완료')
    assert s == '매출금액 확인 필요 :white_check_mark: R4101-JSH · 잔금 10,400,000원 확인 완료'
