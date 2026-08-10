# -*- coding: utf-8 -*-
"""사업자등록증 OCR 이름 정리 회귀 테스트 (2026-08-10).

- PDF 지원: ocr_business_license 가 PDF 를 파일 API(batch_annotate_files)로 분기.
  (Vision 이미지 API 는 PDF 를 'Bad image data' 로 거부 — G3977-JK 계기.)
  → Vision 호출이 필요해 여기선 텍스트→이름 추출 순수 로직만 검증.
- 법인명 꼬리 영문 병기 정리: '한글상호 (English Co., Ltd.)' → 한글 공식 상호만.
  '(주)'·'(지점)' 등 한글 괄호는 보존.
"""
import sys
sys.path.insert(0, '.')

from dashboard.services.business_license_ocr import (
    _clean_name, extract_business_name_from_text,
)


def test_clean_name_strips_trailing_english_paren():
    assert _clean_name('제이에이취엔지니어링주식회사 (JH ENG Co., Ltd') == '제이에이취엔지니어링주식회사'
    assert _clean_name('제이에이취엔지니어링주식회사(JH ENG Co.,Ltd.') == '제이에이취엔지니어링주식회사'
    assert _clean_name('인데코 (INDECO)') == '인데코'
    assert _clean_name('크리스아이티 (Chris IT)') == '크리스아이티'


def test_clean_name_preserves_korean_parens():
    assert _clean_name('(주)크리스아이티') == '(주)크리스아이티'
    assert _clean_name('홍길동상회 (지점)') == '홍길동상회 (지점)'
    assert _clean_name('랩인큐브 주식회사') == '랩인큐브 주식회사'


def test_clean_name_strips_trailing_label():
    # 뒤 라벨(대표자 등) 절단 유지
    assert _clean_name('(주)크리스아이티 대표자 홍길동') == '(주)크리스아이티'


def test_extract_name_from_corp_line():
    text = '사업자등록증\n법인명(단체명) : 제이에이취엔지니어링주식회사 (JH ENG Co., Ltd\n대표자 : 이승제'
    assert extract_business_name_from_text(text) == '제이에이취엔지니어링주식회사'


def test_extract_name_individual_sangho():
    text = '사업자등록증\n상호 : 홍길동상회\n성명 : 홍길동'
    assert extract_business_name_from_text(text) == '홍길동상회'


# ── OCR '상호' 라벨 깨짐 fallback (2026-08-10 R3888-JSH·R3974-YM·G3933-MW) ──
def test_garbled_sangho_label_fallback():
    # '상'이 앞 줄 노이즈에 붙고 값 줄이 '호 : XXX' 만 남는 개인사업자 케이스
    assert extract_business_name_from_text('상성개\n호: 신세계 유통\n명: 서영이') == '신세계 유통'
    assert extract_business_name_from_text('상성개사\n호 : 더밸런스짐\n명 : 민경진') == '더밸런스짐'
    assert extract_business_name_from_text('상성\n호: 꽁지네김밥\n명: 민병구') == '꽁지네김밥'


def test_garbled_fallback_ignores_number_values():
    # '호 :' 뒤 값이 숫자(등록번호 등)면 매치 안 함 — 오탐 방지
    assert extract_business_name_from_text('등록번호: 131-91-11717\n성명: 서영이') == ''


def test_corp_name_wins_over_fallback():
    # 법인명 패턴이 fallback 보다 우선
    text = '법인명(단체명) : (주)크리스아이티\n호 : 노이즈'
    assert extract_business_name_from_text(text) == '(주)크리스아이티'
