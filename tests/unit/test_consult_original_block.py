# -*- coding: utf-8 -*-
"""상담 완료 카드 원본 블록에 주소·연락처 반영 — 재문의 칸은 건드리지 않음 (2026-10-08 L-04199).

재문의 칸의 '방문 주소 : -'(이전 문의 주소)가 먼저 걸려 이번 문의 주소로 덮이고 '원본/변환
주소' 2줄로 바뀌던 버그. 이번 문의 본문('문의시간' 줄)부터만 찾아야 한다.
"""
import sys
sys.path.insert(0, '.')

from dashboard.blueprints.slack_bot import _apply_consult_to_original_block as apply

REPEAT = (">🔁 재문의 감지\n>이전 문의 : 2026.01.22 13:57 / L-00520\n"
          ">이전 이름 : 이창덕 (같은 연락처)\n>방문 주소 : -\n>문의 내용 :\n>현 가게 천장형 4년차\n"
          ">상태 : 상담 대기\n>상담 내용 : -\n>--------------------------------------------\n")
CONV = '의정부 회룡로192번길 41 1층 김장하는날 의정부장암점'
L04199 = (">🔔 새 문의 접수 알림 - 온라인 (홈페이지 · 당근)  L-04199\n>------\n" + REPEAT +
          ">문의시간 : 2026.10.08. 08:49\n>이름 / 상호 : 김장하는날\n>연락처 : 010-9145-3180\n"
          ">원본 주소 : 경기 의정부시 장암동 26-9 1층 김장하는날\n>변환 주소 : " + CONV + "\n"
          ">문의 내용 : \n>고효율 기기 정부지원 가능 한지도\n>------")


def test_repeat_section_untouched_when_body_has_converted_address():
    out = apply(L04199, CONV, False, '010-9145-3180')
    assert REPEAT in out                         # 재문의 칸 그대로 ('방문 주소 : -' 유지)
    assert out.count('원본 주소') == 1 and out.count('변환 주소') == 1


def test_body_visit_address_replaced_not_repeat():
    text = L04199.replace(">원본 주소 : 경기 의정부시 장암동 26-9 1층 김장하는날\n>변환 주소 : " + CONV,
                          ">방문 주소 : 의정부 장암동 26-9")
    out = apply(text, CONV, False, '')
    assert REPEAT in out
    assert ">원본 주소 : 의정부 장암동 26-9\n>변환 주소 : " + CONV in out


def test_kakao_insert_goes_before_body_inquiry_not_repeat():
    kakao = (">🔔 새 문의 (카카오톡)\n" + REPEAT + ">문의시간 : 2026.10.08. 09:10\n"
             ">연락처 : -\n>문의 내용 : \n>견적 문의\n")
    out = apply(kakao, '강남 테헤란로 1', False, '010-1111-2222')
    assert REPEAT in out
    assert ">방문 주소 : 강남 테헤란로 1\n>문의 내용 : \n>견적 문의" in out
    assert ">연락처 : 010-1111-2222" in out


def test_no_repeat_section_keeps_existing_behavior():
    plain = ">문의시간 : x\n>연락처 : -\n>방문 주소 : 서울 A\n>문의 내용 : \n>y"
    out = apply(plain, '서울 A', True, '')
    assert ">방문 주소 : 서울 A  ⚠️ [주소 확인 필요]" in out
