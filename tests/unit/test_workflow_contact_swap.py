# -*- coding: utf-8 -*-
"""전화 워크플로 입력 칸 밀림 교정 (L-04268).

매니저가 이름을 연락처 칸에, 전화번호를 이메일 칸에 넣어 카드에 연락처가 '김보경'으로
나가던 사고. 연락처 칸에 숫자가 없고 이메일·고객명 칸에 전화번호가 있을 때만 교정.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.lead_sync import _fix_swapped_contact_fields as F


def test_name_in_phone_and_phone_in_email():
    assert F('-', '김보경', '010-3435-0924') == {
        '고객 연락처': '010-3435-0924', '이메일': '-', '고객명': '김보경'}


def test_phone_in_email_keeps_existing_name():
    # 고객명이 이미 있으면 유지, 연락처 칸 글자는 버리지 않고 그대로 둠(고객명 미변경)
    assert F('서울교통회관', '담당자', '01027339154') == {
        '고객 연락처': '01027339154', '이메일': '-'}


def test_phone_and_name_swapped():
    assert F('010-1234-5678', '홍길동', '') == {'고객 연락처': '010-1234-5678', '고객명': '홍길동'}


@pytest.mark.parametrize('name, phone, email', [
    ('김보경', '010-3435-0924', '-'),           # 정상
    ('김보경', '', ''),                          # 연락처 비어 있음(입력 진행 중) — 미교정
    ('김보경', '-', '010-3435-0924'),            # '-'는 미교정
    ('김보경', '추후 전달', 'abc@naver.com'),    # 이메일이 전화 아님
    ('김보경', '02-123-4567 내선', '-'),        # 연락처에 숫자 있음
])
def test_no_fix(name, phone, email):
    assert F(name, phone, email) is None


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
