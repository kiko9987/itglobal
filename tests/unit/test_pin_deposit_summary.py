# -*- coding: utf-8 -*-
"""pin_remind._format_deposit_summary 회귀 — 하나 라벨형 거래처 누출 방지.

배경(2026-08-12): #영업_관리 실데이터 백테스트 결과, 하나 '라벨형' SMS
(일시 …/계좌번호 …/적요 거래처)에서 요약 거래처가 '일시 , 계좌번호 적요 더밸런스짐'
처럼 라벨이 섞여 나왔음(146건 중 44건). _format_deposit_summary 가 검증된 파서
(_parse_memo_block)를 재사용하도록 수정 → 거래처만 깔끔히. 이 회귀 보호.

pure 함수 — Slack/Redis 미접촉.
"""
import re
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.pin_remind import _format_deposit_summary, is_invoice_request

_LABEL = re.compile(r'일시|적요|계좌번호')


class TestInvoiceRequestDetect:
    """세금계산서 요청 자동 고정 감지 — 구조 기반(단위 '원' 오타 무관, 2026-09-10).

    금액 단위 오타(운/왼/웜 등 IME 슬립)로 자동 고정이 누락되던 것(새서울공조·최우신)을
    '원' 글자 대신 MM/DD+G/R/N+돈형태금액 구조로 판정하도록 전환.
    """

    @pytest.mark.parametrize('text', [
        '09/08 R 1,805,000원 일산캐리어 R4080-JK 회신 부탁드립니다.',  # 정상
        '09/10 R 1,100,000운 새서울공조 회신 부탁드립니다.',           # 원→운 오타
        '09/10 G 66,000왼 최우신 회신 부탁드립니다.',                # 원→왼 오타
        '09/10 N 500,000웜 홍길동',                               # 또다른 오타
        '9/8 G 66000 최우신',                                     # 단위 없음, 콤마 없음(4자리+)
        '09/10 G 66,000 최우신',                                  # 단위 없음, 콤마
    ])
    def test_matches(self, text):
        assert is_invoice_request(text) is True

    @pytest.mark.parametrize('text', [
        '09/10 G 12 회의건',              # 소액/비금액 (콤마·4자리 아님)
        '09/10 계산서 미발행건 확인 부탁',   # G/R/N 아님
        '계약서 진행 여부 확인해주세요',       # 양식 아님
        '09/10 R 3 건',                  # 한 자리 숫자
    ])
    def test_no_false_positive(self, text):
        assert is_invoice_request(text) is False


class TestNoLabelLeak:
    @pytest.mark.parametrize('raw, expect', [
        # 하나 라벨형 — 핵심 수정 대상
        ('일시 08/07, 10:55\n입금 275,000원\n계좌번호 255******31304\n적요 더밸런스짐',
         '08/07 R 275,000원 더밸런스짐'),
        ('일시 08/12, 09:41\n입금 2,563,000원\n계좌번호 255******31304\n적요 대우인쇄교역',
         '08/12 R 2,563,000원 대우인쇄교역'),
        ('일시 08/11, 18:22\n입금 240,000원\n계좌번호 255******31304\n적요 신명희(희빛데이)',
         '08/11 R 240,000원 신명희(희빛데이)'),
    ])
    def test_hana_label_format(self, raw, expect):
        out = _format_deposit_summary(raw)
        assert not _LABEL.search(out), f'라벨 누출: {out!r}'
        assert out == expect


class TestOtherFormatsStillOK:
    @pytest.mark.parametrize('raw, expect', [
        # 기업 잔액형
        ('[Web발신]\n2026/08/10 14:49\n입금 407,000원\n잔액 144,153,377원\n㈜시프트업\n452***38801011\n기업',
         '08/10 G 407,000원 ㈜시프트업'),
        # 하나 잔액형
        ('[Web발신]\n하나,08/07, 15:33\n255******31304\n입금5,115,000원\n(주)클리어윈코\n잔액239,612,486원',
         '08/07 R 5,115,000원 (주)클리어윈코'),
        # 농협
        ('농협 입금3,825,000원\n08/11 14:48 352-****-1682-33 장동선',
         '08/11 N 3,825,000원 장동선'),
    ])
    def test_format(self, raw, expect):
        out = _format_deposit_summary(raw)
        assert not _LABEL.search(out), f'라벨 누출: {out!r}'
        assert '잔액' not in out and '잔고' not in out   # 잔액 노출 금지
        assert out == expect


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
