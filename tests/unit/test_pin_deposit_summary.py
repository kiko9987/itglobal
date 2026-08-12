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
from dashboard.services.pin_remind import _format_deposit_summary

_LABEL = re.compile(r'일시|적요|계좌번호')


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
