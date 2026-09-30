# -*- coding: utf-8 -*-
"""ITG 사업자 통장 레지스트리 — 계좌 → 은행·사업자 코드·카드 라벨의 단일 진실원천.

2026-09-30: 글로벌그룹(R) 기업은행 '하도급지킴이' 계좌(관공서 공사 전용) 신설로
'은행 = 사업자'(기업=G, 하나=R) 가정이 깨졌다. 사업자는 은행명이 아니라 **계좌 번호**로
판정한다. 계좌가 늘면 ITG_ACCOUNTS 에 한 줄만 추가하면 인입 필터·인입 카드 헤더·
수금 카드 코드·정산 핀 리마인드가 함께 따라온다.

매칭 규칙: 텍스트에서 계좌 토큰(452/255/352 로 시작, 숫자·*·- 조합)을 찾아 하이픈을
뺀 뒤, 마스킹(*)이 있으면 마지막 * 뒤의 '보이는 끝자리'가, 없으면 토큰 전체가 등록 계좌
숫자의 끝과 일치하는지 본다. 보이는 끝자리가 5자리 미만이면 판정하지 않는다(모호).
"""
import re
from typing import NamedTuple, Optional


class ItgAccount(NamedTuple):
    prefix: str   # 앞 3자리 (문자 마스킹에도 노출됨)
    digits: str   # 알고 있는 가장 긴 끝자리(전체 계좌번호면 전체) — 하이픈 제외
    bank: str     # 메모·카드 은행명 (기업/하나/농협)
    code: str     # 사업자 코드 G(글로벌)/R(글로벌그룹)/N(농협 N통장)
    label: str    # #입금_관리 인입 카드 헤더 은행 라벨


ITG_ACCOUNTS = (
    ItgAccount('452', '45203938801011', '기업', 'G', '기업은행 (글로벌)'),         # 452-039388-01-011
    ItgAccount('452', '45204786504039', '기업', 'R', '기업은행 (글로벌그룹)'),     # 452-047865-04-039 하도급지킴이
    ItgAccount('255', '25591001431304', '하나', 'R', '하나은행 (글로벌그룹)'),     # 255-910014-31304
    ItgAccount('352', '168233',         '농협', 'N', '농협은행 (N통장)'),          # 352-****-1682-33
)

_MIN_VISIBLE = 5
_ACCT_TOKEN_RE = re.compile(r'(?<!\d)(?:452|255|352)[\d*\-]{6,}')


def match_account(text: str) -> Optional[ItgAccount]:
    """텍스트 속 ITG 사업자 통장 → ItgAccount (없거나 모호하면 None)."""
    for m in _ACCT_TOKEN_RE.finditer(text or ''):
        tok = m.group(0).replace('-', '').rstrip('*')
        visible = tok.rsplit('*', 1)[-1] if '*' in tok else tok
        if len(visible) < _MIN_VISIBLE:
            continue
        for acct in ITG_ACCOUNTS:
            if tok.startswith(acct.prefix) and acct.digits.endswith(visible):
                return acct
    return None
