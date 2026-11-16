# -*- coding: utf-8 -*-
"""수금 칸 메모 → 금액 합 (2026-11-16 상태 머신 점검 — 값 합 = 메모 합 상시 점검).

계약금·중도금·잔금 칸 값은 입금마다 '값 += 금액, 메모 += 입금 문자'로 함께 바뀐다. 둘이 어긋나면
(메모만 지워짐·값만 고쳐짐·한쪽 쓰기 실패) 수금 카드·미수금이 틀어진다. 정합성 점검이 칸마다
메모 합을 다시 계산해 값과 비교한다.

규칙(블록 = 빈 줄로 나뉜 한 입금):
- '⚠'(이상징후 이력)·'◆'(부담자 표시) 줄은 뺀다.
- 대체수금 블록은 그 금액.
- '카드 실결제 X원'이 있는 블록은 X(칸 값은 수수료 빼기 전 금액으로 적는다).
- 그 밖엔 '입금 X원'을 더하고, '출금 X원'(반환·이체로 빠진 돈)을 뺀다.
- 금액 줄을 하나도 못 찾으면 None — 옛 양식(빈 '입금일:' 틀, 카톡 요약 등)이라 비교하지 않는다.
"""
import re
from typing import Optional

_IN = re.compile(r'입금\s*:?\s*(-?[\d,]+)\s*원')
_OUT = re.compile(r'출금\s*:?\s*-?\s*([\d,]+)\s*원')
_CARD = re.compile(r'카드\s*실결제\s*([\d,]+)\s*원')


def _n(s: str) -> int:
    try:
        return int(str(s).replace(',', '').strip() or 0)
    except ValueError:
        return 0


def note_amount_total(note: Optional[str]) -> Optional[int]:
    """메모 한 칸의 입금 합(출금은 뺌). 금액 줄이 없으면 None."""
    if not note or not str(note).strip():
        return None
    lines = [l for l in str(note).splitlines() if not l.strip().startswith(('⚠', '◆'))]
    text = '\n'.join(lines).strip()
    if not text:
        return None
    try:
        from dashboard.services.payment_sync import _is_substitute_text, _substitute_amount
    except Exception:                                          # pragma: no cover — 순수 함수로도 쓰게
        _is_substitute_text = _substitute_amount = None
    total, found = 0, False
    for block in re.split(r'\n\s*\n', text):
        if not block.strip():
            continue
        if _is_substitute_text and _is_substitute_text(block):
            amt = _substitute_amount(block) or 0
            if amt > 0:
                total += amt
                found = True
                continue
        card = _CARD.search(block)
        if card:
            total += _n(card.group(1))
            found = True
        else:
            for m in _IN.finditer(block):
                total += abs(_n(m.group(1)))
                found = True
        for m in _OUT.finditer(block):
            total -= _n(m.group(1))
            found = True
    return total if found else None


def cell_note_mismatch(value, note) -> Optional[int]:
    """칸 값과 메모 합이 다르면 메모 합, 같거나 비교할 수 없으면 None."""
    tot = note_amount_total(note)
    if tot is None:
        return None
    v = _n(re.sub(r'[^\d\-]', '', str(value if value is not None else '')) or '0')
    return tot if tot != v else None
