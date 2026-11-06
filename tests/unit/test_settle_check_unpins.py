# -*- coding: utf-8 -*-
"""정산 글 ✅ 고정 해제 판정 (2026-11-06 강원포장).

매입 계산서 회신 요청 글은 매니저가 ✅로 "내 거래처"라고 답한다. 체커인 매니저(JW)의 ✅에
고정이 풀려 경영지원 처리 전에 목록에서 빠졌다 → 회신 요청 글은 올린 사람의 ✅만 완료.
"""
import sys
sys.path.insert(0, '.')

from dashboard.blueprints.slack_bot import _settle_check_unpins

SB, JW = 'U_SB', 'U_JW'
INVOICE = {'user': SB, 'text': '11/06 G 79,200원 강원포장 회신 바랍니다.'}


def test_manager_check_on_invoice_request_keeps_pin():
    assert _settle_check_unpins(INVOICE, JW) is False


def test_poster_check_on_invoice_request_unpins():
    assert _settle_check_unpins(INVOICE, SB) is True


def test_deposit_message_any_checker_unpins():
    msg = {'user': SB, 'text': '[기업] 11/06 452***38801011 입금 1,100,000원 홍길동'}
    assert _settle_check_unpins(msg, JW) is True


def test_other_message_any_checker_unpins():
    assert _settle_check_unpins({'user': SB, 'text': '공사 금액 수정 요청'}, JW) is True


def test_unreadable_message_falls_back_to_unpin():
    assert _settle_check_unpins(None, JW) is True


def test_bot_posted_invoice_request_needs_same_user():
    """올린 사람이 없는 글(봇 등)에선 체커 ✅로 풀리지 않음 — 사람이 고정을 직접 풀면 됨."""
    assert _settle_check_unpins({'text': '11/06 G 79,200원 강원포장 회신 바랍니다.'}, JW) is False
