# -*- coding: utf-8 -*-
"""경영지원(정산 확인 담당) 명단 — 슬랙 권한과 경영지원 DM 받는 사람을 한곳에서.

권한: 정산 핀 ✅ · 입금 [✅ 확인 후 기록] · 정산 카드 [🗑 삭제] · 법인 간 이체/자금 이동/반환 버튼 ·
금액 수정·매출 0원 요청 카드 ✅. DM: 이상 징후 입금 지정 · 공사 취소/되돌림 · 금액·총액2 변경 알림.

2026-10-28: SB(황샛별) 후임 MY(이미영) 입사 — 인수인계 동안 둘 다 같은 권한·같은 DM.
SB 가 빠질 때 아래 목록에서 한 줄만 지우면 된다.
env 로 덮어쓰기: SLACK_SETTLEMENT_CHECKER_ID · SLACK_SETTLEMENT_CHECKER_EMAIL (쉼표로 여러 명).
"""
import os

_DEFAULT_CHECKERS = (
    ('U0BHC2JV7U5', 'sb@itg-aircon.com'),   # 황샛별
    ('U0CCV2M1V36', 'my@itg-aircon.com'),   # 이미영 (2026-10-28~)
)


def _env_list(name: str) -> list:
    return [x.strip() for x in os.getenv(name, '').split(',') if x.strip()]


def checker_ids() -> tuple:
    """경영지원 슬랙 user ID 들 (앞이 대표)."""
    return tuple(_env_list('SLACK_SETTLEMENT_CHECKER_ID')) or tuple(i for i, _ in _DEFAULT_CHECKERS)


def checker_emails() -> set:
    """경영지원 이메일 들 (소문자) — PM 본인 편집이면 자기 알림을 건너뛸 때."""
    return {e.lower() for e in (_env_list('SLACK_SETTLEMENT_CHECKER_EMAIL')
                                or [e for _, e in _DEFAULT_CHECKERS])}
