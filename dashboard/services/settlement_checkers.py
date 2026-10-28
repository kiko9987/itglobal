# -*- coding: utf-8 -*-
"""정산 승인 명단 — 슬랙 권한과 경영지원 DM 받는 사람을 한곳에서.

권한: 정산 핀 ✅ · 입금 [✅ 확인 후 기록] · 정산 카드 [🗑 삭제] · 법인 간 이체/자금 이동/반환 버튼 ·
금액 수정·매출 0원 요청 카드 ✅. DM: 이상 징후 입금 지정 · 공사 취소/되돌림 · 금액·총액2 변경 알림.

2026-10-28: SB(황샛별) 후임 MY(이미영) 입사 → 10-29 인수인계 중 퇴사, 명단에서 뺌.
2026-10-29: 경영지원 후임이 없어 SB 권한을 KiKO·JW·YG 에게도 똑같이 (사용자 지시).
  JW·YG 는 영업 겸 승인 — A/S 담당자 목록 같은 '영업이냐' 판단은 이 명단이 아니라 staff_directory 를 본다.
SB 가 빠질 때 아래 목록에서 한 줄만 지우면 된다.
env 로 덮어쓰기: SLACK_SETTLEMENT_CHECKER_ID · SLACK_SETTLEMENT_CHECKER_EMAIL (쉼표로 여러 명, 같은 순서).
"""
import os

_DEFAULT_CHECKERS = (
    ('U0BHC2JV7U5', 'sb@itg-aircon.com'),     # 황샛별 (퇴사 예정)
    ('U04UL2ZLJAX', 'kiko@itg-aircon.com'),   # 고광일 (2026-10-29~)
    ('U051ENK8AAY', 'jw@itg-aircon.com'),     # 박정우 (2026-10-29~)
    ('U050XNG5VBR', 'yg@itg-aircon.com'),     # 박용구 (2026-10-29~)
)


def _env_list(name: str) -> list:
    return [x.strip() for x in os.getenv(name, '').split(',') if x.strip()]


def checker_ids() -> tuple:
    """정산 승인자 슬랙 user ID 들 (앞이 대표)."""
    return tuple(_env_list('SLACK_SETTLEMENT_CHECKER_ID')) or tuple(i for i, _ in _DEFAULT_CHECKERS)


def checker_emails() -> set:
    """정산 승인자 이메일 들 (소문자)."""
    return {e.lower() for e in (_env_list('SLACK_SETTLEMENT_CHECKER_EMAIL')
                                or [e for _, e in _DEFAULT_CHECKERS])}


def checkers() -> list:
    """(슬랙 ID, 이메일 소문자) 짝 — PM 본인 편집이면 그 사람 DM 만 건너뛸 때.

    env 로 덮어쓴 경우 두 목록 길이가 같을 때만 짝짓고, 다르면 이메일은 빈 값(아무도 안 건너뜀)."""
    ids = _env_list('SLACK_SETTLEMENT_CHECKER_ID')
    if not ids:
        return [(i, e.lower()) for i, e in _DEFAULT_CHECKERS]
    emails = _env_list('SLACK_SETTLEMENT_CHECKER_EMAIL')
    if len(emails) != len(ids):
        emails = [''] * len(ids)
    return [(i, e.lower()) for i, e in zip(ids, emails)]
