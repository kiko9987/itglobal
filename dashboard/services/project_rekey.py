# -*- coding: utf-8 -*-
"""공사 코드가 바뀔 때 Redis 연결을 새 코드로 옮김 (2026-10-28 4차).

PM 에서 담당자·사업자를 바꾸면 코드 접미사·앞 글자가 바뀐다(번호는 그대로). 공사 확정 카드와
project_card_msg·project_thread 는 편집 알림(project_slack_notifier)이 이미 옮기지만, 나머지 — 계산서 스레드
(invoice_thread), 입금 카드(payment_slack:ts), 등록증 상태(license_state), 카드 스냅샷(card_field_snap),
취소 스냅샷, 입금 지정 기록 등 — 는 옛 코드에 남아 계산서 요청·입금 카드 갱신이 끊겼다.

키 안에서 코드가 ':' 나 '|' 로 구분된 '한 토큰'일 때만 바꾼다(G4153-JW 가 G4153-JWX 같은 다른 코드의 일부로
잘못 바뀌지 않게). 새 이름이 이미 있으면 덮지 않는다(RENAMENX). TTL 은 RENAME 이 그대로 유지.
"""
import re
from typing import List, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

# 옮기면 안 되는 키 (기록·통계 — 옛 코드 그대로가 사실)
_SKIP_PREFIXES = ('audit', 'invariant_seen', 'log:', 'stats:')


def _token_re(code: str):
    return re.compile(rf'(?<![A-Za-z0-9\-]){re.escape(code)}(?![A-Za-z0-9\-])')


def rekey_project_redis(old_code: str, new_code: str, rc=None) -> List[Tuple[str, str]]:
    """old_code 가 한 토큰으로 들어간 Redis 키를 new_code 로 이름 바꿈. Returns [(옛 키, 새 키)]."""
    old_code, new_code = (old_code or '').strip(), (new_code or '').strip()
    if not old_code or not new_code or old_code == new_code:
        return []
    if rc is None:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    pat = _token_re(old_code)
    moved = []
    for k in rc.scan_iter(match=f'*{old_code}*', count=1000):
        k = k.decode() if isinstance(k, bytes) else k
        if k.startswith(_SKIP_PREFIXES) or not pat.search(k):
            continue
        nk = pat.sub(new_code, k)
        try:
            if rc.renamenx(k, nk):
                moved.append((k, nk))
            else:
                logger.warning(f'[REKEY] 새 키가 이미 있어 건너뜀: {k} → {nk}')
        except Exception as exc:
            logger.warning(f'[REKEY] 이름 바꾸기 실패 {k}: {exc}')
    # 역방향 카드 매핑 값(project_thread:{ch}|{ts} = code)도 새 코드로
    try:
        m = rc.get(f'project_card_msg:{new_code}')
        if m:
            m = m.decode() if isinstance(m, bytes) else m
            cur = rc.get(f'project_thread:{m}')
            cur = cur.decode() if isinstance(cur, bytes) else cur
            if cur == old_code:
                ttl = rc.ttl(f'project_thread:{m}')
                rc.set(f'project_thread:{m}', new_code, ex=ttl if ttl and ttl > 0 else None)
    except Exception as exc:
        logger.debug(f'[REKEY] 역방향 카드 매핑 갱신 실패: {exc}')
    if moved:
        logger.info(f'[REKEY] {old_code} → {new_code}: Redis 키 {len(moved)}개 옮김 '
                    f'({", ".join(k for k, _ in moved[:8])}{" …" if len(moved) > 8 else ""})')
    return moved
