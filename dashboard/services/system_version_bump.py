"""시스템이 PM 화면 밖에서 공사 칸을 고치며 _version 을 올린 기록 (2026-11-20 R4297-YG 계기).

거래처명 자동 채우기(partner_names)는 오래 열린 PM 화면이 옛 값으로 덮어쓰지 못하게 버전을 올린다.
그런데 그 자동 채우기는 사용자가 사업자명을 저장한 직후에 돌아, 저장한 사람의 화면이 곧바로 낡아져
다음 저장이 "다른 사용자가 먼저 수정했습니다"로 거절됐다(박용구 대표 11-20 08:06·08:08 두 번).

여기 남긴 기록으로 저장 때 버전 검사가, 그 사이 올라간 버전이 전부 시스템 기록이고 화면이 그 칸을
안 건드렸으면(옛 값 그대로거나 이미 새 값) 그 칸만 빼고 저장을 통과시킨다. 사람이 고친 버전이 끼어
있거나 화면이 그 칸을 다른 값으로 바꿨으면 지금처럼 충돌로 막는다.
"""
import json
from typing import Dict, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_TTL = 60 * 60 * 24 * 2


def _key(code: str, version) -> str:
    return f'project_sysbump:{code}:{int(version)}'


def _rc():
    from dashboard.utils.redis_client import get_redis_client
    return get_redis_client().redis


def record(code: str, version, changes: Dict[str, tuple]) -> None:
    """changes = {칸: (옛 값, 새 값)}."""
    try:
        _rc().set(_key(code, version),
                  json.dumps({k: [str(a or ''), str(b or '')] for k, (a, b) in changes.items()},
                             ensure_ascii=False), ex=_TTL)
    except Exception as exc:
        logger.warning(f'[SYSBUMP] 기록 실패 {code} v{version}: {exc}')


def get(code: str, version) -> Optional[Dict[str, list]]:
    try:
        v = _rc().get(_key(code, version))
    except Exception:
        return None
    if not v:
        return None
    try:
        return json.loads(v.decode() if isinstance(v, bytes) else v)
    except Exception:
        return None


def _norm(v) -> str:
    s = str(v if v is not None else '').strip()
    return '' if s in ('-', 'None', 'nan') else s


def rebase(code: str, expected, current, data: dict) -> bool:
    """expected → current 사이 버전이 전부 시스템 기록이고 화면(data)이 그 칸을 안 건드렸으면 True.

    True 면 그 칸들을 data 에서 빼(시스템이 쓴 값을 옛 값으로 되돌리지 않게) 저장을 이어 간다."""
    try:
        e, c = int(str(expected or '0')), int(str(current or '0'))
    except ValueError:
        return False
    if c <= e or data is None:
        return False
    drop = set()
    for ver in range(e + 1, c + 1):
        bump = get(code, ver)
        if bump is None:
            return False                       # 사람이 고친 버전(또는 기록 없음) — 진짜 충돌
        for field, pair in bump.items():
            if field not in data:
                continue
            old, new = (pair + ['', ''])[:2]
            v = _norm(data.get(field))
            if v in (_norm(old), _norm(new)):
                drop.add(field)
            else:
                return False                   # 화면이 그 칸을 다른 값으로 바꿈 — 진짜 충돌
    for f in drop:
        data.pop(f, None)
    logger.info(f'[SYSBUMP] {code} v{e}→v{c} 시스템 기록만 — 화면 저장 통과 (뺀 칸: {sorted(drop) or "없음"})')
    return True
