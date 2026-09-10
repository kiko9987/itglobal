"""시트 직접수정 감지 리컨사일러 (2026-09-10).

경영지원(샛별) 등이 PM을 안 거치고 구글 시트를 직접 수정하면 공사확정 카드가 stale.
PM 편집 경로(projects.py)는 notify_project_field_changes 로 카드 재렌더+댓글 로그가 되지만,
시트 직접수정은 훅이 없어 누락된다. 이 폴러가 주기적으로 시트 vs 스냅샷을 대조해
직접수정을 감지 → 기존 notify_project_field_changes(카드 재렌더 + '[코드 데이터 수정 알림]'
댓글) 를 재활용해 반영한다.

PM 편집 중복 방지: PM 편집 시 projects.py 가 project_pm_edit:{code} 마커(15분)를 세팅 →
폴러는 그 건은 조용히 스냅샷만 맞추고 재발송하지 않는다(write-behind 지연 레이스도 차단).

안전장치(payment_sync 패턴): 첫 감지=베이스라인만 무발송, 한 폴링에 변경 과다면 재베이스라인만.
대상=공사확정 카드 매핑(project_card_msg:{code})이 살아있는 프로젝트만.
"""
import os

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_SNAP_PREFIX = 'card_field_snap:'
_SNAP_TTL = 60 * 60 * 24 * 180        # 180일 (카드 매핑 TTL과 정렬)
_MUTEX_KEY = 'project_reconcile:running'
_MUTEX_TTL = 300
_PM_MARK_PREFIX = 'project_pm_edit:'  # PM 편집 최근 마커 (projects.py 가 세팅)
_MAX_PER_TICK = 30                    # 이 이상 변경 = 비정상(시트 대량수정 등) → 재베이스라인만
# 리컨사일러 감시 제외 필드 — 공사확정 카드에 표시 안 되고 전용 채널(#수금_관리)이 담당.
#   계산서(Y)는 유지: 미발행→'잔금 - 발행완료' 변화로 발행 완료를 인지할 수 있어 유용.
#   수금 3단계·수금날짜는 매 입금마다 바뀌어 노이즈 → 리컨사일러에서만 제외(PM편집 _NOTIFY_FIELDS는 유지).
_RECONCILE_EXCLUDE = {'계약금', '중도금', '잔금', '수금 날짜'}


def _dec(v):
    return v.decode() if isinstance(v, bytes) else v


def store_field_snapshot(rc, code: str, data: dict, fields) -> None:
    """프로젝트의 _NOTIFY_FIELDS 원본값을 Redis 해시 스냅샷으로 저장(카드가 반영하는 최신 상태)."""
    snap = {f: str(data.get(f) if data.get(f) is not None else '') for f in fields}
    try:
        rc.delete(_SNAP_PREFIX + code)
        if snap:
            rc.hset(_SNAP_PREFIX + code, mapping=snap)
        rc.expire(_SNAP_PREFIX + code, _SNAP_TTL)
    except Exception as exc:
        logger.debug(f'[RECONCILE] 스냅샷 저장 실패 ({code}): {exc}')


def reconcile_project_cards() -> dict:
    """공사확정 카드 있는 프로젝트의 시트 직접수정 감지 → 카드 반영 + 로그. 스케줄러 진입점."""
    result = {'checked': 0, 'reflected': 0, 'baseline': 0}
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception as exc:
        logger.warning(f'[RECONCILE] Redis 불가: {exc}')
        return result

    # 동시 실행 방지
    try:
        if not rc.set(_MUTEX_KEY, '1', nx=True, ex=_MUTEX_TTL):
            return result
    except Exception:
        pass

    try:
        from dashboard.services.project_service import get_project_records
        from dashboard.services.project_slack_notifier import (
            _NOTIFY_FIELDS, _fmt_field, notify_project_field_changes,
        )
        watch = _NOTIFY_FIELDS - _RECONCILE_EXCLUDE   # 카드 표시 정보 + 계산서 (수금 3단계·수금날짜 제외)

        # 카드 매핑 살아있는 코드 집합
        live = set()
        try:
            for k in rc.scan_iter(match='project_card_msg:*', count=500):
                ks = _dec(k)
                live.add(ks.split(':', 1)[1])
        except Exception as exc:
            logger.warning(f'[RECONCILE] 카드 매핑 스캔 실패: {exc}')
            return result
        if not live:
            return result

        recs = get_project_records(force_refresh=True) or []
        drift = []
        for r in recs:
            code = str(r.get('프로젝트 코드', '')).strip()
            if not code or code not in live:
                continue
            result['checked'] += 1
            cur = {f: str(r.get(f) if r.get(f) is not None else '') for f in watch}
            try:
                prev_raw = rc.hgetall(_SNAP_PREFIX + code)
            except Exception:
                prev_raw = {}
            prev = {_dec(kk): _dec(vv) for kk, vv in (prev_raw or {}).items()}

            if not prev:
                store_field_snapshot(rc, code, r, watch)  # 첫 감지 = 베이스라인
                result['baseline'] += 1
                continue

            diffs = [f for f in watch
                     if _fmt_field(f, prev.get(f, '')) != _fmt_field(f, cur.get(f, ''))]
            if not diffs:
                continue

            # PM 편집 최근 마커 → 조용히 스냅샷만 (중복 발송 방지)
            try:
                pm_recent = rc.get(_PM_MARK_PREFIX + code)
            except Exception:
                pm_recent = None
            if pm_recent:
                store_field_snapshot(rc, code, r, watch)
                continue

            drift.append((code, r, prev, diffs))

        if not drift:
            return result

        # 안전장치 — 한 폴링에 변경 과다면 발송 없이 재베이스라인만
        if len(drift) > _MAX_PER_TICK:
            logger.warning(
                f'[RECONCILE] 변경 과다({len(drift)}) — 재베이스라인만, 발송 skip')
            for code, r, prev, diffs in drift:
                store_field_snapshot(rc, code, r, watch)
            return result

        for code, r, prev, diffs in drift:
            field_changes = [
                {'field_name': f,
                 'old_value': prev.get(f, ''),
                 'new_value': str(r.get(f) if r.get(f) is not None else '')}
                for f in diffs
            ]
            try:
                notify_project_field_changes(code, field_changes, latest_data=r,
                                             editor='시트 직접수정')
                result['reflected'] += 1
                logger.info(f'[RECONCILE] 시트 직접수정 감지 → 카드 반영+로그: {code} {diffs}')
            except Exception as exc:
                logger.warning(f'[RECONCILE] 처리 실패 ({code}): {exc}')
            # notify 성공/실패와 무관하게 스냅샷 갱신 (반복 발송 방지)
            store_field_snapshot(rc, code, r, watch)

        return result
    finally:
        try:
            rc.delete(_MUTEX_KEY)
        except Exception:
            pass
