"""Slack 카드 매핑 orphan 정리 스크립트.

시트에서 사라진 프로젝트 코드의 Redis 매핑 (project_card_msg + project_thread) 을
스캔·정리한다. 매니저가 시트에서 프로젝트를 수동으로 완전 삭제한 경우
자동 정리 경로가 없어 매핑이 180일 TTL 로 남는다.

사용:
    python scripts/cleanup_orphan_slack_mappings.py           # dry-run (기본)
    python scripts/cleanup_orphan_slack_mappings.py --apply   # 실제 삭제

배포·주기:
    수동 실행. 매니저가 시트 정리 후 이 스크립트를 돌리도록 안내.
"""
from __future__ import annotations

import argparse
import io
import os
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
os.chdir(PROJECT_ROOT)

from dotenv import load_dotenv
load_dotenv(PROJECT_ROOT / '.env')


def main() -> int:
    parser = argparse.ArgumentParser(description='Slack 카드 매핑 orphan 정리')
    parser.add_argument('--apply', action='store_true', help='실제 삭제 (기본은 dry-run)')
    args = parser.parse_args()

    from dashboard.services.payment_sync import _get_payment_service
    from dashboard.utils.redis_client import get_redis_client

    rc = get_redis_client()
    if rc is None:
        print('❌ Redis 연결 불가')
        return 1

    # 1. 시트에서 현존 프로젝트 코드 수집
    svc = _get_payment_service()
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    if not (sid and name):
        print('❌ GOOGLE_SHEET_ID/NAME 미설정')
        return 1

    resp = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{name}'!A2:A10000",
    ).execute()
    sheet_codes = {r[0].strip() for r in resp.get('values', []) if r and r[0]}
    print(f'시트 존재 코드: {len(sheet_codes)}개')

    # 2. Redis 에서 project_card_msg:* 전체 스캔
    orphans_fwd: list[tuple[str, str]] = []  # (fwd_key, mapping_value)
    for key in rc.scan_iter(match='project_card_msg:*', count=500):
        code = key.split(':', 1)[1] if ':' in key else key
        if code in sheet_codes:
            continue
        mapping = rc.get(key) or ''
        orphans_fwd.append((key, mapping))

    print(f'\n정방향 orphan (시트에 없는 코드): {len(orphans_fwd)}개')
    for k, v in orphans_fwd[:20]:
        print(f'  {k} → {v}')
    if len(orphans_fwd) > 20:
        print(f'  ... 외 {len(orphans_fwd) - 20}개')

    # 3. Redis 에서 project_thread:* 전체 스캔 (값이 시트에 없는 code 인 경우)
    orphans_rev: list[tuple[str, str]] = []  # (rev_key, code_value)
    for key in rc.scan_iter(match='project_thread:*', count=500):
        code = rc.get(key) or ''
        if not code or code in sheet_codes:
            continue
        orphans_rev.append((key, code))

    print(f'\n역방향 orphan (값이 시트에 없는 코드): {len(orphans_rev)}개')
    for k, v in orphans_rev[:20]:
        print(f'  {k} = {v}')
    if len(orphans_rev) > 20:
        print(f'  ... 외 {len(orphans_rev) - 20}개')

    total = len(orphans_fwd) + len(orphans_rev)
    if total == 0:
        print('\n✅ orphan 없음. 정합성 유지 중.')
        return 0

    if not args.apply:
        print(f'\n[DRY-RUN] 총 {total}개 삭제 대상. 실제 삭제하려면 --apply 옵션 사용.')
        return 0

    # 4. 실제 삭제
    deleted = 0
    for key, _ in orphans_fwd:
        deleted += rc.delete(key)
    for key, _ in orphans_rev:
        deleted += rc.delete(key)
    print(f'\n✅ 삭제 완료: {deleted}개')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
