"""플렉스 결재(내 문서함 내보내기) → 그룹웨어 「플렉스 보관함」 표(flex_approval_archive) — 2026-11-10.

플렉스 해지(12/7) 전에 옛 결재를 우리 서버에서 읽기 전용으로 볼 수 있게 옮긴다.
- 원본: instance/flex_export/flex_workflow_export_*.json 중 가장 최근 것(--file 로 지정 가능)
- 카드번호 칸은 뒤 4자리만 남기고 저장. 다시 돌려도 같은 문서 번호는 덮어씀(중복 없음).
- 첨부 파일 자체는 옮기지 않는다(개수만).
- 인사카드가 없는 작성자가 있으면 쓰지 않는다(2026-11-15 보안 점검 5차 #9). 보관함은 작성자 이름으로만 주인을
  가려서, 카드 없는 옛 작성자와 이름이 같은 직원이 그 문서를 보게 된다. 카드를 먼저 만들고 다시 돌리거나,
  동명이인이 아님을 확인했으면 --allow-no-card.

사용:  .venv/Scripts/python.exe scripts/gw_import_flex_approvals.py            # 미리 보기(쓰지 않음)
       .venv/Scripts/python.exe scripts/gw_import_flex_approvals.py --apply    # 운영 사용자 DB 에 쓰기
       .venv/Scripts/python.exe scripts/gw_import_flex_approvals.py --db <경로> --apply
"""
import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--file', help='플렉스 내보내기 JSON(기본: 가장 최근 것)')
    ap.add_argument('--db', help='쓸 사용자 DB 경로(기본: 운영 instance/users.db)')
    ap.add_argument('--apply', action='store_true', help='실제로 쓰기')
    ap.add_argument('--allow-no-card', action='store_true', help='인사카드 없는 작성자가 있어도 쓰기(확인 뒤에만)')
    args = ap.parse_args()

    from dashboard.services import gw_approval as ga
    if args.db:
        from dashboard.utils.user_database import UserDatabase
        d = UserDatabase(db_path=args.db)
        ga._db = lambda: d
        from dashboard.services import hr
        hr._db = lambda: d          # 인사카드 확인(작성자 카드 있는지)도 같은 DB 에서

    path = args.file or ga.latest_export(str(ROOT / 'instance' / 'flex_export'))
    if not path or not os.path.exists(path):
        print('플렉스 내보내기 파일을 찾지 못했습니다.')
        return 1
    out = ga.import_flex_export(path, apply=args.apply, allow_no_card=args.allow_no_card)
    print(f"파일: {os.path.basename(path)} (내보낸 때 {out['exported_at']})")
    print(f"문서 {out['count']}건 — 상태 {out['by_status']}")
    print('양식(많은 순):', ', '.join(f'{k} {v}' for k, v in out['by_form'].items()))
    if out['no_card_writers']:
        print(f"⚠️ 인사카드 없는 작성자 {len(out['no_card_writers'])}명: {', '.join(out['no_card_writers'])}")
    if out['blocked']:
        print('→ 쓰지 않았습니다. 인사카드를 먼저 만들거나, 동명이인이 아님을 확인했으면 --allow-no-card 로 다시.')
        return 2
    print('→ 보관함에 썼습니다.' if args.apply else '→ 미리 보기만 했습니다(--apply 로 쓰기).')
    return 0


if __name__ == '__main__':
    sys.exit(main())
