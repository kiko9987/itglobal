"""개인 원가 탭을 새 형식(본인 공사만·코드 고정·SB 는 원가 4칸만)으로 다시 만든다.

  python scripts/rebuild_personal_cost_tabs.py            # 비교만 (시트 안 바꿈)
  python scripts/rebuild_personal_cost_tabs.py --apply    # 백업 후 다시 만듦
  python scripts/rebuild_personal_cost_tabs.py --apply 박용구   # 한 탭만

옛 탭 값·수식·서식 정보는 ITG-Project/backups/ 에 JSON 으로 남는다. create_app 을 부르지 않는다.
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.stdout.reconfigure(encoding='utf-8')
from dotenv import load_dotenv  # noqa: E402

load_dotenv(os.path.join(ROOT, '.env'))

from dashboard.services.personal_cost_tabs import rebuild_tab, tab_names  # noqa: E402

BACKUP_DIR = os.path.join(os.path.dirname(ROOT), 'backups')

if __name__ == '__main__':
    apply = '--apply' in sys.argv
    names = [a for a in sys.argv[1:] if not a.startswith('--')] or tab_names()
    for t in names:
        rep = rebuild_tab(t, BACKUP_DIR, apply=apply)
        print(f"== {t}: 줄 {rep['rows']} · PM 에만 원가 {rep['main_only']} · 탭에만 원가 {len(rep['tab_only'])} "
              f"· 둘 다 있는데 다름 {len(rep['differ'])}")
        for code, o in rep['tab_only']:
            print(f'   탭에만 {code} {o}')
        for code, o, m in rep['differ']:
            print(f'   다름 {code} 탭 {o} / PM {m}')
        if rep.get('backup'):
            print('   백업:', rep['backup'])
