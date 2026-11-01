"""#전자_결재 맨 위에 [💸 자금집행 요청] 메뉴 메시지를 올리고 고정 (설치 후 한 번, 2026-10-30).

전자결재 봇 설치·.env 설정·서버 재시작·채널 초대가 끝난 뒤 실행.
create_app() 을 부르지 않는다(스케줄러 이중 기동 방지) — 슬랙 처리 모듈만 쓴다.

  .venv/Scripts/python.exe scripts/eapproval_post_menu.py           # 미리보기
  .venv/Scripts/python.exe scripts/eapproval_post_menu.py --apply   # 실제로 올림
  .venv/Scripts/python.exe scripts/eapproval_post_menu.py --update 1793366788.175229   # 올린 메뉴 고치기(새 양식 추가 때)
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from dotenv import load_dotenv  # noqa: E402

load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '.env'))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--apply', action='store_true')
    ap.add_argument('--update', metavar='TS', help='이미 올린 메뉴 메시지 ts')
    args = ap.parse_args()

    from dashboard.services import e_approval as ea
    from dashboard.blueprints import slack_eapproval as se

    ch = ea.channel_id()
    print(f'채널: {ch or "(SLACK_EAPPROVAL_CHANNEL 없음)"} / 봇 설정: {se.is_configured()}')
    print(ea.menu_blocks()[0]['text']['text'])
    if args.update:
        se.update_menu(args.update)
        print(f'메뉴를 고쳤습니다: ts={args.update}')
        return
    if not args.apply:
        print('\n미리보기만 했습니다. 올리려면 --apply')
        return
    if not se.is_configured():
        sys.exit('SLACK_EAPPROVAL_BOT_TOKEN·SLACK_EAPPROVAL_CHANNEL 을 먼저 넣어 주세요.')
    ts = se.post_menu()
    print(f'올리고 고정했습니다: ts={ts}')


if __name__ == '__main__':
    main()
