# 그룹웨어 주소 바꾸기 — pm.itg-aircon.com → gw.itg-aircon.com (2026-11-19 준비)

부장님 결정(11-19): 사람이 보는 주소는 **gw.itg-aircon.com**. **pm.itg-aircon.com 은 끄지 않는다** —
슬랙 앱 4개(메인·휴가·전자결재·문서변환) 이벤트 주소, 은행 문자 포워딩(/sms), 아임웹 측정(/track),
채널톡 웹훅(/channeltalk), 세콤(/secom)이 그대로 pm 으로 들어온다. 사람이 브라우저로 여는 화면만 gw 로 넘긴다.

## 이미 한 것 (11-19)

- 아임웹 DNS: `gw.itg-aircon.com` A `115.21.161.76` (부장님). 공용·호스트코코아·8.8.8.8 모두 확인.
- Caddy(`C:\Users\SECOM\caddy\Caddyfile`): 사이트 주소 줄을 `pm.itg-aircon.com, gw.itg-aircon.com {` 로.
  백업 `Caddyfile.bak-20261119-before-gw`. `caddy reload --address localhost:2019`(관리자 권한 없이) → gw 인증서 자동 발급,
  gw·pm `/health` 200. **아직 넘기기 규칙 없음** — 두 주소가 똑같이 열린다.
- 코드: 로그인 복귀 주소를 들어온 주소에 맞게(`auth._callback_uri`), 실시간 연결 허용 주소에 gw 추가(`allowed_origins`).
  `.env` 를 바꾸기 전까지는 동작이 예전과 같다(gw 에서 로그인하면 pm 으로 돌아가 실패 — 바꾸는 날 ②에서 풀림).

## 바꾸는 날 순서

1. **구글 클라우드 콘솔**(로그인용 OAuth 클라이언트, 권한 있는 분)
   - API 및 서비스 → 사용자 인증 정보 → 로그인에 쓰는 OAuth 2.0 클라이언트 ID
   - 승인된 JavaScript 원본에 `https://gw.itg-aircon.com` 추가
   - 승인된 리디렉션 URI에 `https://gw.itg-aircon.com/auth/callback` 추가 (pm 줄은 지우지 않음)
   - 저장 — 반영까지 몇 분 걸릴 수 있음
2. **`.env`**
   - `GOOGLE_OAUTH_REDIRECT_URIS=https://pm.itg-aircon.com/auth/callback,https://gw.itg-aircon.com/auth/callback`
   - `GW_PUBLIC_URL=https://gw.itg-aircon.com` (1회용 로그인 링크·휴가·입찰 알림 속 링크)
   - `GOOGLE_OAUTH_DEFAULT_REDIRECT_URI` 는 그대로 둠(목록에 없는 주소로 들어왔을 때만 씀)
3. **재시작**(부장님) → gw 에서 로그인·홈·프로젝트·휴가 열기, pm 에서도 로그인 되는지 확인
4. **Caddy 넘기기 규칙** — 사이트 블록 안, `header { ... }` 아래에 넣고 `caddy validate` → `caddy reload --address localhost:2019`

   ```caddy
   # pm 으로 들어온 '사람이 여는 화면'만 gw 로(11-19). 슬랙·웹훅·API·공개 회사 정보는 pm 그대로
   @pm_page {
       host pm.itg-aircon.com
       method GET
       header Accept *text/html*
       not path /slack/* /sms/* /track/* /channeltalk/* /api/* /secom/* /health* /static/* /auth/callback*
       not path /company /company.md /llms.txt /robots.txt /.well-known/* /favicon.ico /manifest.webmanifest /sw.js
   }
   redir @pm_page https://gw.itg-aircon.com{uri} 302
   ```

   - 처음 한두 주는 302(임시). 문제 없으면 301 로 바꿔도 됨.
   - 확인: `curl -I -H "Accept: text/html" https://pm.itg-aircon.com/projects` → 302 Location gw /
     `curl -I https://pm.itg-aircon.com/health` → 200 그대로 / 슬랙 버튼 한 번 눌러 보기.
5. **공지** — 주소 바뀜, 처음 한 번 다시 로그인, 휴대폰 홈 화면 아이콘은 지우고 새로 만들기
6. 시공·미화 휴대폰 등록(로그인 번호)은 **gw 주소에서** — 등록 쿠키가 주소마다 따로라 pm 에서 등록하면 다시 해야 함

## 직원이 겪는 일

- 로그인 쿠키는 주소마다 따로 → 모두 한 번 다시 로그인.
- 홈 화면에 설치한 앱(아이콘)은 예전 주소로 열림 → 지우고 새 주소로 다시 설치.
- 예전 슬랙 카드의 pm 링크·즐겨찾기는 gw 같은 화면으로 넘어감.

## 되돌리기

- Caddy: `Caddyfile.bak-20261119-before-gw` 로 덮고 reload (또는 넘기기 규칙 블록만 지우고 reload).
- `.env`: `GW_PUBLIC_URL` 지우기 → pm 기본값.
