# ECC(Everything Claude Code) 일부 설치 메모 (아이티글로벌)

- 출처: https://github.com/affaan-m/ECC 커밋 `4adeee6732a12d1540f49a4d50a38ce8c97f9a53` (2026-11-20, 2.2.3 이후), MIT (LICENSE 동봉).
- 2026-11-24 설치(KiKO 승인). 스킬 328개·에이전트 72개 중 우리 시스템에 맞고 기존 스킬과 겹치지 않는 3개만, **문서만** 고정 사본으로 가져옴. 같은 메모가 두 스킬 폴더에 하나씩 있음.
  - `.claude/agents/silent-failure-hunter.md` — 조용한 실패(오류 삼킴·성공처럼 보이는 실패·재시도 없는 버림) 점검 에이전트. 라이선스 `.claude/agents/LICENSE-ECC.txt`.
  - `.claude/skills/click-path-audit/` — 버튼 하나의 상태 변화를 끝까지 따라가 서로 되돌리는 버그 찾기.
  - `.claude/skills/change-impact/` — 바꾸기 전 영향 범위, 바꾼 뒤 실제 diff 대조.
- 플러그인·설치 프로그램(`npx ecc-universal`)으로 깔지 않은 이유: 훅(도구 쓸 때마다 Node 스크립트 자동 실행)이 같이 들어오고, 스킬 수백 개가 목록을 차지하며, 자동 업데이트됨.

## 원본에서 바꾼 것

- 세 파일 모두 끝에 '아이티글로벌 프로젝트' 절을 덧붙임(Flask·Redis·구글 시트·슬랙 카드에 맞춘 읽는 법, 한국어 보고, 읽기 전용).
- description 끝에 한국어 쓰임새 한 줄 추가(부를 때 잘 잡히도록).
- silent-failure-hunter: `model: sonnet` 줄 삭제(세션 모델을 따름).
- 본문 나머지는 원본 그대로.

## 일부러 뺀 것

- 훅 전부(gateguard·delivery-gate·config-health·자동 학습·기억 저장), MCP 설정, 설치 스크립트, 실행 파일.
- 기존 스킬과 겹치는 것(코드·보안 리뷰, 정리, 완료 전 검증, 접근성, 계획, SEO, 구글 문서), 우리 환경과 안 맞는 것(PostgreSQL·Node·Next.js·Stripe 전용, CI/CD, 영업 리드 발굴, git-workflow).
- 후보로 남겨 둔 것: production-audit(gw 주소 전환 전 점검), pr-test-analyzer(테스트가 바뀐 동작을 덮는지).

## 업데이트

- 새 커밋을 같은 방식(문서만, 훅 없이)으로 다시 검토해서 파일을 바꾸고, 덧붙인 절을 다시 붙인다.
