# impeccable 설치 메모 (아이티글로벌)

- 출처: https://github.com/pbakaus/impeccable 커밋 `2b86d9c07b97d4bed7de23583e1bd1e58dd2baae` (2026-11-14, 스킬 버전 4.9.1), Apache 2.0 (LICENSE 동봉).
- 2026-11-15 설치. 원본 `.claude/skills/impeccable/` 중 **문서만** 가져옴: `SKILL.md`, `reference/*.md`.

## 일부러 뺀 것

| 원본 | 뺀 이유 |
|---|---|
| `scripts/impeccable`, `impeccable.cmd` | 처음 실행 때 GitHub 릴리스에서 실행 파일(engine 0.1.24)을 받아 `~/.impeccable/bin` 에 두고 실행함. 체크섬 확인은 하지만 외부 바이너리 실행이라 제외 |
| `scripts/live-browser*.js`, `modern-screenshot.umd.js` | live 모드에서 화면에 주입하는 스크립트. live 모드 [Apply]는 `codex`/`claude` 를 권한 확인 끈 상태(`--dangerously-bypass-approvals-and-sandbox`, `bypassPermissions`)로 실행할 수 있음 |
| `.claude/settings.json` (훅) | 세션 시작·파일 수정·종료 때마다 위 실행 파일을 돌리는 훅. 프로젝트 설정을 바꾸므로 제외 |
| `.claude/agents/*` | live·asset 제작용 보조 에이전트. 점검(critique·audit)에는 필요 없음 |

## 쓰는 법

- 런처가 없으므로 SKILL.md 의 "Launcher unavailable" 경로로 동작: PRODUCT.md·DESIGN.md 를 직접 읽고 `reference/` 지침을 따름.
- `impeccable detect`(규칙 62개 자동 검사)는 못 씀 → critique·audit 의 B 평가는 브라우저에서 직접 잰 값(대비·누를 자리 크기·가로 넘침·글자 크기)으로 대신함.
- `.impeccable/critique/` 스냅샷 저장도 런처 기능이라 하지 않음.
- 업데이트할 때는 새 커밋을 같은 방식(문서만)으로 다시 검토해서 바꿈.
