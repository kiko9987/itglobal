# EGM(Elephant-Goldfish Model) 설치 메모 (아이티글로벌)

- 출처: https://github.com/google/elephant-goldfish-model 커밋 `72d06f52a2482885d9a64bd24b291c6e8628942c` (2026-11-15, 플러그인 버전 0.1.1), Apache 2.0 (LICENSE 동봉). 구글이 만들었지만 "공식 지원 제품 아님".
- 2026-11-17 설치. 이 폴더는 **고정 사본**이다. 깃허브에서 바로 설치하지 않으므로 새 버전이 저절로 들어오지 않는다.

## 가져온 것 (문서만)

`.claude-plugin/`(플러그인 정의·검토 담당 10명), `agents/`(검토 담당 설명 원본), `skills/`(스킬 7개), `resources/`(문서 기준), `docs/*.md`·`docs/designs/*.md`, `LICENSE`, `README.md`.
실행 파일·스크립트는 없다(.py·.js·.sh·.ps1 0개 확인).

## 일부러 뺀 것

| 원본 | 뺀 이유 |
|---|---|
| `site/`, `tests/`, `.github/`, `scripts/`, `adapters/` | 홈페이지·테스트·배포 스크립트. 플러그인이 불러오지 않음 |
| `.cursor-plugin/`, `.codex-plugin/`, `.windsurf/`, `com.github.copilot/`, `.gemini/`, `.agents/`, `.plugin/`, 루트 `plugin.json` | 다른 AI 도구용 |

## 검토한 안전성

- 훅·MCP 서버·명령 없음. `plugin.json` 은 검토 담당(agents) 목록만 등록.
- 검토 담당은 `tools: [Read, Grep, Glob]` — 읽기만 함. `omitClaudeMd: true` 라 우리 CLAUDE.md 를 일부러 안 봄(맥락 없는 눈으로 보는 게 목적).
- 스킬은 `disable-model-invocation: true` — Claude 가 혼자 시작하지 못하고 사람이 `/egm:egm-review 경로` 처럼 불러야 함.

## 쓰는 법·주의

- 쓸 곳: 큰 설계 전 설계서 검토(`/egm:egm-review`), 문서와 코드 어긋남 찾기(`/egm:egm-drift-report docs/…`).
- 비용이 큼(단계마다 새 에이전트 여러 명). 작은 수정에는 쓰지 않는다.
- 스킬이 "결과물은 쉬운 영어로"를 강제함 → 부장님께는 한국어로 옮겨 보고.
- 결과 파일(`design_draft.md`, `*_report.md`)을 작업 폴더에 바로 씀 → 쓸 때마다 `docs/` 아래로 위치 지정, 경로 지정 커밋.
- 설치(이 PC, 사용자 범위):
  `claude plugin marketplace add "<저장소>/tools/claude-plugins/egm"` → `claude plugin install egm@egm`.
  설치 때 사본이 `~/.claude/plugins/cache` 로 복사되므로, 이 폴더를 바꾼 뒤에는 `claude plugin marketplace update egm` 로 다시 반영.
- 업데이트할 때는 새 커밋을 같은 방식(문서만)으로 다시 검토해서 이 폴더를 바꾼다.
