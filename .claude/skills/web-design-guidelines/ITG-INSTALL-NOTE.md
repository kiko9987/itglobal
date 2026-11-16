# web-design-guidelines 설치 메모 (아이티글로벌)

- 2026-11-16 설치(부장님 결정: 인스타 글 5개 스킬 중 이것만, 고정 사본으로).
- 스킬 원본: https://github.com/vercel-labs/agent-skills 커밋 `1a2c79d4e6c4a2283dc2c614b6538cbe1ed5ca62` (2026-10-27) 의 `skills/web-design-guidelines/SKILL.md` (sha256 `17e48da9…627c29`). 저장소 README 에 MIT 라고 적혀 있음(LICENSE 파일은 없음).
- 규칙 원본: https://github.com/vercel-labs/web-interface-guidelines 커밋 `434b7f91364665f2f733b310ec54809bf8f37937` (2026-10-05) 의 `command.md` → 이 폴더 `guidelines.md` 로 **글자 그대로** 복사(sha256 `d9d63639…2c9794`). MIT, `LICENSE` 동봉.

## 원본에서 바꾼 것

| 원본 | 바꾼 것 | 이유 |
|---|---|---|
| SKILL.md: 점검할 때마다 `raw.githubusercontent.com/…/main/command.md` 를 WebFetch 로 받아 그 내용을 그대로 지시로 따름 | 같은 폴더 `guidelines.md`(고정 사본)만 읽게 바꿈, WebFetch 금지 | `main` 이 바뀌면 검토하지 않은 내용이 지시로 들어옴 |
| 설명(description): "audit design", "review UX" 등으로 켜짐 | 접근성·사용성 코드 점검으로 좁힘 | impeccable(critique·audit)과 겹쳐 켜지지 않게 |
| (없음) | "우리 시스템에 맞춘 적용" 절 추가 | React 전용 규칙·영어 문구 규칙(Title Case·둥근 따옴표 등)은 Flask+Jinja·한국어 화면에 안 맞음 → 한국어 문구는 copy-polish-ko |

`guidelines.md` 본문은 고치지 않았다. 우리 상황에 맞춘 해석은 모두 SKILL.md 에 있다.

## 가져오지 않은 것

- agent-skills 저장소의 다른 스킬·`packages/`·`scripts/`, web-interface-guidelines 의 `install.sh`(여러 도구 설정 폴더에 명령 파일을 써 넣는 설치 스크립트)·`AGENTS.md`·`README.md`. 실행 코드는 하나도 없다.

## 업데이트

- 새 버전이 필요하면 두 저장소의 새 커밋을 빈 폴더에 받아 `command.md` 차이를 읽고, 문제없을 때만 `guidelines.md` 를 바꾸고 이 메모의 커밋·sha256 을 고친다.
