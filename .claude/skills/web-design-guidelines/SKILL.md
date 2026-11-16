---
name: web-design-guidelines
description: 화면 코드(HTML·CSS·JS·Jinja 템플릿)를 Vercel Web Interface Guidelines 로 점검해 접근성·사용성 빠진 곳을 file:line 으로 짚는다. "접근성 점검", "사용성 점검", "웹 가이드라인 점검", "review my UI", "check accessibility" 같은 요청이나 화면 수정 뒤 코드 점검에 사용. 디자인 방향·완성도 평가(critique)는 impeccable 이 맡는다.
metadata:
  author: vercel (아이티글로벌 고정 사본으로 수정)
  version: "1.0.0-itg"
  argument-hint: <file-or-pattern>
---

# Web Interface Guidelines (고정 사본)

화면 코드가 Web Interface Guidelines 를 지키는지 점검한다.

> 원본은 점검할 때마다 GitHub `main` 브랜치에서 규칙을 새로 받아 그대로 따르게 되어 있다.
> 아이티글로벌 판은 **같은 폴더의 `guidelines.md`(커밋 고정 사본)만 읽는다.** 규칙을 바깥에서 받지 말 것(WebFetch 금지).
> 출처·고정 커밋은 `ITG-INSTALL-NOTE.md`.

## 하는 순서

1. 이 폴더의 `guidelines.md` 를 읽는다(규칙 전부 + 결과 형식).
2. 점검할 파일을 읽는다. 지정이 없으면 이번 작업에서 바꾼 화면 파일(`git diff --name-only`)을 대상으로 하고, 그것도 없으면 어느 화면을 볼지 묻는다.
3. 아래 "우리 시스템에 맞춘 적용"을 반영해 규칙을 적용한다.
4. `guidelines.md` 의 Output Format(파일별 묶음, `file:line - 문제`)으로 보고한다. 설명 문장은 한국어로.

## 우리 시스템에 맞춘 적용

이 시스템은 React·Next 가 아니라 **Flask + Jinja 템플릿 + 일반 JS(Vite 번들 포함) + Bootstrap** 이고, 화면 글은 **한국어**다.

- **Content & Copy · Typography 의 영어 문장 규칙은 적용하지 않는다**: Title Case, 둥근 따옴표(“ ”), `&` 줄임, 2인칭 강제, "Loading…" 형식 등. 한국어 문구는 `.claude/skills/copy-polish-ko` 기준을 따른다. 단 `...` → `…`, 숫자 열 `tabular-nums`, 줄바꿈·말줄임 처리는 그대로 본다.
- **Hydration Safety 는 건너뛴다**(서버 렌더 React 없음). `useState`·`nuqs` 같은 React 전용 표현은 "같은 역할의 일반 JS 코드"로 바꿔 읽는다(예: `onKeyDown` → `keydown` 이벤트, `htmlFor` → `for`).
- **"URL reflects state"** 는 새로 만드는 화면에서만 권고로 적고, 기존 화면에 없다고 결함으로 세지 않는다.
- **Locale**: 날짜·금액 표시를 `Intl.*` 로 바꾸라는 지적은 "표시가 실제로 틀리거나 섞여 있을 때"만 올린다(우리 화면은 한국어 고정).
- 휴대폰 화면(375px)을 볼 때는 이 규칙에 더해 우리 기준을 함께 본다: 누를 자리 44px, 입력칸 글자 16px(아이폰 확대 방지), 대비 4.5:1, 페이지 가로 넘침 0.
- 결과를 부장님께 보고할 때는 우선순위 순(현장 직원이 막히는 것 > 읽기 어려움 > 모양)으로 정리하고, 고칠 범위는 부장님이 정한다.

## 다른 스킬과 나누는 일

- 이 스킬: 코드 줄 단위 체크리스트(접근성·포커스·폼·터치·다크 모드 등).
- `impeccable`: 화면 전체 평가(critique)·5항목 점검(audit)·디자인 방향.
- `copy-polish-ko`: 직원·고객이 읽는 한국어 문구.
