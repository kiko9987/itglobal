# marketingskills 설치 메모 (아이티글로벌)

- 출처: https://github.com/coreyhaines31/marketingskills 커밋 `1f837447b20bb6f0e72d87b80ece1c221c362c7d` (2026-11-11, 2.14.12), MIT (LICENSE 동봉).
- 2026-11-17 설치(KiKO 승인). 스킬 53개 중 우리 일(상업용 냉난방 시공 B2B·광고·홈페이지)에 맞는 11개만, **문서만** 고정 사본으로 가져옴. 같은 메모가 11개 폴더에 하나씩 있음.
  - 문구: copywriting, copy-editing / 광고: ads, ad-creative / 검색 노출: ai-seo, seo-audit, schema / 전환·고객 이해: cro, customer-research, competitor-profiling, marketing-psychology
- 깃허브 플러그인으로 바로 설치하지 않은 이유: 자동 업데이트 방지, 필요 없는 미국 SaaS용 스킬 42개(결제·가입 흐름·이탈 방지·콜드 메일·앱스토어 등)가 같이 걸리는 것 방지.

## 일부러 뺀 것

| 원본 | 뺀 이유 |
|---|---|
| 각 스킬의 `evals/` | 원본 저장소의 자체 시험 데이터. 스킬 동작에 안 쓰임 |
| 나머지 스킬 42개, `tools/`(외부 서비스 연동 CLI), `scripts/` | 우리 일과 안 맞음 / 외부 실행 코드 |

- 가져온 11개에는 실행 파일·스크립트가 없음. 참고 문서의 `curl`·`npx` 예시는 설명일 뿐 자동 실행 없음.
- 유료 파트너 표시(◆, 예: Ploy) 2줄 있음(ads/references/abm-playbook.md, ai-seo/references/agent-readiness.md). 원본이 '추천 아님'이라 밝혀 둠 — 우리에게 권할 때 근거로 쓰지 말 것.

## 쓰는 법 (우리 규칙이 먼저)

- 회사 정보는 `.claude/product-marketing.md` 를 읽음(원본 스킬이 찾는 위치 중 하나). 사실이 바뀌면 그 파일과 copy-polish-ko '쓸 수 있는 사실', 광고 플레이북 메모리를 같이 고친다.
- 한국어 문구 결과물은 **copy-polish-ko 가 마지막 기준**(프로젝트 CLAUDE.md). 이 스킬들의 영어권 관행(긴급성·희소성 문구, 가격 앵커링, 할인 미끼, 'Start Free Trial' 식 버튼 등)이 우리 표현 규칙·표시광고법과 부딪치면 우리 규칙을 따른다.
- 광고 계정 변경은 여전히 KiKO 가 직접(클로드는 읽기 전용 자문). ads 스킬의 '조치' 제안도 제안까지만.
- 업데이트할 때는 새 커밋을 같은 방식(문서만)으로 다시 검토해서 폴더를 바꿈.
