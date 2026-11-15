# PM 화면 휴대폰 리뉴얼 계획 (2026-11-15)

부장님 결정(11-15): "PM은 모바일을 전혀 고려 안 하고 만들었다 → UI를 모바일용으로 만져서 리뉴얼". 사용 기록은 볼 필요 없음(휴대폰으로 쓴 적이 없어 기록이 없다).
그룹웨어 화면(홈·휴가·결재함·인사)은 같은 날 휴대폰 정리를 마쳤다 — 이 문서는 **예전 PM(Vite) 화면**용.

## 1. 원칙

1. **PC 화면은 바꾸지 않는다.** 모든 변경은 `@media (max-width: 767.98px)` 안, 또는 휴대폰 폭에서만 도는 JS 분기. 단계마다 1280px 화면이 고치기 전과 같은지 확인(스크린샷 비교).
2. **그룹웨어 새 틀(body.gw-shell) 기준**으로 만든다(정식 오픈 뒤 전원이 이 틀). 예전 틀(modern_base 머리글)은 **지금보다 나빠지지만 않게**.
3. 휴대폰에서 할 일 = **찾아보기 + 자주 하는 간단한 처리**(전화 걸기, 주소 지도, 상태 보기, 메모 보기). 금액·계산서·일괄 편집 같은 무거운 편집은 "되기는 하게"(창이 화면 가득, 칸이 안 깨짐)까지만, 모양은 PC 그대로.
4. 모양 규칙은 그룹웨어와 같게(당근 비즈니스 따라가기): 위쪽 = 제목 한 줄 + 주 단추 하나(나머지는 ⋯ 메뉴), 필터 = 검색 칸 + [필터 n] 단추로 펼치기, 목록 = 카드, 카드를 누르면 상세가 화면 가득. 누를 자리 44px, 입력칸 글자 16px(아이폰 확대 방지), 대비 4.5:1.
   - 참고 구현: `dashboard/static/css/gw-pages.css` 끝부분(결재함 `#apRows` 카드·`.gw-chip-more` 필터 접기), `dashboard/static/js/gw-approval.js` 앞부분(필터 접기 JS).
5. 새로 쓰는 직원용 문구는 `.claude/skills/copy-polish-ko` 기준.

## 2. 지금 상태(운영 375px 실측, 11-15)

| 화면 | 문제 |
|---|---|
| 프로젝트 `/projects` | 제목 '프로젝트관리'가 한 글자씩 세로로 쌓임 · [새로고침][캘린더 동기화][+새 프로젝트] 단추가 화면 밖(438px)으로 잘림 · 상태 칩(공사대기 23건 등) 글자 잘림 · 필터 10개가 5줄 · '내 프로젝트만' 줄이 521~648px 로 넘침 → **페이지 가로 넘침** · **기존 `MobileCardView`(768px 이하에서 켜지게 만든 카드 목록)가 카드 0장**, 컨테이너가 페이지 맨 아래(4,900px)에 있고 PC 표(23줄)가 그대로 보임 |
| 고객 리드 `/leads` | 제목 '고객 리/드 관리' 꺾임 · [테이블/캘린더] 단추 글자 세로 · DataTables 표 칸 글자가 서로 겹침(읽을 수 없음) · 필터 3줄 |
| 월별 수금 `/settlement/collected` | '정산 월'·'기간' 이름표 두 줄로 꺾임 · 표 가로 스크롤(정도 약함) |
| 매출통계 `/stats`, 시공사례 `/case-posts` | 괜찮음(확인만) |

## 3. 코드 지도

| 화면 | 템플릿 | JS(Vite, `dashboard/src/js`) | CSS |
|---|---|---|---|
| 프로젝트 | `templates/modern_project_list.html`(611줄, 머리·필터·표·`mobile-card-container` 마크업) | `pages/project-list.js`, `components/ProjectTable.js`, `ModernProjectFilters.js`, `ProjectRowAccordion.js`(7,975줄 — 상세), `MobileCardView.js`(872줄), `ModernProjectModal.js`(새·수정 창), `FieldMemoButton.js`, `InvoiceLicense.js`, `AsView.js` | `src/css/components/mobile-card-view.css`, `responsive.css`, `main.css` |
| 고객 리드 | `templates/leads.html`(159줄) | `pages/leads.js`, `components/LeadTable.js`, `ModernLeadsFilters.js`, `LeadRowAccordion.js`, FullCalendar(캘린더 보기) | `src/css/pages/leads.css` |
| 월별 수금 | `templates/collected_report.html`(인라인 스크립트·스타일) | — | 템플릿 안 |
| 공통 틀 | `templates/modern_base.html`, `components/_gw_shell.html` | — | `static/css/groupware-shell.css`(body.gw-shell 한정 — PM 화면에도 걸림) |

- 휴대폰 링크 주의: 폴더 열기 `itgfolder://` 는 **윈도우 PC 전용**(utils/itgfolder.js) → 휴대폰에선 숨기거나 드라이브 웹 링크로.

## 4. 빌드·배포(반드시)

- Vite 소스 변경은 **빌드해야 반영**. `dashboard/` 에서 `NODE_ENV=production npm run build` (안 하면 동적 import 경로 깨짐).
- 빌드 출력 `dashboard/static/dist` 는 **git 에 커밋돼 있다** → worktree 에서 빌드 → dist 까지 커밋 → main ff 병합·push.
  먼저 임시 outDir(`--outDir <tmp> --emptyOutDir`)로 빌드 성공 확인 후 실제 빌드.
- worktree 엔 `node_modules` 가 없을 수 있다 → `dashboard/` 에서 `npm ci`(또는 main 의 node_modules 를 쓰는 방법 확인).
- JS 번들은 manifest mtime 으로 재시작 없이 반영되지만 **템플릿(HTML)을 바꾸면 재시작 필요**("재시작 필요" 한 줄만, 재시작은 부장님이).
- 줄 끝: 저장소 파일은 CRLF. 파이썬으로 고칠 땐 CRLF 유지(11-15 에 LF 로 바꿔 diff 가 수천 줄 됐던 일).

## 5. 시험 방법

- 운영 사이트는 **보기만**(부장님이 내장 브라우저에 로그인해 둔 상태면 375px·다크·1280px 로 보기 가능). 저장·처리 단추는 누르지 않는다.
- 고친 코드는 **시험 서버**에서 먼저: 메모리 `project_groupware_plan` 끝 "시험 서버 만드는 법"(운영 users.db mode=ro backup, 소켓 127.0.0.1 만, 슬랙 클라이언트만 가짜, before_request 를 field_gatekeeper 보다 앞에, /api 404 는 JSON). 프로젝트·리드 화면은 데이터가 필요 → 운영 Redis `cache:current_sheet_data`(`__BYTES__:`+base64 pickle)·리드 캐시를 **GET 만** 해서 사본 파일로(이전 하네스가 80건 사본 사용). 스크립트에서 `create_app()` 금지.
- 점검 도구: `.claude/skills/impeccable`(critique·audit, 문서만 설치 — 엔진 없음, 브라우저 실측으로 대신). 단계마다 375px·다크·1280px.
- `tests/unit` 전체 pytest 통과, 끝나면 시험 DB 사본 삭제.

## 6. 단계

### 0단계 — 준비
- 시험 서버(프로젝트·리드 데이터 사본 포함), worktree 빌드 환경, 1280px 기준 스크린샷(PC 회귀 비교용).

### 1단계 — 프로젝트 (가장 많이 씀)
1. 머리: 제목 한 줄, [+ 새 프로젝트] 하나만 보이고 [새로고침]·[캘린더 동기화]는 ⋯ 메뉴로.
2. 상태 칩(공사대기·공사진행·수금필요·확인필요·계산서 미발행): 한 줄 가로 스크롤 또는 2칸 격자, 글자 안 잘리게.
3. 필터: 검색 칸 한 줄 + [필터 n] 단추로 10개 칩 펼치기, '내 프로젝트만'·건수 줄 넘침 고치기, [프리셋 저장]은 펼친 안으로.
4. 목록: **MobileCardView 를 살린다**(왜 0장인지·왜 표가 안 숨는지·컨테이너 위치부터). 카드 = 상태 배지·코드 / 현장(거래처) / 금액·미수금 / 담당·날짜. 페이지 가로 넘침 0.
5. 상세: 카드를 누르면 ProjectRowAccordion 내용을 화면 가득(아래에서 올라오는 시트 또는 전체 화면). 고객 연락처 `tel:` 전화, 주소 지도 링크, 폴더 열기는 휴대폰에서 숨김.
6. 새·수정 창(ModernProjectModal): 휴대폰에선 화면 가득, 칸 1열.
7. 부장님 실기기 확인 → 반영.

### 2단계 — 고객 리드
1. 머리·[테이블/캘린더] 전환 정리(아이콘+글자 한 줄).
2. 필터 접기(1단계와 같은 방식).
3. 목록: 휴대폰에선 DataTables 대신 카드(리드 No·상태 / 고객명·플랫폼 / 방문 주소·방문일 / 상담자·영업 담당), 전화번호 `tel:`.
4. 카드 → LeadRowAccordion 화면 가득.
5. 캘린더: 휴대폰에선 FullCalendar 목록 보기(listWeek) 등 좁은 폭용 보기.
6. 실기기 확인 → 반영.

### 3단계 — 나머지
- 월별 수금 이름표 꺾임·표, 매출통계·시공사례 재확인, 공통 창(UserModal·AuditLog·InvoiceLicense·AsView·FieldMemo) 화면 가득, 예전 틀 머리글 점검.

### 각 단계 끝
- impeccable critique·audit(375·다크·1280) 결과표 → 고친 범위 보고 → 부장님 확인 → 빌드·커밋·main 병합·push(ls-remote 비교) → 템플릿 바꿨으면 "재시작 필요".

## 7. 정할 것(부장님)
- 휴대폰에서 금액·상태 **편집까지** 할지, 보기 + 간단 처리만 할지(추천: 보기 + 상태 바꾸기·메모까지, 금액은 PC).
- 프로젝트 카드에 꼭 보여야 할 칸(추천안: 상태·코드·현장·거래처·총액·미수금·담당·공사일).
