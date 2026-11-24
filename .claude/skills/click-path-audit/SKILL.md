---
name: click-path-audit
description: "Trace every user-facing button/touchpoint through its full state change sequence to find bugs where functions individually work but cancel each other out, produce wrong final state, or leave the UI in an inconsistent state. Use when: systematic debugging found no bugs but users report broken buttons, or after any major refactor touching shared state stores. 슬랙 카드 버튼·그룹웨어/PM 화면 단추가 눌러도 안 되거나 누른 뒤 상태가 꼬일 때."
metadata:
  origin: community
---

# /click-path-audit — Behavioural Flow Audit

Find bugs that static code reading misses: state interaction side effects, race conditions between sequential calls, and handlers that silently undo each other.

## The Problem This Solves

Traditional debugging checks:
- Does the function exist? (missing wiring)
- Does it crash? (runtime errors)
- Does it return the right type? (data flow)

But it does NOT check:
- **Does the final UI state match what the button label promises?**
- **Does function B silently undo what function A just did?**
- **Does shared state (Zustand/Redux/context) have side effects that cancel the intended action?**

Real example: A "New Email" button called `setComposeMode(true)` then `selectThread(null)`. Both worked individually. But `selectThread` had a side effect resetting `composeMode: false`. The button did nothing. 54 bugs were found by systematic debugging — this one was missed.

---

## How It Works

For EVERY interactive touchpoint in the target area:

```
1. IDENTIFY the handler (onClick, onSubmit, onChange, etc.)
2. TRACE every function call in the handler, IN ORDER
3. For EACH function call:
   a. What state does it READ?
   b. What state does it WRITE?
   c. Does it have SIDE EFFECTS on shared state?
   d. Does it reset/clear any state as a side effect?
4. CHECK: Does any later call UNDO a state change from an earlier call?
5. CHECK: Is the FINAL state what the user expects from the button label?
6. CHECK: Are there race conditions (async calls that resolve in wrong order)?
```

---

## Execution Steps

### Step 1: Map State Stores

Before auditing any touchpoint, build a side-effect map of every state store action:

```
For each Zustand store / React context in scope:
  For each action/setter:
    - What fields does it set?
    - Does it RESET other fields as a side effect?
    - Document: actionName → {sets: [...], resets: [...]}
```

This is the critical reference. The "New Email" bug was invisible without knowing that `selectThread` resets `composeMode`.

**Output format:**
```
STORE: emailStore
  setComposeMode(bool) → sets: {composeMode}
  selectThread(thread|null) → sets: {selectedThread, selectedThreadId, messages, drafts, selectedDraft, summary} RESETS: {composeMode: false, composeData: null, redraftOpen: false}
  setDraftGenerating(bool) → sets: {draftGenerating}
  ...

DANGEROUS RESETS (actions that clear state they don't own):
  selectThread → resets composeMode (owned by setComposeMode)
  reset → resets everything
```

### Step 2: Audit Each Touchpoint

For each button/toggle/form submit in the target area:

```
TOUCHPOINT: [Button label] in [Component:line]
  HANDLER: onClick → {
    call 1: functionA() → sets {X: true}
    call 2: functionB() → sets {Y: null} RESETS {X: false}  ← CONFLICT
  }
  EXPECTED: User sees [description of what button label promises]
  ACTUAL: X is false because functionB reset it
  VERDICT: BUG — [description]
```

**Check each of these bug patterns:**

#### Pattern 1: Sequential Undo
```
handler() {
  setState_A(true)     // sets X = true
  setState_B(null)     // side effect: resets X = false
}
// Result: X is false. First call was pointless.
```

#### Pattern 2: Async Race
```
handler() {
  fetchA().then(() => setState({ loading: false }))
  fetchB().then(() => setState({ loading: true }))
}
// Result: final loading state depends on which resolves first
```

#### Pattern 3: Stale Closure
```
const [count, setCount] = useState(0)
const handler = useCallback(() => {
  setCount(count + 1)  // captures stale count
  setCount(count + 1)  // same stale count — increments by 1, not 2
}, [count])
```

#### Pattern 4: Missing State Transition
```
// Button says "Save" but handler only validates, never actually saves
// Button says "Delete" but handler sets a flag without calling the API
// Button says "Send" but the API endpoint is removed/broken
```

#### Pattern 5: Conditional Dead Path
```
handler() {
  if (someState) {        // someState is ALWAYS false at this point
    doTheActualThing()    // never reached
  }
}
```

#### Pattern 6: useEffect Interference
```
// Button sets stateX = true
// A useEffect watches stateX and resets it to false
// User sees nothing happen
```

### Step 3: Report

For each bug found:

```
CLICK-PATH-NNN: [severity: CRITICAL/HIGH/MEDIUM/LOW]
  Touchpoint: [Button label] in [file:line]
  Pattern: [Sequential Undo / Async Race / Stale Closure / Missing Transition / Dead Path / useEffect Interference]
  Handler: [function name or inline]
  Trace:
    1. [call] → sets {field: value}
    2. [call] → RESETS {field: value}  ← CONFLICT
  Expected: [what user expects]
  Actual: [what actually happens]
  Fix: [specific fix]
```

---

## Scope Control

This audit is expensive. Scope it appropriately:

- **Full app audit:** Use when launching or after major refactor. Launch parallel agents per page.
- **Single page audit:** Use after building a new page or after a user reports a broken button.
- **Store-focused audit:** Use after modifying a Zustand store — audit all consumers of the changed actions.

### Recommended agent split for full app:

```
Agent 1: Map ALL state stores (Step 1) — this is shared context for all other agents
Agent 2: Dashboard (Tasks, Notes, Journal, Ideas)
Agent 3: Chat (DanteChatColumn, JustChatPage)
Agent 4: Emails (ThreadList, DraftArea, EmailsPage)
Agent 5: Projects (ProjectsPage, ProjectOverviewTab, NewProjectWizard)
Agent 6: CRM (all sub-tabs)
Agent 7: Profile, Settings, Vault, Notifications
Agent 8: Management Suite (all pages)
```

Agent 1 MUST complete first. Its output is input for all other agents.

---

## When to Use

- After systematic debugging finds "no bugs" but users report broken UI
- After modifying any Zustand store action (check all callers)
- After any refactor that touches shared state
- Before release, on critical user flows
- When a button "does nothing" — this is THE tool for that

## When NOT to Use

- For API-level bugs (wrong response shape, missing endpoint) — use systematic-debugging
- For styling/layout issues — visual inspection
- For performance issues — profiling tools

---

## Integration with Other Skills

- Run AFTER `/superpowers:systematic-debugging` (which finds the other 54 bug types)
- Run BEFORE `/superpowers:verification-before-completion` (which verifies fixes work)
- Feeds into `/superpowers:test-driven-development` — every bug found here should get a test

---

## Example: The Bug That Inspired This Skill

**ThreadList.tsx "New Email" button:**
```
onClick={() => {
  useEmailStore.getState().setComposeMode(true)   // ✓ sets composeMode = true
  useEmailStore.getState().selectThread(null)      // ✗ RESETS composeMode = false
}}
```

Store definition:
```
selectThread: (thread) => set({
  selectedThread: thread,
  selectedThreadId: thread?.id ?? null,
  messages: [],
  drafts: [],
  selectedDraft: null,
  summary: null,
  composeMode: false,     // ← THIS silent reset killed the button
  composeData: null,
  redraftOpen: false,
})
```

**Systematic debugging missed it** because:
- The button has an onClick handler (not dead)
- Both functions exist (no missing wiring)
- Neither function crashes (no runtime error)
- The data types are correct (no type mismatch)

**Click-path audit catches it** because:
- Step 1 maps `selectThread` resets `composeMode`
- Step 2 traces the handler: call 1 sets true, call 2 resets false
- Verdict: Sequential Undo — final state contradicts button intent

---

## 아이티글로벌 프로젝트에 맞춰 읽기

(설치 메모: affaan-m/ECC 커밋 `4adeee6732a12d1540f49a4d50a38ce8c97f9a53` (2026-11-20), MIT. 원본은 React·Zustand 예시라 이 절을 덧붙임. 같은 폴더 ITG-INSTALL-NOTE.md 참고)

| 원본 말 | 우리 시스템에서 |
|---|---|
| 상태 저장소(Zustand·Redux·context) | Redis 키(write-behind 큐, SET NX 선점 키, 캐시), 구글 시트 행·열, SQLite 표, **이미 나간 슬랙 카드의 블록 자체**(카드 문구·버튼도 상태) |
| 핸들러(onClick·onSubmit) | 슬랙 `block_actions`·`view_submission` 핸들러, Flask 라우트(`/api/...`), 그룹웨어·PM 화면 JS 의 fetch |
| 비동기 경합 | 폴러(카드 리컨사일러 10분 등)가 사람 클릭과 같은 카드·행을 다시 씀, 서버 재시작이 후속 단계를 자름, 같은 버튼 두 번 누르기 |
| 버튼 이름이 약속한 최종 상태 | 버튼을 누른 뒤 **장부(DB·시트)·카드 문구·DM·화면**이 모두 같은 상태를 말하는지 |

- Step 1 의 '저장소 지도'는 함수마다 **어떤 Redis 키·시트 열·DB 칸·카드를 읽고 쓰는지**, 부수로 지우거나 되돌리는 것이 있는지 적는다.
- 큰 범위는 화면(흐름)별로 나눠 서브에이전트에 맡길 수 있다. 저장소 지도를 먼저 만든다.
- 점검은 코드 읽기로 한다. 운영 슬랙·DB 에서 실제로 버튼을 누르지 않는다. 시험이 필요하면 테스트(REDIS_DB=15)나 운영 DB 사본 서버에서.
- 원본의 `/superpowers:...` 스킬은 우리에게 없다. 대신 code-review, verification-before-completion 을 쓴다.
- 보고는 한국어, 위치는 `파일:줄`.
