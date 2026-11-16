/**
 * 프로젝트 화면 위 「상태 칸」 — 당근 비즈니스 광고 관리의 `운영중 8개 / 심사거절 0개` 칸처럼
 * 상태별 공사 수를 보여 주고, 누르면 그 조건으로 바로 거른다(다시 누르면 해제). 2026-11-10.
 *
 * - 숫자는 지금 걸린 다른 필터(담당자·사업자·검색·내 공사만 등)는 그대로 두고,
 *   칸이 담당하는 두 조건(상태·세금계산서)만 빼고 센다 → 칸끼리 비교가 됨.
 * - 거르는 규칙은 ModernProjectFilters.filterRows 하나만 쓴다(목록과 숫자가 어긋나지 않게).
 * - 그룹웨어 새 화면(시범 명단)에서만 칸 자리(#projectStatusCards)가 그려진다. 없으면 아무것도 안 함.
 */
import ProjectStatusCalculator from '../utils/ProjectStatusCalculator.js';
import logger from '../utils/logger.js';
import { PAID_UNISSUED, PAID_UNISSUED_LABEL, matchesInvoice } from '../utils/invoiceFilter.js';

// key: 칸 이름표, kind: 어느 필터를 쓰는지, value: 그 필터 값, tone: 점 색
// ⚠️ 함수로 만든다: utils 묶음(청크)의 값을 모듈 맨 위에서 읽으면 components↔utils 청크 순환 때문에
//    "Cannot access before initialization"으로 목록 화면 전체가 멈춘다(2026-11-10 시험 서버에서 확인).
const cardDefs = () => [
  { key: '공사대기', kind: 'status', value: '공사대기', tone: 'wait', tip: '공사 시작 전인 공사(계약금을 받은 공사도 여기)' },
  // 입금이 있어도 시작일 전이면 공사대기(ProjectStatusCalculator 4번) — 11-17 안내 문구 맞춤
  { key: '공사진행', kind: 'status', value: '공사진행', tone: 'run', tip: '공사가 시작됐고 아직 끝나지 않은 공사' },
  { key: '수금필요', kind: 'status', value: '수금필요', tone: 'due', tip: '공사가 끝났는데 입금이 다 안 된 공사' },
  { key: '확인필요', kind: 'status', value: '확인필요', tone: 'check', tip: '금액·날짜가 비었거나 서로 맞지 않는 공사' },
  // 입금이 들어왔는데 세금계산서가 없는 공사만(입금 전 공사·현금(N입금)·카드는 안 셈). 2026-11-10 사용자 결정
  { key: '계산서 미발행', kind: 'invoice', value: PAID_UNISSUED, tone: 'bill', tip: '입금은 들어왔는데 세금계산서가 없는 공사' },
];

const SELECT_ID = { status: 'statusFilter', invoice: 'invoiceFilter' };
const selectOf = (kind) => document.getElementById(SELECT_ID[kind]);
// 칸을 연달아 누를 때 마지막 것만 거르기까지 기다리는 시간(ms) — 한 번 누를 땐 거의 못 느끼는 길이
const APPLY_DELAY = 140;

export default class ProjectStatusCards {
  constructor(container, filters) {
    this.el = container;
    this.filters = filters;
    this.render();
    this.el.addEventListener('click', (e) => this.onClick(e));
    filters.onFilterChange(() => this.update());
    this.update();
  }

  render() {
    this.cards = cardDefs();
    this.el.innerHTML = this.cards.map((c, i) => `
      <button type="button" class="gw-stat" data-i="${i}" title="${c.tip} — 누르면 이 공사만 보기" aria-pressed="false">
        <span class="gw-stat-dot gw-stat-${c.tone}" aria-hidden="true"></span>
        <span class="gw-stat-label">${c.key}</span>
        <span class="gw-stat-num"><b>-</b>건</span>
        <i class="fas fa-chevron-right gw-stat-go" aria-hidden="true"></i>
      </button>`).join('');
    this.buttons = [...this.el.querySelectorAll('.gw-stat')];
  }

  update() {
    const f = this.filters;
    if (!f || !f.isDataLoaded || !Array.isArray(f.currentData)) return;
    try {
      const base = { ...f.filters };
      delete base.status;
      delete base.invoice;
      const rows = f.filterRows(f.currentData, base);
      const counts = {};
      for (const r of rows) {
        const st = ProjectStatusCalculator.calculateStatus(r);
        counts['status:' + st] = (counts['status:' + st] || 0) + 1;
      }
      this.cards.forEach((c, i) => {
        const btn = this.buttons[i];
        // 계산서 칸은 목록 필터와 같은 함수로 센다(숫자 = 누르면 나오는 목록 수)
        const n = c.kind === 'invoice' ? rows.filter((r) => matchesInvoice(r, c.value)).length
          : (counts[c.kind + ':' + c.value] || 0);
        btn.querySelector('.gw-stat-num b').textContent = n.toLocaleString('ko-KR');
        btn.classList.toggle('zero', n === 0);
      });
      if (!this._applyTimer) this.syncActive();   // 칸을 눌러 거르기를 기다리는 중이면 방금 누른 표시를 지우지 않음
    } catch (err) {
      logger.error('[ProjectStatusCards] 숫자 계산 실패:', err);
    }
  }

  // 칸을 누르면: 그 조건으로 거르기(다른 칸 조건은 해제) / 이미 켜진 칸이면 해제.
  // 실제 select 값을 바꾸고 change 를 보내 기존 필터 동작(저장·URL·배지)을 그대로 탄다.
  // 11-17: 빠르게 여러 칸을 오가면 누를 때마다 4천여 건 거르기·표 다시 그리기가 쌓여 화면이 멈칫했다
  //   → 칸 표시·select 값은 바로 바꾸고, 무거운 거르기는 손을 멈춘 뒤(APPLY_DELAY) 한 번만.
  onClick(e) {
    const btn = e.target.closest('.gw-stat');
    if (!btn) return;
    const idx = Number(btn.dataset.i);
    const c = this.cards[idx];
    const turnOff = (selectOf(c.kind)?.value || '') === c.value;   // 아직 안 걸린(대기 중) 값까지 보고 판단
    for (const kind of Object.keys(SELECT_ID)) {
      const sel = selectOf(kind);
      if (!sel) continue;
      const value = kind === c.kind && !turnOff ? c.value : '';
      if (value && ![...sel.options].some(o => o.value === value)) {
        // 아직 목록에 없는 값(0건)이어도 고를 수 있게
        sel.add(new Option(value === PAID_UNISSUED ? PAID_UNISSUED_LABEL : value, value));
      }
      sel.value = value;
    }
    this.buttons.forEach((b, i) => {
      const on = !turnOff && i === idx;
      b.classList.toggle('active', on);
      b.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
    clearTimeout(this._applyTimer);
    this._applyTimer = setTimeout(() => this.applyPending(), APPLY_DELAY);
  }

  /** 기다리던 칸 선택을 실제 필터에 건다 — 바뀐 칸이 여럿이어도 거르기는 한 번 */
  applyPending() {
    this._applyTimer = null;
    const f = this.filters;
    const changed = Object.keys(SELECT_ID).filter((kind) => {
      const sel = selectOf(kind);
      return sel && (sel.value || '') !== (f.filters[kind] || '');
    });
    if (changed.length) {
      // 앞의 칸 값은 조용히 맞추고, 마지막 칸만 change 를 보내 한 번에 거른다(저장·URL·칩·편집 중 확인은 그 경로로)
      changed.slice(0, -1).forEach((kind) => {
        const v = selectOf(kind).value;
        if (v) f.filters[kind] = v; else delete f.filters[kind];
      });
      selectOf(changed[changed.length - 1]).dispatchEvent(new Event('change', { bubbles: true }));
    }
    this.syncActive();   // 편집 중 '이어서 편집'을 골라 필터가 되돌아간 경우에도 칸 표시를 실제 값에 맞춤
  }

  /** 켜진 칸 표시만 실제 필터 값에 맞춤(숫자는 그대로) */
  syncActive() {
    const f = this.filters;
    this.cards.forEach((c, i) => {
      const on = (f.filters[c.kind] || '') === c.value;
      this.buttons[i].classList.toggle('active', on);
      this.buttons[i].setAttribute('aria-pressed', on ? 'true' : 'false');
    });
  }
}
