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

// key: 칸 이름표, kind: 어느 필터를 쓰는지, value: 그 필터 값, tone: 점 색
const CARDS = [
  { key: '공사대기', kind: 'status', value: '공사대기', tone: 'wait', tip: '공사 시작 전인 공사' },
  { key: '공사진행', kind: 'status', value: '공사진행', tone: 'run', tip: '공사가 시작됐거나 입금이 들어온 공사' },
  { key: '수금필요', kind: 'status', value: '수금필요', tone: 'due', tip: '공사가 끝났는데 입금이 다 안 된 공사' },
  { key: '확인필요', kind: 'status', value: '확인필요', tone: 'check', tip: '금액·날짜가 비었거나 서로 맞지 않는 공사' },
  { key: '계산서 미발행', kind: 'invoice', value: '미발행', tone: 'bill', tip: '세금계산서 칸이 「미발행」인 공사' },
];

const SELECT_ID = { status: 'statusFilter', invoice: 'invoiceFilter' };

function invoiceStatus(item) {
  // 시트 Y열 = "{단계} - {상태}" → 상태 부분 (ModernProjectFilters 와 같은 규칙)
  return String(item['계산서'] || '').trim().split(' - ').pop().trim();
}

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
    this.el.innerHTML = CARDS.map((c, i) => `
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
        const inv = invoiceStatus(r);
        counts['invoice:' + inv] = (counts['invoice:' + inv] || 0) + 1;
      }
      CARDS.forEach((c, i) => {
        const btn = this.buttons[i];
        const n = counts[c.kind + ':' + c.value] || 0;
        btn.querySelector('.gw-stat-num b').textContent = n.toLocaleString('ko-KR');
        const on = (f.filters[c.kind] || '') === c.value;
        btn.classList.toggle('active', on);
        btn.setAttribute('aria-pressed', on ? 'true' : 'false');
        btn.classList.toggle('zero', n === 0);
      });
    } catch (err) {
      logger.error('[ProjectStatusCards] 숫자 계산 실패:', err);
    }
  }

  // 칸을 누르면: 그 조건으로 거르기(다른 칸 조건은 해제) / 이미 켜진 칸이면 해제.
  // 실제 select 값을 바꾸고 change 를 보내 기존 필터 동작(저장·URL·배지)을 그대로 탄다.
  onClick(e) {
    const btn = e.target.closest('.gw-stat');
    if (!btn) return;
    const c = CARDS[Number(btn.dataset.i)];
    const f = this.filters;
    const turnOff = (f.filters[c.kind] || '') === c.value;
    const setSelect = (kind, value) => {
      const sel = document.getElementById(SELECT_ID[kind]);
      if (!sel || sel.value === value) return;
      if (value && ![...sel.options].some(o => o.value === value)) {
        sel.add(new Option(value, value));   // 아직 목록에 없는 값(0건)이어도 고를 수 있게
      }
      sel.value = value;
      sel.dispatchEvent(new Event('change', { bubbles: true }));
    };
    for (const kind of Object.keys(SELECT_ID)) {
      if (kind !== c.kind) setSelect(kind, '');
    }
    setSelect(c.kind, turnOff ? '' : c.value);
  }
}
