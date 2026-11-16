import logger from '../utils/logger.js';
import { partnerDisplayName } from '../utils/partnerName.js';
import { escHtml } from '../utils/escapeHtml.js';
import ProjectStatusCalculator from '../utils/ProjectStatusCalculator.js';
import MobileDetailSheet from './MobileDetailSheet.js';

/**
 * 휴대폰 프로젝트 카드 목록 + 상세 화면 (2026-11-15 휴대폰 리뉴얼 1단계).
 *
 * - 휴대폰 폭(767.98px 이하)에서만 켜진다. 넓은 화면에선 컨테이너를 숨겨 PC 표 그대로.
 * - 목록은 페이지 필터(ModernProjectFilters)가 거른 결과를 그대로 받는다(따로 검색·정렬 안 함).
 *   예전 이 파일은 템플릿과 다른 id(#mobile-card-container)를 찾아 컨테이너를 따로 만들고, DataTables 2 의
 *   바깥 틀(.dt-container)을 몰라 표를 못 숨겼고, 카드도 0장이었다(updateUI 가 빈 배열로 덮어씀) → 다시 씀.
 * - 카드를 누르면 기존 펼침(ProjectRowAccordion)을 화면 가득 시트 안에서 연다. 펼침 코드는 그대로 쓰고,
 *   '표 행 다음에 붙이기'만 시트 안의 빈 행에 붙게 바꾼다(휴대폰일 때만). 편집·저장·취소도 PC 와 같은 코드.
 * - 휴대폰 뒤로 가기 = 시트 닫기(history 한 칸).
 */
const MOBILE_MQ = '(max-width: 767.98px)';
const PAGE_SIZE = 30;

const STATUS_TONE = {
  공사대기: 'wait',
  공사진행: 'run',
  수금필요: 'due',
  확인필요: 'check',
  공사완료: 'done',
  공사취소: 'cancel',
};

/** 미수금 색: 받을 돈이 남았으면 빨강. 단 공사대기는 아직 받을 때가 아니라 보통 글자(밀린 돈처럼 보이던 것) */
function dueClass(due, status) {
  if (!(due > 0)) return 'pmm-due-zero';
  return status === '공사대기' ? 'pmm-due-later' : 'pmm-due';
}

function won(n) {
  return `${Math.round(n).toLocaleString('ko-KR')}원`;
}

function num(v) {
  const n = parseFloat(String(v ?? '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function shortDate(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? { y: m[1], md: `${m[2]}.${m[3]}` } : null;
}

/** 공사일: 2026.11.16 / 2026.11.16 ~ 11.30 / 2026.12.30 ~ 2027.01.05 / 미정 */
function workPeriod(start, end) {
  const a = shortDate(start);
  const b = shortDate(end);
  if (!a && !b) return '공사일 미정';
  if (!a || !b) return (a || b).y + '.' + (a || b).md;
  if (a.y === b.y && a.md === b.md) return `${a.y}.${a.md}`;
  return a.y === b.y ? `${a.y}.${a.md} ~ ${b.md}` : `${a.y}.${a.md} ~ ${b.y}.${b.md}`;
}

function telHref(raw) {
  const digits = String(raw || '').replace(/[^0-9+]/g, '');
  return digits.length >= 8 ? `tel:${digits}` : '';
}


/** 날짜 칸('2026-11-17', "'2026-11-17", '2026-11-17 ~ 11-20')에서 첫 날짜만 */
function firstDate(raw) {
  const m = String(raw || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : '';
}

/** '가까운 순': 오늘 이후(가까운 날 먼저) → 지난 날(최근 먼저) → 날짜 없음 */
function nearFirst(a, b, today) {
  if (!a || !b) return (a ? 0 : 1) - (b ? 0 : 1);
  const fa = a >= today, fb = b >= today;
  if (fa !== fb) return fa ? -1 : 1;
  if (fa) return a < b ? -1 : a > b ? 1 : 0;
  return a > b ? -1 : a < b ? 1 : 0;
}

function sortStore(key, value) {
  try {
    if (value === undefined) return localStorage.getItem(key) || '';
    localStorage.setItem(key, value);
  } catch (_) { /* 기억 못 해도 정렬은 됨 */ }
  return value || '';
}

const SORTS = [
  ['', '최신순'],
  ['start', '공사일 가까운 순'],
  ['due', '미수금 많은 순'],
  ['total', '총액 큰 순'],
];

export default class MobileCardView {
  constructor() {
    this.mq = window.matchMedia(MOBILE_MQ);
    this.currentData = [];
    this.shown = PAGE_SIZE;
    this.lastCardCode = null;
    this.sortKey = sortStore('pmm.sort.projects');
    if (!SORTS.some(([k]) => k === this.sortKey)) this.sortKey = '';

    this.createContainer();
    this.createSheet();
    this.bindEvents();
    this.sync();
  }

  get sheetOpen() {
    return !!this.sheet?.isOpen;
  }

  /** 휴대폰 폭인지 (바깥에서 쓰던 이름 유지) */
  get isActive() {
    return this.mq.matches;
  }

  createContainer() {
    let el = document.getElementById('mobileCardContainer');
    if (!el) {
      el = document.createElement('section');
      el.id = 'mobileCardContainer';
      const section = document.querySelector('.table-section');
      (section?.parentNode || document.body).insertBefore(el, section?.nextSibling || null);
    }
    el.classList.add('pmm-list');
    el.setAttribute('aria-label', '프로젝트 목록');
    el.innerHTML = `
      <h2 class="visually-hidden">프로젝트 목록</h2>
      <div class="pmm-sort">
        <label for="pmmSortProjects">정렬</label>
        <select id="pmmSortProjects" class="pmm-sort-select">
          ${SORTS.map(([k, t]) => `<option value="${k}">${t}</option>`).join('')}
        </select>
      </div>
      <div class="pmm-cards"></div>
      <p class="pmm-empty" hidden>조건에 맞는 공사가 없어요. 검색어나 필터를 바꿔 보세요.</p>
      <button type="button" class="pmm-more-btn" hidden></button>`;
    this.container = el;
    this.listEl = el.querySelector('.pmm-cards');
    this.emptyEl = el.querySelector('.pmm-empty');
    this.moreBtn = el.querySelector('.pmm-more-btn');
    this.sortEl = el.querySelector('.pmm-sort-select');
    this.sortEl.value = this.sortKey;
  }

  /** 고른 정렬로 줄 세움(같으면 원래 순서 = 최신순) */
  sortedRows() {
    const rows = this.currentData;
    const k = this.sortKey;
    if (!k) return rows;
    const today = new Date().toISOString().slice(0, 10);
    const val = (r) => (k === 'start' ? firstDate(r['공사 시작'])
      : k === 'due' ? num(r['미수금']) : num(r['총액 2'] ?? r['총액2']));
    return rows.map((r, i) => [r, i, val(r)]).sort((a, b) => {
      const c = k === 'start' ? nearFirst(a[2], b[2], today) : (b[2] - a[2]);
      return c || a[1] - b[1];
    }).map((x) => x[0]);
  }

  createSheet() {
    this.sheet = new MobileDetailSheet({ id: 'pmmSheet', onRequestClose: (fromPop) => this.requestClose(fromPop) });
    this.anchorRow = this.sheet.anchorRow;
  }

  bindEvents() {
    const onMq = () => this.sync();
    if (this.mq.addEventListener) this.mq.addEventListener('change', onMq);
    else this.mq.addListener(onMq);

    this.listEl.addEventListener('click', (e) => {
      const card = e.target.closest('.pmm-card');
      if (card) this.openDetail(card.dataset.code);
    });
    this.listEl.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      const card = e.target.closest('.pmm-card');
      if (!card) return;
      e.preventDefault();
      this.openDetail(card.dataset.code);
    });
    this.moreBtn.addEventListener('click', () => {
      this.shown += PAGE_SIZE;
      this.render();
    });
    this.sortEl.addEventListener('change', () => {
      this.sortKey = sortStore('pmm.sort.projects', this.sortEl.value);
      this.shown = PAGE_SIZE;
      this.render();
    });

    // 메모 저장 등 부분 갱신 → 같은 객체를 고친 경우가 많아 카드만 다시 그림
    document.addEventListener('projectUpdated', () => {
      if (this.isActive) this.render();
    });
  }

  /** 휴대폰 ↔ 넓은 화면 전환 */
  sync() {
    if (this.isActive) {
      this.container.style.display = '';
      this.render();
    } else {
      this.container.style.display = 'none';
      if (this.sheetOpen) this.requestClose();
    }
  }

  /** 필터 결과가 바뀔 때마다 project-list.js 가 부른다 */
  updateCards(data) {
    this.currentData = Array.isArray(data) ? data : [];
    this.shown = PAGE_SIZE;
    if (this.isActive) this.render();
  }

  render() {
    if (!this.isActive) return;
    const rows = this.sortedRows();
    const list = rows.slice(0, this.shown);
    this.listEl.innerHTML = list.map((r) => this.cardHtml(r)).join('');
    this.emptyEl.hidden = rows.length > 0;
    const rest = rows.length - list.length;
    this.moreBtn.hidden = rest <= 0;
    if (rest > 0) this.moreBtn.textContent = `더 보기 (${rest.toLocaleString('ko-KR')}건 남음)`;
  }

  cardHtml(r) {
    const code = String(r['프로젝트 코드'] || '');
    let status = '';
    try {
      status = ProjectStatusCalculator.calculateStatus(r) || '';
    } catch (err) {
      logger.warn('[MobileCardView] 상태 계산 실패:', err);
    }
    const tone = STATUS_TONE[status] || 'wait';
    const total = num(r['총액 2'] ?? r['총액2']);
    const due = num(r['미수금']);
    const partner = partnerDisplayName(r);
    const addr = String(r['현장 주소'] || '').trim();
    const manager = String(r['담당자'] || '').trim();
    const cancelled = status === '공사취소';
    const dueCls = dueClass(due, status);

    return `
      <div class="pmm-card${cancelled ? ' is-cancelled' : ''}" role="button" tabindex="0"
           data-code="${escHtml(code)}" aria-label="${escHtml(`${code} ${status} ${addr}`)} 상세 보기">
        <div class="pmm-card-top">
          <span class="pmm-status pmm-st-${tone}"><span class="pmm-dot" aria-hidden="true"></span>${escHtml(status || '상태 없음')}</span>
          <span class="pmm-code">${escHtml(code)}</span>
        </div>
        <div class="pmm-addr">${escHtml(addr || '현장 주소 없음')}</div>
        ${partner && partner !== '-' ? `<div class="pmm-partner">${escHtml(partner)}</div>` : ''}
        <div class="pmm-money">
          <span>총액 <b>${total > 0 ? won(total) : '-'}</b></span>
          <span class="${dueCls}">미수금 <b>${won(due)}</b></span>
        </div>
        <div class="pmm-meta">
          <span>${escHtml(manager || '담당 미정')}</span>
          <span aria-hidden="true">·</span>
          <span>${escHtml(workPeriod(r['공사 시작'], r['공사 종료']))}</span>
        </div>
      </div>`;
  }

  findProject(code) {
    return this.currentData.find((p) => p['프로젝트 코드'] === code)
      || (window.projectsData || []).find((p) => p['프로젝트 코드'] === code);
  }

  /**
   * 실제로 쓰는 펼침 = 표(ProjectTable)가 만든 것(window.projectRowAccordion).
   * (예전엔 projectListApp.components.accordion 이 따로 있었으나 init 안 된 채 안 쓰여 11-16 에 뺌)
   */
  liveAccordion() {
    return window.projectRowAccordion || window.projectListApp?.components?.table?.accordion || null;
  }

  accordion() {
    const acc = this.liveAccordion();
    if (acc) this.patchAccordion(acc);
    return acc;
  }

  /**
   * 펼침이 표 행 대신 시트 안의 빈 행에 붙게 — 휴대폰일 때만. 인스턴스 메서드만 감싸서 PC 동작은 그대로.
   * (코드가 바뀌어 다시 열 때 reopenAccordion 도 openAccordion 을 거치므로 시트 안으로 온다)
   */
  patchAccordion(acc) {
    if (acc.__pmmPatched) return;
    acc.__pmmPatched = true;

    const open = acc.openAccordion.bind(acc);
    acc.openAccordion = (row, data) => {
      if (this.isActive && data && row !== this.anchorRow) {
        this.showSheet(data);
        return open(this.anchorRow, data);
      }
      if (this.isActive && data) this.showSheet(data);
      return open(row, data);
    };

    const close = acc.closeAccordion.bind(acc);
    acc.closeAccordion = (...args) => {
      const wasInSheet = this.sheetOpen;
      const r = close(...args);
      if (wasInSheet) {
        // 닫기 정리는 원래 '닫힘 애니메이션 끝(animationend)' 또는 1초 뒤에 .accordion-row 를 '전부' 지운다.
        // 시트는 바로 숨겨져 애니메이션이 안 끝나므로 1초를 기다리게 되고, 그 안에 다른 카드를 열면
        // 새 상세까지 지워져 빈 시트가 됐다(11-15 점검) → 정리를 지금 바로 돌린다(뒤의 1초 정리는 할 일 없음).
        acc.accordionContainer?.dispatchEvent(new Event('animationend'));
        this.hideSheet();
      }
      return r;
    };

    const reset = acc.hardResetAccordionState.bind(acc);
    acc.hardResetAccordionState = (...args) => {
      const r = reset(...args);
      if (this.sheetOpen) this.hideSheet();
      return r;
    };

    // 코드가 바뀌면 PC 는 펼침을 새 표 행 아래로 옮긴다 → 시트에선 그 자리에 두고 제목만 바꿈
    const move = acc.moveAccordionRow.bind(acc);
    acc.moveAccordionRow = (newCode) => {
      if (this.sheetOpen) {
        const row = acc.accordionContainer?.closest('.accordion-row');
        if (row && this.sheet.el.contains(row)) {
          row.dataset.projectCode = newCode;
          acc.accordionContainer.classList.add('show');
          acc.isOpen = true;
          this.sheet.setTitle(newCode);
          return true;
        }
      }
      return move(newCode);
    };
  }

  openDetail(code) {
    const data = this.findProject(code);
    const acc = this.accordion();
    if (!data || !acc) {
      logger.warn('[MobileCardView] 상세를 열 수 없음:', code, !!data, !!acc);
      return;
    }
    this.lastCardCode = code;
    acc.openAccordion(this.anchorRow, data);
  }

  showSheet(data) {
    const code = String(data['프로젝트 코드'] || '');
    const card = [...this.listEl.querySelectorAll('.pmm-card')].find((c) => c.dataset.code === this.lastCardCode);
    this.sheet.show({ title: code, quickHtml: this.quickHtml(data), returnFocus: card || null });
  }

  /** 시트 위: 요약(상태·총액·미수금 — 상세 아래쪽까지 안 내려가도 보이게) + 바로 하기(발주처 전화, 현장 지도) */
  quickHtml(data) {
    const tel = telHref(data['발주처 연락처']);
    const who = String(data['발주처 담당자'] || '').trim();
    const addr = String(data['현장 주소'] || '').trim();
    let status = '';
    try { status = ProjectStatusCalculator.calculateStatus(data) || ''; } catch (_) { /* 요약만 빠짐 */ }
    const total = num(data['총액 2'] ?? data['총액2']);
    const due = num(data['미수금']);
    const parts = [`
      <div class="pmm-sheet-sum">
        <span class="pmm-status pmm-st-${STATUS_TONE[status] || 'wait'}"><span class="pmm-dot" aria-hidden="true"></span>${escHtml(status || '상태 없음')}</span>
        <span class="pmm-sum-money">총액 <b>${total > 0 ? won(total) : '-'}</b></span>
        <span class="pmm-sum-money ${dueClass(due, status)}">미수금 <b>${won(due)}</b></span>
      </div>`];
    if (tel) {
      parts.push(`<a class="pmm-quick-btn" href="${escHtml(tel)}">
        <i class="fas fa-phone" aria-hidden="true"></i>${who ? `발주처 전화 · ${escHtml(who)}` : '발주처 전화'}</a>`);
    }
    if (addr) {
      const url = `https://map.kakao.com/link/search/${encodeURIComponent(addr)}`;
      parts.push(`<a class="pmm-quick-btn" href="${escHtml(url)}" target="_blank" rel="noopener">
        <i class="fas fa-map-marker-alt" aria-hidden="true"></i>지도에서 보기</a>`);
    }
    return parts.join('');
  }

  /** 닫기 요청(뒤로 단추·휴대폰 뒤로 가기). 고치던 내용이 있으면 먼저 묻는다 */
  requestClose(fromPop = false) {
    const acc = this.liveAccordion();
    if (acc?.modeManager?.isEditMode?.() && acc.editState?.isActive) {
      let changed = false;
      try {
        changed = Object.keys(acc.editState.collectChanges() || {}).length > 0;
      } catch (_) { /* 못 세면 묻지 않음 */ }
      if (changed && !confirm('저장하지 않은 내용이 있어요.\n닫으면 고친 것이 사라져요. 닫을까요?')) {
        if (fromPop) this.sheet.restoreHistory();
        return;
      }
    }
    this.sheet.closingFromPop = fromPop;
    if (acc?.isOpen) acc.closeAccordion();
    if (this.sheetOpen) this.hideSheet();
    this.sheet.closingFromPop = false;
  }

  hideSheet() {
    this.sheet.hide();   // 보던 카드로 포커스도 되돌림
  }

  destroy() {
    this.sheet?.el.remove();
    this.container.innerHTML = '';
  }
}
