import logger from '../utils/logger.js';
import { partnerDisplayName } from '../utils/partnerName.js';
import { escHtml } from '../utils/escapeHtml.js';
import ProjectStatusCalculator from '../utils/ProjectStatusCalculator.js';

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

export default class MobileCardView {
  constructor() {
    this.mq = window.matchMedia(MOBILE_MQ);
    this.currentData = [];
    this.shown = PAGE_SIZE;
    this.sheetOpen = false;
    this.lastCardCode = null;

    this.createContainer();
    this.createSheet();
    this.bindEvents();
    this.sync();
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
      <div class="pmm-cards" role="list"></div>
      <p class="pmm-empty" hidden>조건에 맞는 공사가 없어요. 검색어나 필터를 바꿔 보세요.</p>
      <button type="button" class="pmm-more-btn" hidden></button>`;
    this.container = el;
    this.listEl = el.querySelector('.pmm-cards');
    this.emptyEl = el.querySelector('.pmm-empty');
    this.moreBtn = el.querySelector('.pmm-more-btn');
  }

  createSheet() {
    const sheet = document.createElement('div');
    sheet.className = 'pmm-sheet';
    sheet.id = 'pmmSheet';
    sheet.setAttribute('role', 'dialog');
    sheet.setAttribute('aria-modal', 'true');
    sheet.setAttribute('aria-labelledby', 'pmmSheetTitle');
    sheet.hidden = true;
    sheet.innerHTML = `
      <div class="pmm-sheet-bar">
        <button type="button" class="pmm-sheet-back" aria-label="목록으로 돌아가기">
          <i class="fas fa-arrow-left" aria-hidden="true"></i>
        </button>
        <h2 class="pmm-sheet-title" id="pmmSheetTitle" tabindex="-1"></h2>
      </div>
      <div class="pmm-sheet-scroll">
        <div class="pmm-sheet-quick"></div>
        <table class="pmm-sheet-table" role="presentation">
          <tbody><tr class="pmm-sheet-anchor"><td></td></tr></tbody>
        </table>
      </div>`;
    document.body.appendChild(sheet);
    this.sheet = sheet;
    this.sheetTitle = sheet.querySelector('.pmm-sheet-title');
    this.sheetQuick = sheet.querySelector('.pmm-sheet-quick');
    this.sheetScroll = sheet.querySelector('.pmm-sheet-scroll');
    this.anchorRow = sheet.querySelector('.pmm-sheet-anchor');
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

    this.sheet.querySelector('.pmm-sheet-back').addEventListener('click', () => this.requestClose());
    // 포커스 가두기: 시트가 열려 있는 동안 Tab 으로 뒤 화면에 가지 않게. 시트 위에 뜬 창(메모·사유 입력 등
    //   부트스트랩 창, 몸통으로 옮겨 띄운 체크 메뉴)은 그대로 둔다. 뒤 화면 inert 는 그 창들이 본문 안에 있어 못 씀
    document.addEventListener('focusin', (e) => {
      if (!this.sheetOpen || this.sheet.contains(e.target)) return;
      if (e.target.closest?.('.modal, .dropdown-menu, [data-portal], .ts-dropdown, .tooltip, .popover')) return;
      this.sheetTitle.focus({ preventScroll: true });
    });
    window.addEventListener('popstate', () => {
      if (this.sheetOpen) this.requestClose(true);
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
    const rows = this.currentData;
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

    return `
      <div class="pmm-card${cancelled ? ' is-cancelled' : ''}" role="listitem" tabindex="0"
           data-code="${escHtml(code)}" aria-label="${escHtml(`${code} ${status} ${addr}`)} 상세 보기">
        <div class="pmm-card-top">
          <span class="pmm-status pmm-st-${tone}"><span class="pmm-dot" aria-hidden="true"></span>${escHtml(status || '상태 없음')}</span>
          <span class="pmm-code">${escHtml(code)}</span>
        </div>
        <div class="pmm-addr">${escHtml(addr || '현장 주소 없음')}</div>
        ${partner && partner !== '-' ? `<div class="pmm-partner">${escHtml(partner)}</div>` : ''}
        <div class="pmm-money">
          <span>총액 <b>${total > 0 ? won(total) : '-'}</b></span>
          <span class="${due > 0 ? 'pmm-due' : 'pmm-due-zero'}">미수금 <b>${won(due)}</b></span>
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
   * ⚠️ projectListApp.components.accordion 은 따로 만들어지기만 하고 init 되지 않은 객체라 쓰면 안 됨.
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
        if (row && this.sheet.contains(row)) {
          row.dataset.projectCode = newCode;
          acc.accordionContainer.classList.add('show');
          acc.isOpen = true;
          this.sheetTitle.textContent = newCode;
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
    this.sheetTitle.textContent = code;
    this.sheetQuick.innerHTML = this.quickHtml(data);
    this.sheetScroll.scrollTop = 0;
    if (!this.sheetOpen) {
      this.sheetOpen = true;
      this.sheet.hidden = false;
      document.documentElement.classList.add('pmm-sheet-open');
      try {
        history.pushState({ ...(history.state || {}), pmmSheet: code }, '');
      } catch (_) { /* 기록을 못 남겨도 닫기 단추로 닫힘 */ }
      requestAnimationFrame(() => this.sheetTitle.focus({ preventScroll: true }));
    }
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
        <span class="pmm-sum-money ${due > 0 ? 'pmm-due' : 'pmm-due-zero'}">미수금 <b>${won(due)}</b></span>
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
        if (fromPop) {
          try { history.pushState({ ...(history.state || {}), pmmSheet: this.sheetTitle.textContent }, ''); } catch (_) { /* noop */ }
        }
        return;
      }
    }
    this.closingFromPop = fromPop;
    if (acc?.isOpen) acc.closeAccordion();
    if (this.sheetOpen) this.hideSheet();
    this.closingFromPop = false;
  }

  hideSheet() {
    if (!this.sheetOpen) return;
    this.sheetOpen = false;
    this.sheet.hidden = true;
    document.documentElement.classList.remove('pmm-sheet-open');
    if (!this.closingFromPop && history.state && history.state.pmmSheet) {
      try { history.back(); } catch (_) { /* noop */ }
    }
    // 보던 카드로 포커스 되돌리기
    const code = this.lastCardCode;
    if (code && this.isActive) {
      const card = [...this.listEl.querySelectorAll('.pmm-card')].find((c) => c.dataset.code === code);
      if (card) card.focus({ preventScroll: true });
    }
  }

  destroy() {
    this.sheet?.remove();
    this.container.innerHTML = '';
  }
}
