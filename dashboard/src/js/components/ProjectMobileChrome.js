/**
 * 프로젝트 화면 휴대폰 정리 — 머리 줄 ⋯ 메뉴, 필터 접기 (2026-11-15 휴대폰 리뉴얼 1단계).
 *
 * 만드는 것은 단추·메뉴 몇 개와 class 이름뿐이고, 보이고 숨기는 건 모두 CSS 의
 * `@media (max-width: 767.98px)` 안에서 한다(pm-mobile.css) → 넓은 화면(PC)은 지금과 똑같다.
 * 템플릿을 고치지 않으려고(서버 재시작 없이 반영) 기존 마크업을 찾아 class 를 붙인다.
 */
export default class ProjectMobileChrome {
  constructor() {
    document.body.classList.add('pm-projects-page');
    this.setupHead();
    this.setupFilters();
  }

  // ── 머리 줄: 제목 + [새 프로젝트] 하나, [새로고침]·[캘린더 동기화]는 ⋯ 메뉴로 ─────────
  setupHead() {
    const refresh = document.getElementById('fullRefreshBtn');
    const head = refresh?.closest('.row');
    const actions = refresh?.parentElement;
    if (!head || !actions || actions.querySelector('.pmm-kebab-wrap')) return;
    head.classList.add('pmm-head');

    const items = [
      { id: 'fullRefreshBtn', icon: 'fa-sync-alt', label: '새로고침' },
      { id: 'calendarSyncBtn', icon: 'fa-calendar-check', label: '캘린더 동기화' },
    ].filter((it) => document.getElementById(it.id));

    const wrap = document.createElement('div');
    wrap.className = 'pmm-kebab-wrap';
    wrap.innerHTML = `
      <button type="button" class="pmm-kebab" aria-haspopup="menu" aria-expanded="false" aria-controls="pmmHeadMenu" aria-label="메뉴 더 보기">
        <i class="fas fa-ellipsis-h" aria-hidden="true"></i>
      </button>
      <div class="pmm-menu" id="pmmHeadMenu" role="menu" hidden>
        ${items.map((it) => `
          <button type="button" role="menuitem" class="pmm-menu-item" data-proxy="${it.id}">
            <i class="fas ${it.icon}" aria-hidden="true"></i>${it.label}
          </button>`).join('')}
      </div>`;
    actions.insertBefore(wrap, actions.firstChild);

    const btn = wrap.querySelector('.pmm-kebab');
    const menu = wrap.querySelector('.pmm-menu');
    const setOpen = (open, focusBtn = false) => {
      menu.hidden = !open;
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
      if (open) menu.querySelector('.pmm-menu-item')?.focus();
      else if (focusBtn) btn.focus();
    };
    btn.addEventListener('click', () => setOpen(menu.hidden));
    menu.addEventListener('click', (e) => {
      const item = e.target.closest('.pmm-menu-item');
      if (!item) return;
      setOpen(false);
      document.getElementById(item.dataset.proxy)?.click();
    });
    menu.addEventListener('keydown', (e) => {
      const list = [...menu.querySelectorAll('.pmm-menu-item')];
      const i = list.indexOf(document.activeElement);
      if (e.key === 'ArrowDown') { e.preventDefault(); list[(i + 1) % list.length]?.focus(); }
      if (e.key === 'ArrowUp') { e.preventDefault(); list[(i - 1 + list.length) % list.length]?.focus(); }
    });
    document.addEventListener('click', (e) => {
      if (!menu.hidden && !wrap.contains(e.target)) setOpen(false);
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !menu.hidden) setOpen(false, true);
    });
  }

  // ── 필터: 검색 칸 + [필터 n] 단추, 나머지 칩·초기화·프리셋 저장은 펼쳤을 때만 ───────────
  setupFilters() {
    const section = document.querySelector('.filter-section');
    const row = section?.querySelector(':scope > .row');
    const search = document.getElementById('searchInput');
    const searchCol = search?.closest('[class*="col-"]');
    if (!section || !row || !searchCol || row.querySelector('.pmm-filter-more')) return;

    section.classList.add('pmm-filters');
    searchCol.classList.add('pmm-col-search');
    document.getElementById('filterResultCount')?.closest('[class*="col-"]')?.classList.add('pmm-col-actions');

    const col = document.createElement('div');
    col.className = 'pmm-more-slot';
    col.innerHTML = `
      <button type="button" class="pmm-filter-more" aria-expanded="false" aria-controls="${row.id || ''}">
        <i class="fas fa-sliders-h" aria-hidden="true"></i>필터<b class="pmm-filter-n" hidden></b>
      </button>`;
    searchCol.insertAdjacentElement('afterend', col);
    const btn = col.querySelector('button');
    const nEl = col.querySelector('.pmm-filter-n');
    if (!row.id) {
      row.id = 'pmmFilterRow';
      btn.setAttribute('aria-controls', row.id);
    }

    // 펼친 안의 고르기 칸 중 값이 있는 것 수 (A/S 상태 칸은 모드 전용이라 뺌)
    const sync = () => {
      const n = [...row.querySelectorAll('select.form-select')]
        .filter((s) => s.id !== 'asStatusFilter' && s.value).length;
      nEl.textContent = n;
      nEl.hidden = n === 0;
      btn.classList.toggle('on', n > 0);
      btn.setAttribute('aria-label', n ? `필터 ${n}개 걸림, 펼치기` : '필터 펼치기');
    };
    btn.addEventListener('click', () => {
      const open = !section.classList.contains('pmm-open');
      section.classList.toggle('pmm-open', open);
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    row.addEventListener('change', sync);
    // [초기화]·상태 칸은 값만 바꾸고 change 를 안 보낼 수 있어 필터가 돌 때마다 맞춤
    this.syncFilterCount = sync;
    sync();
  }
}
