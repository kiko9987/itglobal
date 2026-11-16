/**
 * 휴대폰 정리 — 머리 줄 ⋯ 메뉴, 필터 접기. 프로젝트(1단계 11-15)·고객 리드(2단계 11-16) 화면 공용.
 *
 * 만드는 것은 단추·메뉴 몇 개와 class 이름뿐이고, 보이고 숨기는 건 모두 CSS 의
 * `@media (max-width: 767.98px)` 안에서 한다(pm-mobile.css) → 넓은 화면(PC)은 지금과 똑같다.
 * 템플릿을 고치지 않으려고(서버 재시작 없이 반영) 기존 마크업을 찾아 class 를 붙인다.
 */
const DEFAULT_MENU = [
  { id: 'fullRefreshBtn', icon: 'fa-sync-alt', label: '새로고침' },
  { id: 'calendarSyncBtn', icon: 'fa-calendar-check', label: '캘린더 동기화' },
];

export default class MobileChrome {
  /**
   * @param {object} opts
   * @param {string} [opts.pageClass] 화면별 body class(pmm-projects / pmm-leads)
   * @param {Array}  [opts.menu] ⋯ 메뉴로 옮길 단추 {id, icon, label} — 휴대폰에선 원래 단추를 CSS 로 숨김
   */
  constructor({ pageClass = '', menu = DEFAULT_MENU } = {}) {
    document.body.classList.add('pmm-page');
    if (pageClass) document.body.classList.add(pageClass);
    this.menu = menu;
    this.setupHead();
    this.setupFilters();
  }

  // ── 머리 줄: 제목 + 주 단추 하나, 나머지(새로고침 등)는 ⋯ 메뉴로 ─────────
  setupHead() {
    const refresh = document.getElementById('fullRefreshBtn');
    const head = refresh?.closest('.row');
    const actions = refresh?.parentElement;
    if (!head || !actions || actions.querySelector('.pmm-kebab-wrap')) return;
    head.classList.add('pmm-head');

    const items = this.menu.filter((it) => document.getElementById(it.id));

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
    // 원래 단추를 화면 쪽이 숨겼으면(권한 — StateManager.applyPermissions 가 메뉴를 만든 뒤에 숨김) 메뉴에서도 뺀다.
    //   휴대폰에선 원래 단추가 CSS 로 늘 숨으므로 계산된 모양이 아니라 화면 쪽이 건 표시만 본다 (11-16)
    const hiddenByPage = (el) => !el || el.hidden || el.style.display === 'none' || el.classList.contains('d-none');
    const syncItems = () => {
      menu.querySelectorAll('.pmm-menu-item').forEach((item) => {
        item.hidden = hiddenByPage(document.getElementById(item.dataset.proxy));
      });
    };
    const visibleItems = () => [...menu.querySelectorAll('.pmm-menu-item')].filter((it) => !it.hidden);
    const setOpen = (open, focusBtn = false) => {
      if (open) syncItems();
      menu.hidden = !open;
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
      if (open) visibleItems()[0]?.focus();
      else if (focusBtn) btn.focus();
    };
    btn.addEventListener('click', () => setOpen(menu.hidden));
    menu.addEventListener('click', (e) => {
      const item = e.target.closest('.pmm-menu-item');
      if (!item || item.hidden) return;
      setOpen(false);
      const target = document.getElementById(item.dataset.proxy);
      if (!hiddenByPage(target)) target.click();
    });
    menu.addEventListener('keydown', (e) => {
      const list = visibleItems();
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

    // 걸린 조건 수: 펼친 안의 고르기 칸 중 값이 있는 것(A/S 상태 칸은 모드 전용이라 뺌) + 검색어
    //   (검색어만 넣고 필터를 접어 두면 목록이 왜 줄었는지 단서가 없었음)
    const sync = () => {
      const n = [...row.querySelectorAll('select.form-select')]
        .filter((s) => s.id !== 'asStatusFilter' && s.value).length
        + (search.value.trim() ? 1 : 0);
      nEl.textContent = n;
      nEl.hidden = n === 0;
      btn.classList.toggle('on', n > 0);
    };
    btn.addEventListener('click', () => {
      const open = !section.classList.contains('pmm-open');
      section.classList.toggle('pmm-open', open);
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    row.addEventListener('change', sync);
    search.addEventListener('input', sync);
    // [초기화]·상태 칸은 값만 바꾸고 change 를 안 보낼 수 있어 필터가 돌 때마다 맞춤
    this.syncFilterCount = sync;
    sync();
  }
}
