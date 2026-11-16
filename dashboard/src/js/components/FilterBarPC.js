/**
 * 프로젝트 화면 PC 필터 줄 정리 (11-16 저녁, 부장님 결정 — 당근 비즈니스식)
 *
 * 한 줄: [검색] [상태] [담당자] [수금 상태] [세금계산서] [+ 필터 더보기 n] [더보기 안 조건 칩 ✕] [모두 지우기] …… 건수 [저장한 필터 ▾]
 *      '필터 더보기' 안: 내부 사업자·유입 구분·거래처·데이터(펼치면 다음 줄). '내 프로젝트만'·'취소 제외'는 표 위 체크 그대로
 *      (처음엔 조건 칩을 따로 둘째 줄에 뒀다가 상태 칸까지 세 줄이 돼 한 줄로 합침 — 11-16 저녁)
 *
 * 실제 select·체크 칸·검색칸은 그대로 두고 자리(order)와 보이기만 바꾼다 —
 * 필터 동작·URL 저장·편집 중 확인(ModernProjectFilters._applyFromUser)을 그대로 타게.
 * 모양은 groupware-shell.css 의 `.pfb` 규칙(PC, min-width 768px)만. 휴대폰은 MobileChrome 의 [검색+필터 n] 그대로.
 */
const PRIMARY = ['statusFilter', 'managerFilter', 'outstandingFilter', 'invoiceFilter'];
const MORE = ['companyFilter', 'clientFilter', 'businessNameFilter', 'dataFilter'];

const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (m) => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
})[m]);

export default class FilterBarPC {
  constructor(filters) {
    this.filters = filters;
    this.section = document.querySelector('.filter-section');
    this.row = this.section?.querySelector(':scope > .row');
    if (!this.section || !this.row || this.section.classList.contains('pfb')) return;
    this.build();
  }

  colOf(id) {
    return document.getElementById(id)?.closest('[class*="col-"]') || null;
  }

  labelOf(id) {
    return (document.querySelector(`label[for="${id}"]`)?.textContent || '').trim();
  }

  build() {
    const { section, row } = this;
    section.classList.add('pfb');
    if (!row.id) row.id = 'pmmFilterRow';

    // 자리 매기기: 검색 → 자주 쓰는 4개 → 더보기 단추 → (더보기 안) → 저장한 필터
    this.colOf('searchInput')?.classList.add('pfb-search');
    PRIMARY.forEach((id, i) => this.colOf(id)?.style.setProperty('--pfb-order', String(10 + i)));
    this.colOf('asStatusFilter')?.style.setProperty('--pfb-order', '14');
    MORE.forEach((id) => this.colOf(id)?.classList.add('pfb-more-item'));
    this.colOf('filterResultCount')?.classList.add('pfb-old-actions');

    // [+ 필터 더보기]
    const moreCol = document.createElement('div');
    moreCol.className = 'col-auto pfb-more-col';
    moreCol.innerHTML = `
      <button type="button" class="pfb-more" aria-expanded="false" aria-controls="${row.id}">
        <i class="fas fa-plus" aria-hidden="true"></i><span class="pfb-more-txt">필터 더보기</span><b class="pfb-more-n" hidden></b>
      </button>`;
    row.appendChild(moreCol);
    // 더보기를 펼치면 추가 필터는 새 줄에서 시작(첫 줄 끝에 이어 붙으면 어디까지가 자주 쓰는 것인지 흐려짐)
    const brk = document.createElement('div');
    brk.className = 'pfb-break';
    brk.setAttribute('aria-hidden', 'true');
    row.appendChild(brk);
    this.moreBtn = moreCol.querySelector('.pfb-more');
    this.moreN = moreCol.querySelector('.pfb-more-n');
    this.moreBtn.addEventListener('click', () => this.setMoreOpen(!section.classList.contains('pfb-more-open')));

    // [저장한 필터 ▾] — 둘째 줄 오른쪽 끝(건수 옆). 첫 줄에 두면 노트북 폭에서 혼자 다음 줄로 떨어졌음
    const presetCol = document.createElement('div');
    presetCol.className = 'pfb-preset-col';
    presetCol.innerHTML = `
      <button type="button" class="pfb-preset-btn" aria-haspopup="true" aria-expanded="false" aria-controls="pfbPresetMenu"
              aria-label="저장한 필터" title="저장한 필터">
        <i class="fas fa-bookmark" aria-hidden="true"></i><span class="pfb-preset-txt">저장한 필터</span><i class="fas fa-chevron-down pfb-caret" aria-hidden="true"></i>
      </button>
      <div class="pfb-preset-menu" id="pfbPresetMenu" hidden></div>`;
    this.presetBtn = presetCol.querySelector('.pfb-preset-btn');
    this.presetMenu = presetCol.querySelector('.pfb-preset-menu');
    this.presetBtn.addEventListener('click', () => this.togglePresetMenu());
    this.presetMenu.addEventListener('click', (e) => this.onPresetClick(e));
    document.addEventListener('click', (e) => {
      if (!this.presetMenu.hidden && !presetCol.contains(e.target)) this.togglePresetMenu(false);
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !this.presetMenu.hidden) {
        this.togglePresetMenu(false);
        this.presetBtn.focus();
      }
    });

    // 걸린 조건·건수·저장한 필터도 같은 줄에 — 따로 줄을 두니 상태 칸·필터·조건 세 줄이라 표가 아래로 밀렸음
    //   (부장님 지적 11-16 저녁). 칩은 더보기 안에 접혀 첫 줄에서 안 보이는 조건만(첫 줄 필터는 진한 칩으로 이미 보임)
    this.activeRow = document.createElement('div');
    this.activeRow.className = 'col-auto pfb-tags-col';
    this.activeRow.innerHTML = `
      <div class="pfb-tags" role="list" aria-label="더보기 안에 걸린 조건"></div>
      <button type="button" class="pfb-clear" hidden>모두 지우기</button>`;
    row.appendChild(this.activeRow);
    const endCol = document.createElement('div');
    endCol.className = 'col-auto pfb-end';
    endCol.innerHTML = '<span class="pfb-count" aria-live="polite"></span>';
    endCol.appendChild(presetCol);
    row.appendChild(endCol);
    this.tagsEl = this.activeRow.querySelector('.pfb-tags');
    this.clearBtn = this.activeRow.querySelector('.pfb-clear');
    this.countEl = endCol.querySelector('.pfb-count');
    this.clearBtn.addEventListener('click', () => window.modernFilters?.resetFilters?.());
    this.tagsEl.addEventListener('click', (e) => this.onTagClick(e));

    // 건수: 원래 자리(#filterResultCount)의 글자를 그대로 따라 쓴다(계산은 필터 쪽 하나로)
    const src = document.getElementById('filterResultCount');
    if (src) {
      // 줄 안에는 숫자만('4,134개') — '0개 · 필터 2개로 좁혀졌어요' 같은 긴 글은 노트북 폭에서 줄을 넘김. 전체 글은 title 로
      const copy = () => {
        const full = src.textContent.trim();
        const short = (full.match(/^[\d,]+개/) || [full])[0];
        this.countEl.textContent = short;
        this.countEl.title = [full, src.title].filter(Boolean).join('\n');
      };
      new MutationObserver(copy).observe(src, { childList: true, characterData: true, subtree: true, attributes: true });
      copy();
    }

    // 다시 그리기: 칸을 바꿀 때·검색할 때·필터가 돌 때(상태 칸·[초기화]·저장한 필터는 값만 바꾸고 change 를 안 보낼 수 있음)
    const update = () => { this.update(); this.fit(); };
    row.addEventListener('change', () => setTimeout(update, 0));
    document.getElementById('searchInput')?.addEventListener('input', () => setTimeout(update, 350));
    this.filters?.onFilterChange?.(() => update());

    // 화면 폭이 바뀌면 다시 잼(넓은 화면은 8개 모두 펼침, 좁으면 더보기로)
    let lastW = 0;
    new ResizeObserver(() => {
      const w = Math.round(section.getBoundingClientRect().width);
      if (w !== lastW) { lastW = w; this.fit(); }
    }).observe(section);
    this.update();
    this.fit();
  }

  /**
   * 넓은 화면이면 '더보기' 없이 8개를 한 줄에 다 편다(부장님 지적 11-16 저녁 — 오른쪽이 비는데 접어 둘 이유가 없음).
   * 고정 폭 기준이 아니라 실제로 펴 보고 한 줄을 넘으면(줄바꿈 생기면) 다시 접는다. 고른 값에 따라 칸 폭이 바뀌므로
   * 필터가 바뀔 때마다 다시 잰다.
   */
  fit() {
    const { section, row } = this;
    if (!window.matchMedia('(min-width: 768px)').matches) {
      section.classList.remove('pfb-all');
      return;
    }
    section.classList.add('pfb-all');
    this.update();                      // 다 펴면 칩은 감춰지므로 칩을 뺀 모양으로 잰다
    // 높이가 다른 칸(칩 30px·단추 36px)은 top 이 몇 px 다르므로 '줄이 바뀜' = 20px 넘게 아래
    const tops = [...row.children]
      .filter((el) => el.getBoundingClientRect().width > 0)
      .map((el) => el.getBoundingClientRect().top);
    const fits = tops.length === 0 || Math.max(...tops) - Math.min(...tops) < 20;
    if (!fits) {
      section.classList.remove('pfb-all');
      this.update();
    } else if (section.classList.contains('pfb-more-open')) {
      this.setMoreOpen(false);
    }
  }

  setMoreOpen(open) {
    this.section.classList.toggle('pfb-more-open', open);
    this.moreBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    // 펼친 상태가 단추만 봐도 보이게 '필터 접기'·－ 로(펼친 채 두면 줄이 하나 늘어난 이유를 몰랐음 — 11-16 저녁)
    this.moreBtn.querySelector('.pfb-more-txt').textContent = open ? '필터 접기' : '필터 더보기';
    const icon = this.moreBtn.querySelector('i');
    icon.classList.toggle('fa-plus', !open);
    icon.classList.toggle('fa-minus', open);
    this.update();   // 펼치면 칸이 직접 보이므로 칩은 감춤
  }

  // ── 걸린 조건 칩 ───────────────────────────────────────────
  activeConditions() {
    const out = [];
    const add = (key, label, value) => out.push({ key, label, value });
    const search = document.getElementById('searchInput');
    if (search?.value.trim()) add('search', '검색', `"${search.value.trim()}"`);
    const ids = [...PRIMARY, 'asStatusFilter', ...MORE];
    ids.forEach((id) => {
      const sel = document.getElementById(id);
      if (!sel || !sel.value) return;
      if (id === 'asStatusFilter' && !sel.closest('[class*="col-"]')?.offsetParent) return;
      if (id === 'managerFilter' && document.getElementById('myProjectsOnly')?.checked) return; // '내 프로젝트만'으로 보임
      add(id, this.labelOf(id), sel.options[sel.selectedIndex]?.text || sel.value);
    });
    if (document.getElementById('myProjectsOnly')?.checked) add('myProjectsOnly', '', '내 프로젝트만');
    const hc = document.getElementById('hideCancelledProjects');
    if (hc && !hc.checked) add('hideCancelledProjects', '', '취소된 공사 포함');
    return out;
  }

  update() {
    const all = this.activeConditions();
    // 칩은 더보기 안 조건만, 그것도 접혀 있을 때만(첫 줄 필터·검색칸은 값이 그대로 보여 칩이 겹쳤음)
    const hiddenKeys = new Set(MORE);   // '내 프로젝트만'·'취소 제외'는 표 위 체크로 늘 보임
    const shownAll = this.section.classList.contains('pfb-more-open') || this.section.classList.contains('pfb-all');
    const conds = shownAll ? [] : all.filter((c) => hiddenKeys.has(c.key));
    this.tagsEl.innerHTML = conds.map((c) => `
      <span class="pfb-tag" role="listitem">
        <span>${c.label ? `<span class="pfb-tag-k">${esc(c.label)}</span> ` : ''}${esc(c.value)}</span>
        <button type="button" class="pfb-tag-x" data-key="${esc(c.key)}" aria-label="${esc((c.label ? c.label + ' ' : '') + c.value)} 조건 빼기">
          <i class="fas fa-times" aria-hidden="true"></i>
        </button>
      </span>`).join('');
    this.clearBtn.hidden = all.length === 0;          // [모두 지우기]는 어떤 조건이든 걸리면
    this.activeRow.classList.toggle('has-tags', conds.length > 0);
    this.activeRow.hidden = all.length === 0;

    // 더보기 안에 걸린 조건 수 — 접어 둬도 목록이 왜 줄었는지 보이게
    const n = MORE.filter((id) => document.getElementById(id)?.value).length;
    this.moreN.textContent = n ? String(n) : '';
    this.moreN.hidden = n === 0;
    this.moreBtn.classList.toggle('on', n > 0);
  }

  onTagClick(e) {
    const x = e.target.closest('.pfb-tag-x');
    if (!x) return;
    const key = x.dataset.key;
    if (key === 'search') {
      const s = document.getElementById('searchInput');
      s.value = '';
      s.dispatchEvent(new Event('input', { bubbles: true }));
    } else if (key === 'myProjectsOnly' || key === 'hideCancelledProjects') {
      document.getElementById(key)?.click();
    } else {
      const sel = document.getElementById(key);
      if (!sel) return;
      sel.value = '';
      sel.dispatchEvent(new Event('change', { bubbles: true }));
    }
    setTimeout(() => {
      this.update();
      // 지운 칩 자리에서 키보드가 길을 잃지 않게: 남은 첫 칩 또는 검색칸으로
      (this.tagsEl.querySelector('.pfb-tag-x') || document.getElementById('searchInput'))?.focus();
    }, 400);
  }

  // ── 저장한 필터 ────────────────────────────────────────────
  togglePresetMenu(force) {
    const open = force ?? this.presetMenu.hidden;
    if (open) this.renderPresetMenu();
    this.presetMenu.hidden = !open;
    this.presetBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    if (open) this.presetMenu.querySelector('button')?.focus();
  }

  renderPresetMenu() {
    const presets = this.filters?._getPresets?.() || [];
    const items = presets.map((p) => `
      <div class="pfb-preset-item">
        <button type="button" class="pfb-preset-apply" data-name="${esc(p.name)}">${esc(p.name)}</button>
        <button type="button" class="pfb-preset-del" data-name="${esc(p.name)}" aria-label="${esc(p.name)} 저장한 필터 지우기">
          <i class="fas fa-times" aria-hidden="true"></i>
        </button>
      </div>`).join('');
    this.presetMenu.innerHTML = (items || '<p class="pfb-preset-empty">저장한 필터가 없어요.</p>') + `
      <button type="button" class="pfb-preset-save"><i class="fas fa-plus" aria-hidden="true"></i>지금 조건 저장하기</button>`;
  }

  onPresetClick(e) {
    const f = window.modernFilters;
    if (!f) return;
    const apply = e.target.closest('.pfb-preset-apply');
    const del = e.target.closest('.pfb-preset-del');
    if (apply) {
      this.togglePresetMenu(false);
      f.loadPreset(apply.dataset.name);
      setTimeout(() => this.update(), 300);
    } else if (del) {
      f.deletePreset(del.dataset.name);
      this.renderPresetMenu();
    } else if (e.target.closest('.pfb-preset-save')) {
      this.togglePresetMenu(false);
      f.savePresetPrompt();
    }
  }
}
