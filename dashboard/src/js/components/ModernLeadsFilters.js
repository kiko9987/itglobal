import logger from '../utils/logger.js';
import {
  LEAD_STATUS_OPTIONS, splitOwners, consultSortKey, sortKeyToIso, localIso, shiftIso, isVisitLead,
} from '../utils/leadFormat.js';

/**
 * 고객 리드 필터 — 프로젝트 페이지 ModernProjectFilters 와 같은 구조·동작
 *   드롭다운: 상담자 / 영업 담당 / 플랫폼 / 상태 / 기간(상담일) + 키워드 검색
 *   토글(테이블 상단 바): 내 리드만 보기 / 기타 포함 / 방문 모드
 *   상태 유지: sessionStorage(현재 필터) + localStorage(프리셋)
 */
const SESSION_KEY = 'itg_lead_filters_v1';
const PRESET_KEY = 'itg_lead_filter_presets_v1';

const PERIODS = {
  '7d': { label: '최근 7일', days: 7 },
  '30d': { label: '최근 30일', days: 30 },
  '90d': { label: '최근 90일', days: 90 },
  'year': { label: '올해', days: null },
};

export default class ModernLeadsFilters {
  constructor() {
    this.filters = {};          // consultant / salesOwner / platform / status / period / search
    this.toggles = { myLeadsOnly: false, includeEtc: false, visitMode: false };
    this.callbacks = [];
    this.currentData = null;
    this.searchDebounceTimer = null;

    this.searchableFields = [
      '리드 No', '고객명', '고객 연락처', '이메일', '방문 주소',
      '문의 내용', '상담 내용', '키워드', '플랫폼', '영업 담당자', '온라인 상담자', '_project_code',
    ];
  }

  init() {
    this.el = {
      consultant: document.getElementById('consultantFilter'),
      salesOwner: document.getElementById('salesOwnerFilter'),
      platform: document.getElementById('platformFilter'),
      status: document.getElementById('statusFilter'),
      period: document.getElementById('periodFilter'),
      search: document.getElementById('searchInput'),
      reset: document.getElementById('resetFiltersBtn'),
      savePreset: document.getElementById('savePresetBtn'),
      count: document.getElementById('filterResultCount'),
      presets: document.getElementById('filterPresetsContainer'),
    };
    this.restoreSession();
    this.bindEvents();
    this.renderPresets();
    logger.debug('[ModernLeadsFilters] 초기화 완료');
  }

  bindEvents() {
    ['consultant', 'salesOwner', 'platform', 'status', 'period'].forEach((key) => {
      const el = this.el[key];
      if (!el) return;
      el.addEventListener('change', () => {
        if (el.value) this.filters[key] = el.value;
        else delete this.filters[key];
        this.applyFilters(null, true);
      });
    });

    this.el.search?.addEventListener('input', () => {
      const v = this.el.search.value.trim();
      if (v) this.filters.search = v;
      else delete this.filters.search;
      clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = setTimeout(() => this.applyFilters(null, true), 300);
    });

    this.el.reset?.addEventListener('click', () => this.resetFilters());
    this.el.savePreset?.addEventListener('click', () => this.savePresetPrompt());
  }

  /** 테이블 상단 바 토글(내 리드만·기타 포함·방문 모드)에서 호출 */
  setToggle(name, value) {
    this.toggles[name] = !!value;
    this.applyFilters(null, true);
  }

  applyFilters(data = null, triggerCallbacks = true) {
    if (Array.isArray(data)) {
      this.currentData = data;
      this.populateAllFilters(data);
    }
    if (!this.currentData) {
      this.updateResultCount(0);
      return null;
    }

    const f = this.filters;
    const t = this.toggles;
    const me = (window.userDisplayName || '').trim();
    const term = (f.search || '').toLowerCase();
    const periodFrom = this.periodStart(f.period);

    const result = this.currentData.filter((lead) => {
      const platform = String(lead['플랫폼'] || '').trim();
      if (!t.includeEtc && platform === '기타') return false;
      if (t.visitMode && !isVisitLead(lead)) return false;

      const consultants = splitOwners(lead['온라인 상담자']);
      const sales = splitOwners(lead['영업 담당자']);
      if (t.myLeadsOnly && me && !consultants.includes(me) && !sales.includes(me)) return false;

      if (f.consultant && !consultants.includes(f.consultant)) return false;
      if (f.salesOwner && !sales.includes(f.salesOwner)) return false;
      if (f.platform && platform !== f.platform) return false;
      if (f.status && String(lead['상태'] || '').trim() !== f.status) return false;
      if (periodFrom) {
        const iso = sortKeyToIso(consultSortKey(lead['상담 시간']));
        if (!iso || iso < periodFrom) return false;
      }
      if (term && !this.searchableFields.some((k) => String(lead[k] ?? '').toLowerCase().includes(term))) {
        return false;
      }
      return true;
    });

    this.updateResultCount(result.length);
    this.updateFilterVisualEffects();
    this.saveSession();
    // dataRefresh: 새 데이터 유입(새로고침)인지 — 필터 변경이면 테이블이 1페이지로 돌아감
    const meta = { dataRefresh: Array.isArray(data) };
    if (triggerCallbacks) this.callbacks.forEach((cb) => cb(result, meta));
    return result;
  }

  periodStart(period) {
    const p = PERIODS[period];
    if (!p) return '';
    const today = localIso(new Date());
    return p.days ? shiftIso(today, -(p.days - 1)) : `${today.slice(0, 4)}-01-01`;
  }

  populateAllFilters(data) {
    const uniq = (arr) => [...new Set(arr.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ko'));
    this.populate(this.el.consultant, uniq(data.flatMap((l) => splitOwners(l['온라인 상담자']))));
    this.populate(this.el.salesOwner, uniq(data.flatMap((l) => splitOwners(l['영업 담당자']))));
    this.populate(this.el.platform, uniq(data.map((l) => String(l['플랫폼'] || '').trim())));
    const dataStatuses = data.map((l) => String(l['상태'] || '').trim()).filter(Boolean);
    this.populate(this.el.status, [...new Set([...LEAD_STATUS_OPTIONS, ...dataStatuses])]);
    if (this.el.period && this.el.period.options.length <= 1) {
      this.populate(this.el.period, Object.keys(PERIODS), (k) => PERIODS[k].label);
    }
    this.syncElementsFromFilters();
  }

  populate(el, values, labelOf = (v) => v) {
    if (!el) return;
    const current = el.value;
    el.innerHTML = '<option value="">전체</option>';
    values.forEach((v) => {
      const o = document.createElement('option');
      o.value = v;
      o.textContent = labelOf(v);
      el.appendChild(o);
    });
    if (current && values.includes(current)) el.value = current;
  }

  syncElementsFromFilters() {
    ['consultant', 'salesOwner', 'platform', 'status', 'period'].forEach((k) => {
      if (this.el[k]) this.el[k].value = this.filters[k] || '';
    });
    if (this.el.search) this.el.search.value = this.filters.search || '';
  }

  updateResultCount(count) {
    if (!this.el?.count) return;
    const active = this.activeFilterCount();
    this.el.count.textContent = count === 0 && active
      ? `0개 · 필터 ${active}개로 좁혀졌어요`
      : `${count.toLocaleString()}개 리드`;
  }

  activeFilterCount() {
    return Object.keys(this.filters).length + (this.toggles.myLeadsOnly ? 1 : 0);
  }

  updateFilterVisualEffects() {
    ['consultant', 'salesOwner', 'platform', 'status', 'period'].forEach((k) => {
      this.el[k]?.classList.toggle('filter-selected', !!this.filters[k]);
    });
    this.el.search?.classList.toggle('filter-selected', !!this.filters.search);
    document.querySelector('.filter-section')?.classList.toggle('has-active-filters', this.activeFilterCount() > 0);
  }

  onFilterChange(cb) {
    this.callbacks.push(cb);
  }

  resetFilters() {
    this.filters = {};
    this.syncElementsFromFilters();
    // 모드(방문)·기타 포함은 보기 방식이라 유지, '내 리드만'은 필터라 해제
    this.toggles.myLeadsOnly = false;
    const my = document.getElementById('myLeadsOnly');
    if (my) my.checked = false;
    this.applyFilters(null, true);
  }

  // ── 세션 유지 ─────────────────────────────────────────────
  saveSession() {
    try {
      sessionStorage.setItem(SESSION_KEY, JSON.stringify({ filters: this.filters, toggles: this.toggles }));
    } catch (_) { /* noop */ }
  }

  restoreSession() {
    try {
      const saved = JSON.parse(sessionStorage.getItem(SESSION_KEY) || 'null');
      if (saved) {
        this.filters = saved.filters || {};
        this.toggles = { ...this.toggles, ...(saved.toggles || {}) };
      }
    } catch (_) { /* noop */ }
  }

  // ── 프리셋 (프로젝트 페이지와 같은 방식, 키만 분리) ─────────────
  loadPresets() {
    try {
      return JSON.parse(localStorage.getItem(PRESET_KEY) || '[]');
    } catch (_) {
      return [];
    }
  }

  savePresets(list) {
    try {
      localStorage.setItem(PRESET_KEY, JSON.stringify(list));
    } catch (_) { /* noop */ }
  }

  savePresetPrompt() {
    if (!this.activeFilterCount()) {
      window.showPageAlert?.('저장할 필터를 먼저 선택하세요', 'warning');
      return;
    }
    const name = (prompt('프리셋 이름') || '').trim();
    if (!name) return;
    const list = this.loadPresets().filter((p) => p.name !== name);
    list.push({ name, filters: { ...this.filters }, toggles: { ...this.toggles }, created_at: Date.now() });
    this.savePresets(list);
    this.renderPresets();
  }

  loadPreset(name) {
    const p = this.loadPresets().find((x) => x.name === name);
    if (!p) return;
    this.filters = { ...(p.filters || {}) };
    this.toggles = { ...this.toggles, ...(p.toggles || {}) };
    this.syncElementsFromFilters();
    // 테이블 상단 바 토글 UI(내 리드만·기타 포함·방문 모드)를 프리셋 값으로 맞춤
    document.dispatchEvent(new CustomEvent('leadTogglesRestored', { detail: this.toggles }));
    this.applyFilters(null, true);
  }

  deletePreset(name) {
    this.savePresets(this.loadPresets().filter((p) => p.name !== name));
    this.renderPresets();
  }

  renderPresets() {
    const box = this.el?.presets;
    if (!box) return;
    box.innerHTML = '';
    this.loadPresets().forEach((p) => {
      const chip = document.createElement('span');
      chip.className = 'badge bg-light text-dark border d-inline-flex align-items-center gap-1';
      chip.style.cursor = 'pointer';
      const label = document.createElement('span');
      label.textContent = p.name;
      label.addEventListener('click', () => this.loadPreset(p.name));
      const del = document.createElement('i');
      del.className = 'fas fa-times text-muted';
      del.title = '프리셋 삭제';
      del.addEventListener('click', (e) => {
        e.stopPropagation();
        if (confirm(`'${p.name}' 프리셋을 삭제할까요?`)) this.deletePreset(p.name);
      });
      chip.append(label, del);
      box.appendChild(chip);
    });
  }
}
