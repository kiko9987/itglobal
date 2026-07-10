import ProjectStatusCalculator from '../utils/ProjectStatusCalculator.js';
import UnifiedBadgeSystem from './UnifiedBadgeSystem.js';
import AmountCalculator from '../utils/AmountCalculator.js';

import logger from '../utils/logger.js';
/**
 * Modern Project Filters 컴포넌트
 * 현대화된 프로젝트 목록 필터링 시스템
 * - 동적 드랍박스 생성 (company, client, status, manager)
 * - 정렬은 DataTables에서 처리 (기본: 프로젝트 코드 오름차순)
 * - 페이지네이션 고려
 */
export default class ModernProjectFilters {
  constructor() {
    this.filters = {};
    this.callbacks = [];
    this.currentData = null;
    this.isDataLoaded = false;
    this.sortOrder = { field: '프로젝트 코드', direction: 'asc' }; // 기본 정렬: 프로젝트 코드 오름차순 (실제로는 DataTables에서 처리)
    this.resultCount = 0;
    this.unifiedBadgeSystem = new UnifiedBadgeSystem(); // 통합 뱃지 시스템
    this.searchDebounceTimer = null; // 검색 디바운싱 타이머
    this.resignedManagers = []; // 퇴사 처리된 담당자 목록

    // 검색 대상 필드 정의 (성능 최적화)
    this.searchableFields = [
      '프로젝트 코드',
      '유입 구분',
      '담당자',
      '사업자명',  // E열 신규: 사업자등록증명
      '현장 주소',
      '공사 내용',
      '사업자',
      '발주처 담당자',
      '시공자',
      '계약금_메모',
      '중도금_메모',
      '잔금_메모',
      '수금 관련 특이사항'  // 수금 관련 특이사항 검색 지원
    ];
  }


  /**
   * 경고 로그 출력 (항상 출력)
   */
  warnLog(...args) {
    logger.warn('[ModernProjectFilters]', ...args);
  }

  /**
   * 필터 초기화
   */
  async init() {
    this.setupFilterElements();
    this.bindEvents();
    this.updateResultCount(0);
    await this.fetchResignedManagers();
    // 복원: sessionStorage 로 먼저 채우고, URL query 로 덮어쓴다 (URL 우선).
    //   2026-07-10 fix — 이전엔 URL에 필터 하나라도 있으면 storage restore 를 통째로
    //   스킵했다. URL과 storage 가 서로 다른 키를 갖고 있는 경우 (검색어는 storage에,
    //   프리셋 조건은 URL에) storage 값이 유실됐다.
    //   재현: 검색창에 R3070-TH → 데이터 미완료 프리셋 클릭 → Ctrl+Shift+R
    //     → URL={data:incomplete}, storage={search:R3070-TH, data:incomplete}
    //     → URL restore 만 실행 → filters={data:incomplete} → 검색창 비어보임.
    //   해결: 둘 다 실행. URL 이 나중에 실행되어 겹치는 키는 URL 값이 이김.
    this._restoreFiltersFromStorage();
    this._restoreFiltersFromUrl();
    // 2026-07-10 검색어·select 즉시 세팅 — dropdown 옵션 populate 대기 없이.
    //   증상: 새로고침 직후 데이터 로드(수 초) 사이에 담당자 등 select 필터가
    //     "전체" 로 보임. 매니저 눈엔 필터 안 걸린 것처럼 오인 → 다시 필터 클릭.
    //   해결: text input 은 즉시 세팅. select 는 옵션 없으면 임시 <option> 을
    //     만들어 붙여 값 표시. populate 완료 시 옵션 재생성되면서 자연스럽게 정상화.
    if (this.searchInput && this.filters?.search) {
      this.searchInput.value = this.filters.search;
    }
    const setSelectImmediate = (el, val) => {
      if (!el || !val) return;
      const exists = Array.from(el.options).some(o => o.value === val);
      if (!exists) {
        const opt = document.createElement('option');
        opt.value = val;
        opt.textContent = val;
        opt.dataset.tempPreload = '1';  // populate 시 정리 대상 (필요 시)
        el.appendChild(opt);
      }
      el.value = val;
    };
    setSelectImmediate(this.companyFilter, this.filters?.company);
    setSelectImmediate(this.clientFilter, this.filters?.client);
    setSelectImmediate(this.businessNameFilter, this.filters?.businessName);
    setSelectImmediate(this.statusFilter, this.filters?.status);
    setSelectImmediate(this.dataFilter, this.filters?.data);
    setSelectImmediate(this.managerFilter, this.filters?.manager);
    setSelectImmediate(this.outstandingFilter, this.filters?.outstanding);
    if (this.myProjectsOnlyCheckbox && this.filters?.myProjectsOnly !== undefined) {
      this.myProjectsOnlyCheckbox.checked = !!this.filters.myProjectsOnly;
    }
    // 프리셋 chip 렌더 (localStorage 로부터)
    this.renderPresetChips();
  }

  // ─────────────────────────────────────────────────────────────
  // 필터 프리셋 (localStorage, 브라우저에 영구 저장)
  // 매니저별로 자주 쓰는 필터 조합을 이름 붙여 저장 후 클릭 한번에 적용
  // ─────────────────────────────────────────────────────────────

  _getPresets() {
    try {
      const raw = localStorage.getItem('itg_filter_presets_v1');
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch (_) { return []; }
  }

  _savePresets(list) {
    try { localStorage.setItem('itg_filter_presets_v1', JSON.stringify(list)); }
    catch (_) { /* private mode 등 무시 */ }
  }

  savePresetPrompt() {
    const active = this.getActiveFilters();
    const receivablesOn = !!document.getElementById('receivablesToggle')?.checked;
    if (active.length === 0) {
      const msg = '저장할 활성 필터가 없습니다. 담당자·상태 등 필터를 먼저 설정하세요.';
      if (window.showSystemAlert) window.showSystemAlert(msg, 'info');
      else alert(msg);
      return;
    }
    const defaultName = active.join(' + ').slice(0, 40);
    const name = (prompt(`이 필터 조합의 이름을 지정하세요 (예: ${defaultName})`, defaultName) || '').trim();
    if (!name) return;

    const presets = this._getPresets();
    const existing = presets.findIndex(p => p.name === name);
    const entry = {
      name,
      filters: JSON.parse(JSON.stringify(this.filters || {})),
      receivables: receivablesOn,
      created_at: new Date().toISOString(),
    };
    if (existing >= 0) {
      if (!confirm(`'${name}' 프리셋이 이미 있습니다. 덮어쓸까요?`)) return;
      presets[existing] = entry;
    } else {
      presets.push(entry);
    }
    this._savePresets(presets);
    this.renderPresetChips();
    if (window.showSystemAlert) {
      window.showSystemAlert(`프리셋 저장됨: ${name}`, 'success');
    }
  }

  loadPreset(name) {
    const presets = this._getPresets();
    const preset = presets.find(p => p.name === name);
    if (!preset) return;

    // 편집 모드 중이면 change 이벤트로 확인 다이얼로그가 뜨는 경로에 진입시키기 위해
    // receivablesToggle change dispatch 를 먼저 시도. 취소되면 preset 적용 중단.
    const receivablesToggle = document.getElementById('receivablesToggle');
    const targetReceivables = !!preset.receivables;
    if (receivablesToggle && receivablesToggle.checked !== targetReceivables) {
      receivablesToggle.checked = targetReceivables;
      receivablesToggle.dispatchEvent(new Event('change', { bubbles: true }));
      // change 리스너가 편집 dirty 로 취소하면 checked 원복 → preset 적용 중단
      if (receivablesToggle.checked !== targetReceivables) return;
    }

    // 필터 적용
    this.filters = JSON.parse(JSON.stringify(preset.filters || {}));
    this._syncUIFromFilters();
    this.updateFilterVisualEffects();
    if (this.isDataLoaded) {
      this.applyFilters(null, true);
    }
    if (window.showSystemAlert) {
      window.showSystemAlert(`프리셋 적용됨: ${name}`, 'info');
    }
  }

  deletePreset(name) {
    if (!confirm(`'${name}' 프리셋을 삭제하시겠습니까?`)) return;
    const presets = this._getPresets().filter(p => p.name !== name);
    this._savePresets(presets);
    this.renderPresetChips();
  }

  renderPresetChips() {
    const container = document.getElementById('filterPresetsContainer');
    if (!container) return;
    const presets = this._getPresets();
    if (presets.length === 0) {
      container.innerHTML = '';
      return;
    }
    // 각 chip: [이름 · X]
    const html = presets.map(p => {
      const esc = (s) => String(s).replace(/[&<>"']/g, m => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
      })[m]);
      const escName = esc(p.name);
      const jsName = escName.replace(/'/g, '\\&#39;');
      const filterCount = Object.keys(p.filters || {}).length + (p.receivables ? 1 : 0);
      return `
        <span class="badge bg-light text-dark border" style="cursor: pointer; padding: 0.3rem 0.5rem; font-weight: 500;">
          <span onclick="modernFilters.loadPreset('${jsName}')" title="${escName} (${filterCount}개 필터) — 클릭하여 적용">
            <i class="fas fa-bookmark me-1 text-primary" style="font-size: 0.7rem;"></i>${escName}
          </span>
          <span onclick="event.stopPropagation(); modernFilters.deletePreset('${jsName}')"
                title="프리셋 삭제"
                style="margin-left: 0.4rem; color: #999; cursor: pointer;">
            <i class="fas fa-times" style="font-size: 0.7rem;"></i>
          </span>
        </span>
      `;
    }).join('');
    container.innerHTML = html;
  }

  /**
   * URL query string 에서 필터 상태 복원.
   * 예: ?manager=%EB%B0%95%EC%A0%95%EC%9A%B0&status=%EA%B3%B5%EC%82%AC%20%ED%99%95%EC%A0%95
   * @returns {boolean} 복원 여부 (URL 에 필터 파라미터 있었으면 true)
   */
  _restoreFiltersFromUrl() {
    try {
      const params = new URLSearchParams(window.location.search);
      const knownKeys = ['search', 'company', 'client', 'businessName', 'status',
                         'data', 'manager', 'outstanding', 'myProjectsOnly'];
      const restored = {};
      let hasAny = false;
      knownKeys.forEach(k => {
        const v = params.get(k);
        if (v !== null && v !== '') {
          restored[k] = k === 'myProjectsOnly' ? (v === '1' || v === 'true') : v;
          hasAny = true;
        }
      });
      if (hasAny) {
        // 2026-07-10 fix — 통째로 대체하지 않고 병합. sessionStorage 로 먼저 채운 값을
        //   덮어쓰되, URL 에 없는 키는 storage 값 유지 (검색어가 storage 에만 있어도 복원).
        this.filters = { ...(this.filters || {}), ...restored };
        logger.debug('[Filters] URL query 에서 복원 (병합):', restored);
        return true;
      }
    } catch (err) {
      logger.warn('[Filters] URL 복원 실패:', err);
    }
    return false;
  }

  /**
   * 현재 필터 상태를 URL query string 에 반영 (history.replaceState).
   * 다른 URL params (예: receivables) 은 그대로 유지.
   */
  _persistFiltersToUrl() {
    try {
      const params = new URLSearchParams(window.location.search);
      const knownKeys = ['search', 'company', 'client', 'businessName', 'status',
                         'data', 'manager', 'outstanding', 'myProjectsOnly'];
      knownKeys.forEach(k => {
        const v = this.filters[k];
        if (v && v !== '') {
          params.set(k, k === 'myProjectsOnly' ? (v ? '1' : '') : String(v));
        } else {
          params.delete(k);
        }
      });
      const qs = params.toString();
      const newUrl = window.location.pathname + (qs ? '?' + qs : '') + window.location.hash;
      // pushState 아님 — 필터 변경마다 history 스택 쌓이면 back 버튼 오작동
      window.history.replaceState({}, '', newUrl);
    } catch (_) { /* private mode 등 무시 */ }
  }

  /**
   * sessionStorage 에서 필터 상태 복원.
   * this.filters 만 채우고, UI select 값은 populateAllFilters 후 _syncUIFromFilters 로 반영.
   */
  _restoreFiltersFromStorage() {
    try {
      const raw = sessionStorage.getItem('itg_filters_v1');
      if (!raw) return;
      const saved = JSON.parse(raw);
      if (saved && typeof saved === 'object') {
        this.filters = saved;
        logger.debug('[Filters] sessionStorage 에서 복원:', saved);
      }
    } catch (err) {
      logger.warn('[Filters] 복원 실패:', err);
    }
  }

  /**
   * 현재 this.filters 값을 UI select/input 요소에 반영.
   * populateAllFilters 후 (옵션 채워진 후) 호출해야 select value 세팅 가능.
   */
  _syncUIFromFilters() {
    const f = this.filters || {};
    // 옵션에 값이 실제로 존재할 때만 select value 세팅.
    //   2026-07-10 fix — 기존 로직은 `el.value = wanted` 후 세팅 실패 시 filters 를
    //   삭제했다. 문제는 "옵션이 아직 populate 안 됨"(로딩 지연)과 "저장된 값이 옵션에
    //   없음"(stale) 을 구분 못 해서 로딩 순서 이슈에도 filters 를 지워버림.
    //   결과: storage 에 manager="이상덕" 있고 populate 후 옵션도 있는데 selectVal
    //   빈 상태로 남음.
    //   해결: 옵션 존재 확인 후에만 세팅. 존재 안 하면 filters 유지 (populate 후
    //   다시 sync 될 때 세팅되도록 대기).
    const setSelectSafe = (el, filterKey) => {
      if (!el || !(filterKey in f)) return;
      const wanted = f[filterKey] || '';
      if (!wanted) { el.value = ''; return; }
      const exists = Array.from(el.options).some(opt => opt.value === wanted);
      if (exists) {
        el.value = wanted;
      }
      // 옵션 없음: filters 는 유지. 옵션이 나중에 채워지면 재호출 시 세팅됨.
      //   진짜 stale (매니저 삭제 등) 인 경우엔 데이터 로드 후 populateAllFilters →
      //   restoreFilterSelections 가 별도로 정리.
    };
    if (this.searchInput && 'search' in f) this.searchInput.value = f.search || '';
    setSelectSafe(this.companyFilter, 'company');
    setSelectSafe(this.clientFilter, 'client');
    setSelectSafe(this.businessNameFilter, 'businessName');
    setSelectSafe(this.statusFilter, 'status');
    setSelectSafe(this.dataFilter, 'data');
    setSelectSafe(this.managerFilter, 'manager');
    setSelectSafe(this.outstandingFilter, 'outstanding');
    if (this.myProjectsOnlyCheckbox && f.myProjectsOnly !== undefined) {
      this.myProjectsOnlyCheckbox.checked = !!f.myProjectsOnly;
    }
  }

  /**
   * 현재 필터 상태 sessionStorage 에 저장.
   */
  _persistFiltersToStorage() {
    try {
      // myProjectsOnly 체크박스는 filters 에 별도 표시 (managerFilter 와 다름)
      const snapshot = { ...this.filters };
      if (this.myProjectsOnlyCheckbox && this.myProjectsOnlyCheckbox.checked) {
        snapshot.myProjectsOnly = true;
      }
      sessionStorage.setItem('itg_filters_v1', JSON.stringify(snapshot));
    } catch (_) { /* private mode 등 무시 */ }
  }

  /**
   * 필터 엘리먼트 설정
   */
  setupFilterElements() {
    this.searchInput = document.getElementById('searchInput');
    this.companyFilter = document.getElementById('companyFilter');
    this.clientFilter = document.getElementById('clientFilter');
    this.businessNameFilter = document.getElementById('businessNameFilter');
    this.statusFilter = document.getElementById('statusFilter');
    this.dataFilter = document.getElementById('dataFilter');
    this.managerFilter = document.getElementById('managerFilter');
    this.outstandingFilter = document.getElementById('outstandingFilter');
    this.myProjectsOnlyCheckbox = document.getElementById('myProjectsOnly');
    this.resultCountElement = document.getElementById('filterResultCount');
    this.filterSection = document.querySelector('.filter-section');
    this.resultDisplayContainer = this.filterSection ? this.filterSection.querySelector('.filter-result-display') : null;

    if (this.resultCountElement) {
      this.resultCountElement.classList.add('filter-result-count');
    }
  }

  /**
   * 이벤트 바인딩
   */
  bindEvents() {
    // 검색 입력 (디바운싱 적용)
    if (this.searchInput) {
      this.searchInput.addEventListener('input', (e) => {
        this.filters.search = e.target.value;

        // 기존 타이머 취소
        if (this.searchDebounceTimer) {
          clearTimeout(this.searchDebounceTimer);
        }

        // 300ms 후에 필터 적용 (사용자가 타이핑을 멈출 때까지 대기)
        this.searchDebounceTimer = setTimeout(() => {
          this.applyFilters(null, true);
        }, 300);
      });
    }

    // 사업자 필터
    if (this.companyFilter) {
      this.companyFilter.addEventListener('change', (e) => {
        const value = e.target.value;
        if (value && value !== '' && value !== '전체') {
          this.filters.company = value;
        } else {
          delete this.filters.company;
        }
        this.applyFilters(null, true);
        e.target.blur(); // 포커스 제거
      });
    }

    // 거래처 필터 (유입채널)
    if (this.clientFilter) {
      this.clientFilter.addEventListener('change', (e) => {
        const value = e.target.value;
        if (value && value !== '' && value !== '전체') {
          this.filters.client = value;
        } else {
          delete this.filters.client;
        }
        this.applyFilters(null, true);
        e.target.blur(); // 포커스 제거
      });
    }

    // 사업자명 필터 (E열, 신규)
    if (this.businessNameFilter) {
      this.businessNameFilter.addEventListener('change', (e) => {
        const value = e.target.value;
        if (value && value !== '' && value !== '전체') {
          this.filters.businessName = value;
        } else {
          delete this.filters.businessName;
        }
        this.applyFilters(null, true);
        e.target.blur(); // 포커스 제거
      });
    }

    // 상태 필터
    if (this.statusFilter) {
      this.statusFilter.addEventListener('change', (e) => {
        const value = e.target.value;
        if (value && value !== '' && value !== '전체') {
          this.filters.status = value;
        } else {
          delete this.filters.status;
        }
        this.applyFilters(null, true);
        e.target.blur(); // 포커스 제거
      });
    }

    // 데이터 필터
    if (this.dataFilter) {
      this.dataFilter.addEventListener('change', (e) => {
        const value = e.target.value;
        if (value && value !== '' && value !== '전체') {
          this.filters.data = value;
        } else {
          delete this.filters.data;
        }
        this.applyFilters(null, true);
        e.target.blur(); // 포커스 제거
      });
    }

    // 담당자 필터
    if (this.managerFilter) {
      this.managerFilter.addEventListener('change', (e) => {
        const value = e.target.value;
        if (value && value !== '' && value !== '전체') {
          this.filters.manager = value;
        } else {
          delete this.filters.manager;
        }
        this.applyFilters(null, true);
        e.target.blur(); // 포커스 제거
      });
    }

    // 미수금 필터
    if (this.outstandingFilter) {
      this.outstandingFilter.addEventListener('change', (e) => {
        const value = e.target.value;
        if (value && value !== '' && value !== '전체') {
          this.filters.outstanding = value;
        } else {
          delete this.filters.outstanding;
        }
        this.applyFilters(null, true);
        e.target.blur(); // 포커스 제거
      });
    }

    // 내 공사만 보기 체크박스 — 담당자 필터를 로그인 사용자 이름으로 자동 세팅.
    // 2026-07-07: 옛 로직은 '발주처 이메일'(고객 이메일)과 로그인 이메일을 비교해 항상
    // 0건. 이미 있는 담당자 select 필터와 시맨틱 통일이 더 간단·명확.
    // 2026-07-09: 체크 전 담당자 값 백업 → 언체크 시 복원 (사용자 상태 파괴 방지)
    if (this.myProjectsOnlyCheckbox) {
      this.myProjectsOnlyCheckbox.addEventListener('change', (e) => {
        if (e.target.checked) {
          // 체크: 현재 담당자 값 백업 후 로그인 사용자로 덮어씀
          this._managerBackupBeforeMine = this.filters.manager || '';
          const userName = (window.userDisplayName || '').trim();
          if (userName) {
            this.filters.manager = userName;
            if (this.managerFilter) this.managerFilter.value = userName;
          }
        } else {
          // 언체크: 현재 담당자 필터가 로그인 사용자 이름 그대로면 → 백업 복원 or clear
          //        다른 사람으로 명시적 변경돼 있으면 → 그대로 유지 (사용자 의도 존중)
          const userName = (window.userDisplayName || '').trim();
          const currentManager = (this.filters.manager || '').trim();
          if (currentManager === userName) {
            const restored = this._managerBackupBeforeMine || '';
            if (restored) {
              this.filters.manager = restored;
              if (this.managerFilter) this.managerFilter.value = restored;
            } else {
              delete this.filters.manager;
              if (this.managerFilter) this.managerFilter.value = '';
            }
          }
          // 담당자가 다른 사람으로 바뀌었으면 그대로 (사용자가 이미 명시적으로 다른 필터 세팅)
          this._managerBackupBeforeMine = null;
        }
        this.applyFilters(null, true);
      });
    }
  }

  /**
   * 필터 적용 및 정렬
   * @param {Array} data - 새로운 데이터 (옵션널, 전달 시에만 데이터 업데이트 수행)
   * @param {boolean} triggerCallbacks - 콜백 실행 여부 (기본값: true)
   */
  applyFilters(data = null, triggerCallbacks = true) {
    const startTime = performance.now();

    // 새 데이터가 전달되면 현재 데이터로 저장하고 동적 옵션 생성
    let dataUpdated = false;
    if (data && Array.isArray(data)) {
      this.currentData = data;
      this.isDataLoaded = true;
      this.populateAllFilters(data);
      // 동적 옵션 생성 후 저장된 필터 값을 UI select 에 반영
      // (init 단계에서 복원한 this.filters 값이 dropdown 채워지길 기다렸다가 select value 세팅)
      this._syncUIFromFilters();
      // 시각적 효과도 최신 상태 반영
      this.updateFilterVisualEffects();
      dataUpdated = true;
    }

    // 데이터가 아직 로드되지 않은 경우
    if (!this.isDataLoaded || this.currentData === null) {
      this.warnLog('데이터가 아직 로드되지 않음');
      this.updateResultCount(0);
      return null;
    }

    // 데이터는 로드되었지만 빈 배열인 경우
    if (this.currentData.length === 0) {
      this.warnLog('빈 데이터 배열');
      this.updateResultCount(0);
      if (triggerCallbacks) {
        this.callbacks.forEach(callback => callback([]));
      }
      return [];
    }

    const originalCount = this.currentData.length;
    let filteredData = [...this.currentData];


    // 검색어 필터 (최적화: 주요 필드만 검색)
    if (this.filters.search) {
      const searchTerm = this.filters.search.toLowerCase();
      filteredData = filteredData.filter(item => {
        // 정의된 검색 가능 필드만 검색하여 성능 향상
        return this.searchableFields.some(field => {
          const value = item[field];
          return value && String(value).toLowerCase().includes(searchTerm);
        });
      });
    }

    // 사업자 필터
    if (this.filters.company) {
      filteredData = filteredData.filter(item => {
        const company = item['사업자'] || '';
        return company.toString().trim() === this.filters.company;
      });
    }

    // 유입 구분 필터 (유입 채널 카테고리)
    if (this.filters.client) {
      filteredData = filteredData.filter(item => {
        const client = item['유입 구분'] || '';
        return client.toString().trim() === this.filters.client;
      });
    }

    // 사업자명 필터 (E열 신규)
    if (this.filters.businessName) {
      filteredData = filteredData.filter(item => {
        const businessName = item['사업자명'] || '';
        return businessName.toString().trim() === this.filters.businessName;
      });
    }

    // 상태 필터 (공통 유틸리티 사용)
    if (this.filters.status) {
      filteredData = filteredData.filter(item => {
        const statusText = ProjectStatusCalculator.calculateStatus(item);
        return statusText === this.filters.status;
      });
    }

    // 데이터 완성도 필터
    if (this.filters.data) {
      filteredData = filteredData.filter(item => {
        const isEmpty = this.isDataIncomplete(item);
        if (this.filters.data === 'complete') {
          return !isEmpty; // 완료: 빈 데이터가 없음
        } else if (this.filters.data === 'incomplete') {
          return isEmpty; // 미완료: 빈 데이터가 있음
        }
        return true;
      });
    }

    // 담당자 필터
    if (this.filters.manager) {
      filteredData = filteredData.filter(item => {
        const manager = item['담당자'] || '';
        return manager.toString().trim() === this.filters.manager;
      });
    }

    // '내 공사만 보기' 별도 필터는 없음 — 체크박스 change 핸들러에서 this.filters.manager를
    // 세팅해 위쪽 담당자 필터가 처리하도록 위임 (2026-07-07 변경).

    // 미수금 필터
    if (this.filters.outstanding) {
      const beforeCount = filteredData.length;
      filteredData = filteredData.filter((item, index) => {
        const outstandingData = item['미수금'] || 0;
        const outstandingAmount = AmountCalculator.safeParseCurrency(outstandingData);
        const collectedValue = item['수금 확인'];
        const isCollected = collectedValue === true ||
                           collectedValue === 'TRUE' ||
                           collectedValue === 'true' ||
                           collectedValue === 1 ||
                           collectedValue === '1';

        if (this.filters.outstanding === 'collected') {
          // 수금 완료: 수금확인 체크박스가 체크된 경우만
          return isCollected;
        } else if (this.filters.outstanding === 'outstanding') {
          // 미수금 있음: 미수금이 0보다 크고 수금확인이 true가 아닌 경우
          return outstandingAmount > 0 && !isCollected;
        }
        return true;
      });

      // 수금 관리 모드에서는 취소된 공사 제외 (띄어쓰기 무시)
      // (2026-07-08: item['AG'] fallback 제거 — 컬럼 시프트 후 AG=마진율이라 잘못된 참조)
      filteredData = filteredData.filter(item => {
        const collectionNotes = item['수금 관련 특이사항'] || '';
        return !/공사\s*취소/.test(collectionNotes);
      });
    }

    // 필터링 후 프로젝트 코드 기준 숫자 내림차순 정렬
    filteredData.sort((a, b) => {
      const aCode = a['프로젝트 코드'] || '';
      const bCode = b['프로젝트 코드'] || '';

      // 프로젝트 코드에서 숫자 부분 추출 (예: "G0123-AB" → 123)
      const aNumber = parseInt(aCode.match(/\d+/)?.[0] || '0');
      const bNumber = parseInt(bCode.match(/\d+/)?.[0] || '0');

      // 숫자 내림차순 정렬
      return bNumber - aNumber;
    });


    // 결과 수 업데이트
    this.updateResultCount(filteredData.length);

    // 성능 측정
    const endTime = performance.now();
    const processingTime = endTime - startTime;

    // 필터 시각적 효과 업데이트
    this.updateFilterVisualEffects();

    // 필터 상태 sessionStorage 저장 (F5 시 복원용)
    this._persistFiltersToStorage();
    // 필터 상태 URL 반영 (북마크·공유·F5 복원 우선순위 1)
    this._persistFiltersToUrl();

    // 콜백 함수들 실행 (조건부)
    if (triggerCallbacks) {
      this.callbacks.forEach((callback, index) => {
        try {
          callback(filteredData);
        } catch (error) {
          logger.error('[ModernProjectFilters] 콜백', index, '실행 오류:', error);
        }
      });
    }

    return filteredData;
  }

  /**
   * 데이터 정렬 (기본: 프로젝트 코드 숫자 내림차순)
   */
  sortData(data) {
    return data.sort((a, b) => {
      const aValue = a[this.sortOrder.field] || '';
      const bValue = b[this.sortOrder.field] || '';

      // 프로젝트 코드는 숫자 부분으로 정렬
      if (this.sortOrder.field === '프로젝트 코드') {
        const aNumber = parseInt(aValue.match(/\d+/)?.[0] || '0');
        const bNumber = parseInt(bValue.match(/\d+/)?.[0] || '0');

        if (this.sortOrder.direction === 'desc') {
          return bNumber - aNumber; // 숫자 내림차순
        } else {
          return aNumber - bNumber; // 숫자 오름차순
        }
      }

      // 다른 필드들의 정렬 로직
      if (this.sortOrder.direction === 'desc') {
        return bValue > aValue ? 1 : -1;
      } else {
        return aValue > bValue ? 1 : -1;
      }
    });
  }

  /**
   * 모든 필터 옵션을 동적으로 생성 (기존 선택값 보존)
   */
  populateAllFilters(data) {
    logger.debug('[ModernProjectFilters] 필터 옵션 업데이트 시작');

    // 현재 선택값들 저장
    const currentSelections = {
      company: this.companyFilter?.value || '',
      client: this.clientFilter?.value || '',
      businessName: this.businessNameFilter?.value || '',
      status: this.statusFilter?.value || '',
      manager: this.managerFilter?.value || ''
    };

    // 필터 옵션 재생성
    this.populateCompanyFilter(data);
    this.populateClientFilter(data);
    this.populateBusinessNameFilter(data);  // E열 사업자명 (신규)
    this.populateStatusFilter(data);
    this.populateManagerFilter(data);

    // 이전 선택값 복원 (옵션이 여전히 존재하는 경우만)
    this.restoreFilterSelections(currentSelections);

    logger.debug('[ModernProjectFilters] 필터 옵션 업데이트 완료');
  }

  /**
   * 필터 선택값 복원
   */
  restoreFilterSelections(selections) {
    // 회사 필터 복원
    if (selections.company && this.companyFilter) {
      const companyOption = Array.from(this.companyFilter.options).find(opt => opt.value === selections.company);
      if (companyOption) {
        this.companyFilter.value = selections.company;
        this.filters.company = selections.company;
      } else {
        // 옵션이 더 이상 존재하지 않는 경우 초기화
        this.companyFilter.value = '';
        delete this.filters.company;
      }
    }

    // 거래처 필터 복원
    if (selections.client && this.clientFilter) {
      const clientOption = Array.from(this.clientFilter.options).find(opt => opt.value === selections.client);
      if (clientOption) {
        this.clientFilter.value = selections.client;
        this.filters.client = selections.client;
      } else {
        // 옵션이 더 이상 존재하지 않는 경우 초기화
        this.clientFilter.value = '';
        delete this.filters.client;
      }
    }

    // 사업자명 필터 복원 (E열 신규)
    if (selections.businessName && this.businessNameFilter) {
      const opt = Array.from(this.businessNameFilter.options).find(o => o.value === selections.businessName);
      if (opt) {
        this.businessNameFilter.value = selections.businessName;
        this.filters.businessName = selections.businessName;
      } else {
        this.businessNameFilter.value = '';
        delete this.filters.businessName;
      }
    }

    // 상태 필터 복원
    if (selections.status && this.statusFilter) {
      const statusOption = Array.from(this.statusFilter.options).find(opt => opt.value === selections.status);
      if (statusOption) {
        this.statusFilter.value = selections.status;
        this.filters.status = selections.status;
      } else {
        // 옵션이 더 이상 존재하지 않는 경우 초기화
        this.statusFilter.value = '';
        delete this.filters.status;
      }
    }

    // 담당자 필터 복원
    if (selections.manager && this.managerFilter) {
      const managerOption = Array.from(this.managerFilter.options).find(opt => opt.value === selections.manager);
      if (managerOption) {
        this.managerFilter.value = selections.manager;
        this.filters.manager = selections.manager;
      } else {
        // 옵션이 더 이상 존재하지 않는 경우 초기화
        this.managerFilter.value = '';
        delete this.filters.manager;
      }
    }
  }

  /**
   * 사업자 필터 옵션 동적 생성
   */
  populateCompanyFilter(data) {
    if (!this.companyFilter || !data || !Array.isArray(data)) return;

    // 고유한 사업자 목록 추출 (map 단계에서 바로 trim 적용)
    const companies = [...new Set(data.map(item => (item['사업자'] || '').trim()).filter(company => company))];

    // 기존 옵션 제거 (첫 번째 "전체" 옵션은 유지)
    while (this.companyFilter.children.length > 1) {
      this.companyFilter.removeChild(this.companyFilter.lastChild);
    }

    // 사업자 옵션 추가
    companies.sort().forEach(company => {
      const option = document.createElement('option');
      option.value = company;
      option.textContent = company;
      this.companyFilter.appendChild(option);
    });
  }

  /**
   * 유입 구분 필터 옵션 동적 생성.
   * 기본은 프로젝트 D열 유니크 값으로 즉시 채우고, /api/inflow-options 를 백그라운드로
   * 조회해 시트 dataValidation 값으로 덮어씀 (사용자가 시트에서 바꾼 값 즉시 반영).
   */
  populateClientFilter(data) {
    if (!this.clientFilter || !data || !Array.isArray(data)) return;

    // 1) 프로젝트 D열 유니크 값으로 즉시 렌더 (fallback)
    const clients = [...new Set(data.map(item => (item['유입 구분'] || '').trim()).filter(client => client))];

    while (this.clientFilter.children.length > 1) {
      this.clientFilter.removeChild(this.clientFilter.lastChild);
    }
    clients.sort((a, b) => a.localeCompare(b, 'ko')).forEach(client => {
      const option = document.createElement('option');
      option.value = client;
      option.textContent = client;
      this.clientFilter.appendChild(option);
    });

    // 2) 시트 유효성 검사 값으로 덮어쓰기
    this._refreshClientFilterFromSheet();
  }

  async _refreshClientFilterFromSheet() {
    try {
      const res = await fetch('/api/inflow-options', { credentials: 'same-origin' });
      if (!res.ok) return;
      const json = await res.json();
      const options = Array.isArray(json.options) ? json.options.filter(Boolean) : [];
      if (!options.length || !this.clientFilter) return;

      const currentValue = this.clientFilter.value;
      while (this.clientFilter.children.length > 1) {
        this.clientFilter.removeChild(this.clientFilter.lastChild);
      }
      options.forEach(v => {
        const opt = document.createElement('option');
        opt.value = v;
        opt.textContent = v;
        this.clientFilter.appendChild(opt);
      });
      if (currentValue && options.includes(currentValue)) {
        this.clientFilter.value = currentValue;
      }
    } catch (err) {
      console.warn('[FILTER] 유입 구분 옵션 갱신 실패:', err);
    }
  }

  /**
   * 사업자명 필터 옵션 동적 생성 (E열, 신규)
   */
  populateBusinessNameFilter(data) {
    if (!this.businessNameFilter || !data || !Array.isArray(data)) return;

    // 고유한 사업자명 목록 추출
    const businessNames = [...new Set(data.map(item => (item['사업자명'] || '').trim()).filter(name => name))];

    // 기존 옵션 제거 (첫 번째 "전체" 옵션은 유지)
    while (this.businessNameFilter.children.length > 1) {
      this.businessNameFilter.removeChild(this.businessNameFilter.lastChild);
    }

    // 사업자명 옵션 추가 (가나다순)
    businessNames.sort((a, b) => a.localeCompare(b, 'ko')).forEach(name => {
      const option = document.createElement('option');
      option.value = name;
      option.textContent = name;
      this.businessNameFilter.appendChild(option);
    });
  }

  /**
   * 상태 필터 옵션 동적 생성 (계산된 상태값 기반)
   */
  populateStatusFilter(data) {
    if (!this.statusFilter || !data || !Array.isArray(data)) return;

    // 모든 데이터에 대한 상태값 계산 및 추출 (공통 유틸리티 사용)
    const statusSet = new Set();
    data.forEach(item => {
      const statusText = ProjectStatusCalculator.calculateStatus(item);
      if (statusText) {
        statusSet.add(statusText);
      }
    });

    const statuses = Array.from(statusSet);

    // 기존 옵션 제거 (첫 번째 "전체" 옵션은 유지)
    while (this.statusFilter.children.length > 1) {
      this.statusFilter.removeChild(this.statusFilter.lastChild);
    }

    // 상태 순서 정의 (공통 유틸리티에서 가져옴)
    const statusOrder = ProjectStatusCalculator.getStatusOrder();
    const sortedStatuses = statuses.sort((a, b) => {
      const aIndex = statusOrder.indexOf(a);
      const bIndex = statusOrder.indexOf(b);
      if (aIndex === -1 && bIndex === -1) return a.localeCompare(b);
      if (aIndex === -1) return 1;
      if (bIndex === -1) return -1;
      return aIndex - bIndex;
    });

    // 상태 옵션 추가
    sortedStatuses.forEach(status => {
      const option = document.createElement('option');
      option.value = status;
      option.textContent = status;
      this.statusFilter.appendChild(option);
    });
  }

  /**
   * 퇴사 처리된 담당자 목록 가져오기
   */
  async fetchResignedManagers() {
    try {
      const response = await fetch('/api/resigned-managers');
      const data = await response.json();

      if (data.success && Array.isArray(data.resigned_managers)) {
        // API는 [{name: 'X', email: 'Y'}, ...] 형태의 객체 배열을 반환
        // 이름만 추출하여 문자열 배열로 변환
        this.resignedManagers = data.resigned_managers.map(user => user.name.trim());
        logger.debug(`[ModernProjectFilters] 퇴사자 ${this.resignedManagers.length}명 로드됨:`, this.resignedManagers);
      } else {
        logger.warn('[ModernProjectFilters] 퇴사자 목록 로드 실패:', data);
        this.resignedManagers = [];
      }
    } catch (error) {
      logger.error('[ModernProjectFilters] 퇴사자 목록 로드 오류:', error);
      this.resignedManagers = [];
    }
  }

  /**
   * 담당자 필터 옵션 동적 생성 (퇴사자 제외)
   */
  populateManagerFilter(data) {
    if (!this.managerFilter || !data || !Array.isArray(data)) return;

    // 고유한 담당자 목록 추출 (map 단계에서 바로 trim 적용)
    const allManagers = [...new Set(data.map(item => (item['담당자'] || '').trim()).filter(manager => manager))];

    // 퇴사자 제외 (퇴사자는 키워드 검색으로만 찾을 수 있음)
    const activeManagers = allManagers.filter(manager => !this.resignedManagers.includes(manager));

    // 기존 옵션 제거 (첫 번째 "전체" 옵션은 유지)
    while (this.managerFilter.children.length > 1) {
      this.managerFilter.removeChild(this.managerFilter.lastChild);
    }

    // 담당자 옵션 추가 (퇴사자 제외)
    activeManagers.sort().forEach(manager => {
      const option = document.createElement('option');
      option.value = manager;
      option.textContent = manager;
      this.managerFilter.appendChild(option);
    });

    logger.debug(`[ModernProjectFilters] 담당자 필터 생성: 전체 ${allManagers.length}명 중 ${activeManagers.length}명 표시 (퇴사자 ${this.resignedManagers.length}명 제외)`);
  }

  /**
   * 상태값 계산 (공통 유틸리티 사용) - 중복 제거
   */
  getStatusText(rowData) {
    return ProjectStatusCalculator.calculateStatus(rowData);
  }

  /**
   * 상태 배지 생성 - 통합 뱃지 시스템 사용
   */
  getStatusBadge(rowData) {
    return this.unifiedBadgeSystem.createBadge('status', rowData);
  }

  /**
   * 결과 수 표시 업데이트
   */
  updateResultCount(count) {
    this.resultCount = count;
    if (!this.resultCountElement) return;

    if (count === 0 && !this.isDataLoaded) {
      this.resultCountElement.textContent = '로딩 중...';
      this.resultCountElement.title = '';
      return;
    }

    // 기본 표시
    this.resultCountElement.textContent = `${count.toLocaleString()}개 프로젝트`;
    this.resultCountElement.title = '';

    // 결과 0건이고 활성 필터 있으면 짧은 힌트 (매니저 헤맴 방지).
    //   2026-07-10 fix — 이전엔 활성 필터 이름을 다 나열해서 텍스트가 길어졌고
    //   옆 "N개 활성" 배지를 화면 밖으로 밀어냈다. 개수만 노출하고 상세는 title 로.
    if (count === 0 && this.isDataLoaded) {
      const active = this.getActiveFilters();
      if (active.length > 0) {
        this.resultCountElement.textContent = `0개 · 필터 ${active.length}개로 좁혀졌어요`;
        this.resultCountElement.title =
          `활성 필터: ${active.join(', ')}\n초기화 버튼을 눌러 모두 지우거나 하나씩 해제해 조건을 넓혀보세요.`;
      }
    }
  }

  /**
   * 필터 변경 콜백 등록
   */
  onFilterChange(callback) {
    this.callbacks.push(callback);
  }

  /**
   * 검색 입력창에 포커스
   */
  focusSearch() {
    if (this.searchInput) {
      this.searchInput.focus();
    }
  }

  /**
   * 필터 초기화
   */
  resetFilters() {
    // 수금 관리 토글부터 처리 (편집 모드였을 때 dirty 확인 다이얼로그 우선 노출)
    // 사용자가 취소하면 나머지 필터 리셋도 안 하는 게 자연스러움.
    const receivablesToggle = document.getElementById('receivablesToggle');
    if (receivablesToggle && receivablesToggle.checked) {
      receivablesToggle.checked = false;
      receivablesToggle.dispatchEvent(new Event('change', { bubbles: true }));
      // change 리스너가 편집 모드 확인 다이얼로그에서 '취소' 를 받으면
      // receivablesToggle.checked 를 true 로 원복함 → 이 경우 나머지 필터도 리셋 안 함.
      if (receivablesToggle.checked) {
        return;
      }
    }

    // 이제 나머지 필터 리셋
    this.filters = {};

    if (this.searchInput) this.searchInput.value = '';
    if (this.companyFilter) this.companyFilter.value = '';
    if (this.clientFilter) this.clientFilter.value = '';
    if (this.businessNameFilter) this.businessNameFilter.value = '';
    if (this.statusFilter) this.statusFilter.value = '';
    if (this.dataFilter) this.dataFilter.value = '';
    if (this.managerFilter) this.managerFilter.value = '';
    if (this.outstandingFilter) this.outstandingFilter.value = '';
    if (this.myProjectsOnlyCheckbox) this.myProjectsOnlyCheckbox.checked = false;

    // sessionStorage 삭제 (다음 F5 때 복원되지 않도록)
    try {
      sessionStorage.removeItem('itg_filters_v1');
      sessionStorage.removeItem('itg_receivables_mode');
    } catch (_) {}

    // URL query 도 정리 (북마크 상태 반영)
    try {
      const params = new URLSearchParams(window.location.search);
      ['search', 'company', 'client', 'businessName', 'status', 'data',
       'manager', 'outstanding', 'myProjectsOnly', 'receivables'].forEach(k => params.delete(k));
      const qs = params.toString();
      const newUrl = window.location.pathname + (qs ? '?' + qs : '') + window.location.hash;
      window.history.replaceState({}, '', newUrl);
    } catch (_) {}

    // 시각적 효과 즉시 업데이트 (필터 초기화)
    this.updateFilterVisualEffects();

    // 데이터가 로드된 경우에만 필터 적용 (콜백 실행)
    if (this.isDataLoaded) {
      this.applyFilters(null, true);
    }
  }

  /**
   * 데이터 상태 초기화
   */
  clearData() {
    this.currentData = null;
    this.isDataLoaded = false;
    this.updateResultCount(0);
  }

  /**
   * 현재 필터 상태 확인
   */
  hasFilters() {
    return Object.keys(this.filters).length > 0;
  }

  /**
   * 데이터 로드 상태 확인
   */
  isReady() {
    return this.isDataLoaded && this.currentData !== null;
  }

  /**
   * 현재 설정된 정렬 순서 반환
   */
  getSortOrder() {
    return this.sortOrder;
  }

  /**
   * 정렬 순서 변경
   */
  setSortOrder(field, direction = 'desc') {
    this.sortOrder = { field, direction };
    if (this.isDataLoaded) {
      this.applyFilters(null, true);
    }
  }

  /**
   * 필터 시각적 효과 업데이트
   */
  updateFilterVisualEffects() {
    if (!this.filterSection) return;

    const activeFilters = this.getActiveFilters();
    const hasActiveFilters = activeFilters.length > 0;

    // 필터 섹션 강조 효과
    if (hasActiveFilters) {
      this.filterSection.classList.add('has-active-filters');
      this.addActiveFilterBadge(activeFilters.length);
    } else {
      this.filterSection.classList.remove('has-active-filters');
      this.removeActiveFilterBadge();
    }

    // 개별 필터 요소 강조
    this.updateIndividualFilterEffects();
  }

  /**
   * 활성 필터 목록 반환
   */
  getActiveFilters() {
    const active = [];
    if (this.filters.company) active.push('사업자');
    if (this.filters.client) active.push('유입 구분');
    if (this.filters.businessName) active.push('사업자명');
    if (this.filters.status) active.push('상태');
    if (this.filters.data) active.push('데이터');
    if (this.filters.manager) active.push('담당자');
    if (this.filters.outstanding) active.push('미수금');
    // 수금 관리 모드 토글도 상단에서 활성이면 카운트에 포함 (매니저 인지 편의)
    const receivablesToggle = document.getElementById('receivablesToggle');
    if (receivablesToggle?.checked) active.push('수금 관리');
    if (this.filters.myProjectsOnly) active.push('내 공사');
    if (this.searchInput && this.searchInput.value.trim()) active.push('검색');
    return active;
  }

  /**
   * 활성 필터 배지 추가
   */
  addActiveFilterBadge(count) {
    let badge = this.filterSection.querySelector('.active-filter-badge');
    if (!badge) {
      badge = document.createElement('span');
      badge.className = 'active-filter-badge';
      // 결과 표시 컨테이너 안에 배지 추가
      if (this.resultDisplayContainer) {
        this.resultDisplayContainer.appendChild(badge);
      } else {
        const fallback = this.filterSection?.querySelector('.d-flex.gap-2.align-items-center');
        fallback?.appendChild(badge);
      }
    }
    badge.innerHTML = `<i class="fas fa-filter"></i>${count}개 활성`;
  }

  /**
   * 활성 필터 배지 제거
   */
  removeActiveFilterBadge() {
    const badge = this.filterSection.querySelector('.active-filter-badge');
    if (badge) {
      badge.remove();
    }
  }

  /**
   * 개별 필터 요소 효과 업데이트
   */
  updateIndividualFilterEffects() {
    // 각 필터 요소에 선택된 상태 클래스 추가/제거
    const filterElements = [
      { element: this.companyFilter, key: 'company' },
      { element: this.clientFilter, key: 'client' },
      { element: this.statusFilter, key: 'status' },
      { element: this.dataFilter, key: 'data' },
      { element: this.managerFilter, key: 'manager' },
      { element: this.outstandingFilter, key: 'outstanding' },
      { element: this.myProjectsOnlyCheckbox, key: 'myProjectsOnly', isCheckbox: true },
      { element: this.searchInput, key: 'search', isInput: true }
    ];

    filterElements.forEach(({ element, key, isInput, isCheckbox }) => {
      if (!element) return;

      let hasValue;
      if (isInput) {
        hasValue = element.value.trim() !== '';
      } else if (isCheckbox) {
        hasValue = element.checked;
      } else {
        hasValue = this.filters[key] && this.filters[key] !== '' && this.filters[key] !== '전체';
      }

      if (hasValue) {
        element.classList.add('filter-selected');
      } else {
        element.classList.remove('filter-selected');
      }
    });
  }

  /**
   * 데이터가 미완료인지 확인
   * 중요한 필드들이 비어있으면 미완료로 판단
   */
  isDataIncomplete(item) {
    if (!item) return true;

    // 체크할 중요 필드들 (ProjectTable.js와 동일한 필드명 사용)
    const importantFields = [
      '사업자',
      '발주처 담당자',
      '발주처 연락처',
      '현장 주소',
      '공사 내용',
      '공사 시작',
      '공사 종료',
      '총액 1',
      '총액 2',
      '기계 분류',
      '브랜드',
      '견적서 및 계약서 폴더 경로'
    ];

    // 하나라도 비어있으면 미완료
    return importantFields.some(field => {
      const value = item[field];
      return !value ||
             value.toString().trim() === '' ||
             value.toString().trim() === '-' ||
             value.toString().trim() === 'null' ||
             value.toString().trim() === 'undefined';
    });
  }

  /**
   * 컴포넌트 정리 (2026-07-08 추가)
   * 페이지 전환·컴포넌트 재초기화 시 pending timer 정리로 예상치 못한 필터 적용 방지.
   */
  destroy() {
    if (this.searchDebounceTimer) {
      clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = null;
    }
  }
}