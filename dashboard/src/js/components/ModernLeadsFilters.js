import logger from '../utils/logger.js';

/**
 * Modern Leads Filters 컴포넌트 (온라인 문의 현황 15열 구조)
 * 필터: 상태 / 플랫폼 / 영업 담당자 / 키워드 검색
 */
export default class ModernLeadsFilters {
  constructor() {
    this.filters = {};
    this.callbacks = [];
    this.currentData = null;
    this.isDataLoaded = false;
    this.resultCount = 0;
    this.searchDebounceTimer = null;

    // 검색 대상 필드 (옛 키 alias 포함: 거래처/담당자도 호환)
    this.searchableFields = [
      '고객명',
      '고객 연락처',
      '연락처',          // alias
      '이메일',
      '방문 주소',
      '문의 내용',        // 새 J열 (인입 원본)
      '상담 내용',        // 새 K열 (매니저 처리 결과, 옛 피드백)
      '키워드',
      '플랫폼',
      '거래처',          // alias
      '영업 담당자',
      '온라인 상담자',
      '담당자',          // alias
      '피드백',          // 옛 이름 (호환)
    ];
  }

  async init() {
    this.setupFilterElements();
    this.bindEvents();
    logger.debug('[ModernLeadsFilters] 필터 초기화 완료');
  }

  setupFilterElements() {
    this.elements = {
      statusFilter: document.getElementById('statusFilter'),
      platformFilter: document.getElementById('platformFilter'),
      salesOwnerFilter: document.getElementById('salesOwnerFilter'),
      searchInput: document.getElementById('searchInput'),
      resetFiltersBtn: document.getElementById('resetFiltersBtn'),
      filterResultCount: document.getElementById('filterResultCount'),
    };
  }

  bindEvents() {
    if (this.elements.statusFilter) {
      this.elements.statusFilter.addEventListener('change', () => {
        const value = this.elements.statusFilter.value;
        if (value) this.filters.status = value;
        else delete this.filters.status;
        this.applyFilters(null, true);
      });
    }

    if (this.elements.platformFilter) {
      this.elements.platformFilter.addEventListener('change', () => {
        const value = this.elements.platformFilter.value;
        if (value) this.filters.platform = value;
        else delete this.filters.platform;
        this.applyFilters(null, true);
      });
    }

    if (this.elements.salesOwnerFilter) {
      this.elements.salesOwnerFilter.addEventListener('change', () => {
        const value = this.elements.salesOwnerFilter.value;
        if (value) this.filters.salesOwner = value;
        else delete this.filters.salesOwner;
        this.applyFilters(null, true);
      });
    }

    if (this.elements.searchInput) {
      this.elements.searchInput.addEventListener('input', () => {
        const value = this.elements.searchInput.value.trim();
        if (value) this.filters.search = value;
        else delete this.filters.search;

        if (this.searchDebounceTimer) clearTimeout(this.searchDebounceTimer);
        this.searchDebounceTimer = setTimeout(() => {
          this.applyFilters(null, true);
        }, 300);
      });
    }

    if (this.elements.resetFiltersBtn) {
      this.elements.resetFiltersBtn.addEventListener('click', () => this.resetFilters());
    }
  }

  applyFilters(data = null, triggerCallbacks = true) {
    const startTime = performance.now();

    if (data && Array.isArray(data)) {
      this.currentData = data;
      this.isDataLoaded = true;
      this.populateAllFilters(data);
    }

    if (!this.isDataLoaded || this.currentData === null) {
      logger.warn('[ModernLeadsFilters] 데이터가 아직 로드되지 않음');
      this.updateResultCount(0);
      return null;
    }

    if (this.currentData.length === 0) {
      this.updateResultCount(0);
      if (triggerCallbacks) this.callbacks.forEach((cb) => cb([]));
      return [];
    }

    const originalCount = this.currentData.length;
    let filteredData = [...this.currentData];

    // 검색
    if (this.filters.search) {
      const searchTerm = this.filters.search.toLowerCase();
      filteredData = filteredData.filter((item) =>
        this.searchableFields.some((field) => {
          const value = item[field];
          return value && String(value).toLowerCase().includes(searchTerm);
        })
      );
    }

    // 상태
    if (this.filters.status) {
      filteredData = filteredData.filter(
        (item) => (item['상태'] || '').toString().trim() === this.filters.status
      );
    }

    // 플랫폼
    if (this.filters.platform) {
      filteredData = filteredData.filter((item) => {
        const v = (item['플랫폼'] ?? item['거래처'] ?? '').toString().trim();
        return v === this.filters.platform;
      });
    }

    // 영업 담당자
    if (this.filters.salesOwner) {
      filteredData = filteredData.filter((item) => {
        const v = (item['영업 담당자'] ?? item['담당자'] ?? '').toString().trim();
        return v === this.filters.salesOwner;
      });
    }

    const filteredCount = filteredData.length;
    this.resultCount = filteredCount;
    this.updateResultCount(filteredCount);

    const elapsed = (performance.now() - startTime).toFixed(2);
    logger.debug(`[ModernLeadsFilters] 필터 적용: ${originalCount} → ${filteredCount} (${elapsed}ms)`);

    if (triggerCallbacks) this.callbacks.forEach((cb) => cb(filteredData));

    return filteredData;
  }

  populateAllFilters(data) {
    // 상태 (고정 + 데이터에서 발견된 미정의 상태도 포함)
    const fixedStatuses = [
      '상담 대기', '유선 상담', '부재중',
      '방문 예약', '방문 대기', '방문 완료', '방문 취소',
      '견적 제출', '문의 드랍',
      '공사 확정', '공사 취소', '공사 드랍',
    ];
    const dataStatuses = [...new Set(data.map((l) => (l['상태'] || '').toString().trim()).filter(Boolean))];
    const statuses = [...new Set([...fixedStatuses, ...dataStatuses])];
    this.populateFilter(this.elements.statusFilter, statuses, '전체 상태');

    // 플랫폼 (동적)
    const platforms = [
      ...new Set(
        data
          .map((l) => (l['플랫폼'] ?? l['거래처'] ?? '').toString().trim())
          .filter(Boolean)
      ),
    ].sort();
    this.populateFilter(this.elements.platformFilter, platforms, '전체 플랫폼');

    // 영업 담당자 (동적)
    const owners = [
      ...new Set(
        data
          .map((l) => (l['영업 담당자'] ?? l['담당자'] ?? '').toString().trim())
          .filter(Boolean)
      ),
    ].sort();
    this.populateFilter(this.elements.salesOwnerFilter, owners, '전체 영업 담당자');
  }

  populateFilter(element, options, placeholderText) {
    if (!element) return;
    const currentValue = element.value;
    element.innerHTML = `<option value="">${placeholderText}</option>`;
    options.forEach((option) => {
      const opt = document.createElement('option');
      opt.value = option;
      opt.textContent = option;
      element.appendChild(opt);
    });
    if (currentValue && options.includes(currentValue)) {
      element.value = currentValue;
    }
  }

  updateResultCount(count) {
    if (this.elements.filterResultCount) {
      this.elements.filterResultCount.textContent = `${count}개 리드 표시`;
    }
  }

  onFilterChange(callback) {
    this.callbacks.push(callback);
  }

  resetFilters() {
    this.filters = {};
    if (this.elements.statusFilter) this.elements.statusFilter.value = '';
    if (this.elements.platformFilter) this.elements.platformFilter.value = '';
    if (this.elements.salesOwnerFilter) this.elements.salesOwnerFilter.value = '';
    if (this.elements.searchInput) this.elements.searchInput.value = '';

    if (this.isDataLoaded) this.applyFilters(null, true);
    logger.debug('[ModernLeadsFilters] 필터 초기화 완료');
  }

  getCurrentFilters() {
    return { ...this.filters };
  }

  isReady() {
    return this.isDataLoaded && this.currentData !== null;
  }
}
