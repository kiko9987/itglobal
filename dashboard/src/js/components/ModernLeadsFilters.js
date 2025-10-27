import logger from '../utils/logger.js';

/**
 * Modern Leads Filters 컴포넌트
 * 리드 관리 페이지 필터링 시스템 (프로젝트 관리 패턴 준수)
 */
export default class ModernLeadsFilters {
  constructor() {
    this.filters = {};
    this.callbacks = [];
    this.currentData = null;
    this.isDataLoaded = false;
    this.resultCount = 0;
    this.searchDebounceTimer = null;

    // 검색 대상 필드 정의
    this.searchableFields = [
      '고객명',
      '연락처',
      '방문 주소',
      '상담 내용',
      '거래처',
      '담당자'
    ];
  }

  /**
   * 필터 초기화
   */
  async init() {
    this.setupFilterElements();
    this.bindEvents();
    logger.debug('[ModernLeadsFilters] 필터 초기화 완료');
  }

  /**
   * 필터 요소 설정
   */
  setupFilterElements() {
    this.elements = {
      statusFilter: document.getElementById('statusFilter'),
      companyFilter: document.getElementById('companyFilter'),
      ownerFilter: document.getElementById('ownerFilter'),
      searchInput: document.getElementById('searchInput'),
      resetFiltersBtn: document.getElementById('resetFiltersBtn'),
      filterResultCount: document.getElementById('filterResultCount')
    };
  }

  /**
   * 이벤트 바인딩
   */
  bindEvents() {
    // 상태 필터
    if (this.elements.statusFilter) {
      this.elements.statusFilter.addEventListener('change', () => {
        const value = this.elements.statusFilter.value;
        if (value && value !== '') {
          this.filters.status = value;
        } else {
          delete this.filters.status;
        }
        this.applyFilters(null, true);
      });
    }

    // 거래처 필터
    if (this.elements.companyFilter) {
      this.elements.companyFilter.addEventListener('change', () => {
        const value = this.elements.companyFilter.value;
        if (value && value !== '') {
          this.filters.company = value;
        } else {
          delete this.filters.company;
        }
        this.applyFilters(null, true);
      });
    }

    // 담당자 필터
    if (this.elements.ownerFilter) {
      this.elements.ownerFilter.addEventListener('change', () => {
        const value = this.elements.ownerFilter.value;
        if (value && value !== '') {
          this.filters.owner = value;
        } else {
          delete this.filters.owner;
        }
        this.applyFilters(null, true);
      });
    }

    // 검색어 필터 (디바운싱)
    if (this.elements.searchInput) {
      this.elements.searchInput.addEventListener('input', () => {
        const value = this.elements.searchInput.value.trim();
        if (value) {
          this.filters.search = value;
        } else {
          delete this.filters.search;
        }

        // 디바운싱: 300ms 후에 필터 적용
        if (this.searchDebounceTimer) {
          clearTimeout(this.searchDebounceTimer);
        }
        this.searchDebounceTimer = setTimeout(() => {
          this.applyFilters(null, true);
        }, 300);
      });
    }

    // 필터 초기화 버튼
    if (this.elements.resetFiltersBtn) {
      this.elements.resetFiltersBtn.addEventListener('click', () => {
        this.resetFilters();
      });
    }
  }

  /**
   * 필터 적용 및 콜백 실행
   * @param {Array} data - 새로운 데이터 (옵션널)
   * @param {boolean} triggerCallbacks - 콜백 실행 여부
   */
  applyFilters(data = null, triggerCallbacks = true) {
    const startTime = performance.now();

    // 새 데이터가 전달되면 현재 데이터로 저장하고 동적 옵션 생성
    if (data && Array.isArray(data)) {
      this.currentData = data;
      this.isDataLoaded = true;
      this.populateAllFilters(data);
    }

    // 데이터가 아직 로드되지 않은 경우
    if (!this.isDataLoaded || this.currentData === null) {
      logger.warn('[ModernLeadsFilters] 데이터가 아직 로드되지 않음');
      this.updateResultCount(0);
      return null;
    }

    // 데이터는 로드되었지만 빈 배열인 경우
    if (this.currentData.length === 0) {
      logger.warn('[ModernLeadsFilters] 빈 데이터 배열');
      this.updateResultCount(0);
      if (triggerCallbacks) {
        this.callbacks.forEach(callback => callback([]));
      }
      return [];
    }

    const originalCount = this.currentData.length;
    let filteredData = [...this.currentData];

    // 검색어 필터
    if (this.filters.search) {
      const searchTerm = this.filters.search.toLowerCase();
      filteredData = filteredData.filter(item => {
        return this.searchableFields.some(field => {
          const value = item[field];
          return value && String(value).toLowerCase().includes(searchTerm);
        });
      });
    }

    // 상태 필터
    if (this.filters.status) {
      filteredData = filteredData.filter(item => {
        const status = item['상태'] || '';
        return status.toString().trim() === this.filters.status;
      });
    }

    // 거래처 필터
    if (this.filters.company) {
      filteredData = filteredData.filter(item => {
        const company = item['거래처'] || '';
        return company.toString().trim() === this.filters.company;
      });
    }

    // 담당자 필터
    if (this.filters.owner) {
      filteredData = filteredData.filter(item => {
        const owner = item['담당자'] || '';
        return owner.toString().trim() === this.filters.owner;
      });
    }

    const filteredCount = filteredData.length;
    this.resultCount = filteredCount;
    this.updateResultCount(filteredCount);

    const elapsed = (performance.now() - startTime).toFixed(2);
    logger.debug(`[ModernLeadsFilters] 필터 적용 완료: ${originalCount}개 → ${filteredCount}개 (${elapsed}ms)`);

    // 콜백 실행 (테이블 업데이트)
    if (triggerCallbacks) {
      this.callbacks.forEach(callback => callback(filteredData));
    }

    return filteredData;
  }

  /**
   * 필터 옵션 동적 생성
   */
  populateAllFilters(data) {
    // 상태 목록 (고정값)
    const statuses = ['유선 상담', '방문 예약', '견적 제출', '방문 취소', '공사 확정', '공사 드랍'];
    this.populateFilter(this.elements.statusFilter, statuses, '전체 상태');

    // 거래처 목록 (동적)
    const companies = [...new Set(data.map(lead => lead['거래처']).filter(Boolean))].sort();
    this.populateFilter(this.elements.companyFilter, companies, '전체 거래처');

    // 담당자 목록 (동적)
    const owners = [...new Set(data.map(lead => lead['담당자']).filter(Boolean))].sort();
    this.populateFilter(this.elements.ownerFilter, owners, '전체 담당자');
  }

  /**
   * 개별 필터 옵션 생성
   */
  populateFilter(element, options, placeholderText) {
    if (!element) return;

    const currentValue = element.value;
    element.innerHTML = `<option value="">${placeholderText}</option>`;

    options.forEach(option => {
      const opt = document.createElement('option');
      opt.value = option;
      opt.textContent = option;
      element.appendChild(opt);
    });

    // 이전 선택값 복원
    if (currentValue && options.includes(currentValue)) {
      element.value = currentValue;
    }
  }

  /**
   * 결과 카운트 업데이트
   */
  updateResultCount(count) {
    if (this.elements.filterResultCount) {
      this.elements.filterResultCount.textContent = `${count}개 리드 표시`;
    }
  }

  /**
   * 콜백 등록 (테이블 업데이트용)
   */
  onFilterChange(callback) {
    this.callbacks.push(callback);
  }

  /**
   * 필터 초기화
   */
  resetFilters() {
    // 필터 객체 초기화
    this.filters = {};

    // UI 초기화
    if (this.elements.statusFilter) this.elements.statusFilter.value = '';
    if (this.elements.companyFilter) this.elements.companyFilter.value = '';
    if (this.elements.ownerFilter) this.elements.ownerFilter.value = '';
    if (this.elements.searchInput) this.elements.searchInput.value = '';

    // 필터 재적용
    if (this.isDataLoaded) {
      this.applyFilters(null, true);
    }

    logger.debug('[ModernLeadsFilters] 필터 초기화 완료');
  }

  /**
   * 현재 필터 상태 반환
   */
  getCurrentFilters() {
    return { ...this.filters };
  }

  /**
   * 데이터 상태 확인
   */
  isReady() {
    return this.isDataLoaded && this.currentData !== null;
  }
}
