import logger from '../utils/logger.js';

/**
 * MobileCardView 컴포넌트
 * 모바일 및 작은 화면을 위한 카드 기반 프로젝트 뷰
 * 반응형 디자인과 터치 친화적 인터페이스 제공
 */
export default class MobileCardView {
  constructor(options = {}) {
    // 설정
    this.config = {
      containerSelector: '#mobile-card-container',
      breakpoint: 768, // px 이하에서 모바일 뷰 활성화
      cardsPerRow: {
        xs: 1,   // < 576px
        sm: 2,   // 576px - 768px
        md: 3    // 768px - 992px
      },
      enableSwipe: true,
      enableInfiniteScroll: false,
      cardAnimation: 'slideIn',
      sortBy: '프로젝트 코드',
      sortOrder: 'desc',
      showFilters: true,
      enableSearch: true,
      ...options
    };

    // 상태 관리
    this.container = null;
    this.isActive = false;
    this.currentData = [];
    this.filteredData = [];
    this.currentSort = { field: this.config.sortBy, order: this.config.sortOrder };

    // 검색 및 필터
    this.searchTerm = '';
    this.activeFilters = new Set();

    // 컴포넌트
    this.badgeSystem = null;

    // 터치 이벤트
    this.touchStartX = 0;
    this.touchStartY = 0;
    this.swipeThreshold = 50;

    // 무한 스크롤
    this.currentPage = 1;
    this.itemsPerPage = 20;
    this.isLoading = false;

    this.init();
  }

  /**
   * 컴포넌트 초기화
   */
  async init() {
    this.createContainer();
    await this.initializeComponents();
    this.setupEventListeners();
    this.checkViewportAndToggle();
  }

  /**
   * 컨테이너 생성
   */
  createContainer() {
    // 기존 컨테이너 확인
    this.container = document.querySelector(this.config.containerSelector);

    if (!this.container) {
      // 새 컨테이너 생성
      this.container = document.createElement('div');
      this.container.id = 'mobile-card-container';
      this.container.className = 'mobile-card-view d-none';

      // 테이블 컨테이너 다음에 삽입
      const tableContainer = document.querySelector('.table-responsive, #projectsTable')?.parentNode;
      if (tableContainer) {
        tableContainer.parentNode.insertBefore(this.container, tableContainer.nextSibling);
      } else {
        document.body.appendChild(this.container);
      }
    }

    this.createLayout();
  }

  /**
   * 레이아웃 생성
   */
  createLayout() {
    this.container.innerHTML = `
      <div class="mobile-view-header">
        <div class="mobile-search-bar">
          <div class="input-group">
            <input type="text"
                   class="form-control mobile-search-input"
                   placeholder="프로젝트 검색..."
                   aria-label="프로젝트 검색">
            <button class="btn btn-outline-secondary mobile-search-btn" type="button">
              <i class="fas fa-search"></i>
            </button>
          </div>
        </div>

        <div class="mobile-controls">
          <div class="mobile-sort-controls">
            <select class="form-select form-select-sm mobile-sort-select">
              <option value="프로젝트 코드">프로젝트 코드</option>
              <option value="공사 시작">공사 시작일</option>
              <option value="공사상태">상태별</option>
              <option value="총액2">금액별</option>
              <option value="미수금">미수금별</option>
            </select>
            <button class="btn btn-sm btn-outline-secondary mobile-sort-order"
                    data-order="desc" title="정렬 순서">
              <i class="fas fa-sort-amount-down"></i>
            </button>
          </div>

          <div class="mobile-filter-controls">
            <button class="btn btn-sm btn-outline-primary mobile-filter-toggle"
                    data-bs-toggle="collapse"
                    data-bs-target="#mobile-quick-filters">
              <i class="fas fa-filter"></i>
              <span class="filter-count d-none">0</span>
            </button>
          </div>
        </div>
      </div>

      <div class="collapse" id="mobile-quick-filters">
        <div class="mobile-quick-filters p-3 bg-light">
          <div class="quick-filter-group">
            <h6 class="mb-2">빠른 필터</h6>
            <div class="btn-group-vertical w-100" role="group">
              <button type="button" class="btn btn-outline-secondary btn-sm quick-filter"
                      data-filter="completed">
                <i class="fas fa-check-circle me-2"></i>완료된 프로젝트
              </button>
              <button type="button" class="btn btn-outline-warning btn-sm quick-filter"
                      data-filter="outstanding">
                <i class="fas fa-exclamation-triangle me-2"></i>미수금 있는 프로젝트
              </button>
              <button type="button" class="btn btn-outline-primary btn-sm quick-filter"
                      data-filter="in-progress">
                <i class="fas fa-play-circle me-2"></i>진행 중인 프로젝트
              </button>
              <button type="button" class="btn btn-outline-danger btn-sm quick-filter"
                      data-filter="urgent">
                <i class="fas fa-exclamation-circle me-2"></i>수금필요 프로젝트
              </button>
            </div>
          </div>
        </div>
      </div>

      <div class="mobile-stats">
        <div class="row text-center">
          <div class="col-4">
            <div class="stat-item">
              <div class="stat-number" id="mobile-total-count">0</div>
              <div class="stat-label">총 프로젝트</div>
            </div>
          </div>
          <div class="col-4">
            <div class="stat-item">
              <div class="stat-number" id="mobile-filtered-count">0</div>
              <div class="stat-label">필터 결과</div>
            </div>
          </div>
          <div class="col-4">
            <div class="stat-item">
              <div class="stat-number" id="mobile-outstanding-sum">0원</div>
              <div class="stat-label">총 미수금</div>
            </div>
          </div>
        </div>
      </div>

      <div class="mobile-cards-container">
        <div class="row" id="mobile-cards-grid">
          <!-- 카드들이 여기에 동적으로 생성됩니다 -->
        </div>
      </div>

      <div class="mobile-load-more text-center" style="display: none;">
        <button class="btn btn-outline-primary load-more-btn">
          <i class="fas fa-chevron-down me-2"></i>더 보기
        </button>
      </div>

      <div class="mobile-loading text-center" style="display: none;">
        <div class="spinner-border text-primary" role="status">
          <span class="visually-hidden">로딩중...</span>
        </div>
      </div>
    `;
  }

  /**
   * 하위 컴포넌트 초기화
   */
  async initializeComponents() {
    try {
      // UnifiedBadgeSystem 컴포넌트 초기화 (모듈형 뱃지 시스템)
      const { default: UnifiedBadgeSystem } = await import('./UnifiedBadgeSystem.js');
      this.badgeSystem = new UnifiedBadgeSystem();

      // 하위 컴포넌트 초기화 완료
    } catch (error) {
      logger.error('[ERROR] MobileCardView 컴포넌트 초기화 실패:', error);
    }
  }

  /**
   * 이벤트 리스너 설정
   */
  setupEventListeners() {
    // projectUpdated 이벤트 구독 (필드 메모 업데이트 등)
    document.addEventListener('projectUpdated', (e) => {
      const { projectCode, memoKey, memo, updateType } = e.detail || {};

      if (updateType === 'fieldMemo' && projectCode && memoKey) {
        // currentData에서 해당 프로젝트 찾아서 메모 업데이트
        const targetProject = this.currentData.find(p => p['프로젝트 코드'] === projectCode);

        if (targetProject) {
          targetProject[memoKey] = memo;
          logger.debug(`[MobileCardView] 메모 동기화 완료: ${projectCode} / ${memoKey}`);

          // 활성화된 상태라면 카드 다시 렌더링
          if (this.isActive) {
            this.applyFiltersAndSort();
            this.renderCards();
          }
        }
      }
    });

    // 뷰포트 크기 변경 감지
    window.addEventListener('resize', this.debounce(() => {
      this.checkViewportAndToggle();
    }, 250));

    // 컨테이너 내 이벤트 위임
    if (this.container) {
      // 검색
      this.container.addEventListener('input', (e) => {
        if (e.target.matches('.mobile-search-input')) {
          this.handleSearch(e.target.value);
        }
      });

      // 정렬
      this.container.addEventListener('change', (e) => {
        if (e.target.matches('.mobile-sort-select')) {
          this.handleSort(e.target.value);
        }
      });

      this.container.addEventListener('click', (e) => {
        if (e.target.matches('.mobile-sort-order') || e.target.closest('.mobile-sort-order')) {
          this.toggleSortOrder(e.target.closest('.mobile-sort-order'));
        }
      });

      // 빠른 필터
      this.container.addEventListener('click', (e) => {
        if (e.target.matches('.quick-filter') || e.target.closest('.quick-filter')) {
          this.handleQuickFilter(e.target.closest('.quick-filter'));
        }
      });

      // 카드 클릭
      this.container.addEventListener('click', (e) => {
        const card = e.target.closest('.project-card');
        if (card && !e.target.closest('.card-actions')) {
          this.handleCardClick(card);
        }
      });

      // 터치 이벤트 (스와이프)
      if (this.config.enableSwipe) {
        this.container.addEventListener('touchstart', this.handleTouchStart.bind(this));
        this.container.addEventListener('touchmove', this.handleTouchMove.bind(this));
        this.container.addEventListener('touchend', this.handleTouchEnd.bind(this));
      }

      // 더 보기 버튼
      this.container.addEventListener('click', (e) => {
        if (e.target.matches('.load-more-btn')) {
          this.loadMoreCards();
        }
      });
    }
  }

  /**
   * 뷰포트 크기 확인 및 뷰 전환
   */
  checkViewportAndToggle() {
    const isMobileSize = window.innerWidth <= this.config.breakpoint;

    if (isMobileSize && !this.isActive) {
      this.activate();
    } else if (!isMobileSize && this.isActive) {
      this.deactivate();
    }
  }

  /**
   * 모바일 뷰 활성화
   */
  activate() {
    this.isActive = true;

    // 테이블 숨기기
    const tableContainer = document.querySelector('.table-responsive, .dataTables_wrapper');
    if (tableContainer) {
      tableContainer.style.display = 'none';
    }

    // 모바일 뷰 표시
    this.container.classList.remove('d-none');
    this.container.classList.add('d-block');

    // 현재 데이터로 카드 렌더링
    if (this.currentData.length > 0) {
      this.updateCards(this.currentData);
    }

    logger.debug('📱 모바일 카드 뷰 활성화');
  }

  /**
   * 모바일 뷰 비활성화
   */
  deactivate() {
    this.isActive = false;

    // 모바일 뷰 숨기기
    this.container.classList.remove('d-block');
    this.container.classList.add('d-none');

    // 테이블 표시
    const tableContainer = document.querySelector('.table-responsive, .dataTables_wrapper');
    if (tableContainer) {
      tableContainer.style.display = '';
    }

    logger.debug('🖥️ 데스크톱 테이블 뷰 활성화');
  }

  /**
   * 카드 데이터 업데이트
   */
  updateCards(data) {
    this.currentData = Array.isArray(data) ? data : [];
    this.applyFiltersAndSort();
    this.renderCards();
    this.updateStats();
  }

  /**
   * 필터 및 정렬 적용
   */
  applyFiltersAndSort() {
    let filtered = [...this.currentData];

    // 검색어 적용
    if (this.searchTerm) {
      filtered = filtered.filter(item => {
        const searchFields = ['프로젝트 코드', '담당자', '유입 구분', '현장 주소', '공사 내용'];
        return searchFields.some(field =>
          String(item[field] || '').toLowerCase().includes(this.searchTerm.toLowerCase())
        );
      });
    }

    // 빠른 필터 적용
    this.activeFilters.forEach(filter => {
      switch (filter) {
        case 'completed':
          filtered = filtered.filter(item => item['공사상태'] === '공사완료');
          break;
        case 'outstanding':
          filtered = filtered.filter(item => {
            const outstanding = parseFloat(item['미수금'] || item['미수금W'] || 0);
            return outstanding > 0;
          });
          break;
        case 'in-progress':
          filtered = filtered.filter(item => item['공사상태'] === '공사진행');
          break;
        case 'urgent':
          filtered = filtered.filter(item => item['공사상태'] === '수금필요');
          break;
      }
    });

    // 정렬 적용
    filtered.sort((a, b) => {
      let aVal = a[this.currentSort.field] || '';
      let bVal = b[this.currentSort.field] || '';

      // 프로젝트 코드의 경우 숫자만 추출하여 비교
      if (this.currentSort.field === '프로젝트 코드') {
        const aMatch = aVal.match(/[RG](\d+)/);
        const bMatch = bVal.match(/[RG](\d+)/);
        aVal = aMatch ? parseInt(aMatch[1], 10) : 0;
        bVal = bMatch ? parseInt(bMatch[1], 10) : 0;
      }

      let comparison = 0;
      if (aVal < bVal) comparison = -1;
      if (aVal > bVal) comparison = 1;

      return this.currentSort.order === 'desc' ? -comparison : comparison;
    });

    this.filteredData = filtered;
  }

  /**
   * 카드 렌더링
   */
  renderCards() {
    const grid = this.container.querySelector('#mobile-cards-grid');
    if (!grid) return;

    // 기존 카드 제거
    grid.innerHTML = '';

    // 페이지네이션 적용
    const startIndex = 0;
    const endIndex = this.currentPage * this.itemsPerPage;
    const cardsToShow = this.filteredData.slice(startIndex, endIndex);

    cardsToShow.forEach((item, index) => {
      const card = this.createCard(item, index);
      grid.appendChild(card);
    });

    // 더 보기 버튼 표시/숨기기
    const loadMoreBtn = this.container.querySelector('.mobile-load-more');
    if (loadMoreBtn) {
      const hasMore = this.filteredData.length > endIndex;
      loadMoreBtn.style.display = hasMore ? 'block' : 'none';
    }

    // 애니메이션 적용
    this.animateCards();
  }

  /**
   * 개별 카드 생성
   */
  createCard(data, index) {
    const projectCode = data['프로젝트 코드'] || '';
    const status = data['공사상태'] || '';
    const totalAmount = parseFloat(data['총액 2'] || data['총액2'] || data['S'] || 0);
    const outstanding = parseFloat(data['미수금'] || data['미수금W'] || 0);

    // 컬럼 클래스 결정
    const colClass = this.getColumnClass();

    const cardDiv = document.createElement('div');
    cardDiv.className = `col-${colClass} mb-3`;
    cardDiv.innerHTML = `
      <div class="card project-card h-100"
           data-project-code="${projectCode}"
           data-index="${index}">
        <div class="card-header d-flex justify-content-between align-items-center">
          <h6 class="card-title mb-0">
            <i class="fas fa-folder-open me-2 text-primary"></i>
            ${projectCode}
          </h6>
          <div class="card-status">
            ${this.renderStatusBadge(status)}
          </div>
        </div>

        <div class="card-body">
          <div class="project-info">
            <div class="info-row">
              <span class="info-label">
                <i class="fas fa-user me-1"></i>담당자
              </span>
              <span class="info-value">${this.badgeSystem ? this.badgeSystem.createManagerBadge(data['담당자']) : (data['담당자'] || '-')}</span>
            </div>

            <div class="info-row">
              <span class="info-label">
                <i class="fas fa-building me-1"></i>유입 구분
              </span>
              <span class="info-value">${this.badgeSystem ? this.badgeSystem.createCompanyBadge(data['유입 구분']) : (data['유입 구분'] || '-')}</span>
            </div>

            <div class="info-row">
              <span class="info-label">
                <i class="fas fa-map-marker-alt me-1"></i>현장
              </span>
              <span class="info-value text-truncate" title="${data['현장 주소'] || '-'}">
                ${data['현장 주소'] || '-'}
              </span>
            </div>

            <div class="info-row">
              <span class="info-label">
                <i class="fas fa-tools me-1"></i>내용
              </span>
              <span class="info-value text-truncate" title="${data['공사 내용'] || '-'}">
                ${data['공사 내용'] || '-'}
              </span>
            </div>

            <div class="info-row">
              <span class="info-label">
                <i class="fas fa-calendar me-1"></i>기간
              </span>
              <span class="info-value">
                ${this.formatDate(data['공사 시작'])} ~ ${this.formatDate(data['공사 종료'])}
              </span>
            </div>
          </div>

          <div class="amount-info mt-3">
            <div class="row">
              <div class="col-6">
                <div class="amount-item">
                  <small class="text-muted">총액</small>
                  <div class="amount-value text-primary">
                    ${totalAmount > 0 ? totalAmount.toLocaleString() + '원' : '-'}
                  </div>
                </div>
              </div>
              <div class="col-6">
                <div class="amount-item">
                  <small class="text-muted">미수금</small>
                  <div class="amount-value ${outstanding > 0 ? 'text-warning' : 'text-success'}">
                    ${outstanding > 0 ? outstanding.toLocaleString() + '원' : '0원'}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div class="card-footer">
          <div class="card-actions">
            ${this.renderActionButtons(data, projectCode)}
          </div>
        </div>
      </div>
    `;

    return cardDiv;
  }

  /**
   * 상태 뱃지 렌더링
   */
  renderStatusBadge(status) {
    if (this.badgeSystem && status) {
      return this.badgeSystem.createStatusBadge(status);
    }
    return `<span class="badge status-waiting">미설정</span>`;
  }

  /**
   * 액션 버튼 렌더링
   */
  renderActionButtons(data, projectCode) {
    // 기본 버튼들만 표시 (ActionButtons 제거됨)
    return `
      <div class="btn-group btn-group-sm w-100" role="group">
        <button type="button" class="btn btn-outline-primary btn-edit flex-fill"
                data-project-code="${projectCode}">
          <i class="fas fa-edit"></i>
        </button>
        <button type="button" class="btn btn-outline-info btn-view flex-fill"
                data-project-code="${projectCode}">
          <i class="fas fa-eye"></i>
        </button>
        <button type="button" class="btn btn-outline-secondary btn-copy flex-fill"
                data-project-code="${projectCode}">
          <i class="fas fa-copy"></i>
        </button>
      </div>
    `;
  }

  /**
   * 컬럼 클래스 결정 (반응형)
   */
  getColumnClass() {
    const width = window.innerWidth;
    if (width < 576) return this.config.cardsPerRow.xs === 1 ? '12' : '6';
    if (width < 768) return this.config.cardsPerRow.sm === 1 ? '12' : '6';
    return this.config.cardsPerRow.md === 1 ? '12' : this.config.cardsPerRow.md === 2 ? '6' : '4';
  }

  /**
   * 날짜 포맷팅
   */
  formatDate(dateStr) {
    if (!dateStr) return '-';

    try {
      if (typeof dateStr === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
        return dateStr;
      }

      const date = new Date(dateStr);
      if (isNaN(date.getTime())) return '-';

      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    } catch {
      return '-';
    }
  }

  /**
   * 카드 애니메이션
   */
  animateCards() {
    if (this.config.cardAnimation === 'slideIn') {
      const cards = this.container.querySelectorAll('.project-card');
      cards.forEach((card, index) => {
        card.style.opacity = '0';
        card.style.transform = 'translateY(20px)';

        setTimeout(() => {
          card.style.transition = 'all 0.3s ease-out';
          card.style.opacity = '1';
          card.style.transform = 'translateY(0)';
        }, index * 50);
      });
    }
  }

  /**
   * 통계 업데이트
   */
  updateStats() {
    const totalCount = this.currentData.length;
    const filteredCount = this.filteredData.length;
    const totalOutstanding = this.filteredData.reduce((sum, item) => {
      return sum + (parseFloat(item['미수금'] || item['미수금W'] || 0) || 0);
    }, 0);

    // DOM 업데이트
    const totalEl = this.container.querySelector('#mobile-total-count');
    const filteredEl = this.container.querySelector('#mobile-filtered-count');
    const outstandingEl = this.container.querySelector('#mobile-outstanding-sum');

    if (totalEl) totalEl.textContent = totalCount.toLocaleString();
    if (filteredEl) filteredEl.textContent = filteredCount.toLocaleString();
    if (outstandingEl) outstandingEl.textContent = totalOutstanding.toLocaleString() + '원';

    // 필터 카운트 업데이트
    const filterCountEl = this.container.querySelector('.filter-count');
    if (filterCountEl) {
      if (this.activeFilters.size > 0) {
        filterCountEl.textContent = this.activeFilters.size;
        filterCountEl.classList.remove('d-none');
      } else {
        filterCountEl.classList.add('d-none');
      }
    }
  }

  /**
   * 검색 처리
   */
  handleSearch(searchTerm) {
    this.searchTerm = searchTerm.trim();
    this.currentPage = 1;
    this.applyFiltersAndSort();
    this.renderCards();
    this.updateStats();
  }

  /**
   * 정렬 처리
   */
  handleSort(field) {
    this.currentSort.field = field;
    this.applyFiltersAndSort();
    this.renderCards();
  }

  /**
   * 정렬 순서 토글
   */
  toggleSortOrder(button) {
    const currentOrder = button.dataset.order;
    const newOrder = currentOrder === 'desc' ? 'asc' : 'desc';

    button.dataset.order = newOrder;
    this.currentSort.order = newOrder;

    // 아이콘 업데이트
    const icon = button.querySelector('i');
    if (icon) {
      icon.className = newOrder === 'desc' ? 'fas fa-sort-amount-down' : 'fas fa-sort-amount-up';
    }

    this.applyFiltersAndSort();
    this.renderCards();
  }

  /**
   * 빠른 필터 처리
   */
  handleQuickFilter(button) {
    const filter = button.dataset.filter;

    if (this.activeFilters.has(filter)) {
      this.activeFilters.delete(filter);
      button.classList.remove('active');
    } else {
      this.activeFilters.add(filter);
      button.classList.add('active');
    }

    this.currentPage = 1;
    this.applyFiltersAndSort();
    this.renderCards();
    this.updateStats();
  }

  /**
   * 카드 클릭 처리
   */
  handleCardClick(card) {
    const projectCode = card.dataset.projectCode;

    // 테이블의 아코디언 열기와 동일한 효과
    if (window.projectsTable && window.projectsTable.accordion) {
      const projectData = this.currentData.find(item => item['프로젝트 코드'] === projectCode);
      if (projectData) {
        // 데스크톱 뷰로 전환하고 아코디언 열기
        this.deactivate();
        setTimeout(() => {
          window.projectsTable.accordion.openAccordion(null, projectData);
        }, 100);
      }
    }
  }

  /**
   * 터치 이벤트 처리
   */
  handleTouchStart(e) {
    this.touchStartX = e.touches[0].clientX;
    this.touchStartY = e.touches[0].clientY;
  }

  handleTouchMove(e) {
    // 스크롤 방지는 하지 않음 (필요시 구현)
  }

  handleTouchEnd(e) {
    if (!this.touchStartX || !this.touchStartY) return;

    const touchEndX = e.changedTouches[0].clientX;
    const touchEndY = e.changedTouches[0].clientY;

    const deltaX = touchEndX - this.touchStartX;
    const deltaY = touchEndY - this.touchStartY;

    // 스와이프 제스처 감지 (좌우)
    if (Math.abs(deltaX) > Math.abs(deltaY) && Math.abs(deltaX) > this.swipeThreshold) {
      if (deltaX > 0) {
        // 오른쪽 스와이프 - 이전
        this.handleSwipeRight();
      } else {
        // 왼쪽 스와이프 - 다음
        this.handleSwipeLeft();
      }
    }

    this.touchStartX = 0;
    this.touchStartY = 0;
  }

  handleSwipeRight() {
    // NOTE: 미구현 - 이전 페이지/필터 변경 기능 (추후 구현 예정)
    logger.debug('오른쪽 스와이프');
  }

  handleSwipeLeft() {
    // NOTE: 미구현 - 다음 페이지/필터 변경 기능 (추후 구현 예정)
    logger.debug('왼쪽 스와이프');
  }

  /**
   * 더 보기 카드 로드
   */
  loadMoreCards() {
    if (this.isLoading) return;

    this.isLoading = true;
    this.currentPage++;

    // 로딩 표시
    const loadingEl = this.container.querySelector('.mobile-loading');
    if (loadingEl) loadingEl.style.display = 'block';

    // 시뮬레이션된 지연 (실제로는 즉시 렌더링)
    setTimeout(() => {
      this.renderCards();
      this.isLoading = false;

      if (loadingEl) loadingEl.style.display = 'none';
    }, 300);
  }

  /**
   * 디바운스 유틸리티
   */
  debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func(...args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }

  /**
   * 컴포넌트 정리
   */
  destroy() {
    // 이벤트 리스너 제거
    window.removeEventListener('resize', this.checkViewportAndToggle);

    // 하위 컴포넌트 정리
    this.badgeSystem = null;

    // 컨테이너 제거
    if (this.container) {
      this.container.remove();
      this.container = null;
    }

    // 데스크톱 뷰 복원
    const tableContainer = document.querySelector('.table-responsive, .dataTables_wrapper');
    if (tableContainer) {
      tableContainer.style.display = '';
    }
  }
}