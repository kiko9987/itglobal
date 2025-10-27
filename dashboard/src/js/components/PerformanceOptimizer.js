/**
 * PerformanceOptimizer 컴포넌트
 * 대용량 데이터 처리를 위한 가상화, 메모이제이션, 성능 최적화 시스템
 * Virtual Scrolling, Memoization, Lazy Loading 등의 고급 최적화 기법 제공
 */
import logger from '../utils/logger.js';

export default class PerformanceOptimizer {
  constructor(options = {}) {
    // 설정
    this.config = {
      enableVirtualScrolling: true,
      enableMemoization: true,
      enableLazyLoading: true,
      virtualScrollItemHeight: 60, // px
      virtualScrollContainerHeight: 600, // px
      virtualScrollBuffer: 5, // 보이는 영역 위아래 추가 렌더링할 아이템 수
      memoizationCacheSize: 100,
      debounceDelay: 300, // ms
      batchSize: 50, // 한 번에 처리할 데이터 청크 크기
      performanceMonitoring: true,
      ...options
    };

    // 가상 스크롤링 관련
    this.virtualScrollContainer = null;
    this.virtualScrollData = [];
    this.visibleStartIndex = 0;
    this.visibleEndIndex = 0;
    this.scrollTop = 0;
    this.totalHeight = 0;

    // 메모이제이션 캐시
    this.renderCache = new Map();
    this.computationCache = new Map();
    this.filterCache = new Map();

    // 성능 모니터링
    this.performanceMetrics = {
      renderTime: [],
      filterTime: [],
      scrollTime: [],
      memoryUsage: []
    };

    // 디바운스된 함수들
    this.debouncedFunctions = new Map();

    // 교차점 관찰자 (Intersection Observer)
    this.intersectionObserver = null;
    this.lazyElements = new Set();

    // 웹 워커 (대용량 계산용)
    this.webWorker = null;

    this.init();
  }

  /**
   * 컴포넌트 초기화
   */
  init() {
    this.setupIntersectionObserver();
    this.setupPerformanceMonitoring();
    this.initializeWebWorker();
  }

  /**
   * 가상 스크롤링 설정
   */
  setupVirtualScrolling(container, data, renderItemCallback) {
    if (!this.config.enableVirtualScrolling) {
      return this.setupRegularScrolling(container, data, renderItemCallback);
    }

    this.virtualScrollContainer = container;
    this.virtualScrollData = data;
    this.renderItemCallback = renderItemCallback;

    // 컨테이너 스타일 설정
    container.style.height = `${this.config.virtualScrollContainerHeight}px`;
    container.style.overflow = 'auto';
    container.style.position = 'relative';

    // 총 높이 계산
    this.totalHeight = data.length * this.config.virtualScrollItemHeight;

    // 가상 스크롤 래퍼 생성
    this.createVirtualScrollWrapper(container);

    // 스크롤 이벤트 바인딩
    this.bindVirtualScrollEvents(container);

    // 초기 렌더링
    this.updateVirtualScrollView();

    logger.debug(`[ROCKET] 가상 스크롤링 활성화: ${data.length}개 아이템`);
  }

  /**
   * 가상 스크롤 래퍼 생성
   */
  createVirtualScrollWrapper(container) {
    // 기존 내용 정리
    container.innerHTML = '';

    // 스크롤러블 영역 생성
    const scrollableArea = document.createElement('div');
    scrollableArea.style.height = `${this.totalHeight}px`;
    scrollableArea.style.position = 'relative';
    scrollableArea.className = 'virtual-scroll-area';

    // 가시 영역 컨테이너 생성
    const visibleContainer = document.createElement('div');
    visibleContainer.style.position = 'absolute';
    visibleContainer.style.top = '0';
    visibleContainer.style.left = '0';
    visibleContainer.style.right = '0';
    visibleContainer.className = 'virtual-scroll-visible';

    scrollableArea.appendChild(visibleContainer);
    container.appendChild(scrollableArea);

    this.visibleContainer = visibleContainer;
  }

  /**
   * 가상 스크롤 이벤트 바인딩
   */
  bindVirtualScrollEvents(container) {
    const debouncedScrollHandler = this.debounce((e) => {
      this.handleVirtualScroll(e);
    }, 16); // 60fps

    container.addEventListener('scroll', debouncedScrollHandler);

    // 리사이즈 이벤트도 처리
    window.addEventListener('resize', this.debounce(() => {
      this.updateVirtualScrollView();
    }, 250));
  }

  /**
   * 가상 스크롤 처리
   */
  handleVirtualScroll(e) {
    const startTime = performance.now();

    this.scrollTop = e.target.scrollTop;
    this.updateVirtualScrollView();

    // 성능 측정
    if (this.config.performanceMonitoring) {
      const duration = performance.now() - startTime;
      this.performanceMetrics.scrollTime.push(duration);
      this.limitMetricsArray(this.performanceMetrics.scrollTime);
    }
  }

  /**
   * 가상 스크롤 뷰 업데이트
   */
  updateVirtualScrollView() {
    const containerHeight = this.virtualScrollContainer.clientHeight;
    const itemHeight = this.config.virtualScrollItemHeight;
    const buffer = this.config.virtualScrollBuffer;

    // 보이는 영역 계산
    const startIndex = Math.max(0, Math.floor(this.scrollTop / itemHeight) - buffer);
    const endIndex = Math.min(
      this.virtualScrollData.length - 1,
      Math.ceil((this.scrollTop + containerHeight) / itemHeight) + buffer
    );

    // 변경이 없으면 렌더링 스킵
    if (startIndex === this.visibleStartIndex && endIndex === this.visibleEndIndex) {
      return;
    }

    this.visibleStartIndex = startIndex;
    this.visibleEndIndex = endIndex;

    // 가시 영역 위치 조정
    this.visibleContainer.style.transform = `translateY(${startIndex * itemHeight}px)`;

    // 아이템 렌더링
    this.renderVirtualItems(startIndex, endIndex);
  }

  /**
   * 가상 아이템 렌더링
   */
  renderVirtualItems(startIndex, endIndex) {
    const startTime = performance.now();

    // 기존 아이템 정리
    this.visibleContainer.innerHTML = '';

    const fragment = document.createDocumentFragment();

    for (let i = startIndex; i <= endIndex; i++) {
      const item = this.virtualScrollData[i];
      if (!item) continue;

      // 메모이제이션 확인
      const cacheKey = this.generateCacheKey('render', item, i);
      let element = this.renderCache.get(cacheKey);

      if (!element) {
        // 새로 렌더링
        element = this.renderItemCallback(item, i);
        element.style.height = `${this.config.virtualScrollItemHeight}px`;
        element.style.overflow = 'hidden';

        // 캐시에 저장
        if (this.renderCache.size < this.config.memoizationCacheSize) {
          this.renderCache.set(cacheKey, element.cloneNode(true));
        }
      } else {
        // 캐시된 요소 복제
        element = element.cloneNode(true);
      }

      fragment.appendChild(element);
    }

    this.visibleContainer.appendChild(fragment);

    // 성능 측정
    if (this.config.performanceMonitoring) {
      const duration = performance.now() - startTime;
      this.performanceMetrics.renderTime.push(duration);
      this.limitMetricsArray(this.performanceMetrics.renderTime);
    }
  }

  /**
   * 일반 스크롤링 (폴백)
   */
  setupRegularScrolling(container, data, renderItemCallback) {
    logger.debug('[REFRESH] 일반 스크롤링 모드');

    const fragment = document.createDocumentFragment();

    data.forEach((item, index) => {
      const element = renderItemCallback(item, index);
      fragment.appendChild(element);
    });

    container.appendChild(fragment);
  }

  /**
   * 메모이제이션된 함수 생성
   */
  memoize(fn, keyGenerator) {
    if (!this.config.enableMemoization) {
      return fn;
    }

    return (...args) => {
      const key = keyGenerator ? keyGenerator(...args) : JSON.stringify(args);

      if (this.computationCache.has(key)) {
        return this.computationCache.get(key);
      }

      const result = fn(...args);

      if (this.computationCache.size < this.config.memoizationCacheSize) {
        this.computationCache.set(key, result);
      }

      return result;
    };
  }

  /**
   * 필터 최적화
   */
  optimizeFiltering(data, filterFunctions) {
    const startTime = performance.now();

    // 필터 캐시 키 생성
    const filterKey = this.generateFilterCacheKey(data, filterFunctions);

    if (this.filterCache.has(filterKey)) {
      logger.debug('[INFO] 필터 캐시 히트');
      return this.filterCache.get(filterKey);
    }

    // 배치 처리로 필터링
    const result = this.batchProcess(data, (batch) => {
      return batch.filter(item => {
        return filterFunctions.every(filterFn => filterFn(item));
      });
    }).flat();

    // 캐시에 저장
    if (this.filterCache.size < this.config.memoizationCacheSize) {
      this.filterCache.set(filterKey, result);
    }

    // 성능 측정
    if (this.config.performanceMonitoring) {
      const duration = performance.now() - startTime;
      this.performanceMetrics.filterTime.push(duration);
      this.limitMetricsArray(this.performanceMetrics.filterTime);
      logger.debug(`⚡ 필터링 완료: ${duration.toFixed(2)}ms`);
    }

    return result;
  }

  /**
   * 배치 처리
   */
  batchProcess(data, processFn) {
    const results = [];
    const batchSize = this.config.batchSize;

    for (let i = 0; i < data.length; i += batchSize) {
      const batch = data.slice(i, i + batchSize);
      results.push(processFn(batch));
    }

    return results;
  }

  /**
   * 레이지 로딩 설정
   */
  setupLazyLoading(elements) {
    if (!this.config.enableLazyLoading) return;

    elements.forEach(element => {
      this.lazyElements.add(element);
      element.classList.add('lazy-loading');

      // 플레이스홀더 설정
      if (!element.dataset.loaded) {
        this.setLoadingPlaceholder(element);
      }

      this.intersectionObserver.observe(element);
    });
  }

  /**
   * 교차점 관찰자 설정
   */
  setupIntersectionObserver() {
    if (!window.IntersectionObserver) return;

    this.intersectionObserver = new IntersectionObserver(
      (entries) => {
        entries.forEach(entry => {
          if (entry.isIntersecting) {
            this.loadLazyElement(entry.target);
          }
        });
      },
      {
        root: null,
        rootMargin: '50px',
        threshold: 0.1
      }
    );
  }

  /**
   * 레이지 요소 로드
   */
  async loadLazyElement(element) {
    if (element.dataset.loaded) return;

    try {
      // 로딩 상태 표시
      element.classList.add('loading');

      // 실제 콘텐츠 로드 (예: 이미지, 컴포넌트 등)
      await this.loadElementContent(element);

      // 로딩 완료 처리
      element.dataset.loaded = 'true';
      element.classList.remove('lazy-loading', 'loading');
      element.classList.add('loaded');

      // 관찰 중지
      this.intersectionObserver.unobserve(element);
      this.lazyElements.delete(element);

    } catch (error) {
      logger.error('레이지 로딩 실패:', error);
      element.classList.add('error');
    }
  }

  /**
   * 요소 콘텐츠 로드
   */
  async loadElementContent(element) {
    return new Promise((resolve) => {
      // 시뮬레이션된 지연 (실제로는 API 호출, 이미지 로드 등)
      setTimeout(() => {
        // 플레이스홀더 제거
        this.removeLoadingPlaceholder(element);
        resolve();
      }, 100);
    });
  }

  /**
   * 로딩 플레이스홀더 설정
   */
  setLoadingPlaceholder(element) {
    const placeholder = document.createElement('div');
    placeholder.className = 'loading-placeholder';
    placeholder.innerHTML = `
      <div class="placeholder-content">
        <div class="placeholder-line"></div>
        <div class="placeholder-line short"></div>
      </div>
    `;

    // 원본 콘텐츠 숨기기
    const originalContent = Array.from(element.children);
    originalContent.forEach(child => {
      child.style.display = 'none';
    });

    element.appendChild(placeholder);
  }

  /**
   * 로딩 플레이스홀더 제거
   */
  removeLoadingPlaceholder(element) {
    const placeholder = element.querySelector('.loading-placeholder');
    if (placeholder) {
      placeholder.remove();
    }

    // 원본 콘텐츠 표시
    Array.from(element.children).forEach(child => {
      child.style.display = '';
    });
  }

  /**
   * 디바운스 함수
   */
  debounce(func, wait) {
    const key = func.toString();

    if (this.debouncedFunctions.has(key)) {
      return this.debouncedFunctions.get(key);
    }

    let timeout;
    const debouncedFn = function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func(...args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };

    this.debouncedFunctions.set(key, debouncedFn);
    return debouncedFn;
  }

  /**
   * 웹 워커 초기화
   */
  initializeWebWorker() {
    if (!window.Worker) return;

    // 웹 워커 코드 생성
    const workerCode = `
      self.onmessage = function(e) {
        const { type, data } = e.data;

        switch(type) {
          case 'FILTER':
            const filtered = data.items.filter(item => {
              // 필터 로직 실행
              return item.status === data.filter.status;
            });
            self.postMessage({ type: 'FILTER_RESULT', data: filtered });
            break;

          case 'SORT':
            const sorted = data.items.sort((a, b) => {
              const aVal = a[data.field] || '';
              const bVal = b[data.field] || '';
              return data.order === 'asc' ?
                aVal.localeCompare(bVal) :
                bVal.localeCompare(aVal);
            });
            self.postMessage({ type: 'SORT_RESULT', data: sorted });
            break;

          case 'CALCULATE_STATS':
            const stats = {
              total: data.items.length,
              totalAmount: data.items.reduce((sum, item) => sum + (parseFloat(item.amount) || 0), 0),
              avgAmount: 0
            };
            stats.avgAmount = stats.total > 0 ? stats.totalAmount / stats.total : 0;
            self.postMessage({ type: 'STATS_RESULT', data: stats });
            break;
        }
      };
    `;

    const blob = new Blob([workerCode], { type: 'application/javascript' });
    this.webWorker = new Worker(URL.createObjectURL(blob));

    this.webWorker.onmessage = (e) => {
      this.handleWorkerMessage(e.data);
    };
  }

  /**
   * 웹 워커 메시지 처리
   */
  handleWorkerMessage(message) {
    const { type, data } = message;

    switch(type) {
      case 'FILTER_RESULT':
        this.emit('filterComplete', data);
        break;
      case 'SORT_RESULT':
        this.emit('sortComplete', data);
        break;
      case 'STATS_RESULT':
        this.emit('statsComplete', data);
        break;
    }
  }

  /**
   * 웹 워커로 작업 전송
   */
  executeInWorker(type, data) {
    if (!this.webWorker) {
      logger.warn('웹 워커가 사용 불가합니다.');
      return Promise.reject(new Error('웹 워커 없음'));
    }

    return new Promise((resolve) => {
      const handler = (e) => {
        const { type: responseType, data: responseData } = e.data;
        if (responseType === `${type}_RESULT`) {
          this.webWorker.removeEventListener('message', handler);
          resolve(responseData);
        }
      };

      this.webWorker.addEventListener('message', handler);
      this.webWorker.postMessage({ type, data });
    });
  }

  /**
   * 성능 모니터링 설정
   */
  setupPerformanceMonitoring() {
    if (!this.config.performanceMonitoring) return;

    // 메모리 사용량 모니터링
    setInterval(() => {
      if (window.performance && window.performance.memory) {
        this.performanceMetrics.memoryUsage.push({
          used: window.performance.memory.usedJSHeapSize,
          total: window.performance.memory.totalJSHeapSize,
          timestamp: Date.now()
        });
        this.limitMetricsArray(this.performanceMetrics.memoryUsage, 50);
      }
    }, 5000);

    // 캐시 크기 모니터링
    setInterval(() => {
      this.cleanupCaches();
    }, 30000);
  }

  /**
   * 캐시 정리
   */
  cleanupCaches() {
    const maxSize = this.config.memoizationCacheSize;

    if (this.renderCache.size > maxSize) {
      const excess = this.renderCache.size - maxSize;
      const keysToDelete = Array.from(this.renderCache.keys()).slice(0, excess);
      keysToDelete.forEach(key => this.renderCache.delete(key));
    }

    if (this.computationCache.size > maxSize) {
      const excess = this.computationCache.size - maxSize;
      const keysToDelete = Array.from(this.computationCache.keys()).slice(0, excess);
      keysToDelete.forEach(key => this.computationCache.delete(key));
    }

    if (this.filterCache.size > maxSize) {
      const excess = this.filterCache.size - maxSize;
      const keysToDelete = Array.from(this.filterCache.keys()).slice(0, excess);
      keysToDelete.forEach(key => this.filterCache.delete(key));
    }
  }

  /**
   * 성능 메트릭스 배열 제한
   */
  limitMetricsArray(array, maxLength = 100) {
    if (array.length > maxLength) {
      array.splice(0, array.length - maxLength);
    }
  }

  /**
   * 캐시 키 생성
   */
  generateCacheKey(type, ...args) {
    return `${type}_${JSON.stringify(args)}`;
  }

  /**
   * 필터 캐시 키 생성
   */
  generateFilterCacheKey(data, filterFunctions) {
    const dataHash = this.hashCode(JSON.stringify(data.map(item => item.id || item.key)));
    const filterHash = this.hashCode(filterFunctions.map(fn => fn.toString()).join(''));
    return `filter_${dataHash}_${filterHash}`;
  }

  /**
   * 해시 코드 생성
   */
  hashCode(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
      const char = str.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // 32bit int로 변환
    }
    return hash;
  }

  /**
   * 이벤트 발생
   */
  emit(eventName, data) {
    const event = new CustomEvent(eventName, { detail: data });
    document.dispatchEvent(event);
  }

  /**
   * 성능 리포트 생성
   */
  generatePerformanceReport() {
    const metrics = this.performanceMetrics;

    const report = {
      renderTime: {
        avg: this.calculateAverage(metrics.renderTime),
        max: Math.max(...metrics.renderTime),
        min: Math.min(...metrics.renderTime),
        samples: metrics.renderTime.length
      },
      filterTime: {
        avg: this.calculateAverage(metrics.filterTime),
        max: Math.max(...metrics.filterTime),
        min: Math.min(...metrics.filterTime),
        samples: metrics.filterTime.length
      },
      scrollTime: {
        avg: this.calculateAverage(metrics.scrollTime),
        max: Math.max(...metrics.scrollTime),
        min: Math.min(...metrics.scrollTime),
        samples: metrics.scrollTime.length
      },
      cacheStatus: {
        renderCache: this.renderCache.size,
        computationCache: this.computationCache.size,
        filterCache: this.filterCache.size
      },
      memoryUsage: metrics.memoryUsage.length > 0 ?
        metrics.memoryUsage[metrics.memoryUsage.length - 1] : null
    };

    return report;
  }

  /**
   * 평균 계산
   */
  calculateAverage(array) {
    return array.length > 0 ? array.reduce((a, b) => a + b, 0) / array.length : 0;
  }

  /**
   * 성능 통계 출력 (개발 모드에서만)
   */
  logPerformanceStats() {
    // 개발 모드 체크 (localhost 또는 DEBUG 모드)
    const isDevelopment = window.location.hostname === 'localhost' ||
                         window.location.hostname === '127.0.0.1' ||
                         window.DEBUG === true;

    if (!isDevelopment) {
      return; // 프로덕션에서는 로그 출력 안 함
    }

    const report = this.generatePerformanceReport();
    console.table(report);
  }

  /**
   * 컴포넌트 정리
   */
  destroy() {
    // 웹 워커 정리
    if (this.webWorker) {
      this.webWorker.terminate();
      this.webWorker = null;
    }

    // 교차점 관찰자 정리
    if (this.intersectionObserver) {
      this.intersectionObserver.disconnect();
      this.intersectionObserver = null;
    }

    // 캐시 정리
    this.renderCache.clear();
    this.computationCache.clear();
    this.filterCache.clear();

    // 이벤트 리스너 제거
    this.debouncedFunctions.clear();
    this.lazyElements.clear();

    logger.debug('🧹 PerformanceOptimizer 정리 완료');
  }
}

// 유틸리티 함수들
export const PerformanceUtils = {
  /**
   * FPS 측정
   */
  measureFPS(callback, duration = 1000) {
    let frames = 0;
    let lastTime = performance.now();

    function count() {
      frames++;
      const currentTime = performance.now();

      if (currentTime >= lastTime + duration) {
        const fps = Math.round((frames * 1000) / (currentTime - lastTime));
        callback(fps);
        frames = 0;
        lastTime = currentTime;
      }

      requestAnimationFrame(count);
    }

    requestAnimationFrame(count);
  },

  /**
   * 메모리 사용량 확인
   */
  getMemoryUsage() {
    if (window.performance && window.performance.memory) {
      return {
        used: window.performance.memory.usedJSHeapSize,
        total: window.performance.memory.totalJSHeapSize,
        limit: window.performance.memory.jsHeapSizeLimit
      };
    }
    return null;
  },

  /**
   * 렌더링 시간 측정
   */
  measureRenderTime(renderFunction) {
    const start = performance.now();
    const result = renderFunction();
    const end = performance.now();

    logger.debug(`렌더링 시간: ${(end - start).toFixed(2)}ms`);
    return { result, duration: end - start };
  }
};