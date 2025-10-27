import logger from '../utils/logger.js';

/**
 * Lazy Loading 유틸리티
 * 이미지, 컴포넌트, 데이터를 필요에 따라 동적으로 로드
 */
export class LazyLoader {
  constructor() {
    this.loadedModules = new Map();
    this.loadingPromises = new Map();
    this.observers = new Map();
  }

  /**
   * 이미지 레이지 로딩
   */
  initImageLazyLoading() {
    if (!('IntersectionObserver' in window)) {
      // 폴백: 모든 이미지 즉시 로드
      this.loadAllImages();
      return;
    }

    const imageObserver = new IntersectionObserver((entries, observer) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          const img = entry.target;
          this.loadImage(img);
          observer.unobserve(img);
        }
      });
    }, {
      rootMargin: '50px 0px',
      threshold: 0.01
    });

    this.observers.set('images', imageObserver);

    // data-lazy-src 속성을 가진 이미지들 관찰
    document.querySelectorAll('img[data-lazy-src]').forEach(img => {
      imageObserver.observe(img);
    });
  }

  /**
   * 개별 이미지 로드
   */
  loadImage(img) {
    const src = img.dataset.lazySrc;
    if (!src) return;

    // 로딩 인디케이터 표시
    img.classList.add('loading');

    const imageLoader = new Image();
    imageLoader.onload = () => {
      img.src = src;
      img.classList.remove('loading');
      img.classList.add('loaded');
      delete img.dataset.lazySrc;
    };

    imageLoader.onerror = () => {
      img.classList.remove('loading');
      img.classList.add('error');
      logger.error('이미지 로드 실패:', src);
    };

    imageLoader.src = src;
  }

  /**
   * 모든 이미지 즉시 로드 (폴백)
   */
  loadAllImages() {
    document.querySelectorAll('img[data-lazy-src]').forEach(img => {
      this.loadImage(img);
    });
  }

  /**
   * 동적 모듈 로드 (코드 스플리팅)
   */
  async loadModule(modulePath, cacheable = true) {
    // 이미 로드된 모듈 확인
    if (cacheable && this.loadedModules.has(modulePath)) {
      return this.loadedModules.get(modulePath);
    }

    // 현재 로딩 중인 모듈 확인
    if (this.loadingPromises.has(modulePath)) {
      return this.loadingPromises.get(modulePath);
    }

    // 새로운 모듈 로드
    const loadPromise = this.dynamicImport(modulePath);
    this.loadingPromises.set(modulePath, loadPromise);

    try {
      const module = await loadPromise;

      if (cacheable) {
        this.loadedModules.set(modulePath, module);
      }

      this.loadingPromises.delete(modulePath);
      return module;

    } catch (error) {
      this.loadingPromises.delete(modulePath);
      logger.error(`모듈 로드 실패: ${modulePath}`, error);
      throw error;
    }
  }

  /**
   * 동적 import 래퍼
   */
  async dynamicImport(modulePath) {
    try {
      return await import(modulePath);
    } catch (error) {
      // 재시도 로직
      logger.warn(`모듈 로드 재시도: ${modulePath}`);
      await this.delay(1000);
      return await import(modulePath);
    }
  }

  /**
   * 컴포넌트 레이지 로딩
   */
  async loadComponent(componentName, container) {
    const loadingElement = this.createLoadingElement();
    container.appendChild(loadingElement);

    try {
      const modulePath = `../components/${componentName}.js`;
      const { default: Component } = await this.loadModule(modulePath);

      const component = new Component();
      await component.init?.();

      container.removeChild(loadingElement);
      return component;

    } catch (error) {
      container.removeChild(loadingElement);
      this.showErrorElement(container, `컴포넌트 로드 실패: ${componentName}`);
      throw error;
    }
  }

  /**
   * 페이지별 스크립트 레이지 로딩
   */
  async loadPageScript(pageName) {
    const scriptPath = `../pages/${pageName}.js`;
    return this.loadModule(scriptPath);
  }

  /**
   * 조건부 폴리필 로딩
   */
  async loadPolyfills() {
    const polyfills = [];

    // IntersectionObserver 폴리필
    if (!('IntersectionObserver' in window)) {
      polyfills.push(import('intersection-observer'));
    }

    // Fetch 폴리필
    if (!('fetch' in window)) {
      polyfills.push(import('whatwg-fetch'));
    }

    // Promise 폴리필
    if (!window.Promise) {
      polyfills.push(import('es6-promise/auto'));
    }

    if (polyfills.length > 0) {
      logger.debug('폴리필 로딩 중...');
      await Promise.all(polyfills);
      logger.debug('폴리필 로딩 완료');
    }
  }

  /**
   * 테이블 행 가상화 (대량 데이터 성능 최적화)
   */
  initVirtualScrolling(tableContainer, data, rowHeight = 50, visibleRows = 20) {
    const totalHeight = data.length * rowHeight;
    const viewportHeight = visibleRows * rowHeight;

    let scrollTop = 0;
    let startIndex = 0;
    let endIndex = Math.min(visibleRows, data.length);

    const scrollContainer = document.createElement('div');
    scrollContainer.style.height = `${viewportHeight}px`;
    scrollContainer.style.overflow = 'auto';

    const contentContainer = document.createElement('div');
    contentContainer.style.height = `${totalHeight}px`;
    contentContainer.style.position = 'relative';

    const visibleContainer = document.createElement('div');
    visibleContainer.style.position = 'absolute';
    visibleContainer.style.top = '0px';
    visibleContainer.style.width = '100%';

    contentContainer.appendChild(visibleContainer);
    scrollContainer.appendChild(contentContainer);

    const updateVisibleRows = () => {
      const newStartIndex = Math.floor(scrollTop / rowHeight);
      const newEndIndex = Math.min(newStartIndex + visibleRows, data.length);

      if (newStartIndex !== startIndex || newEndIndex !== endIndex) {
        startIndex = newStartIndex;
        endIndex = newEndIndex;

        visibleContainer.style.top = `${startIndex * rowHeight}px`;
        visibleContainer.innerHTML = '';

        for (let i = startIndex; i < endIndex; i++) {
          const row = this.createTableRow(data[i], i);
          visibleContainer.appendChild(row);
        }
      }
    };

    scrollContainer.addEventListener('scroll', () => {
      scrollTop = scrollContainer.scrollTop;
      updateVisibleRows();
    });

    updateVisibleRows();
    tableContainer.appendChild(scrollContainer);

    return {
      container: scrollContainer,
      updateData: (newData) => {
        data = newData;
        updateVisibleRows();
      }
    };
  }

  /**
   * 로딩 엘리먼트 생성
   */
  createLoadingElement() {
    const loading = document.createElement('div');
    loading.className = 'lazy-loading';
    loading.innerHTML = `
      <div class="text-center py-3">
        <div class="spinner-border spinner-border-sm text-primary" role="status">
          <span class="visually-hidden">로딩 중...</span>
        </div>
        <div class="mt-2 text-muted">로딩 중...</div>
      </div>
    `;
    return loading;
  }

  /**
   * 에러 엘리먼트 표시
   */
  showErrorElement(container, message) {
    const error = document.createElement('div');
    error.className = 'lazy-error alert alert-warning';
    error.innerHTML = `
      <i class="fas fa-exclamation-triangle me-2"></i>
      ${message}
      <button class="btn btn-sm btn-outline-warning ms-2" onclick="location.reload()">
        다시 시도
      </button>
    `;
    container.appendChild(error);
  }

  /**
   * 테이블 행 생성 (가상 스크롤용)
   */
  createTableRow(data, index) {
    const row = document.createElement('div');
    row.className = 'virtual-row';
    row.style.height = '50px';
    row.style.display = 'flex';
    row.style.alignItems = 'center';
    row.style.borderBottom = '1px solid #dee2e6';
    row.style.padding = '0 1rem';

    // 데이터에 따라 행 내용 구성
    row.innerHTML = `
      <div class="flex-grow-1">${data.name || `Row ${index + 1}`}</div>
      <div class="text-muted">${data.value || ''}</div>
    `;

    return row;
  }

  /**
   * 지연 함수
   */
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * 리소스 프리로딩
   */
  preloadResources(resources = []) {
    resources.forEach(resource => {
      const link = document.createElement('link');
      link.rel = 'prefetch';
      link.href = resource;
      document.head.appendChild(link);
    });
  }

  /**
   * 정리 함수
   */
  cleanup() {
    // 옵저버들 정리
    this.observers.forEach(observer => observer.disconnect());
    this.observers.clear();

    // 캐시 정리
    this.loadedModules.clear();
    this.loadingPromises.clear();
  }
}

// 전역 인스턴스 생성
export const lazyLoader = new LazyLoader();

// CSS 스타일 추가
const style = document.createElement('style');
style.textContent = `
  img[data-lazy-src] {
    background: #f0f0f0;
    min-height: 100px;
    transition: opacity 0.3s ease;
  }

  img.loading {
    opacity: 0.6;
    background: #f0f0f0 url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><circle cx="50" cy="50" r="20" fill="none" stroke="%23ccc" stroke-width="4"><animate attributeName="r" values="20;25;20" dur="1s" repeatCount="indefinite"/></circle></svg>') center no-repeat;
    background-size: 30px 30px;
  }

  img.loaded {
    opacity: 1;
  }

  img.error {
    opacity: 0.5;
    background: #f8f9fa url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="%23dc3545"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/></svg>') center no-repeat;
    background-size: 24px 24px;
  }

  .lazy-loading {
    min-height: 100px;
    display: flex;
    align-items: center;
    justify-content: center;
  }

  .virtual-row:hover {
    background-color: rgba(0, 0, 0, 0.05);
  }
`;

if (!document.getElementById('lazy-loader-styles')) {
  style.id = 'lazy-loader-styles';
  document.head.appendChild(style);
}

export default LazyLoader;