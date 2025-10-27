import logger from '../utils/logger.js';

/**
 * DataManager - 데이터 로딩과 캐싱 전담
 * 전문가 리뷰: "데이터 로딩을 각각 별도 클래스로 나눠 app.init()이 orchestration만 담당하게"
 */

export default class DataManager {
  constructor() {
    // 캐시 설정
    this.CACHE_CONFIG = {
      namespace: 'itg_dashboard',
      version: '1.0',
      ttl: 30 * 1000, // 30초
      maxSize: 4 * 1024 * 1024, // 4MB
      key: 'project_data'
    };
  }

  /**
   * 프로젝트 데이터 가져오기 (캐시 우선)
   */
  async loadProjectData(forceRefresh = false) {
    try {
      // 강제 새로고침이 아닌 경우 캐시 먼저 확인
      if (!forceRefresh) {
        const cachedData = this.getCachedData();
        if (cachedData) {
          logger.debug('[DataManager] 캐시된 데이터 사용');
          return { data: cachedData, fromCache: true };
        }
      }

      // 서버에서 최신 데이터 가져오기
      const freshData = await this.fetchFromServer(forceRefresh);
      if (freshData && freshData.length > 0) {
        this.cacheData(freshData);
        logger.debug('[DataManager] 서버에서 새 데이터 로드');
        return { data: freshData, fromCache: false };
      }

      throw new Error('데이터를 불러올 수 없습니다');
    } catch (error) {
      logger.error('[DataManager] 데이터 로딩 실패:', error);
      throw error;
    }
  }

  /**
   * 서버에서 프로젝트 데이터 가져오기
   */
  async fetchFromServer(forceRefresh = false) {
    try {
      // 브라우저 캐시 활용: forceRefresh가 아닐 때는 고정 URL 사용
      let apiUrl;
      const cacheMode = forceRefresh ? 'no-cache' : 'default';

      if (forceRefresh) {
        // 강제 새로고침: timestamp로 캐시 무효화
        const timestamp = new Date().getTime();
        apiUrl = `/api/projects/list?refresh=true&t=${timestamp}`;
      } else {
        // 일반 조회: 고정 URL로 브라우저 캐시 활용 (Cache-Control: max-age=30)
        apiUrl = `/api/projects/list`;
      }

      // API 요청 중복 제거를 위한 전역 promise 캐시
      window.apiCache = window.apiCache || {};

      // 캐시 키는 고정된 값을 사용 (타임스탬프 제외)
      const cacheKey = `/api/projects/list${forceRefresh ? '_force' : ''}`;

      // 강제 새로고침이 아니고 이미 진행 중인 API 요청이 있는지 확인
      if (!forceRefresh && window.apiCache[cacheKey]) {
        logger.debug('[DataManager] 기존 API 요청 재사용');
        const cachedResponse = await window.apiCache[cacheKey];
        return cachedResponse;
      }

      logger.debug('[DataManager] 새로운 API 요청 시작');

      // 새로운 요청 시작
      window.apiCache[cacheKey] = (async () => {
        const response = await fetch(apiUrl, {
          cache: cacheMode,
          headers: forceRefresh ? {
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache'
          } : {}
        });

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const response_data = await response.json();

        if (response_data.error || !response_data.success) {
          throw new Error(response_data.error || 'API 호출 실패');
        }

        const data = response_data.data;
        return Array.isArray(data) ? data : [];
      })();

      // 요청 완료 후 캐시 정리 (5초 후)
      window.apiCache[cacheKey].finally(() => {
        setTimeout(() => {
          delete window.apiCache[cacheKey];
        }, 5000);
      });

      return await window.apiCache[cacheKey];
    } catch (error) {
      logger.error('[DataManager] 서버 데이터 로드 실패:', error);
      throw error;
    }
  }

  /**
   * 캐시 관리 메서드들
   */
  getCacheKey(key) {
    return `${this.CACHE_CONFIG.namespace}_v${this.CACHE_CONFIG.version}_${key}`;
  }

  validateCacheData(data) {
    if (!Array.isArray(data)) {
      throw new Error('캐시 데이터는 배열이어야 합니다');
    }

    if (data.length > 0) {
      const sample = data[0];
      if (!sample || typeof sample !== 'object') {
        throw new Error('잘못된 프로젝트 데이터 형식');
      }
    }

    return true;
  }

  getCachedData() {
    try {
      const dataKey = this.getCacheKey(this.CACHE_CONFIG.key);
      const timestampKey = this.getCacheKey('timestamp');
      const versionKey = this.getCacheKey('version');

      const cached = localStorage.getItem(dataKey);
      const timestamp = localStorage.getItem(timestampKey);
      const version = localStorage.getItem(versionKey);

      // 버전 체크
      if (version !== this.CACHE_CONFIG.version) {
        logger.warn('[DataManager] 캐시 버전 불일치, 클리어합니다');
        this.clearCache();
        return null;
      }

      if (cached && timestamp) {
        const age = Date.now() - parseInt(timestamp);

        if (age < this.CACHE_CONFIG.ttl) {
          const data = JSON.parse(cached);
          this.validateCacheData(data);
          return data;
        } else {
          this.clearCache();
        }
      }
    } catch (error) {
      logger.warn('[DataManager] 캐시 데이터 읽기 실패:', error.message);
      this.clearCache();
    }
    return null;
  }

  cacheData(data) {
    try {
      this.validateCacheData(data);

      const serialized = JSON.stringify(data);
      const estimatedSize = serialized.length * 2;

      if (estimatedSize > this.CACHE_CONFIG.maxSize) {
        logger.warn('[DataManager] 데이터가 너무 큼, 캐싱 건너뜀', {
          size: Math.round(estimatedSize / 1024) + 'KB',
          limit: Math.round(this.CACHE_CONFIG.maxSize / 1024) + 'KB'
        });
        return false;
      }

      this.clearCache();

      const dataKey = this.getCacheKey(this.CACHE_CONFIG.key);
      const timestampKey = this.getCacheKey('timestamp');
      const versionKey = this.getCacheKey('version');

      localStorage.setItem(dataKey, serialized);
      localStorage.setItem(timestampKey, Date.now().toString());
      localStorage.setItem(versionKey, this.CACHE_CONFIG.version);

      return true;

    } catch (error) {
      if (error.name === 'QuotaExceededError') {
        logger.error('[DataManager] localStorage 용량 초과:', error);
        this.clearCache();
      } else {
        logger.warn('[DataManager] 캐시 저장 실패:', error.message);
      }
      return false;
    }
  }

  clearCache() {
    try {
      const namespace = this.CACHE_CONFIG.namespace;
      const keysToRemove = [];

      for (let key in localStorage) {
        if (key.startsWith(namespace)) {
          keysToRemove.push(key);
        }
      }

      keysToRemove.forEach(key => localStorage.removeItem(key));

    } catch (error) {
      logger.warn('[DataManager] 캐시 클리어 실패:', error.message);
    }
  }
}