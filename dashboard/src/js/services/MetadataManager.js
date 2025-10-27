import logger from '../utils/logger.js';

/**
 * MetadataManager - 초기 페이지 메타데이터 관리
 * 전문가 리뷰: "초기 페이지에서는 최소 메타 정보만 제공하고 상세 데이터는 API로만 받도록"
 */

export default class MetadataManager {
  constructor() {
    this.metadata = null;
    this.statistics = null;
  }

  /**
   * 페이지 메타데이터 로드 (사용자 권한, 기본 설정 등)
   */
  loadPageMetadata() {
    // DOM에서 서버가 제공한 최소 메타데이터 추출
    const metaElement = document.getElementById('page-metadata');

    if (metaElement) {
      try {
        this.metadata = JSON.parse(metaElement.textContent);
        logger.debug('[MetadataManager] 페이지 메타데이터 로드 완료:', this.metadata);
        return this.metadata;
      } catch (error) {
        logger.error('[MetadataManager] 메타데이터 파싱 실패:', error);
      }
    }

    // 폴백: window 객체에서 추출
    this.metadata = {
      user_role: window.userRole || 'viewer',
      user_email: window.userEmail || '',
      user_can_create: window.userCanCreate || false,
      csrf_token: window.csrfToken || '',
      app_version: window.appVersion || '1.0.0'
    };

    logger.debug('[MetadataManager] 폴백 메타데이터 사용:', this.metadata);
    return this.metadata;
  }

  /**
   * 프로젝트 통계 정보 API 로드 (분리된 엔드포인트)
   */
  async loadProjectStatistics() {
    try {
      const response = await fetch('/api/projects/statistics', {
        cache: 'no-cache',
        headers: {
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const result = await response.json();

      if (result.success) {
        this.statistics = result.data;
        return this.statistics;
      } else {
        throw new Error(result.message || '통계 로드 실패');
      }

    } catch (error) {
      logger.error('[MetadataManager] 통계 로드 실패:', error);

      // 폴백 통계
      this.statistics = {
        total_projects: 0,
        total_amount: 0,
        pending_amount: 0,
        completed_projects: 0,
        last_updated: new Date().toISOString()
      };

      return this.statistics;
    }
  }

  /**
   * 사용자 권한 확인
   */
  getUserPermissions() {
    if (!this.metadata) {
      this.loadPageMetadata();
    }

    const userRole = this.metadata.user_role.toLowerCase();

    return {
      canEdit: userRole === 'admin' || userRole === 'editor',
      canDelete: userRole === 'admin',
      canCreate: this.metadata.user_can_create || userRole === 'admin' || userRole === 'editor',
      userRole: userRole,
      userEmail: this.metadata.user_email
    };
  }

  /**
   * 통계 UI 업데이트
   */
  updateStatisticsUI(statistics = null) {
    const stats = statistics || this.statistics;
    if (!stats) return;

    // 통계 카드 업데이트
    this.updateStatCard('total-projects', stats.total_projects, '개');
    this.updateStatCard('total-amount', this.formatCurrency(stats.total_amount), '');
    this.updateStatCard('pending-amount', this.formatCurrency(stats.pending_amount), '');
    this.updateStatCard('completed-projects', stats.completed_projects, '개');

    // 마지막 업데이트 시간
    if (stats.last_updated) {
      const updateElement = document.getElementById('stats-last-updated');
      if (updateElement) {
        const updateTime = new Date(stats.last_updated);
        updateElement.textContent = `마지막 업데이트: ${updateTime.toLocaleString('ko-KR')}`;
      }
    }
  }

  /**
   * 개별 통계 카드 업데이트
   */
  updateStatCard(elementId, value, suffix = '') {
    const element = document.getElementById(elementId);
    if (element) {
      element.textContent = `${value}${suffix}`;
    }
  }

  /**
   * 통화 포맷팅
   */
  formatCurrency(amount) {
    if (!amount || amount === 0) return '0원';

    return new Intl.NumberFormat('ko-KR', {
      style: 'currency',
      currency: 'KRW',
      maximumFractionDigits: 0
    }).format(amount);
  }

  /**
   * 현재 메타데이터 반환
   */
  getMetadata() {
    return this.metadata;
  }

  /**
   * 현재 통계 반환
   */
  getStatistics() {
    return this.statistics;
  }

  /**
   * 메타데이터 디버그 정보
   */
  getDebugInfo() {
    return {
      hasMetadata: !!this.metadata,
      hasStatistics: !!this.statistics,
      metadata: this.metadata,
      statistics: this.statistics,
      permissions: this.getUserPermissions()
    };
  }

  /**
   * 정리 작업
   */
  cleanup() {
    this.metadata = null;
    this.statistics = null;
    logger.debug('[MetadataManager] 정리 작업 완료');
  }
}