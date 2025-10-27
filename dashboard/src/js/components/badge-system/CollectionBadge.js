/**
 * Collection Badge Module
 * 수금확인 뱃지 (true/false 토글 상태)
 */
import BaseBadge from './BaseBadge.js';

export default class CollectionBadge extends BaseBadge {
  constructor() {
    super();

    // 수금확인 상태별 설정
    this.statusConfig = {
      true: {
        text: '완료',
        cssClass: 'collection-completed',
        style: 'background-color: var(--success-bg, #d1edcc); color: var(--success-text, #0f5132);'
      },
      false: {
        text: '대기',
        cssClass: 'collection-pending',
        style: 'background-color: var(--warning-bg, #fff3cd); color: var(--warning-text, #664d03);'
      },
      '-': {
        text: '-',
        cssClass: 'collection-empty',
        style: 'background-color: var(--gray-100, #f8f9fa); color: var(--gray-600, #6c757d);'
      }
    };
  }

  /**
   * 수금확인 뱃지 생성
   * @param {any} status - 수금확인 상태 (true, false, '완료', '미완료', '-' 등)
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  create(status, options = {}) {
    const normalizedStatus = this.normalizeStatus(status);
    const config = this.statusConfig[normalizedStatus] || this.statusConfig['-'];

    return this.createBadgeHtml(config.text, `badge collection-badge ${config.cssClass}`, {
      style: config.style,
      title: `수금확인: ${config.text}`,
      dataAttributes: {
        'collection-status': normalizedStatus
      },
      ...options
    });
  }

  /**
   * 상태 정규화
   * @param {any} status - 원본 상태값
   * @returns {string} 정규화된 상태 ('true', 'false', '-')
   */
  normalizeStatus(status) {
    // null, undefined, 빈 문자열 처리
    if (status === null || status === undefined || status === '') {
      return '-';
    }

    // 문자열 변환 후 소문자로 처리
    const statusStr = String(status).toLowerCase().trim();

    // true 값들
    if (status === true || statusStr === 'true' || statusStr === '완료' ||
        statusStr === 'y' || statusStr === 'yes' || statusStr === '✓' ||
        statusStr === '1' || status === 1) {
      return 'true';
    }

    // false 값들
    if (status === false || statusStr === 'false' || statusStr === '미완료' ||
        statusStr === '대기' || statusStr === 'n' || statusStr === 'no' ||
        statusStr === '0' || status === 0) {
      return 'false';
    }

    // 기본값: 빈 상태
    return '-';
  }

  /**
   * 뱃지에서 상태값 추출
   * @param {string} badgeHtml - 뱃지 HTML
   * @returns {string} 상태값 ('true', 'false', '-')
   */
  extractStatus(badgeHtml) {
    if (!badgeHtml) return '-';

    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = badgeHtml;

    // data-collection-status 속성에서 추출
    const badgeElement = tempDiv.querySelector('[data-collection-status]');
    if (badgeElement) {
      return badgeElement.getAttribute('data-collection-status');
    }

    // 텍스트 내용으로 판단
    const text = tempDiv.textContent.trim();
    if (text === '완료') return 'true';
    if (text === '미완료') return 'false';

    return '-';
  }

  /**
   * 상태 유효성 검증
   * @param {any} status - 상태값
   * @returns {boolean} 유효한지 여부
   */
  isValidStatus(status) {
    const normalized = this.normalizeStatus(status);
    return ['true', 'false', '-'].includes(normalized);
  }

  /**
   * 불린 값으로 변환
   * @param {any} status - 상태값
   * @returns {boolean|null} 불린값 또는 null (빈 상태일 때)
   */
  toBoolean(status) {
    const normalized = this.normalizeStatus(status);

    if (normalized === 'true') return true;
    if (normalized === 'false') return false;
    return null; // 빈 상태
  }

  /**
   * 수금 완료 여부 확인 (기존 로직과 호환)
   * @param {any} status - 상태값
   * @returns {boolean} 완료 여부
   */
  isCompleted(status) {
    return this.normalizeStatus(status) === 'true';
  }

  /**
   * 상태별 색상 정보 반환
   * @param {any} status - 상태값
   * @returns {Object} 색상 정보
   */
  getStatusColorInfo(status) {
    const normalizedStatus = this.normalizeStatus(status);
    const config = this.statusConfig[normalizedStatus] || this.statusConfig['-'];

    const colorMap = {
      'true': {
        background: 'var(--success-bg, #d1edcc)',
        text: 'var(--success-text, #0f5132)',
        border: 'var(--success-border, #badbcc)'
      },
      'false': {
        background: 'var(--warning-bg, #fff3cd)',
        text: 'var(--warning-text, #664d03)',
        border: 'var(--warning-border, #ffecb5)'
      },
      '-': {
        background: 'var(--gray-100, #f8f9fa)',
        text: 'var(--gray-600, #6c757d)',
        border: 'var(--gray-300, #dee2e6)'
      }
    };

    return colorMap[normalizedStatus] || colorMap['-'];
  }

  /**
   * 수금확인 통계 정보 반환
   * @param {Array} projectData - 프로젝트 데이터 배열
   * @returns {Object} 수금확인 통계
   */
  getCollectionStats(projectData) {
    const stats = {
      completed: 0,
      pending: 0,
      empty: 0,
      total: 0
    };

    projectData.forEach(project => {
      const status = project['수금 확인'] || project.Z;
      const normalized = this.normalizeStatus(status);

      stats.total++;

      if (normalized === 'true') {
        stats.completed++;
      } else if (normalized === 'false') {
        stats.pending++;
      } else {
        stats.empty++;
      }
    });

    return stats;
  }
}