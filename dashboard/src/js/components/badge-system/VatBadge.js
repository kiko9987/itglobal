/**
 * VAT Badge Module
 * 부가세 뱃지 (true/false 토글 상태)
 */
import BaseBadge from './BaseBadge.js';

export default class VatBadge extends BaseBadge {
  constructor() {
    super();

    // 부가세 상태별 설정
    this.statusConfig = {
      true: {
        text: '포함',
        cssClass: 'vat-included',
        style: 'background-color: #cfe2ff; color: #084298;'
      },
      false: {
        text: '미포함',
        cssClass: 'vat-excluded',
        style: 'background-color: #f8f9fa; color: #6c757d;'
      },
      '-': {
        text: '-',
        cssClass: 'vat-empty',
        style: 'background-color: var(--gray-100, #f8f9fa); color: var(--gray-600, #6c757d);'
      }
    };
  }

  /**
   * 부가세 뱃지 생성 (수금확인과 비슷한 시각적 스타일)
   * @param {any} status - 부가세 상태 (true, false, '포함', '별도', '-' 등)
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  create(status, options = {}) {
    const normalizedStatus = this.normalizeStatus(status);
    const config = this.statusConfig[normalizedStatus] || this.statusConfig['-'];

    const result = this.createBadgeHtml(config.text, `badge vat-badge ${config.cssClass}`, {
      style: config.style,
      title: `부가세: ${config.text}`,
      dataAttributes: {
        'vat-status': normalizedStatus
      },
      ...options
    });

    return result;
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

    // true 값들 (포함)
    if (status === true || statusStr === 'true' || statusStr === '포함' ||
        statusStr === 'included' || statusStr === 'y' || statusStr === 'yes' ||
        statusStr === '1' || status === 1) {
      return 'true';
    }

    // false 값들 (별도)
    if (status === false || statusStr === 'false' || statusStr === '별도' ||
        statusStr === '미포함' || statusStr === 'excluded' || statusStr === 'n' ||
        statusStr === 'no' || statusStr === '0' || status === 0) {
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

    // data-vat-status 속성에서 추출
    const badgeElement = tempDiv.querySelector('[data-vat-status]');
    if (badgeElement) {
      return badgeElement.getAttribute('data-vat-status');
    }

    // 텍스트 내용으로 판단
    const text = tempDiv.textContent.trim();
    if (text === '포함') return 'true';
    if (text === '별도') return 'false';

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
   * 부가세 포함 여부 확인
   * @param {any} status - 상태값
   * @returns {boolean} 포함 여부
   */
  isIncluded(status) {
    return this.normalizeStatus(status) === 'true';
  }

  /**
   * 부가세 별도 여부 확인
   * @param {any} status - 상태값
   * @returns {boolean} 별도 여부
   */
  isSeparate(status) {
    return this.normalizeStatus(status) === 'false';
  }

  /**
   * 상태별 색상 정보 반환
   * @param {any} status - 상태값
   * @returns {Object} 색상 정보
   */
  getStatusColorInfo(status) {
    const normalizedStatus = this.normalizeStatus(status);

    const colorMap = {
      'true': {
        background: 'var(--info-bg, #d1ecf1)',
        text: 'var(--info-text, #055160)',
        border: 'var(--info-border, #b6d4d9)'
      },
      'false': {
        background: 'var(--secondary-bg, #e2e3e5)',
        text: 'var(--secondary-text, #41464b)',
        border: 'var(--secondary-border, #c4c8cb)'
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
   * 부가세 통계 정보 반환
   * @param {Array} projectData - 프로젝트 데이터 배열
   * @returns {Object} 부가세 통계
   */
  getVatStats(projectData) {
    const stats = {
      included: 0,
      excluded: 0,
      empty: 0,
      total: 0
    };

    projectData.forEach(project => {
      const status = project['부가세'] || project.vat;
      const normalized = this.normalizeStatus(status);

      stats.total++;

      if (normalized === 'true') {
        stats.included++;
      } else if (normalized === 'false') {
        stats.excluded++;
      } else {
        stats.empty++;
      }
    });

    return stats;
  }
}