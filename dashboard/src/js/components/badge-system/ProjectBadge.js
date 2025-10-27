/**
 * Project Badge Module
 * 프로젝트 코드 뱃지 (R/G/P/I 타입 + 전체 텍스트)
 */
import BaseBadge from './BaseBadge.js';

export default class ProjectBadge extends BaseBadge {
  constructor() {
    super();
    this.typeMap = {
      'R': 'r-type',
      'G': 'g-type',
      'P': 'p-type',
      'I': 'i-type'
    };
  }

  /**
   * 프로젝트 코드 뱃지 생성 (기존 형태 완전 보존: 뱃지 + 텍스트)
   * @param {string} projectCode - 프로젝트 코드
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  create(projectCode, options = {}) {
    if (!projectCode) {
      return options.showEmpty ? this.createEmptyBadge() : '';
    }

    const code = String(projectCode).trim();
    const firstChar = code.charAt(0).toUpperCase();
    const typeClass = this.typeMap[firstChar];

    if (!typeClass) {
      // 타입이 없는 경우 단순 텍스트 반환
      return code;
    }

    // 기존 형태 완전 보존: project-badge + 타입 클래스 + ms-1 텍스트
    const badgeHtml = this.createBadgeHtml(firstChar, `project-badge ${typeClass}`, {
      dataAttributes: {
        'badge-type': 'project'
      }
    });
    const textHtml = `<span class="ms-1">${code}</span>`;

    return badgeHtml + textHtml;
  }

  /**
   * 프로젝트 코드 HTML에서 실제 코드 추출
   * @param {string} badgeHtml - 뱃지 HTML
   * @returns {string} 프로젝트 코드
   */
  extractCode(badgeHtml) {
    if (!badgeHtml) return '';

    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = badgeHtml;

    // .ms-1 요소에서 프로젝트 코드 추출 (우선순위)
    const codeElement = tempDiv.querySelector('.ms-1');
    if (codeElement) {
      return codeElement.textContent.trim();
    }

    // .ms-1이 없으면 전체 텍스트에서 추출
    const fullText = tempDiv.textContent.trim();

    // 단일 문자(R/G/P/I) 제거하고 실제 코드만 추출
    if (fullText.length > 1) {
      const firstChar = fullText.charAt(0).toUpperCase();
      if (this.typeMap[firstChar]) {
        // 첫 문자가 타입이면 제거
        return fullText.substring(1).trim();
      }
    }

    return fullText;
  }

  /**
   * 프로젝트 코드에서 타입 추출
   * @param {string} projectCode - 프로젝트 코드
   * @returns {string|null} 타입 ('R', 'G', 'P', 'I' 또는 null)
   */
  extractType(projectCode) {
    if (!projectCode) return null;

    const firstChar = String(projectCode).charAt(0).toUpperCase();
    return this.typeMap[firstChar] ? firstChar : null;
  }

  /**
   * 타입별 색상 정보 반환
   * @param {string} type - 타입 ('R', 'G', 'P', 'I')
   * @returns {Object} 색상 정보
   */
  getTypeColorInfo(type) {
    const colorMap = {
      'R': {
        background: 'var(--project-r-type)',
        text: 'var(--project-r-type-text)',
        border: 'var(--project-r-type)'
      },
      'G': {
        background: 'var(--project-g-type)',
        text: 'var(--project-g-type-text)',
        border: 'var(--project-g-type)'
      },
      'P': {
        background: 'var(--project-p-type)',
        text: 'var(--project-p-type-text)',
        border: 'var(--project-p-type)'
      },
      'I': {
        background: 'var(--project-i-type)',
        text: 'var(--project-i-type-text)',
        border: 'var(--project-i-type)'
      }
    };

    return colorMap[type] || null;
  }

  /**
   * 프로젝트 코드 유효성 검증
   * @param {string} projectCode - 프로젝트 코드
   * @returns {boolean} 유효한지 여부
   */
  isValid(projectCode) {
    if (!projectCode || typeof projectCode !== 'string') return false;

    const trimmed = projectCode.trim();
    if (trimmed.length === 0) return false;

    // 기본적으로 모든 문자열 허용, 단지 타입 여부만 구분
    return true;
  }

  /**
   * 프로젝트 코드 정규화
   * @param {string} projectCode - 프로젝트 코드
   * @returns {string} 정규화된 프로젝트 코드
   */
  normalize(projectCode) {
    if (!projectCode) return '';

    return String(projectCode).trim().toUpperCase();
  }
}