/**
 * Base Badge Class
 * 모든 뱃지의 공통 기능을 제공하는 베이스 클래스
 */
export default class BaseBadge {
  constructor() {
    this.cssPrefix = 'ub-'; // unified-badge prefix
  }

  /**
   * 공통 HTML 생성 메서드
   * @param {string} content - 뱃지 내용
   * @param {string|Array} cssClasses - CSS 클래스들
   * @param {Object} options - 추가 옵션들
   * @returns {string} HTML 문자열
   */
  createBadgeHtml(content, cssClasses, options = {}) {
    const classes = Array.isArray(cssClasses) ? cssClasses : [cssClasses];
    const classString = classes.filter(Boolean).join(' ');

    const attributes = this.buildAttributes(options);

    return `<span class="${classString}"${attributes}>${content}</span>`;
  }

  /**
   * HTML 속성 생성
   * @param {Object} options - 옵션 객체
   * @returns {string} 속성 문자열
   */
  buildAttributes(options = {}) {
    const attrs = [];

    if (options.title) {
      attrs.push(` title="${this.escapeHtml(options.title)}"`);
    }

    if (options.style) {
      attrs.push(` style="${this.escapeHtml(options.style)}"`);
    }

    if (options.dataAttributes) {
      Object.entries(options.dataAttributes).forEach(([key, value]) => {
        attrs.push(` data-${key}="${this.escapeHtml(value)}"`);
      });
    }

    if (options.id) {
      attrs.push(` id="${this.escapeHtml(options.id)}"`);
    }

    return attrs.join('');
  }

  /**
   * HTML 이스케이프
   * @param {string} text - 이스케이프할 텍스트
   * @returns {string} 이스케이프된 텍스트
   */
  escapeHtml(text) {
    if (typeof text !== 'string') return text;

    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * 검색 하이라이트 적용
   * @param {string} content - 원본 내용
   * @param {string} searchTerm - 검색어
   * @returns {string} 하이라이트된 내용
   */
  applySearchHighlight(content, searchTerm) {
    if (!searchTerm || !content) return content;

    const regex = new RegExp(`(${this.escapeRegex(searchTerm)})`, 'gi');
    return content.replace(regex, '<span class="search-highlight">$1</span>');
  }

  /**
   * 정규식 특수문자 이스케이프
   * @param {string} text - 이스케이프할 텍스트
   * @returns {string} 이스케이프된 텍스트
   */
  escapeRegex(text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /**
   * 뱃지 HTML에서 값 추출
   * @param {string} badgeHtml - 뱃지 HTML
   * @returns {string} 추출된 값
   */
  extractValueFromBadge(badgeHtml) {
    if (!badgeHtml) return '';

    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = badgeHtml;

    // .ms-1 요소가 있으면 우선 (프로젝트 코드용)
    const textElement = tempDiv.querySelector('.ms-1');
    if (textElement) {
      return textElement.textContent.trim();
    }

    // 뱃지 요소에서 직접 추출
    const badgeElement = tempDiv.querySelector(`[class*="${this.cssPrefix}"], .badge`);
    if (badgeElement) {
      return badgeElement.textContent.trim();
    }

    // 전체 텍스트 반환
    return tempDiv.textContent.trim();
  }

  /**
   * CSS 클래스명으로 안전한 문자열 생성
   * @param {string} text - 원본 텍스트
   * @returns {string} 안전한 클래스명
   */
  sanitizeClassName(text) {
    return String(text)
      .replace(/[^a-zA-Z0-9가-힣]/g, '') // 특수문자 제거
      .replace(/\s+/g, '') // 공백 제거
      .trim();
  }

  /**
   * 해시 기반 색상 인덱스 생성
   * @param {string} text - 해시할 텍스트
   * @param {number} colorCount - 색상 개수
   * @returns {number} 색상 인덱스
   */
  getHashIndex(text, colorCount) {
    let hash = 0;
    for (let i = 0; i < text.length; i++) {
      hash = ((hash << 5) - hash) + text.charCodeAt(i);
      hash = hash & hash; // 32비트 정수로 변환
    }
    return Math.abs(hash) % colorCount;
  }

  /**
   * 빈 뱃지 생성
   * @returns {string} 빈 뱃지 HTML
   */
  createEmptyBadge() {
    return this.createBadgeHtml('-', `${this.cssPrefix}empty`, {
      style: 'opacity: 0.5; color: #6c757d;'
    });
  }

  /**
   * 오류 뱃지 생성
   * @param {string} message - 오류 메시지
   * @returns {string} 오류 뱃지 HTML
   */
  createErrorBadge(message = 'Error') {
    return this.createBadgeHtml(message, `${this.cssPrefix}error`, {
      style: 'background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb;'
    });
  }
}