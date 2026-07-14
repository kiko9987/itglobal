/**
 * Badge Factory
 * 모든 뱃지 생성을 중앙에서 관리하는 팩토리 클래스
 */
import ProjectBadge from './ProjectBadge.js';
import StatusBadge from './StatusBadge.js';
import ManagerBadge from './ManagerBadge.js';
import CompanyBadge from './CompanyBadge.js';
import CollectionBadge from './CollectionBadge.js';
import VatBadge from './VatBadge.js';
import InvoiceBadge from './InvoiceBadge.js';

import logger from '../../utils/logger.js';
export default class BadgeFactory {
  constructor() {
    // 각 뱃지 모듈 인스턴스 생성
    this.projectBadge = new ProjectBadge();
    this.statusBadge = new StatusBadge();
    this.managerBadge = new ManagerBadge();
    this.companyBadge = new CompanyBadge();
    this.collectionBadge = new CollectionBadge();
    this.vatBadge = new VatBadge();
    this.invoiceBadge = new InvoiceBadge();

    // 뱃지 타입별 매핑
    this.badgeMap = {
      'project': this.projectBadge,
      'status': this.statusBadge,
      'manager': this.managerBadge,
      'company': this.companyBadge,
      'collection': this.collectionBadge,
      'vat': this.vatBadge,
      'invoice': this.invoiceBadge
    };
  }

  /**
   * 통합 뱃지 생성 메서드
   * @param {string} type - 뱃지 타입 ('project', 'status', 'manager', 'company', 'collection', 'vat')
   * @param {any} data - 뱃지 데이터
   * @param {Object} options - 추가 옵션
   * @returns {string} HTML 문자열
   */
  createBadge(type, data, options = {}) {
    const badgeModule = this.badgeMap[type];

    if (!badgeModule) {
      logger.warn(`[BadgeFactory] 알 수 없는 뱃지 타입: ${type}`);
      return this.createErrorBadge(`Unknown type: ${type}`);
    }

    try {
      return badgeModule.create(data, options);
    } catch (error) {
      logger.error(`[BadgeFactory] ${type} 뱃지 생성 오류:`, error);
      return this.createErrorBadge('Error');
    }
  }

  /**
   * 프로젝트 코드 뱃지 생성
   * @param {string} projectCode - 프로젝트 코드
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  createProjectBadge(projectCode, options = {}) {
    return this.projectBadge.create(projectCode, options);
  }

  /**
   * 상태 뱃지 생성
   * @param {Object|string} projectDataOrStatus - 프로젝트 데이터 또는 상태
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  createStatusBadge(projectDataOrStatus, options = {}) {
    return this.statusBadge.create(projectDataOrStatus, options);
  }

  /**
   * 담당자 뱃지 생성
   * @param {string} managerName - 담당자명
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  createManagerBadge(managerName, options = {}) {
    return this.managerBadge.create(managerName, options);
  }

  /**
   * 회사 뱃지 생성
   * @param {string} companyName - 회사명
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  createCompanyBadge(companyName, options = {}) {
    return this.companyBadge.create(companyName, options);
  }

  /**
   * 수금확인 뱃지 생성
   * @param {any} status - 수금확인 상태
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  createCollectionBadge(status, options = {}) {
    return this.collectionBadge.create(status, options);
  }

  /**
   * 부가세 뱃지 생성
   * @param {any} status - 부가세 상태
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  createVatBadge(status, options = {}) {
    return this.vatBadge.create(status, options);
  }

  /**
   * 뱃지에서 값 추출 (타입별 처리)
   * @param {string} badgeHtml - 뱃지 HTML
   * @param {string} type - 뱃지 타입
   * @returns {string} 추출된 값
   */
  extractValueFromBadge(badgeHtml, type = 'auto') {
    if (!badgeHtml) return '';

    // 타입 자동 감지
    if (type === 'auto') {
      type = this.detectBadgeType(badgeHtml);
    }

    const badgeModule = this.badgeMap[type];

    // 특별 처리가 필요한 뱃지들
    if (type === 'project' && badgeModule.extractCode) {
      return badgeModule.extractCode(badgeHtml);
    }

    if (type === 'collection' && badgeModule.extractStatus) {
      return badgeModule.extractStatus(badgeHtml);
    }

    if (type === 'vat' && badgeModule.extractStatus) {
      return badgeModule.extractStatus(badgeHtml);
    }

    // 기본 값 추출
    if (badgeModule && badgeModule.extractValueFromBadge) {
      return badgeModule.extractValueFromBadge(badgeHtml);
    }

    // 폴백: BaseBadge의 기본 추출 메서드
    return this.projectBadge.extractValueFromBadge(badgeHtml);
  }

  /**
   * 뱃지 타입 자동 감지
   * @param {string} badgeHtml - 뱃지 HTML
   * @returns {string} 감지된 타입
   */
  detectBadgeType(badgeHtml) {
    if (!badgeHtml) return 'unknown';

    const html = badgeHtml.toLowerCase();

    // CSS 클래스로 타입 감지
    if (html.includes('project-badge')) return 'project';
    if (html.includes('status-')) return 'status';
    if (html.includes('manager-badge')) return 'manager';
    if (html.includes('company-badge')) return 'company';
    if (html.includes('collection-badge')) return 'collection';
    if (html.includes('vat-badge')) return 'vat';

    // data 속성으로 타입 감지
    if (html.includes('data-collection-status')) return 'collection';
    if (html.includes('data-vat-status')) return 'vat';

    // 내용으로 타입 감지
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = badgeHtml;
    const text = tempDiv.textContent.trim();

    if (['완료', '미완료', '포함', '별도'].includes(text)) {
      return html.includes('collection') ? 'collection' : 'vat';
    }

    return 'unknown';
  }

  /**
   * 검색 하이라이트 적용
   * @param {string} text - 텍스트
   * @param {string} searchTerm - 검색어
   * @returns {string} 하이라이트된 HTML
   */
  applySearchHighlight(text, searchTerm) {
    return this.projectBadge.applySearchHighlight(text, searchTerm);
  }

  /**
   * 오류 뱃지 생성
   * @param {string} message - 오류 메시지
   * @returns {string} 오류 뱃지 HTML
   */
  createErrorBadge(message = 'Error') {
    return this.projectBadge.createErrorBadge(message);
  }

  /**
   * 빈 뱃지 생성
   * @returns {string} 빈 뱃지 HTML
   */
  createEmptyBadge() {
    return this.projectBadge.createEmptyBadge();
  }

  /**
   * 뱃지 모듈 직접 접근 (고급 사용법)
   * @param {string} type - 뱃지 타입
   * @returns {Object|null} 뱃지 모듈 인스턴스
   */
  getBadgeModule(type) {
    return this.badgeMap[type] || null;
  }

  /**
   * 지원되는 뱃지 타입 목록 반환
   * @returns {Array} 지원되는 타입 배열
   */
  getSupportedTypes() {
    return Object.keys(this.badgeMap);
  }

  /**
   * 팩토리 정보 반환
   * @returns {Object} 팩토리 정보
   */
  getInfo() {
    return {
      version: '1.0.0',
      supportedTypes: this.getSupportedTypes(),
      modules: Object.entries(this.badgeMap).map(([type, module]) => ({
        type,
        className: module.constructor.name
      }))
    };
  }
}