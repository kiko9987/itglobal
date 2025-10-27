/**
 * Unified Badge System
 * Centralizes all badge creation and management using modular architecture
 * Maintains backward compatibility while using new BadgeFactory internally
 * Singleton pattern: 여러 컴포넌트에서 사용해도 하나의 인스턴스만 생성
 */
import BadgeFactory from './badge-system/BadgeFactory.js';

class UnifiedBadgeSystem {
  constructor() {
    // Singleton pattern: 이미 인스턴스가 있으면 그것을 반환
    if (UnifiedBadgeSystem.instance) {
      return UnifiedBadgeSystem.instance;
    }

    // Initialize the modular badge factory
    this.badgeFactory = new BadgeFactory();

    // Expose badge modules for direct access if needed
    this.projectBadge = this.badgeFactory.projectBadge;
    this.statusBadge = this.badgeFactory.statusBadge;
    this.managerBadge = this.badgeFactory.managerBadge;
    this.companyBadge = this.badgeFactory.companyBadge;
    this.collectionBadge = this.badgeFactory.collectionBadge;
    this.vatBadge = this.badgeFactory.vatBadge;

    // Legacy compatibility properties (delegated to modules)
    this.knownManagers = this.managerBadge.knownManagers;
    this.managementTeam = this.managerBadge.managementTeam;
    this.resignedTeam = this.managerBadge.resignedTeam;
    this.statusOrder = this.statusBadge.statusOrder;

    // 싱글톤 인스턴스 저장
    UnifiedBadgeSystem.instance = this;
  }

  /**
   * Main factory method for creating badges
   * @param {string} type - Badge type (project, status, manager, company, collection, vat)
   * @param {any} value - Value to display
   * @param {Object} options - Additional options
   * @returns {string} HTML string for the badge
   */
  createBadge(type, value, options = {}) {
    // Delegate to BadgeFactory
    const result = this.badgeFactory.createBadge(type, value, options);
    return result;
  }

  /**
   * Create project code badge - delegate to ProjectBadge module
   */
  createProjectBadge(projectCode, options = {}) {
    return this.projectBadge.create(projectCode, options);
  }

  /**
   * Create status badge - delegate to StatusBadge module
   */
  createStatusBadge(projectData, options = {}) {
    return this.statusBadge.create(projectData, options);
  }

  /**
   * Create manager badge - delegate to ManagerBadge module
   */
  createManagerBadge(managerName, options = {}) {
    return this.managerBadge.create(managerName, options);
  }

  /**
   * Create company badge - delegate to CompanyBadge module
   */
  createCompanyBadge(companyName, options = {}) {
    return this.companyBadge.create(companyName, options);
  }

  /**
   * Create collection status badge - delegate to CollectionBadge module
   */
  createCollectionBadge(status, options = {}) {
    return this.collectionBadge.create(status, options);
  }

  /**
   * Create VAT badge - delegate to VatBadge module
   */
  createVatBadge(status, options = {}) {
    const result = this.vatBadge.create(status, options);
    return result;
  }

  /**
   * Create generic badge for unknown types
   */
  createGenericBadge(value, options = {}) {
    return this.badgeFactory.createErrorBadge(value);
  }

  /**
   * Create empty badge placeholder
   */
  createEmptyBadge() {
    return this.badgeFactory.createEmptyBadge();
  }

  /**
   * Calculate project status - delegate to StatusBadge module
   */
  calculateStatus(projectData) {
    return this.statusBadge.calculateStatus(projectData);
  }

  /**
   * Helper methods - delegate to appropriate modules
   */
  getStatusClass(status) {
    return this.statusBadge.getStatusClass(status);
  }

  isCollectionCompleted(status) {
    return this.collectionBadge.isCompleted(status);
  }

  isVatIncluded(status) {
    return this.vatBadge.isIncluded(status);
  }

  getCompanyColor(companyName) {
    const colorInfo = this.companyBadge.getCompanyColorInfo(companyName);
    return colorInfo.background;
  }

  sanitizeClassName(name) {
    return this.managerBadge.sanitizeClassName(name);
  }

  getCompanyHash(company) {
    return this.companyBadge.getCompanyHash(company);
  }

  getManagerHash(manager) {
    return this.managerBadge.getManagerHash(manager);
  }

  getDataAttributes(options) {
    // Legacy support - not needed in new architecture
    return '';
  }

  /**
   * Migration helper: Extract value from badge HTML
   */
  extractValueFromBadge(badgeHtml) {
    return this.badgeFactory.extractValueFromBadge(badgeHtml);
  }

  /**
   * Apply search highlighting to badge content
   */
  applySearchHighlight(badgeHtml, searchTerm) {
    return this.badgeFactory.applySearchHighlight(badgeHtml, searchTerm);
  }
}

export default UnifiedBadgeSystem;