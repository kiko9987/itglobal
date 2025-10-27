/**
 * Company Badge Module
 * 거래처/회사 뱃지 (해시 기반 색상)
 */
import BaseBadge from './BaseBadge.js';

export default class CompanyBadge extends BaseBadge {
  constructor() {
    super();

    // 거래처용 색상 팔레트 (기존 시스템과 동일)
    this.companyColors = [
      '#e8f5e8',  // 연한 녹색
      '#e0f2f1',  // 연한 틸색
      '#f1f8e9',  // 연한 라임색
      '#e8f5e8',  // 매우 연한 녹색
      '#fff9c4',  // 연한 레몬색
      '#ffecb3',  // 연한 앰버색
      '#f9fbe7',  // 연한 라임 그린색
      '#fff8e1',  // 연한 앰버 화이트색
      '#f0f4c3',  // 연한 라임 옐로우색
      '#e1f5fe',  // 매우 연한 시안색
      '#e0f7fa',  // 연한 틸 블루색
      '#fce4ec',  // 연한 로즈색
      '#f3e5f5',  // 연한 자주색
      '#fef7ff',  // 매우 연한 보라색
      '#ffebee',  // 연한 핑크 화이트색
      '#fafafa'   // 연한 그레이색
    ];
  }

  /**
   * 회사 뱃지 생성
   * @param {string} companyName - 회사명
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  create(companyName, options = {}) {
    if (!companyName) {
      if (options.showEmpty) {
        return this.createBadgeHtml('미등록', 'badge company-badge text-muted', {
          dataAttributes: {
            'badge-type': 'company'
          }
        });
      }
      return '';
    }

    const name = String(companyName).trim();
    const hashIndex = this.getCompanyHash(name);
    const backgroundColor = this.companyColors[hashIndex];

    return this.createBadgeHtml(name, 'badge company-badge', {
      style: `background-color: ${backgroundColor};`,
      dataAttributes: {
        'badge-type': 'company'
      },
      ...options
    });
  }

  /**
   * 회사명의 해시값 계산 (색상 팔레트용)
   * @param {string} company - 회사명
   * @returns {number} 해시값
   */
  getCompanyHash(company) {
    // FNV-1a 해시 알고리즘 (더 나은 분산)
    let hash = 2166136261;
    for (let i = 0; i < company.length; i++) {
      hash ^= company.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return Math.abs(hash) % this.companyColors.length;
  }

  /**
   * 회사별 색상 정보 반환
   * @param {string} companyName - 회사명
   * @returns {Object} 색상 정보
   */
  getCompanyColorInfo(companyName) {
    const name = String(companyName).trim();
    const hashIndex = this.getCompanyHash(name);

    return {
      background: this.companyColors[hashIndex],
      text: 'var(--gray-700)',
      border: 'var(--black-alpha-10)'
    };
  }

  /**
   * 회사명 유효성 검증
   * @param {string} companyName - 회사명
   * @returns {boolean} 유효한지 여부
   */
  isValid(companyName) {
    if (!companyName || typeof companyName !== 'string') return false;

    const trimmed = companyName.trim();
    return trimmed.length > 0;
  }

  /**
   * 회사명 정규화
   * @param {string} companyName - 회사명
   * @returns {string} 정규화된 회사명
   */
  normalize(companyName) {
    if (!companyName) return '';

    return String(companyName).trim();
  }

  /**
   * 회사 통계 정보 반환
   * @param {Array} projectData - 프로젝트 데이터 배열
   * @returns {Object} 회사별 통계
   */
  getCompanyStats(projectData) {
    const stats = {};

    projectData.forEach(project => {
      const company = project.거래처 || project.company;
      if (company) {
        const normalizedName = this.normalize(company);
        stats[normalizedName] = (stats[normalizedName] || 0) + 1;
      }
    });

    return stats;
  }
}