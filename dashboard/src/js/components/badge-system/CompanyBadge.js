/**
 * Company Badge Module
 * 유입 구분 뱃지 — 알려진 값은 지정 색상, 나머지는 해시 기반 색상
 */
import BaseBadge from './BaseBadge.js';

// 유입 구분 값별 고정 색상 매핑 (플랫폼 정체성 반영)
const INFLOW_COLORS = {
  '홈페이지':   { bg: '#dbeafe', fg: '#1e40af' },  // 파란색
  '카카오톡':   { bg: '#fef08a', fg: '#713f12' },  // 카카오 노랑
  '채널톡':     { bg: '#e0f2fe', fg: '#075985' },  // 하늘색
  '전화':       { bg: '#e5e7eb', fg: '#374151' },  // 회색
  '메일':       { bg: '#e0e7ff', fg: '#3730a3' },  // 인디고
  '당근':       { bg: '#fed7aa', fg: '#9a3412' },  // 당근 주황
  '숨고':       { bg: '#dcfce7', fg: '#14532d' },  // 숨고 초록
  '큐플레이스': { bg: '#e9d5ff', fg: '#6b21a8' },  // 보라
  '거래처':     { bg: '#c7d2fe', fg: '#312e81' },  // 진한 파랑
  '온라인':     { bg: '#ccfbf1', fg: '#0f766e' },  // 청록
  '소개':       { bg: '#fce7f3', fg: '#9d174d' },  // 핑크
};

export default class CompanyBadge extends BaseBadge {
  constructor() {
    super();

    // 미정의 값 fallback 색상 팔레트
    this.companyColors = [
      '#e8f5e8', '#e0f2f1', '#f1f8e9', '#e8f5e8',
      '#fff9c4', '#ffecb3', '#f9fbe7', '#fff8e1',
      '#f0f4c3', '#e1f5fe', '#e0f7fa', '#fce4ec',
      '#f3e5f5', '#fef7ff', '#ffebee', '#fafafa',
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
    const known = INFLOW_COLORS[name];
    let style;
    if (known) {
      style = `background-color: ${known.bg}; color: ${known.fg}; font-weight: 600;`;
    } else {
      const hashIndex = this.getCompanyHash(name);
      style = `background-color: ${this.companyColors[hashIndex]};`;
    }

    return this.createBadgeHtml(name, 'badge company-badge', {
      style,
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
    const known = INFLOW_COLORS[name];
    if (known) {
      return { background: known.bg, text: known.fg, border: 'var(--black-alpha-10)' };
    }
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