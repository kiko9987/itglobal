/**
 * Manager Badge Module
 * 담당자 뱃지 (CSS 토큰 기반 + 해시 기반 색상)
 */
import BaseBadge from './BaseBadge.js';

import logger from '../../utils/logger.js';
export default class ManagerBadge extends BaseBadge {
  constructor() {
    super();

    // 기존 담당자들 (CSS 토큰 사용) - 관리팀 포함
    this.knownManagers = [
      '박정우', '강성환', '박용구', '박민우', '권태훈',
      '주영민', '빈승정', '박민재', '조성헌', '황해승', '강민석',
      '고광일', '황샛별'
    ];

    // 퇴사팀 (API에서 동적으로 로드)
    this.resignedTeam = [];

    // 신규 담당자용 예비 색상들 (기존 색상과 겹치지 않게)
    this.fallbackManagerColors = [
      '#f1f8e9',  // 연한 라임색
      '#fce4ec',  // 연한 로즈색
      '#e0f2f1',  // 연한 틸색
      '#fff9c4',  // 연한 레몬색
      '#f3e5f5',  // 연한 자주색
      '#e8f5e8',  // 매우 연한 녹색
      '#ffecb3',  // 연한 앰버색
      '#e1f5fe',  // 매우 연한 시안색
      '#f9fbe7',  // 연한 라임 그린색
      '#fef7ff',  // 매우 연한 보라색
      '#e0f7fa',  // 연한 틸 블루색
      '#fff8e1',  // 연한 앰버 화이트색
      '#f0f4c3',  // 연한 라임 옐로우색
      '#fafafa',  // 연한 그레이색
      '#ffebee'   // 연한 핑크 화이트색
    ];

    // API에서 퇴사자 목록 로드
    this.loadResignedManagers();
  }

  /**
   * API에서 퇴사자 목록 로드
   */
  async loadResignedManagers() {
    try {
      const response = await fetch('/api/resigned-managers', {
        credentials: 'include'
      });

      if (response.ok) {
        const data = await response.json();
        if (data.success && Array.isArray(data.resigned_managers)) {
          this.resignedTeam = data.resigned_managers;
          logger.debug('[ManagerBadge] 퇴사자 목록 로드 완료:', this.resignedTeam);
        }
      }
    } catch (error) {
      logger.warn('[ManagerBadge] 퇴사자 목록 로드 실패:', error);
      // 실패 시 빈 배열 유지
    }
  }

  /**
   * 담당자 뱃지 생성 (기존 로직 완전 복원)
   * @param {string} managerName - 담당자명
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  create(managerName, options = {}) {
    if (!managerName) {
      return options.showEmpty ? this.createEmptyBadge() : '';
    }

    const name = String(managerName).trim();
    const isResigned = this.resignedTeam.some(user => user.name === name);

    if (isResigned) {
      return this.createBadgeHtml(`${name} (퇴사)`, 'badge manager-badge resigned', {
        dataAttributes: {
          'badge-type': 'manager'
        }
      });
    }

    // 기존 담당자는 CSS 토큰 사용
    if (this.knownManagers.includes(name)) {
      const safeManagerName = this.sanitizeClassName(name);
      const managerClass = `manager-${safeManagerName}`;
      return this.createBadgeHtml(name, `badge manager-badge ${managerClass}`, {
        dataAttributes: {
          'badge-type': 'manager'
        }
      });
    }

    // 신규 담당자는 해시 기반 인라인 스타일
    const hashIndex = this.getManagerHash(name);
    const backgroundColor = this.fallbackManagerColors[hashIndex];

    return this.createBadgeHtml(name, 'badge manager-badge', {
      style: `background-color: ${backgroundColor};`,
      title: '신규 담당자',
      dataAttributes: {
        'badge-type': 'manager'
      }
    });
  }

  /**
   * 담당자 타입 판별
   * @param {string} managerName - 담당자명
   * @returns {string} 타입 ('management', 'resigned', 'known', 'new')
   */
  getManagerType(managerName) {
    if (!managerName) return 'unknown';

    const name = String(managerName).trim();

    if (this.resignedTeam.some(user => user.name === name)) return 'resigned';
    if (this.knownManagers.includes(name)) return 'known';

    return 'new';
  }

  /**
   * 담당자명의 해시값 계산 (예비 색상 팔레트용)
   * @param {string} manager - 담당자명
   * @returns {number} 해시값
   */
  getManagerHash(manager) {
    // FNV-1a 해시 알고리즘 (더 나은 분산)
    let hash = 2166136261;
    for (let i = 0; i < manager.length; i++) {
      hash ^= manager.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return Math.abs(hash) % this.fallbackManagerColors.length;
  }

  /**
   * 담당자별 색상 정보 반환
   * @param {string} managerName - 담당자명
   * @returns {Object} 색상 정보
   */
  getManagerColorInfo(managerName) {
    const type = this.getManagerType(managerName);
    const name = String(managerName).trim();

    switch (type) {
      case 'resigned':
        return {
          background: 'var(--gray-100)',
          text: 'var(--gray-500)',
          border: 'var(--gray-300)',
          opacity: 'var(--black-alpha-70)'
        };

      case 'known':
        const safeName = this.sanitizeClassName(name);
        return {
          background: `var(--manager-${safeName}-bg)`,
          text: 'var(--gray-700)',
          border: 'var(--black-alpha-10)'
        };

      case 'new':
        const hashIndex = this.getManagerHash(name);
        return {
          background: this.fallbackManagerColors[hashIndex],
          text: 'var(--gray-700)',
          border: 'var(--black-alpha-10)'
        };

      default:
        return {
          background: 'var(--gray-100)',
          text: 'var(--gray-700)',
          border: 'var(--black-alpha-10)'
        };
    }
  }

  /**
   * 담당자 유효성 검증
   * @param {string} managerName - 담당자명
   * @returns {boolean} 유효한지 여부
   */
  isValid(managerName) {
    if (!managerName || typeof managerName !== 'string') return false;

    const trimmed = managerName.trim();
    return trimmed.length > 0;
  }

  /**
   * 담당자명 정규화
   * @param {string} managerName - 담당자명
   * @returns {string} 정규화된 담당자명
   */
  normalize(managerName) {
    if (!managerName) return '';

    return String(managerName).trim();
  }

  /**
   * 모든 알려진 담당자 목록 반환
   * @returns {Array} 담당자 목록
   */
  getAllKnownManagers() {
    return [
      ...this.resignedTeam.map(user => user.name),
      ...this.knownManagers
    ];
  }

  /**
   * 활성 담당자 목록 반환 (퇴사자 제외)
   * @returns {Array} 활성 담당자 목록
   */
  getActiveManagers() {
    return [...this.knownManagers];
  }

  /**
   * 담당자 통계 정보 반환
   * @param {Array} projectData - 프로젝트 데이터 배열
   * @returns {Object} 담당자별 통계
   */
  getManagerStats(projectData) {
    const stats = {};

    projectData.forEach(project => {
      const manager = project.담당자 || project.manager;
      if (manager) {
        const normalizedName = this.normalize(manager);
        stats[normalizedName] = (stats[normalizedName] || 0) + 1;
      }
    });

    return stats;
  }
}