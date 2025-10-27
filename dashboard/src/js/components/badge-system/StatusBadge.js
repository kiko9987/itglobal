/**
 * Status Badge Module
 * 프로젝트 상태 계산 및 뱃지 생성
 */
import BaseBadge from './BaseBadge.js';

import logger from '../../utils/logger.js';
export default class StatusBadge extends BaseBadge {
  constructor() {
    super();

    // 상태 순서 (필터링 및 정렬용)
    this.statusOrder = ['수금필요', '확인필요', '공사대기', '공사진행', '공사완료', '공사취소'];

    // 상태별 CSS 클래스 매핑
    this.statusClassMap = {
      '수금필요': 'status-payment-needed',
      '공사완료': 'status-completed',
      '확인필요': 'status-review',
      '공사진행': 'status-in-progress',
      '공사대기': 'status-waiting',
      '공사취소': 'status-cancelled'
    };

    // 기본 상태
    this.defaultStatus = '확인필요';
  }

  /**
   * 상태 뱃지 생성
   * @param {Object|string} projectDataOrStatus - 프로젝트 데이터 또는 직접 상태값
   * @param {Object} options - 옵션
   * @returns {string} HTML 문자열
   */
  create(projectDataOrStatus, options = {}) {
    let status;

    if (typeof projectDataOrStatus === 'string') {
      // 직접 상태값이 전달된 경우
      status = projectDataOrStatus;
    } else if (projectDataOrStatus && projectDataOrStatus.calculatedStatus) {
      // 미리 계산된 상태가 있는 경우
      status = projectDataOrStatus.calculatedStatus;
    } else {
      // 프로젝트 데이터에서 상태 계산
      status = this.calculateStatus(projectDataOrStatus);
    }

    const statusClass = this.getStatusClass(status);

    // 기존 클래스명 사용 (.badge + status-*)
    return this.createBadgeHtml(status, `badge ${statusClass}`, options);
  }

  /**
   * 프로젝트 데이터에서 상태 계산 (기존 로직 완전 복제)
   * @param {Object} projectData - 프로젝트 데이터
   * @returns {string} 계산된 상태
   */
  calculateStatus(projectData) {
    if (!projectData || typeof projectData !== 'object') {
      return this.defaultStatus;
    }

    // 공사 취소 여부 확인 (최우선, 띄어쓰기 무시)
    const collectionNotes = projectData['수금 관련 특이사항'] || projectData['AG'] || '';
    if (collectionNotes && /공사\s*취소/.test(collectionNotes)) {
      return '공사취소';
    }

    // 필드 추출
    const startDate = projectData['공사 시작'] || '';
    const endDate = projectData['공사 종료'] || '';
    const contractAmount = parseFloat(projectData.계약금 || 0);
    const intermediateAmount = parseFloat(projectData.중도금 || 0);
    const finalAmount = parseFloat(projectData.잔금 || 0);
    const outstandingAmount = parseFloat(projectData.미수금 || projectData.미수금W || projectData.W || 0);
    const totalAmount = parseFloat(projectData['총액 2'] || projectData.총액2 || projectData.S || projectData.총액 || 0);
    const collectionConfirmed = projectData['수금 확인'] || projectData.Z || '';

    // 수금확인 여부 판단
    const isCollectionConfirmed = this.isCollectionCompleted(collectionConfirmed);
    const totalPaid = contractAmount + intermediateAmount + finalAmount;

    // 날짜 파싱
    let parsedStartDate = null;
    let parsedEndDate = null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    if (startDate) {
      try {
        parsedStartDate = new Date(startDate);
        parsedStartDate.setHours(0, 0, 0, 0);
      } catch (error) {
        logger.warn('시작일 파싱 오류:', error);
      }
    }

    if (endDate) {
      try {
        parsedEndDate = new Date(endDate);
        parsedEndDate.setHours(0, 0, 0, 0);
      } catch (error) {
        logger.warn('종료일 파싱 오류:', error);
      }
    }

    // 상태 계산 로직 (기존과 동일)
    if (isCollectionConfirmed || (outstandingAmount === 0 && finalAmount > 0)) {
      return '공사완료';
    }

    if (parsedEndDate && today > parsedEndDate &&
        ((totalAmount > 0 && totalPaid < totalAmount) || (totalAmount === 0 && outstandingAmount > 0))) {
      return '수금필요';
    }

    if ((totalAmount > 0 && totalPaid === totalAmount && outstandingAmount === 0) ||
        (totalAmount === 0 && outstandingAmount === 0 && finalAmount > 0)) {
      return '공사완료';
    }

    if ((totalAmount > 0 && totalPaid === totalAmount && outstandingAmount > 0) ||
        (totalAmount > 0 && totalPaid < totalAmount && contractAmount > 0 && intermediateAmount > 0 && finalAmount > 0) ||
        (totalAmount === 0 && outstandingAmount > 0 && contractAmount > 0 && intermediateAmount > 0 && finalAmount > 0) ||
        (totalPaid === 0 && outstandingAmount === 0) ||
        (!startDate && !endDate)) {
      return '확인필요';
    }

    if (parsedStartDate && today < parsedStartDate) {
      return '공사대기';
    }

    return '공사진행';
  }

  /**
   * 상태에 대응하는 CSS 클래스 반환
   * @param {string} status - 상태
   * @returns {string} CSS 클래스
   */
  getStatusClass(status) {
    return this.statusClassMap[status] || this.statusClassMap[this.defaultStatus];
  }

  /**
   * 수금 완료 여부 판단
   * @param {any} status - 수금확인 값
   * @returns {boolean} 완료 여부
   */
  isCollectionCompleted(status) {
    return status === true || status === 'true' || status === '✓' ||
           status === 'Y' || status === 'y' || status === '완료' ||
           status === 'TRUE' || status === '1' || status === 1;
  }

  /**
   * 상태 순서 반환 (필터링용)
   * @returns {Array} 상태 순서
   */
  getStatusOrder() {
    return [...this.statusOrder];
  }

  /**
   * 상태 유효성 검증
   * @param {string} status - 상태
   * @returns {boolean} 유효한지 여부
   */
  isValidStatus(status) {
    return this.statusOrder.includes(status);
  }

  /**
   * 상태별 우선순위 반환 (정렬용)
   * @param {string} status - 상태
   * @returns {number} 우선순위 (낮을수록 우선)
   */
  getStatusPriority(status) {
    const index = this.statusOrder.indexOf(status);
    return index === -1 ? 999 : index;
  }

  /**
   * 상태별 색상 정보 반환
   * @param {string} status - 상태
   * @returns {Object} 색상 정보
   */
  getStatusColorInfo(status) {
    const colorMap = {
      '수금필요': {
        background: 'var(--status-payment-bg)',
        text: 'var(--status-payment-text)',
        border: 'var(--status-payment-bg)'
      },
      '공사완료': {
        background: 'var(--status-completed-bg)',
        text: 'var(--status-completed-text)',
        border: 'var(--status-completed-bg)'
      },
      '확인필요': {
        background: 'var(--status-review-bg)',
        text: 'var(--status-review-text)',
        border: 'var(--status-review-bg)'
      },
      '공사진행': {
        background: 'var(--status-progress-bg)',
        text: 'var(--status-progress-text)',
        border: 'var(--status-progress-bg)'
      },
      '공사대기': {
        background: 'var(--status-waiting-bg)',
        text: 'var(--status-waiting-text)',
        border: 'var(--status-waiting-bg)'
      },
      '공사취소': {
        background: '#8B7E7D',  // 어두운 회갈색 파스텔
        text: '#FFE5D9',        // 밝은 복숭아색 파스텔
        border: '#8B7E7D'
      }
    };

    return colorMap[status] || colorMap[this.defaultStatus];
  }

  /**
   * 프로젝트 데이터의 완성도 체크
   * @param {Object} projectData - 프로젝트 데이터
   * @returns {boolean} 데이터가 불완전한지 여부
   */
  isDataIncomplete(projectData) {
    if (!projectData) return true;

    const requiredFields = [
      '사업자', '현장 담당자', '담당자 연락처', '현장 주소', '공사 내용',
      '공사 시작', '공사 종료', '총액 1', '총액 2', '기계 분류', '브랜드',
      '견적서 및 계약서 폴더 경로'
    ];

    return requiredFields.some(field => {
      const value = projectData[field];
      return !value ||
             value.toString().trim() === '' ||
             value.toString().trim() === '-' ||
             value.toString().trim() === 'null' ||
             value.toString().trim() === 'undefined';
    });
  }
}