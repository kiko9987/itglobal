import AmountCalculator from './AmountCalculator.js';

import logger from '../utils/logger.js';
/**
 * 프로젝트 상태 계산 공통 유틸리티
 * ProjectTable과 ModernProjectFilters에서 공통으로 사용
 */
export default class ProjectStatusCalculator {
  /**
   * 상태값 계산 (HTML 태그 없는 순수 텍스트)
   * @param {Object} rowData - 프로젝트 데이터
   * @returns {string} 상태 텍스트
   */
  static calculateStatus(rowData) {
    // 공사 취소 여부 확인 (최우선, 띄어쓰기 무시)
    const collectionNotes = rowData['수금 관련 특이사항'] || rowData['AG'] || '';
    if (collectionNotes && /공사\s*취소/.test(collectionNotes)) {
      return '공사취소';
    }

    const startDateStr = rowData['공사 시작'] || '';
    const endDateStr = rowData['공사 종료'] || '';

    // 금액 정보
    const contractAmount = AmountCalculator.safeParseCurrency(rowData['계약금'] || 0);
    const midAmount = AmountCalculator.safeParseCurrency(rowData['중도금'] || 0);
    const finalAmount = AmountCalculator.safeParseCurrency(rowData['잔금'] || 0);
    const outstandingAmount = AmountCalculator.safeParseCurrency(rowData['미수금'] || rowData['미수금W'] || rowData['W'] || 0);
    const totalAmount = AmountCalculator.safeParseCurrency(rowData['총액 2'] || rowData['총액2'] || rowData['S'] || rowData['총액'] || 0);

    // 수금 확인 (Z열)
    const paymentConfirmed = rowData['수금 확인'] || rowData['Z'] || '';
    const isPaymentConfirmed = (paymentConfirmed === true || paymentConfirmed === 'TRUE' ||
                              paymentConfirmed === '✓' || paymentConfirmed === 'true' ||
                              paymentConfirmed === 'Y' || paymentConfirmed === 'y' ||
                              paymentConfirmed === '1' || paymentConfirmed === 1);

    // 총 입금액 계산
    const totalPaidAmount = contractAmount + midAmount + finalAmount;

    // 날짜 정보 처리
    let startDate = null, endDate = null, today = new Date();
    today.setHours(0, 0, 0, 0);

    if (startDateStr) {
        try {
            startDate = new Date(startDateStr);
            startDate.setHours(0, 0, 0, 0);
        } catch (e) {
            logger.warn('시작일 파싱 오류:', e);
        }
    }

    if (endDateStr) {
        try {
            endDate = new Date(endDateStr);
            endDate.setHours(0, 0, 0, 0);
        } catch (e) {
            logger.warn('종료일 파싱 오류:', e);
        }
    }


    // 0. 🟢 수금완료 체크박스 우선 (최우선) - 사용자가 수동으로 완료 표시한 경우
    if (isPaymentConfirmed) {
        return '공사완료';
    }

    // 0-1. 🟢 미수금 0원이면 공사완료 (두번째 우선) - 실제 수금상태 반영
    if (outstandingAmount === 0 && finalAmount > 0) {
        return '공사완료';
    }

    // 1. 🔴 수금필요: 공사 종료일 지났는데 입금 완료 안됨
    if (endDate && today > endDate) {
        if ((totalAmount > 0 && totalPaidAmount < totalAmount) ||
            (totalAmount === 0 && outstandingAmount > 0)) {
            return '수금필요';
        }
    }

    // 2. 🟢 공사완료: 모든 입금 완료 (기존 로직)
    if ((totalAmount > 0 && totalPaidAmount === totalAmount && outstandingAmount === 0) ||
        (totalAmount === 0 && outstandingAmount === 0 && finalAmount > 0)) {
        return '공사완료';
    }

    // 3. 🟠 확인필요: 이상한 상황들 + 데이터 부족
    if ((totalAmount > 0 && totalPaidAmount === totalAmount && outstandingAmount > 0) ||
        (totalAmount > 0 && totalPaidAmount < totalAmount && contractAmount > 0 && midAmount > 0 && finalAmount > 0) ||
        (totalAmount === 0 && outstandingAmount > 0 && contractAmount > 0 && midAmount > 0 && finalAmount > 0) ||
        (totalPaidAmount === 0 && outstandingAmount === 0) || // 모든 금액 정보 없음
        (!startDateStr && !endDateStr)) { // 날짜 정보 없음
        return '확인필요';
    }

    // 4. ⚫ 공사대기: 오직 공사 시작 전만
    if (startDate && today < startDate) {
        return '공사대기';
    }

    // 5. 🔵 공사진행: 나머지 모든 경우 (입금이 있거나 공사가 시작됨)
    return '공사진행';
  }

  /**
   * 상태 배지 생성
   * @param {Object} rowData - 프로젝트 데이터
   * @returns {string} HTML 배지 문자열
   */
  static createStatusBadge(rowData) {
    const status = this.calculateStatus(rowData);
    const statusClasses = {
      '수금필요': 'status-payment-needed',
      '공사완료': 'status-completed',
      '확인필요': 'status-pending',
      '공사대기': 'status-waiting',
      '공사진행': 'status-in-progress',
      '공사취소': 'status-cancelled'
    };

    const cssClass = statusClasses[status] || 'status-in-progress';
    return `<span class="badge ${cssClass}">${status}</span>`;
  }

  /**
   * 상태별 우선순위 정의
   */
  static getStatusOrder() {
    return ['수금필요', '확인필요', '공사대기', '공사진행', '공사완료', '공사취소'];
  }
}