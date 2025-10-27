/**
 * AmountCalculator - 금액 계산 및 변환 유틸리티
 */
export default class AmountCalculator {
  /**
   * 숫자를 한글로 변환
   * @param {number} num - 변환할 숫자
   * @returns {string} 한글 금액
   */
  static numberToKorean(num) {
    const koreanNumbers = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
    const koreanUnits = ['', '십', '백', '천', '만', '십만', '백만', '천만', '억', '십억', '백억', '천억', '조'];

    if (num === 0) return '영';

    let result = '';
    const numStr = Math.floor(num).toString();
    const len = numStr.length;

    for (let i = 0; i < len; i++) {
      const digit = parseInt(numStr[i]);
      const unitIndex = len - i - 1;

      if (digit > 0) {
        // '일십', '일십만', '일십억' -> '십', '십만', '십억' 최적화
        if (digit === 1 && (unitIndex === 1 || unitIndex === 5 || unitIndex === 9)) {
          result += koreanUnits[unitIndex];
        } else {
          result += koreanNumbers[digit] + koreanUnits[unitIndex];
        }
      }
    }

    return result || '영';
  }

  /**
   * 숫자를 콤마 포맷으로 변환
   * @param {number} num - 변환할 숫자
   * @returns {string} 콤마 포맷 문자열
   */
  static formatWithComma(num) {
    return new Intl.NumberFormat('ko-KR').format(num);
  }

  /**
   * 콤마 포맷 문자열을 숫자로 변환
   * @param {string} str - 콤마 포맷 문자열
   * @returns {number} 숫자
   */
  static parseCommaString(str) {
    if (!str) return 0;
    return parseFloat(str.replace(/,/g, '')) || 0;
  }

  /**
   * 통화 문자열을 숫자로 안전하게 변환
   * 콤마, "원", 공백 등 모든 비숫자 문자를 제거
   * @param {string|number} value - 통화 문자열 또는 숫자
   * @returns {number} 숫자
   * @example
   * safeParseCurrency("5,000,000원") // 5000000
   * safeParseCurrency("3,200,000") // 3200000
   * safeParseCurrency("1234") // 1234
   * safeParseCurrency(5000) // 5000
   */
  static safeParseCurrency(value) {
    if (!value) return 0;
    if (typeof value === 'number') return value;

    // 숫자, 마이너스, 소수점만 남기고 모두 제거
    const cleaned = String(value).replace(/[^0-9.-]/g, '');
    return parseFloat(cleaned) || 0;
  }

  /**
   * 부가세 포함 금액 계산
   * @param {number} amount - 원금
   * @param {boolean} includeVAT - 부가세 포함 여부
   * @returns {number} 계산된 금액
   */
  static calculateWithVAT(amount, includeVAT) {
    if (!amount) return 0;
    return includeVAT ? amount * 1.1 : amount;
  }

  /**
   * 금액 표시 문자열 생성
   * @param {number} amount - 금액
   * @param {boolean} includeVAT - 부가세 포함 여부
   * @returns {string} 표시 문자열
   */
  static formatAmountDisplay(amount, includeVAT) {
    if (!amount || amount === 0) {
      return {
        formatted: '0원',
        korean: '(금 영원정)',
        vat: ''
      };
    }

    const total = this.calculateWithVAT(amount, includeVAT);
    const formatted = this.formatWithComma(total);
    const korean = this.numberToKorean(total);
    const vatStatus = includeVAT ? 'VAT 포함' : 'VAT 별도';

    return {
      formatted: `${formatted}원`,
      korean: `(금 ${korean}원정)`,
      vat: vatStatus,
      full: `${formatted}원 (금 ${korean}원정) / ${vatStatus}`
    };
  }

  /**
   * 금액 버튼 값 추가
   * @param {number} currentAmount - 현재 금액
   * @param {number} addAmount - 추가할 금액
   * @returns {number} 새 금액
   */
  static addAmount(currentAmount, addAmount) {
    return (currentAmount || 0) + addAmount;
  }
}
