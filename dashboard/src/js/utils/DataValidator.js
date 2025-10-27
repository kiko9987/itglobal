/**
 * 데이터 검증 및 변환 유틸리티
 * 날짜, 숫자, 텍스트 등의 데이터 형식을 표준화
 */
export class DataValidator {
  /**
   * 날짜 형식 정규화
   * @param {string} dateStr
   * @returns {string} YYYY-MM-DD 형식
   */
  static normalizeDate(dateStr) {
    if (!dateStr || dateStr === '-') return '';

    // YYYY/M/D 형식을 YYYY-MM-DD로 변환
    if (dateStr.includes('/')) {
      const parts = dateStr.split('/');
      if (parts.length === 3) {
        const year = parts[0];
        const month = parts[1].padStart(2, '0');
        const day = parts[2].padStart(2, '0');
        return `${year}-${month}-${day}`;
      }
    }

    // YYYY.M.D 형식을 YYYY-MM-DD로 변환
    if (dateStr.includes('.')) {
      const parts = dateStr.split('.');
      if (parts.length === 3) {
        const year = parts[0];
        const month = parts[1].padStart(2, '0');
        const day = parts[2].padStart(2, '0');
        return `${year}-${month}-${day}`;
      }
    }

    // 이미 YYYY-MM-DD 형식인 경우
    const isoDateRegex = /^\d{4}-\d{2}-\d{2}$/;
    if (isoDateRegex.test(dateStr)) {
      return dateStr;
    }

    return '';
  }

  /**
   * 한국어 날짜를 표준 형식으로 변환
   * @param {string} koreanDate
   * @returns {string}
   */
  static parseKoreanDate(koreanDate) {
    if (!koreanDate) return '';

    // "2024년 1월 15일" -> "2024-01-15"
    const match = koreanDate.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
    if (match) {
      const [, year, month, day] = match;
      return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
    }

    return koreanDate;
  }

  /**
   * 숫자 포맷팅 (쉼표 추가)
   * @param {number|string} value
   * @returns {string}
   */
  static formatNumber(value) {
    if (!value && value !== 0) return '';

    const num = typeof value === 'string' ? parseFloat(value.replace(/,/g, '')) : value;
    if (isNaN(num)) return '';

    return num.toLocaleString('ko-KR');
  }

  /**
   * 문자열에서 숫자만 추출
   * @param {string} value
   * @returns {number}
   */
  static parseNumber(value) {
    if (!value) return 0;

    const cleaned = value.toString().replace(/[^\d.-]/g, '');
    const num = parseFloat(cleaned);

    return isNaN(num) ? 0 : num;
  }

  /**
   * 퍼센트 값 검증 및 포맷팅
   * @param {string|number} value
   * @returns {string}
   */
  static formatPercent(value) {
    const num = this.parseNumber(value);
    if (num === 0) return '';

    return `${num}%`;
  }

  /**
   * 이메일 형식 검증
   * @param {string} email
   * @returns {boolean}
   */
  static isValidEmail(email) {
    if (!email) return false;

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }

  /**
   * 전화번호 형식 정규화
   * @param {string} phone
   * @returns {string}
   */
  static normalizePhone(phone) {
    if (!phone) return '';

    // 숫자만 추출
    const numbers = phone.replace(/\D/g, '');

    // 11자리 휴대폰 번호 형식
    if (numbers.length === 11) {
      return numbers.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
    }

    // 10자리 일반 전화번호 형식
    if (numbers.length === 10) {
      return numbers.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    }

    return phone;
  }

  /**
   * 사업자등록번호 형식 정규화
   * @param {string} businessNumber
   * @returns {string}
   */
  static normalizeBusinessNumber(businessNumber) {
    if (!businessNumber) return '';

    const numbers = businessNumber.replace(/\D/g, '');

    if (numbers.length === 10) {
      return numbers.replace(/(\d{3})(\d{2})(\d{5})/, '$1-$2-$3');
    }

    return businessNumber;
  }

  /**
   * 문자열 길이 제한
   * @param {string} text
   * @param {number} maxLength
   * @returns {string}
   */
  static truncateText(text, maxLength = 100) {
    if (!text) return '';

    if (text.length <= maxLength) return text;

    return text.substring(0, maxLength - 3) + '...';
  }

  /**
   * HTML 태그 제거
   * @param {string} html
   * @returns {string}
   */
  static stripHtml(html) {
    if (!html) return '';

    const temp = document.createElement('div');
    temp.innerHTML = html;
    return temp.textContent || temp.innerText || '';
  }

  /**
   * 빈 값 체크 (null, undefined, '', ' ')
   * @param {any} value
   * @returns {boolean}
   */
  static isEmpty(value) {
    if (value === null || value === undefined) return true;
    if (typeof value === 'string' && value.trim() === '') return true;
    if (Array.isArray(value) && value.length === 0) return true;
    if (typeof value === 'object' && Object.keys(value).length === 0) return true;

    return false;
  }

  /**
   * 안전한 문자열 변환
   * @param {any} value
   * @returns {string}
   */
  static safeString(value) {
    if (this.isEmpty(value)) return '';
    return String(value).trim();
  }
}

// 레거시 호환성을 위한 글로벌 객체
window.DataValidator = {
  normalizeDate: DataValidator.normalizeDate,
  parseKoreanDate: DataValidator.parseKoreanDate,
  formatNumber: DataValidator.formatNumber,
  parseNumber: DataValidator.parseNumber,
  formatPercent: DataValidator.formatPercent,
  isValidEmail: DataValidator.isValidEmail,
  normalizePhone: DataValidator.normalizePhone,
  normalizeBusinessNumber: DataValidator.normalizeBusinessNumber,
  truncateText: DataValidator.truncateText,
  stripHtml: DataValidator.stripHtml,
  isEmpty: DataValidator.isEmpty,
  safeString: DataValidator.safeString
};

export default DataValidator;