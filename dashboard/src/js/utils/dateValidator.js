/**
 * 날짜 형식 검증 및 변환 유틸리티
 * 다양한 입력 형식을 YYYY-MM-DD로 정규화
 */

/**
 * 날짜 형식 검증
 * @param {string} dateStr - 검증할 날짜 문자열
 * @returns {Object} - { valid: boolean, normalized: string|null, error: string|null }
 */
export function validateAndNormalizeDate(dateStr) {
  // 빈 값은 허용
  if (!dateStr || dateStr.trim() === '') {
    return { valid: true, normalized: '', error: null };
  }

  const trimmed = dateStr.trim();

  // 1. YYYY-MM-DD 형식 (정규 형식)
  const dashFormat = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
  const dashMatch = trimmed.match(dashFormat);
  if (dashMatch) {
    const [, year, month, day] = dashMatch;
    const normalized = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;

    if (isValidDate(year, month, day)) {
      return { valid: true, normalized, error: null };
    } else {
      return { valid: false, normalized: null, error: '유효하지 않은 날짜입니다' };
    }
  }

  // 2. YYYY/MM/DD 형식
  const slashFormat = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/;
  const slashMatch = trimmed.match(slashFormat);
  if (slashMatch) {
    const [, year, month, day] = slashMatch;
    const normalized = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;

    if (isValidDate(year, month, day)) {
      return { valid: true, normalized, error: null };
    } else {
      return { valid: false, normalized: null, error: '유효하지 않은 날짜입니다' };
    }
  }

  // 3. YYYY.MM.DD 형식
  const dotFormat = /^(\d{4})\.(\d{1,2})\.(\d{1,2})$/;
  const dotMatch = trimmed.match(dotFormat);
  if (dotMatch) {
    const [, year, month, day] = dotMatch;
    const normalized = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;

    if (isValidDate(year, month, day)) {
      return { valid: true, normalized, error: null };
    } else {
      return { valid: false, normalized: null, error: '유효하지 않은 날짜입니다' };
    }
  }

  // 4. YYYYMMDD 형식 (구분자 없음)
  const compactFormat = /^(\d{4})(\d{2})(\d{2})$/;
  const compactMatch = trimmed.match(compactFormat);
  if (compactMatch) {
    const [, year, month, day] = compactMatch;
    const normalized = `${year}-${month}-${day}`;

    if (isValidDate(year, month, day)) {
      return { valid: true, normalized, error: null };
    } else {
      return { valid: false, normalized: null, error: '유효하지 않은 날짜입니다' };
    }
  }

  // 5. 형식 불일치
  return {
    valid: false,
    normalized: null,
    error: '날짜 형식이 올바르지 않습니다 (YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD 형식만 지원)'
  };
}

/**
 * 날짜 유효성 검증 (윤년 고려)
 * @param {string|number} year - 연도
 * @param {string|number} month - 월 (1-12)
 * @param {string|number} day - 일 (1-31)
 * @returns {boolean} - 유효한 날짜 여부
 */
function isValidDate(year, month, day) {
  const y = parseInt(year, 10);
  const m = parseInt(month, 10);
  const d = parseInt(day, 10);

  // 기본 범위 검증
  if (y < 1900 || y > 2100) return false;
  if (m < 1 || m > 12) return false;
  if (d < 1 || d > 31) return false;

  // 월별 일수 검증
  const daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

  // 윤년 검증
  const isLeapYear = (y % 4 === 0 && y % 100 !== 0) || (y % 400 === 0);
  if (isLeapYear) {
    daysInMonth[1] = 29; // 2월을 29일로 변경
  }

  return d <= daysInMonth[m - 1];
}

/**
 * 날짜 범위 검증
 * @param {string} startDate - 시작 날짜 (YYYY-MM-DD)
 * @param {string} endDate - 종료 날짜 (YYYY-MM-DD)
 * @returns {Object} - { valid: boolean, error: string|null }
 */
export function validateDateRange(startDate, endDate) {
  // 둘 다 비어있으면 허용
  if ((!startDate || startDate.trim() === '') && (!endDate || endDate.trim() === '')) {
    return { valid: true, error: null };
  }

  // 하나만 비어있으면 허용 (선택적)
  if (!startDate || startDate.trim() === '' || !endDate || endDate.trim() === '') {
    return { valid: true, error: null };
  }

  // 둘 다 있을 때 순서 검증
  const start = new Date(startDate);
  const end = new Date(endDate);

  if (start > end) {
    return {
      valid: false,
      error: '시작 날짜가 종료 날짜보다 늦을 수 없습니다'
    };
  }

  return { valid: true, error: null };
}

/**
 * 날짜 input 필드에 실시간 검증 적용
 * @param {HTMLInputElement} inputElement - 날짜 input 요소
 * @param {Function} onValidChange - 유효한 값 변경 시 콜백
 */
export function attachDateValidator(inputElement, onValidChange = null) {
  if (!inputElement) return;

  // 기존 리스너 제거 (중복 방지)
  if (inputElement._dateValidatorAttached) return;
  inputElement._dateValidatorAttached = true;

  // blur 핸들러 정의 및 참조 저장
  const blurHandler = (e) => {
    const result = validateAndNormalizeDate(e.target.value);

    if (!result.valid) {
      // 에러 표시
      e.target.classList.add('is-invalid');

      // 에러 메시지 표시 (기존 feedback 요소 찾거나 생성)
      let feedbackEl = e.target.parentElement.querySelector('.invalid-feedback');
      if (!feedbackEl) {
        feedbackEl = document.createElement('div');
        feedbackEl.className = 'invalid-feedback';
        e.target.parentElement.appendChild(feedbackEl);
      }
      feedbackEl.textContent = result.error;
    } else {
      // 정규화된 값으로 업데이트
      if (result.normalized !== e.target.value) {
        e.target.value = result.normalized;
      }

      // 에러 제거
      e.target.classList.remove('is-invalid');
      const feedbackEl = e.target.parentElement.querySelector('.invalid-feedback');
      if (feedbackEl) {
        feedbackEl.remove();
      }

      // 콜백 호출
      if (onValidChange && result.normalized) {
        onValidChange(result.normalized);
      }
    }
  };

  // input 핸들러 정의 및 참조 저장
  const inputHandler = (e) => {
    // 입력 중에는 단순히 invalid 클래스만 제거
    e.target.classList.remove('is-invalid');
  };

  // 핸들러 저장 (cleanup용)
  inputElement._dateBlurHandler = blurHandler;
  inputElement._dateInputHandler = inputHandler;

  // 이벤트 리스너 등록
  inputElement.addEventListener('blur', blurHandler);
  inputElement.addEventListener('input', inputHandler);
}

/**
 * 오늘 날짜를 YYYY-MM-DD 형식으로 반환
 * @returns {string} - YYYY-MM-DD
 */
export function getTodayFormatted() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 날짜 비교 (a < b: -1, a === b: 0, a > b: 1)
 * @param {string} dateA - YYYY-MM-DD
 * @param {string} dateB - YYYY-MM-DD
 * @returns {number} - -1, 0, 1
 */
export function compareDates(dateA, dateB) {
  const a = new Date(dateA);
  const b = new Date(dateB);

  if (a < b) return -1;
  if (a > b) return 1;
  return 0;
}

export default {
  validateAndNormalizeDate,
  validateDateRange,
  attachDateValidator,
  getTodayFormatted,
  compareDates
};
