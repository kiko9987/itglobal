/**
 * FormValidator - 폼 검증 유틸리티
 */
export default class FormValidator {
  /**
   * 필수 필드 검증
   * @param {string} fieldName - 필드명
   * @param {any} value - 검증할 값
   * @returns {object} { isValid: boolean, message: string }
   */
  static validateRequired(fieldName, value) {
    if (!value || value.toString().trim() === '') {
      return {
        isValid: false,
        message: `${fieldName}은(는) 필수 입력 항목입니다.`
      };
    }
    return { isValid: true, message: '' };
  }

  /**
   * 이메일 형식 검증
   * @param {string} email - 이메일 주소
   * @returns {object} { isValid: boolean, message: string }
   */
  static validateEmail(email) {
    if (!email || email.trim() === '') {
      return { isValid: true, message: '' }; // 선택 필드이므로 빈 값 허용
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return {
        isValid: false,
        message: '올바른 이메일 형식을 입력해주세요.'
      };
    }

    return { isValid: true, message: '' };
  }

  /**
   * 전화번호 형식 검증 (한국 전화번호)
   * @param {string} phone - 전화번호
   * @returns {object} { isValid: boolean, message: string }
   */
  static validatePhone(phone) {
    if (!phone || phone.trim() === '') {
      return { isValid: true, message: '' }; // 선택 필드이므로 빈 값 허용
    }

    // 010-1234-5678, 02-1234-5678, 031-123-4567 등의 형식
    const phoneRegex = /^(0\d{1,2})-?(\d{3,4})-?(\d{4})$/;
    if (!phoneRegex.test(phone.replace(/\s/g, ''))) {
      return {
        isValid: false,
        message: '올바른 전화번호 형식을 입력해주세요. (예: 010-1234-5678)'
      };
    }

    return { isValid: true, message: '' };
  }

  /**
   * 날짜 범위 검증 (시작일 <= 종료일)
   * @param {string} startDate - 시작일 (YYYY-MM-DD)
   * @param {string} endDate - 종료일 (YYYY-MM-DD)
   * @returns {object} { isValid: boolean, message: string }
   */
  static validateDateRange(startDate, endDate) {
    if (!startDate || !endDate) {
      return { isValid: true, message: '' }; // 둘 다 있을 때만 검증
    }

    const start = new Date(startDate);
    const end = new Date(endDate);

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      return {
        isValid: false,
        message: '올바른 날짜 형식을 입력해주세요.'
      };
    }

    if (end < start) {
      return {
        isValid: false,
        message: '종료일은 시작일보다 과거일 수 없습니다.'
      };
    }

    return { isValid: true, message: '' };
  }

  /**
   * 숫자 범위 검증
   * @param {number} value - 검증할 값
   * @param {number} min - 최소값 (선택)
   * @param {number} max - 최대값 (선택)
   * @returns {object} { isValid: boolean, message: string }
   */
  static validateNumberRange(value, min = null, max = null) {
    const num = parseFloat(value);

    if (isNaN(num)) {
      return {
        isValid: false,
        message: '올바른 숫자를 입력해주세요.'
      };
    }

    if (min !== null && num < min) {
      return {
        isValid: false,
        message: `${min} 이상의 값을 입력해주세요.`
      };
    }

    if (max !== null && num > max) {
      return {
        isValid: false,
        message: `${max} 이하의 값을 입력해주세요.`
      };
    }

    return { isValid: true, message: '' };
  }

  /**
   * 체크박스 최대 선택 개수 검증
   * @param {NodeList} checkboxes - 체크박스 NodeList
   * @param {number} maxCount - 최대 선택 개수
   * @returns {object} { isValid: boolean, message: string, selectedCount: number }
   */
  static validateCheckboxLimit(checkboxes, maxCount) {
    const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;

    if (checkedCount > maxCount) {
      return {
        isValid: false,
        message: `최대 ${maxCount}개까지 선택 가능합니다.`,
        selectedCount: checkedCount
      };
    }

    return { isValid: true, message: '', selectedCount: checkedCount };
  }

  /**
   * 폼 데이터 전체 검증
   * @param {FormData} formData - 검증할 폼 데이터
   * @param {object} rules - 검증 규칙 { fieldName: { required: true, type: 'email' | 'phone' | 'number', ... } }
   * @returns {object} { isValid: boolean, errors: { fieldName: message } }
   */
  static validateForm(formData, rules) {
    const errors = {};
    let isValid = true;

    for (const [fieldName, rule] of Object.entries(rules)) {
      const value = formData.get(fieldName);

      // 필수 필드 검증
      if (rule.required) {
        const result = this.validateRequired(fieldName, value);
        if (!result.isValid) {
          errors[fieldName] = result.message;
          isValid = false;
          continue;
        }
      }

      // 타입별 검증
      if (value && value.toString().trim() !== '') {
        let result = { isValid: true, message: '' };

        switch (rule.type) {
          case 'email':
            result = this.validateEmail(value);
            break;
          case 'phone':
            result = this.validatePhone(value);
            break;
          case 'number':
            result = this.validateNumberRange(value, rule.min, rule.max);
            break;
        }

        if (!result.isValid) {
          errors[fieldName] = result.message;
          isValid = false;
        }
      }
    }

    return { isValid, errors };
  }

  /**
   * 에러 메시지 표시
   * @param {HTMLElement} field - 필드 엘리먼트
   * @param {string} message - 에러 메시지
   */
  static showFieldError(field, message) {
    // 기존 에러 메시지 제거
    this.clearFieldError(field);

    // 필드에 에러 클래스 추가
    field.classList.add('is-invalid');

    // 에러 메시지 엘리먼트 생성
    const errorDiv = document.createElement('div');
    errorDiv.className = 'invalid-feedback';
    errorDiv.textContent = message;

    // 필드 바로 다음에 에러 메시지 삽입
    field.parentNode.insertBefore(errorDiv, field.nextSibling);
  }

  /**
   * 에러 메시지 제거
   * @param {HTMLElement} field - 필드 엘리먼트
   */
  static clearFieldError(field) {
    field.classList.remove('is-invalid');

    const errorDiv = field.parentNode.querySelector('.invalid-feedback');
    if (errorDiv) {
      errorDiv.remove();
    }
  }

  /**
   * 폼 전체 에러 메시지 제거
   * @param {HTMLFormElement} form - 폼 엘리먼트
   */
  static clearAllErrors(form) {
    const invalidFields = form.querySelectorAll('.is-invalid');
    invalidFields.forEach(field => this.clearFieldError(field));
  }

  /**
   * 프로젝트 폼 특화 검증 (레거시 로직 기반)
   * @param {object} formDataObj - 폼 데이터 객체
   * @returns {object} { isValid: boolean, errors: object, firstError: string }
   */
  static validateProjectForm(formDataObj) {
    const errors = {};
    let firstError = '';

    // 필수 필드
    const requiredFields = {
      '사업자': formDataObj['company'] || formDataObj['사업자'],
      '담당자': formDataObj['owner'] || formDataObj['담당자'],
      '거래처': formDataObj['client'] || formDataObj['거래처']
    };

    for (const [fieldName, value] of Object.entries(requiredFields)) {
      if (!value || value.trim() === '') {
        errors[fieldName] = `${fieldName}을(를) 선택해주세요.`;
        if (!firstError) firstError = errors[fieldName];
      }
    }

    // 날짜 범위 검증
    const startDate = formDataObj['공사 시작'] || formDataObj['start-date'];
    const endDate = formDataObj['공사 종료'] || formDataObj['end-date'];

    if (startDate && endDate) {
      const dateValidation = this.validateDateRange(startDate, endDate);
      if (!dateValidation.isValid) {
        errors['날짜'] = dateValidation.message;
        if (!firstError) firstError = dateValidation.message;
      }
    }

    // 이메일 검증
    const email = formDataObj['담당자 이메일'] || formDataObj['manager-email'];
    if (email) {
      const emailValidation = this.validateEmail(email);
      if (!emailValidation.isValid) {
        errors['이메일'] = emailValidation.message;
        if (!firstError) firstError = emailValidation.message;
      }
    }

    // 금액 검증 (0 이상)
    const amount = parseFloat(formDataObj['총액 1'] || formDataObj['total-amount'] || 0);
    if (amount < 0) {
      errors['금액'] = '금액은 0 이상이어야 합니다.';
      if (!firstError) firstError = errors['금액'];
    }

    return {
      isValid: Object.keys(errors).length === 0,
      errors,
      firstError
    };
  }
}
