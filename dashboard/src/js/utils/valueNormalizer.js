/**
 * 값 정규화 유틸리티
 * collectAllChanges()의 분산된 정규화 로직을 통합 관리
 */

/**
 * 금액 필드 값 정규화
 * @param {string} value - 입력값 (콤마 포함 가능)
 * @param {boolean} isOriginal - 원본 값 여부 (원본이면 "원" 제거도 수행)
 * @returns {string} 정규화된 값
 */
export function normalizeMoneyValue(value, isOriginal = false) {
  if (value === null || value === undefined) {
    return '0';
  }

  let normalized = String(value).replace(/,/g, '');

  // 원본 값이면 "원" 제거
  if (isOriginal) {
    normalized = normalized.replace(/원/g, '');
  }

  // 빈 문자열이면 0으로 변환 (구글 시트 수식 호환)
  if (normalized === '') {
    return '0';
  }

  return normalized;
}

/**
 * 체크박스 값 정규화 (부가세, 수금 확인)
 * @param {boolean|string} value - 체크박스 상태 또는 문자열
 * @returns {string} "true" 또는 "false"
 */
export function normalizeCheckboxValue(value) {
  if (typeof value === 'boolean') {
    return value ? 'true' : 'false';
  }

  if (typeof value === 'string') {
    const lower = value.toLowerCase();
    if (lower === 'true' || lower === '포함') {
      return 'true';
    }
    if (lower === 'false' || lower === '미포함') {
      return 'false';
    }
  }

  return 'false';
}

/**
 * 빈 값 정규화
 * null, undefined, 'null', 'undefined', '-' → 빈 문자열
 * @param {any} value - 입력값
 * @returns {string} 정규화된 값
 */
export function normalizeEmptyValue(value) {
  if (
    value === null ||
    value === undefined ||
    value === 'null' ||
    value === 'undefined' ||
    value === '-' ||
    value === ''
  ) {
    return '';
  }

  return String(value);
}

/**
 * 필드 타입에 따른 값 정규화
 * @param {any} value - 입력값
 * @param {string} fieldType - 필드 타입 ('money', 'checkbox', 'vat', 'collection', 'text')
 * @param {boolean} isOriginal - 원본 값 여부
 * @returns {string} 정규화된 값
 */
export function normalizeFieldValue(value, fieldType, isOriginal = false) {
  switch (fieldType) {
    case 'money':
      return normalizeMoneyValue(value, isOriginal);

    case 'checkbox':
    case 'vat':
    case 'collection':
      return normalizeCheckboxValue(value);

    case 'text':
    default:
      return normalizeEmptyValue(value);
  }
}

/**
 * 두 값이 정규화 후 동일한지 비교
 * @param {any} value1 - 비교할 값 1
 * @param {any} value2 - 비교할 값 2
 * @param {string} fieldType - 필드 타입
 * @returns {boolean} 동일 여부
 */
export function isValueEqual(value1, value2, fieldType = 'text') {
  const normalized1 = normalizeFieldValue(value1, fieldType, false);
  const normalized2 = normalizeFieldValue(value2, fieldType, true); // 원본으로 간주

  return normalized1 === normalized2;
}

/**
 * 값이 비어있는지 확인 (정규화 후)
 * @param {any} value - 확인할 값
 * @returns {boolean} 비어있는지 여부
 */
export function isEmptyValue(value) {
  const normalized = normalizeEmptyValue(value);
  return normalized === '';
}
