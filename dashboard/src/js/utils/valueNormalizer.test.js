import logger from '../utils/logger.js';

/**
 * 값 정규화 유틸리티 테스트
 *
 * 실행 방법:
 * node --experimental-modules valueNormalizer.test.js
 *
 * 또는 브라우저 콘솔에서 직접 실행:
 * <script type="module" src="valueNormalizer.test.js"></script>
 */

import {
  normalizeMoneyValue,
  normalizeCheckboxValue,
  normalizeEmptyValue,
  normalizeFieldValue,
  isValueEqual,
  isEmptyValue
} from './valueNormalizer.js';

// 테스트 헬퍼
function assert(condition, message) {
  if (!condition) {
    logger.error(`❌ FAIL: ${message}`);
    throw new Error(message);
  } else {
    logger.debug(`✅ PASS: ${message}`);
  }
}

function assertEqual(actual, expected, testName) {
  assert(
    actual === expected,
    `${testName}: expected "${expected}", got "${actual}"`
  );
}

// 테스트 시작
logger.debug('\n========================================');
logger.debug('값 정규화 유틸리티 테스트');
logger.debug('========================================\n');

// 1. normalizeMoneyValue 테스트
console.group('1. normalizeMoneyValue() 테스트');

assertEqual(normalizeMoneyValue('1,000,000'), '1000000', '콤마 제거');
assertEqual(normalizeMoneyValue('1000000'), '1000000', '콤마 없는 경우');
assertEqual(normalizeMoneyValue(''), '0', '빈 문자열 → 0');
assertEqual(normalizeMoneyValue(null), '0', 'null → 0');
assertEqual(normalizeMoneyValue(undefined), '0', 'undefined → 0');
assertEqual(normalizeMoneyValue('1,000,000원', true), '1000000', '원본: 콤마와 "원" 제거');
assertEqual(normalizeMoneyValue('0'), '0', '0은 그대로');
assertEqual(normalizeMoneyValue('0원', true), '0', '원본: 0원 → 0');

console.groupEnd();

// 2. normalizeCheckboxValue 테스트
console.group('\n2. normalizeCheckboxValue() 테스트');

assertEqual(normalizeCheckboxValue(true), 'true', 'boolean true → "true"');
assertEqual(normalizeCheckboxValue(false), 'false', 'boolean false → "false"');
assertEqual(normalizeCheckboxValue('true'), 'true', '문자열 "true"');
assertEqual(normalizeCheckboxValue('false'), 'false', '문자열 "false"');
assertEqual(normalizeCheckboxValue('TRUE'), 'true', '대문자 TRUE');
assertEqual(normalizeCheckboxValue('FALSE'), 'false', '대문자 FALSE');
assertEqual(normalizeCheckboxValue('포함'), 'true', '포함 → true');
assertEqual(normalizeCheckboxValue('미포함'), 'false', '미포함 → false');
assertEqual(normalizeCheckboxValue(''), 'false', '빈 문자열 → false');
assertEqual(normalizeCheckboxValue(null), 'false', 'null → false');

console.groupEnd();

// 3. normalizeEmptyValue 테스트
console.group('\n3. normalizeEmptyValue() 테스트');

assertEqual(normalizeEmptyValue(null), '', 'null → ""');
assertEqual(normalizeEmptyValue(undefined), '', 'undefined → ""');
assertEqual(normalizeEmptyValue('null'), '', '"null" → ""');
assertEqual(normalizeEmptyValue('undefined'), '', '"undefined" → ""');
assertEqual(normalizeEmptyValue('-'), '', '"-" → ""');
assertEqual(normalizeEmptyValue(''), '', '"" → ""');
assertEqual(normalizeEmptyValue('Hello'), 'Hello', '일반 문자열은 그대로');
assertEqual(normalizeEmptyValue('0'), '0', '0은 빈 값 아님');
assertEqual(normalizeEmptyValue(0), '0', '숫자 0은 문자열로 변환');

console.groupEnd();

// 4. normalizeFieldValue 테스트
console.group('\n4. normalizeFieldValue() 테스트');

// 금액 필드
assertEqual(normalizeFieldValue('1,000,000', 'money'), '1000000', '금액: 콤마 제거');
assertEqual(normalizeFieldValue('', 'money'), '0', '금액: 빈 값 → 0');
assertEqual(normalizeFieldValue('1,000원', 'money', true), '1000', '금액 원본: 원 제거');

// 체크박스 필드
assertEqual(normalizeFieldValue(true, 'checkbox'), 'true', '체크박스: true');
assertEqual(normalizeFieldValue('포함', 'vat'), 'true', '부가세: 포함');
assertEqual(normalizeFieldValue('미포함', 'vat'), 'false', '부가세: 미포함');

// 텍스트 필드
assertEqual(normalizeFieldValue(null, 'text'), '', '텍스트: null');
assertEqual(normalizeFieldValue('-', 'text'), '', '텍스트: -');
assertEqual(normalizeFieldValue('Hello', 'text'), 'Hello', '텍스트: 일반 값');

console.groupEnd();

// 5. isValueEqual 테스트
console.group('\n5. isValueEqual() 테스트');

assert(isValueEqual('1,000,000', '1000000원', 'money'), '금액: 콤마 vs 원 표시');
assert(isValueEqual('', '0', 'money'), '금액: 빈 값 vs 0');
assert(isValueEqual(true, '포함', 'vat'), '부가세: boolean vs 한글');
assert(isValueEqual('true', 'true', 'checkbox'), '체크박스: 문자열 동일');
assert(isValueEqual(null, '-', 'text'), '텍스트: null vs -');
assert(isValueEqual('', undefined, 'text'), '텍스트: 빈 값 vs undefined');
assert(!isValueEqual('Hello', 'World', 'text'), '텍스트: 다른 값');
assert(!isValueEqual('1000', '2000', 'money'), '금액: 다른 값');

console.groupEnd();

// 6. isEmptyValue 테스트
console.group('\n6. isEmptyValue() 테스트');

assert(isEmptyValue(null), 'null은 빈 값');
assert(isEmptyValue(undefined), 'undefined는 빈 값');
assert(isEmptyValue(''), '빈 문자열은 빈 값');
assert(isEmptyValue('-'), '-는 빈 값');
assert(isEmptyValue('null'), '"null"은 빈 값');
assert(isEmptyValue('undefined'), '"undefined"는 빈 값');
assert(!isEmptyValue('0'), '0은 빈 값 아님');
assert(!isEmptyValue('Hello'), '일반 문자열은 빈 값 아님');
assert(!isEmptyValue('false'), '"false"는 빈 값 아님');

console.groupEnd();

// 7. 엣지 케이스 테스트
console.group('\n7. 엣지 케이스 테스트');

// 금액에 특수 문자
assertEqual(normalizeMoneyValue('1,234,567.89'), '1234567.89', '소수점 포함 금액');
assertEqual(normalizeMoneyValue('$1,000'), '$1000', '$는 제거 안 함');

// 체크박스에 예상치 못한 값
assertEqual(normalizeCheckboxValue('yes'), 'false', 'yes → false');
assertEqual(normalizeCheckboxValue('no'), 'false', 'no → false');
assertEqual(normalizeCheckboxValue(1), 'false', '숫자 1 → false');

// 빈 값 경계 케이스
assertEqual(normalizeEmptyValue('  '), '  ', '공백만 있는 경우 그대로 (trim 안 함)');
assertEqual(normalizeEmptyValue('0'), '0', '문자열 0은 빈 값 아님');

console.groupEnd();

logger.debug('\n========================================');
logger.debug('✅ 모든 테스트 통과!');
logger.debug('========================================\n');
