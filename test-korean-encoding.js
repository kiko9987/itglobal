/**
 * 한글 인코딩 테스트 - valueNormalizer.js
 * Node.js 환경에서 실행
 */

// normalizeCheckboxValue 함수 복사
function normalizeCheckboxValue(value) {
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

// 테스트 케이스
const testCases = [
  { input: true, expected: 'true', desc: 'boolean true' },
  { input: false, expected: 'false', desc: 'boolean false' },
  { input: 'true', expected: 'true', desc: "string 'true'" },
  { input: 'false', expected: 'false', desc: "string 'false'" },
  { input: '포함', expected: 'true', desc: "string '포함' (Korean)" },
  { input: '미포함', expected: 'false', desc: "string '미포함' (Korean)" },
  { input: 'True', expected: 'true', desc: "string 'True' (uppercase)" },
  { input: 'FALSE', expected: 'false', desc: "string 'FALSE' (uppercase)" },
  { input: '포함'.toUpperCase(), expected: 'false', desc: "string '포함'.toUpperCase() (should fail)" },
];

console.log('\n=== 한글 인코딩 테스트 - valueNormalizer.js ===\n');

let passCount = 0;
let failCount = 0;

testCases.forEach(({ input, expected, desc }, index) => {
  const actual = normalizeCheckboxValue(input);
  const match = actual === expected;

  if (match) {
    passCount++;
    console.log(`✅ PASS: ${desc}`);
  } else {
    failCount++;
    console.log(`❌ FAIL: ${desc}`);
  }

  console.log(`   입력: ${JSON.stringify(input)} (타입: ${typeof input})`);
  console.log(`   기대: ${JSON.stringify(expected)}`);
  console.log(`   실제: ${JSON.stringify(actual)}`);

  // 한글 문자열의 경우 바이트 표시
  if (typeof input === 'string' && /[\u3131-\uD79D]/.test(input)) {
    const bytes = Buffer.from(input, 'utf8')
      .toString('hex')
      .match(/.{2}/g)
      .join(' ');
    console.log(`   바이트: ${bytes}`);
  }

  console.log('');
});

console.log(`\n=== 테스트 결과 ===`);
console.log(`통과: ${passCount}/${testCases.length}`);
console.log(`실패: ${failCount}/${testCases.length}`);

if (failCount === 0) {
  console.log('\n🎉 모든 테스트 통과! 한글 인코딩이 정상적으로 동작합니다.\n');
} else {
  console.log('\n⚠️  일부 테스트 실패. 한글 인코딩을 확인해주세요.\n');
}

// 한글 문자열 바이트 비교
console.log('\n=== 한글 문자열 바이트 검증 ===\n');

const koreanStrings = [
  { str: '포함', expected: 'ed 8f ac ed 95 a8' },
  { str: '미포함', expected: 'eb af b8 ed 8f ac ed 95 a8' },
];

koreanStrings.forEach(({ str, expected }) => {
  const actual = Buffer.from(str, 'utf8')
    .toString('hex')
    .match(/.{2}/g)
    .join(' ');

  const match = actual === expected;
  const status = match ? '✅' : '❌';

  console.log(`${status} "${str}"`);
  console.log(`   기대: ${expected}`);
  console.log(`   실제: ${actual}`);
  console.log('');
});

process.exit(failCount > 0 ? 1 : 0);
