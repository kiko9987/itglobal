/**
 * 계산서(세금계산서) 단계별 상태 계산 — 테이블·아코디언 공용.
 *
 * (2026-09-06) 단일 문자열이 못 담던 "단계 × 결제방법 × 발행여부"를 단계별로 환원.
 * 금액과 결합해 "입금됐는데 계산서 없음(미발행)"을 ⚠️로 잡아내는 게 핵심.
 *
 * 입력값(계산서 Y열)을 세 형식 모두 처리:
 *  - 하이픈 조합(명시적, 신규 편집기): "일반-계약금, 카드-잔금"
 *  - 레거시 단일 stage 토큰: 계약금/중도금/잔금 = "그 단계까지 세금계산서(일반)" 누적
 *  - 레거시 결제방법: N입금(현금)/카드결제 = 금액 있는 단계에 방법 표시(계산서 불필요)
 *  - 레거시 혼합: 단계별 방법 불명 → 금액 있는 단계에 '혼합'(확인 필요)
 *  - 미발행/공란: 금액 있는 단계는 '미발행'(입금됐으나 계산서 없음 경고)
 *
 * 반환: { 계약금, 중도금, 잔금 } 각 값 ∈
 *   '일반' | 'N입금' | '카드' | '미발행' | '혼합' | 'none'
 *   ('none' = 표시 없음: 금액도 없고 발행정보도 없음)
 */

export const BILL_STAGES = ['계약금', '중도금', '잔금'];

function toNum(v) {
  const n = parseFloat(String(v == null ? '' : v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function normalizeCategory(cat) {
  const c = String(cat || '').trim();
  if (c === '카드결제') return '카드';
  if (!c) return '일반';
  return c; // 일반 / N입금 / 카드 / 혼합
}

/**
 * @param {string} billValue - 계산서 Y열 값
 * @param {object} amounts - { 계약금, 중도금, 잔금 } 금액 (문자열/숫자 허용)
 * @returns {{계약금:string, 중도금:string, 잔금:string}}
 */
export function computeBillStages(billValue, amounts) {
  const result = { 계약금: 'none', 중도금: 'none', 잔금: 'none' };
  const amt = {
    계약금: toNum(amounts && amounts['계약금']),
    중도금: toNum(amounts && amounts['중도금']),
    잔금: toNum(amounts && amounts['잔금']),
  };
  const hasAmt = (s) => amt[s] > 0;
  const v = String(billValue == null ? '' : billValue).trim();

  // 미발행 / 공란 → 금액 있는 단계는 미발행 경고
  if (!v || v === '-' || v === '미발행') {
    BILL_STAGES.forEach((s) => { if (hasAmt(s)) result[s] = '미발행'; });
    return result;
  }

  // 하이픈 조합(명시적) — 각 단계에 명시된 카테고리, 명시 안 됐는데 금액 있으면 미발행
  if (v.includes('-')) {
    const explicit = {};
    v.split(',').map((x) => x.trim()).filter(Boolean).forEach((item) => {
      if (item.includes('-')) {
        const parts = item.split('-');
        const stage = parts[1] && parts[1].trim();
        if (stage) explicit[stage] = normalizeCategory(parts[0]);
      }
    });
    BILL_STAGES.forEach((s) => {
      if (explicit[s]) result[s] = explicit[s];
      else if (hasAmt(s)) result[s] = '미발행';
    });
    return result;
  }

  // 레거시 단일 stage → 누적 일반 (그 단계까지 세금계산서), 이후 금액 있으면 미발행
  if (BILL_STAGES.includes(v)) {
    const idx = BILL_STAGES.indexOf(v);
    BILL_STAGES.forEach((s, i) => {
      if (i <= idx) result[s] = '일반';
      else if (hasAmt(s)) result[s] = '미발행';
    });
    return result;
  }

  // 레거시 결제방법 — 금액 있는 단계에 방법 표시 (현금·카드는 계산서 불필요)
  if (v === 'N입금') {
    BILL_STAGES.forEach((s) => { if (hasAmt(s)) result[s] = 'N입금'; });
    return result;
  }
  if (v === '카드결제') {
    BILL_STAGES.forEach((s) => { if (hasAmt(s)) result[s] = '카드'; });
    return result;
  }

  // 레거시 혼합 — 단계별 방법 불명 → 금액 있는 단계에 '확인 필요'
  if (v === '혼합') {
    BILL_STAGES.forEach((s) => { if (hasAmt(s)) result[s] = '혼합'; });
    return result;
  }

  return result; // 미지 토큰 → 표시 없음
}
