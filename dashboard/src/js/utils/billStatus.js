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

// ─────────────────────────────────────────────────────────────
// (2026-09-07) 단계별 컬럼(Z/AA/AB) = source of truth 로 전환.
//   계약금 계산서 / 중도금 계산서 / 잔금 계산서 각 셀 = 일반/N입금/카드/미발행/혼합/공란.
//   Y(계산서)는 하위호환 롤업으로만 유지.
// ─────────────────────────────────────────────────────────────

export const BILL_STAGE_COL = {
  '계약금': '계약금 계산서',
  '중도금': '중도금 계산서',
  '잔금': '잔금 계산서',
};

// 수금 확인 체크 여부 (수금완료 확정 신호). TRUE/true/1 등 허용.
export function isCollected(v) {
  return v === true || v === 'TRUE' || v === 'true' || v === 1 || v === '1';
}

export function normalizeToken(t) {
  const s = String(t == null ? '' : t).trim();
  if (s === '' || s === '-') return '';   // 빈값·대시(-) = 없음
  if (s === '카드결제') return '카드';
  if (s === '일반') return '발행';         // 레거시 '일반' → '발행'(세금계산서 발행됨)
  if (s === '혼합') return '확인필요';     // 레거시 '혼합' → '확인필요'
  return s; // 발행 / N입금 / 카드 / 미발행 / 확인필요
}

/**
 * 단계별 컬럼 우선으로 상태 계산. 컬럼 비었으면 금액>0→미발행.
 * 3열 전부 비었는데 Y에 값 있으면 레거시 Y 파싱 폴백(미마이그레이션·구경로 방어).
 * @returns {{계약금:string, 중도금:string, 잔금:string}}
 */
export function computeBillStagesFromColumns(row) {
  const result = { 계약금: 'none', 중도금: 'none', 잔금: 'none' };
  // 수금확인 체크(수금완료 확정)면 미발행은 '미발행'(⚠️ 요청 필요), 아니면(진행중) '발행예정'🕒
  const uninvoiced = isCollected(row && row['수금 확인']) ? '미발행' : '발행예정';
  let anyCol = false;
  BILL_STAGES.forEach((s) => {
    const raw = String((row && row[BILL_STAGE_COL[s]]) == null ? '' : row[BILL_STAGE_COL[s]]).trim();
    const v = normalizeToken(raw);
    if (v) { result[s] = (v === '미발행') ? uninvoiced : v; anyCol = true; }
    else if (raw === '-') { anyCol = true; /* 명시적 '-' = 계산서 불필요/전체발행 포함(covered) → none, ⚠️ 아님 */ }
    else if (toNum(row && row[s]) > 0) { result[s] = uninvoiced; }
  });
  if (!anyCol) {
    const y = String((row && row['계산서']) == null ? '' : row['계산서']).trim();
    if (y && y !== '미발행' && y !== '-') {
      return computeBillStages(y, row);
    }
  }
  return result;
}

/**
 * 단계별 상태 → Y(계산서) 롤업 문자열 (하위호환·필터용).
 *   일반/N입금/카드 있는 단계는 "카테고리-단계" 로, 전부 미발행/none 이면 '미발행'.
 */
export function rollupBillStages(stages) {
  const parts = [];
  BILL_STAGES.forEach((s) => {
    const v = normalizeToken(stages && stages[s]);
    if (v && v !== 'none' && v !== '미발행') parts.push(`${v}-${s}`);
  });
  return parts.length ? parts.join(', ') : '미발행';
}

/**
 * 단계별 상태 → Y(계산서) 프로젝트 단위 요약 (2026-09-07).
 * 값: 미발행 / 발행중 / 발행완료 / N입금 / 카드결제 / 확인필요 / '-'(입금없음)
 *   우선순위: 미발행(⚠️) > 확인필요 > 완료도/방법.
 *   - 미발행 단계 있음 → 미발행
 *   - 확인필요(혼합) 있음 → 확인필요
 *   - 처리된 단계 없음 → '-'
 *   - 잔금까지 처리: 발행 있으면 발행완료 / 전부 카드 카드결제 / 전부 현금 N입금 / 그 외 확인필요
 *   - 잔금 미처리(진행중) → 발행중
 */
export function computeYSummary(stages, collected) {
  const isPending = !collected; // 수금확인 미체크 = 진행중
  const vals = BILL_STAGES.map((s) => {
    const v = normalizeToken(stages && stages[s]);
    return (v === '미발행' && isPending) ? '발행예정' : v; // 진행중 미발행 → 발행예정
  });
  if (vals.includes('미발행')) return '미발행';   // 수금완료 미발행만 남음 (요청 필요)
  if (vals.includes('확인필요')) return '확인필요';
  const handled = vals.filter((v) => v === '발행' || v === 'N입금' || v === '카드');
  const hasPending = vals.includes('발행예정');
  if (handled.length === 0 && !hasPending) return '-';
  // 발행(세금계산서)이 있거나 발행예정이면 발행중/발행완료. 순수 현금/카드는 방법 라벨.
  // 발행완료 = 진행중(발행예정) 단계 없음 = 금액 있는 모든 단계 처리됨(미발행·발행예정 없음).
  //   잔금 유무로 판정하면 전액 계약금 등 잔금 금액 없는 건이 발행중으로 오판(2026-09-07 G3721-YG).
  if (vals.includes('발행') || hasPending) {
    return hasPending ? '발행중' : '발행완료';
  }
  if (handled.every((v) => v === '카드')) return '카드결제';
  if (handled.every((v) => v === 'N입금')) return 'N입금';
  return '확인필요'; // 발행 없이 카드+현금 혼재
}

/**
 * 세금계산서 발행 금액 = '발행' 단계의 결제금액 합 (현금/카드는 세금계산서 아니라 제외).
 * @returns {number}
 */
export function computeInvoicedAmount(row) {
  let sum = 0;
  BILL_STAGES.forEach((s) => {
    const raw = String((row && row[BILL_STAGE_COL[s]]) == null ? '' : row[BILL_STAGE_COL[s]]).trim();
    const amt = toNum(row && row[s]);
    // '발행' = 이 단계 세금계산서 발행. '-'+금액 = 전체발행(총액2 요청)에 포함된 단계(covered).
    // 현금(N입금)·카드는 세금계산서 아니라 제외.
    if (normalizeToken(raw) === '발행' || (raw === '-' && amt > 0)) sum += amt;
  });
  return sum;
}
