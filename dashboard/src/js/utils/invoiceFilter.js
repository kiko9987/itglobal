/**
 * 프로젝트 목록 「세금계산서」 거르기 규칙 — 목록 필터와 위쪽 상태 칸이 함께 쓴다(숫자와 목록이 어긋나지 않게).
 *
 * 시트 Y열(계산서) = "{단계} - {상태}" 또는 바로 "{상태}". 상태 ∈ 미발행/발행완료/N입금/카드결제/기타.
 * Y열은 아무 일도 없으면 처음부터 「미발행」이라, 입금 전인 공사까지 섞인다(2026-11-10 기준 119건 중 93건).
 * → 「입금된 미발행」(PAID_UNISSUED) = 미발행이면서 계약금·중도금·잔금 중 들어온 돈이 있는 공사만.
 *   현금·N통장 건은 Y가 'N입금'이라 원래 안 섞인다.
 */
export const PAID_UNISSUED = '미발행·입금';
export const PAID_UNISSUED_LABEL = '미발행 · 입금된 건';

const STAGES = ['계약금', '중도금', '잔금'];

function toNum(v) {
  const n = parseFloat(String(v == null ? '' : v).replace(/[,원\s]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

export function invoiceStatusOf(item) {
  return String((item && item['계산서']) || '').trim().split(' - ').pop().trim();
}

export function paidSum(item) {
  return STAGES.reduce((a, s) => a + toNum(item && item[s]), 0);
}

/** 필터 값 하나에 맞는 공사인지. value 가 비면 전부 통과. */
export function matchesInvoice(item, value) {
  if (!value) return true;
  const st = invoiceStatusOf(item);
  if (value === PAID_UNISSUED) return st === '미발행' && paidSum(item) > 0;
  return st === value;
}
