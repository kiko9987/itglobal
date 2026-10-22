/**
 * 거래처명(공사 현황 '거래처명' 열) 규칙 — 서버 dashboard/services/partner_names.py 미러 (2026-10-22).
 *
 * - 거래처명: 일을 맡긴 곳. 거래처(인테리어 등)거나 두 번 이상 공사한 곳이면 이름, 아니면 유입 묶음
 *   (온라인·전화·소개·중고판매·기타). 묶음 값은 이름이 아니므로 이름처럼 묶거나 고르게 하면 안 됨.
 * - 사업자명: 사업자등록증 상호 — 계산서 기준.
 */

export const PARTNER_GROUP_LABELS = ['온라인', '전화', '소개', '중고판매', '기타'];
const ONLINE_INFLOWS = ['온라인', '당근', '홈페이지', '숨고', '큐플레이스', '카카오톡', '채널톡', '메일'];
const NOT_NAMES = new Set([...PARTNER_GROUP_LABELS, ...ONLINE_INFLOWS, '거래처', '-', '']);
// business_license_handler._PLACEHOLDER_BIZ 와 같은 값 (이름으로 안 봄)
const PLACEHOLDER = new Set([
  '온라인', '거래처', '전화', '소개', '기타', '카톡', '카카오', '방문',
  '미정', '개인', '없음', 'none', 'null', 'na', 'n/a',
]);
const LEGAL_MARK_RE = /\[\s*폐업\s*\]|주식회사|\(주\)|㈜|유한회사|\(유\)/g;

/** 유입 구분 → 거래처가 아닐 때 거래처명에 들어가는 묶음 (온라인 계열·빈칸은 '온라인') */
export function inflowGroup(inflow) {
  const s = String(inflow ?? '').trim();
  if (PARTNER_GROUP_LABELS.includes(s)) return s;
  return !s || ONLINE_INFLOWS.includes(s) ? '온라인' : '기타';
}

/** 부르는 이름 — '주식회사 메이드바이디자인' → '메이드바이디자인' */
export function shortPartnerName(name) {
  return String(name ?? '').replace(LEGAL_MARK_RE, ' ').replace(/\s+/g, ' ').trim();
}

function normName(s) {
  const n = String(s ?? '').replace(/（/g, '(').replace(/）/g, ')').replace(/\s+/g, '').toLowerCase();
  return PLACEHOLDER.has(n) || n.length <= 1 ? '' : n;
}

/** 실제 거래처 이름인가 (온라인·전화 같은 분류값·'-'·자리표시가 아니면 true) */
export function isPartnerName(value) {
  const s = shortPartnerName(value);
  return !!s && !NOT_NAMES.has(s) && !!normName(s);
}

/** 같은 곳 판단 키 (띄어쓰기·대소문자 무시). 이름이 아니면 '' */
export function partnerPlaceKey(name) {
  return isPartnerName(name) ? normName(shortPartnerName(name)) : '';
}

/** 공사를 부르는 이름 — 거래처명이 이름이면 그것, 아니면 사업자명 */
export function partnerDisplayName(row) {
  const p = String(row?.['거래처명'] ?? '').trim();
  if (isPartnerName(p)) return p;
  const b = String(row?.['사업자명'] ?? '').trim();
  return b && b !== '-' ? b : '-';
}

/** 이 사업자명과 같은 곳의 기존 공사가 쓰는 거래처 이름 (없으면 '') — 서버 place_name 의 간단판 */
export function suggestPartnerForBiz(biz, rows) {
  const key = partnerPlaceKey(biz);
  if (!key) return '';
  const counts = new Map();
  (rows || []).forEach((r) => {
    if (partnerPlaceKey(r?.['사업자명']) !== key) return;
    const p = String(r?.['거래처명'] ?? '').trim();
    if (isPartnerName(p) && partnerPlaceKey(p) === key) counts.set(p, (counts.get(p) || 0) + 1);
  });
  let best = '';
  let bestN = 0;
  counts.forEach((n, p) => { if (n > bestN) { best = p; bestN = n; } });
  return best;
}
