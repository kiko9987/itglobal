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
  '당근', '당근마켓', '숨고', '홈페이지', '중고판매', '큐플레이스', '채널톡', '카카오톡', '메일',
]);
// 부르는 이름에서 빼는 것 — 서버 partner_names.short_name 과 같은 규칙 (10-23 넓힘)
const LEGAL_MARK_RE = new RegExp(
  '\\[\\s*폐업\\s*\\]|주식회사|유한책임회사|유한회사|합자회사|합명회사|재단법인|사단법인|의료법인|학교법인'
  + '|사회복지법인|농업회사법인|\\(\\s*(?:주|유|재|사|의|학|복|합)\\s*\\)|㈜|㈔|㈐', 'g',
);
const PAREN_NO_HANGUL_RE = /\(([^()가-힣]*)\)/g;          // (PK) (suum) (AB STUDI0) (完)
const OPEN_PAREN_TAIL_RE = /\([^()가-힣]*$/;               // 닫히지 않은 꼬리 '(JH ENG Co.,Ltd.'
const FOREIGN_TAIL_RE = /^(.*[가-힣].*?)\s+([^가-힣]+)$/;  // '완디자인 完 Design Interior'
const ROMAN_RE = /^[IVX]+$/;
// 앞이 업종 이름뿐이면 영문 꼬리가 실제 상호 — '건축사사무소 VINZIP'
const GENERIC_HEAD_RE = /(사무소|사무실|컨설팅|법무법인|회계법인|특허법인|세무법인|건축사)$/;
// 건물 관리 단체 꼬리 — '관악산 휴먼시아아파트 입주자대표회의' → '관악산 휴먼시아아파트'
const ORG_SUFFIX_RE = /\s*(입주자\s*대표\s*회의?|관리\s*사무소|관리단|관리\s*위원회|운영\s*위원회)$/;

/** 유입 구분 → 거래처가 아닐 때 거래처명에 들어가는 묶음 (온라인 계열·빈칸은 '온라인') */
export function inflowGroup(inflow) {
  const s = String(inflow ?? '').trim();
  if (PARTNER_GROUP_LABELS.includes(s)) return s;
  return !s || ONLINE_INFLOWS.includes(s) ? '온라인' : '기타';
}

/**
 * 부르는 이름 — '주식회사 메이드바이디자인' → '메이드바이디자인'.
 * 법인 표기·한글 없는 괄호·닫히지 않은 영문 괄호·한글 뒤 영문/한자 꼬리·건물 관리 단체 꼬리를 뺌.
 * 지점·한글 괄호는 남김.
 */
export function shortPartnerName(name) {
  const base = String(name ?? '').replace(LEGAL_MARK_RE, ' ').replace(/\s+/g, ' ').trim();
  let s = base.replace(PAREN_NO_HANGUL_RE, '').replace(OPEN_PAREN_TAIL_RE, '').replace(/\s+/g, ' ').trim();
  const m = FOREIGN_TAIL_RE.exec(s);
  if (m) {
    const tail = m[2].trim();
    if (!ROMAN_RE.test(tail) && !GENERIC_HEAD_RE.test(m[1].trim())
        && ((tail.match(/[A-Za-z]/g) || []).length >= 2 || /[一-鿿]/.test(tail))) {
      s = m[1];
    }
  }
  const t = s.replace(ORG_SUFFIX_RE, '').trim();
  if ((t.match(/[가-힣]/g) || []).length >= 2) s = t;   // '관리단'만 있으면 그대로
  s = s.replace(/^[\s.,\-_/·]+|[\s.,\-_/·]+$/g, '');
  return /[가-힣A-Za-z0-9]/.test(s) ? s : base;
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

/**
 * 이 거래처 공사들에 쓰인 사업자명 — 거래처 자신의 상호 먼저, 그다음 그 거래처를 통해 돈 낸 고객 상호
 * (각각 공사 수 많은 순). 거래처 유입 등록 때 사업자명 목록을 좁히는 데 씀 (2026-10-23 사용자 요청).
 */
export function businessNamesForPartner(partner, rows) {
  const key = partnerPlaceKey(partner);
  if (!key) return [];
  const own = new Map();
  const others = new Map();
  (rows || []).forEach((r) => {
    if (partnerPlaceKey(r?.['거래처명']) !== key) return;
    const biz = String(r?.['사업자명'] ?? '').trim();
    if (!isPartnerName(biz)) return;
    const m = partnerPlaceKey(biz) === key ? own : others;
    m.set(biz, (m.get(biz) || 0) + 1);
  });
  const byCount = (m) => [...m.keys()].sort((a, b) => m.get(b) - m.get(a) || a.localeCompare(b, 'ko'));
  return [...byCount(own), ...byCount(others)];
}

/**
 * 연락처로 기존 거래처명 찾기 — 같은 발주처 연락처를 쓴 공사들의 거래처명이 한 곳으로 모일 때만.
 * 여러 곳으로 갈리면(직원 번호 등) '' (2026-10-23 '북서울꿈의숲' 리드 = 디자인킴스 번호. 거래처 리드 대조:
 * 제안 38건 중 37건 일치, 갈린 8건은 안 채움).
 */
export function suggestPartnerFromPhone(phone, rows) {
  const digits = (s) => String(s ?? '').replace(/\D/g, '');
  const p = digits(phone);
  if (p.length < 9) return '';
  const byKey = new Map();          // 같은 곳 → { 이름: 횟수 }
  (rows || []).forEach((r) => {
    if (digits(r?.['발주처 연락처']) !== p) return;
    const n = String(r?.['거래처명'] ?? '').trim();
    if (!isPartnerName(n)) return;
    const k = partnerPlaceKey(n);
    if (!byKey.has(k)) byKey.set(k, new Map());
    byKey.get(k).set(n, (byKey.get(k).get(n) || 0) + 1);
  });
  if (byKey.size !== 1) return '';
  const names = [...byKey.values()][0];
  return [...names.keys()].sort((a, b) => names.get(b) - names.get(a))[0];
}

/**
 * 자유 입력 이름(방문 현장 고객명 등)에서 기존 거래처명 찾기 — 확실할 때만.
 * ① 같은 곳(띄어쓰기·법인 표기 무시) ② 기존 거래처명(3자 이상)이 그대로 들어 있음('양준혁/코나아이' → 코나아이),
 * 후보가 여럿이면 가장 긴 것, 같은 길이로 갈리면 '' (2026-10-23 거래처 리드 70건 대조: 제안 18건 중 13건 일치,
 * 나머지도 대부분 공사 쪽 거래처명이 고객으로 잘못 들어간 경우). 고객명이 사람·고객이면 '' → 직접 고름.
 */
export function suggestPartnerFromText(text, rows) {
  const counts = new Map();
  (rows || []).forEach((r) => {
    const p = String(r?.['거래처명'] ?? '').trim();
    if (isPartnerName(p)) counts.set(p, (counts.get(p) || 0) + 1);
  });
  const names = [...counts.keys()].sort((a, b) => counts.get(b) - counts.get(a));
  const key = partnerPlaceKey(text);
  if (key) {
    const same = names.find((n) => partnerPlaceKey(n) === key);
    if (same) return same;
  }
  const flat = (s) => shortPartnerName(s).replace(/\s+/g, '').toLowerCase();
  const q = flat(text);
  if (!q) return '';
  const cands = names.filter((n) => flat(n).length >= 3 && q.includes(flat(n)));
  if (!cands.length) return '';
  const best = Math.max(...cands.map((n) => flat(n).length));
  const top = [...new Set(cands.filter((n) => flat(n).length === best).map((n) => partnerPlaceKey(n)))];
  return top.length === 1 ? cands.find((n) => flat(n).length === best) : '';
}
