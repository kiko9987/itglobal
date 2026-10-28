/**
 * 직원 명단 한곳 (2026-10-28 3차) — GET /api/staff-directory (서버 staff_directory.py).
 *
 * 예전엔 담당자 접미사 표·비영업 제외 목록이 화면 곳곳(펼침·새 공사 창·통계·코드 미리보기)에 따로 박혀 있어
 * 입·퇴사 때마다 빠지는 곳이 생겼다. 이제 서버 한 곳에서 받아 쓴다. 서버를 못 부르면 아래 예비값으로 동작.
 */
import logger from './logger.js';

// 예비값 — 서버 응답이 없을 때만 (2026-10-28 기준 설정과 같음)
const FALLBACK = {
  suffix_map: {
    박정우: 'JW', 강성환: 'SH', 박용구: 'YG', 박민우: 'MW', 이근혁: 'GH', 아이티: 'IT', 김단이: 'DN',
    권태훈: 'TH', 주영민: 'YM', 빈승정: 'SJ', 심장원: 'SJW', 박민재: 'MJ', 조성헌: 'JSH', 황해승: 'HS',
    강민석: 'MS', 황샛별: 'SB', 강정권: 'JK', 이상덕: 'SD', 김지훈: 'JH', 이미영: 'MY', 김호중: 'HJ',
  },
  company_prefix_map: { 글로벌: 'G', 글로벌그룹: 'R', 플렌트: 'P' },
  non_sales: ['김단이', '심장원', '아이티', '이근혁', '황샛별', '이미영', '황해승'],
  inactive: [],
  active_sales: null,   // null = 모름 → 예비 규칙(비영업만 뺌)
};

let _dir = { ...FALLBACK };
let _promise = null;

/** 한 번만 불러 옴 (여러 곳에서 불러도 같은 요청) */
export function loadStaffDirectory() {
  if (_promise) return _promise;
  _promise = fetch('/api/staff-directory', { credentials: 'same-origin' })
    .then((r) => (r.ok ? r.json() : null))
    .then((d) => {
      if (d && d.ok) {
        _dir = {
          suffix_map: d.suffix_map || FALLBACK.suffix_map,
          company_prefix_map: d.company_prefix_map || FALLBACK.company_prefix_map,
          non_sales: d.non_sales || FALLBACK.non_sales,
          inactive: d.inactive || [],
          active_sales: Array.isArray(d.active_sales) ? d.active_sales : null,
        };
      }
      return _dir;
    })
    .catch((e) => {
      logger.warn('[STAFF] 직원 명단 불러오기 실패 — 예비값 사용', e);
      return _dir;
    });
  return _promise;
}

/** 담당자 이름 → 접미사('JW', '-' 없음). 없으면 '' */
export function ownerSuffix(name) {
  return (_dir.suffix_map && _dir.suffix_map[String(name || '').trim()]) || '';
}

/** 내부 사업자 → 코드 앞 글자. 없으면 '' */
export function companyPrefix(company) {
  return (_dir.company_prefix_map && _dir.company_prefix_map[String(company || '').trim()]) || '';
}

/** 영업 담당자 목록·통계에서 빼는 이름인가 (비영업 + 퇴사·비활성) */
export function isNonSales(name) {
  const n = String(name || '').trim();
  return _dir.non_sales.includes(n) || _dir.inactive.includes(n);
}

/** 매출 통계에서만 빼는 이름 (비영업). 퇴사자는 옛 매출이 있으니 통계엔 남김 — 예전 규칙과 같음 */
export function isStatsExcluded(name) {
  return _dir.non_sales.includes(String(name || '').trim());
}

/**
 * 지금 공사를 맡길 수 있는 영업 담당자. 서버 값이 없으면 candidates 에서 비영업·퇴사만 뺀다.
 * @param {string[]} candidates 예비 후보(화면 표에 있는 이름 등)
 */
export function activeSalesManagers(candidates = []) {
  if (Array.isArray(_dir.active_sales) && _dir.active_sales.length) return [..._dir.active_sales];
  return Array.from(new Set(candidates)).filter((n) => n && !isNonSales(n));
}

// 페이지가 뜨면 바로 불러 둠 (첫 사용 때 이미 와 있도록)
if (typeof window !== 'undefined') loadStaffDirectory();
