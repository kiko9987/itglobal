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

// 수금완료 자동 판정: 미수금≈0 (총액2>0 & 입금>0 가드) 또는 수금확인 체크(보조).
//   매니저가 수금확인·계산서요청을 깜빡해도 시스템이 "다 받았음"을 스스로 인식 (2026-09-07).
export function isFullyCollected(row) {
  const total2 = toNum(row && row['총액 2']);
  const paidSum = BILL_STAGES.reduce((a, s) => a + toNum(row && row[s]), 0);
  const unpaid = toNum(row && row['미수금']);
  const auto = total2 > 0 && paidSum > 0 && Math.abs(unpaid) < 1;
  return auto || isCollected(row && row['수금 확인']);
}

export function normalizeToken(t) {
  const s = String(t == null ? '' : t).trim();
  if (s === '' || s === '-') return '';   // 빈값·대시(-) = 없음
  if (s === '카드결제') return '카드';
  if (s === '일반') return '발행';         // 레거시 '일반' → '발행'(세금계산서 발행됨)
  if (s === '혼합' || s === '확인필요') return '기타'; // 레거시 혼합·확인필요 → '기타'(특이, 메모 참고)
  return s; // 발행 / N입금 / 카드 / 미발행 / 기타
}

// 마지막 '발행' 단계 인덱스(없으면 -1). '-'가 (a)전체발행 covered와 (b)빈 행 두 의미로
// 겹쳐 쓰이므로, covered는 **마지막 발행 앵커보다 앞 단계**의 '-'만 인정한다.
// (전체발행 한 장=잔금에 발행, 계약금·중도금 '-' covered. 계약금만 발행이면 뒤 중도금·잔금
//  '-'는 빈 행=미발행. 2026-09-16 R4112-SJ 선발행-부분 오인 fix.)
export function lastIssuedIndex(row) {
  let idx = -1;
  BILL_STAGES.forEach((s, i) => {
    const raw = String((row && row[BILL_STAGE_COL[s]]) == null ? '' : row[BILL_STAGE_COL[s]]).trim();
    if (normalizeToken(raw) === '발행') idx = i;
  });
  return idx;
}

/**
 * 단계별 컬럼 우선으로 상태 계산. 컬럼 비었으면 금액>0→미발행.
 * 3열 전부 비었는데 Y에 값 있으면 레거시 Y 파싱 폴백(미마이그레이션·구경로 방어).
 * @returns {{계약금:string, 중도금:string, 잔금:string}}
 */
export function computeBillStagesFromColumns(row) {
  const result = { 계약금: 'none', 중도금: 'none', 잔금: 'none' };
  // 수금완료(미수금0 자동 or 수금확인)면 미발행='미발행'(⚠️ 요청 필요), 진행중이면 '발행예정'🕒(조용)
  const uninvoiced = isFullyCollected(row) ? '미발행' : '발행예정';
  // covered('-')는 그 행에 실제 '발행'이 있을 때만 유효(전체발행 한 장에 포함됨). '발행' 없이
  // '-'만 있으면 빈칸처럼 취급 → 입금 시 미발행으로 잡음 (Y요약 _bill_y_summary와 동일 규칙).
  const rawOf = (s) => String((row && row[BILL_STAGE_COL[s]]) == null ? '' : row[BILL_STAGE_COL[s]]).trim();
  const lastIdx = lastIssuedIndex(row);
  // covered = 마지막 발행 앵커보다 앞 단계의 '-'만 (뒤 '-'는 빈 행=미발행). '-' 겹침 존중.
  const isCovered = (s) => rawOf(s) === '-' && BILL_STAGES.indexOf(s) < lastIdx;
  // 마지막 진행단계(입금 or 계산서 토큰). 수금완료 + 그 단계가 발행이면 전체발행 완료 → 앞 미발행 covered.
  const active = BILL_STAGES.filter((s) => toNum(row && row[s]) > 0 || normalizeToken(rawOf(s)));
  const anchor = active.length ? active[active.length - 1] : null;
  const anchorInvoiced = anchor && (normalizeToken(rawOf(anchor)) === '발행' || isCovered(anchor));
  const fullDone = isFullyCollected(row) && anchorInvoiced;
  let anyCol = false;
  BILL_STAGES.forEach((s) => {
    const raw = rawOf(s);
    const v = normalizeToken(raw);
    // 전체발행 완료면 앞 단계 미발행/빈칸은 covered(none) — ⚠️ 안 뜸
    if (v) { result[s] = (v === '미발행') ? (fullDone ? 'none' : uninvoiced) : v; anyCol = true; }
    else if (isCovered(s)) { anyCol = true; /* covered(마지막 발행 앞) = 발행됨 → none, ⚠️ 아님 */ }
    else if (toNum(row && row[s]) > 0) { result[s] = fullDone ? 'none' : uninvoiced; } // 빈칸+입금: 전체발행완료면 covered, 아니면 미발행
    else if (raw === '-') { anyCol = true; /* 발행없는/뒤쪽 '-'+금액0 = 빈 행 */ }
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
 * 단계별 상태 → Y(계산서) 요약 "{마지막 입금단계} - {상태}" (2026-09-07 개편).
 *   Y = 진행도 요약(어디까지 왔나). 상세는 3열(Z/AA/AB)에. '발행중' 폐기.
 *   앵커 = 금액 입금된 마지막 단계(계약금<중도금<잔금).
 *   상태: 미발행 / 발행완료 / N입금 / 카드결제 / 혼합.  입금 없으면 '-'.
 *     - 입금됐는데 미발행(계산서 없음) 단계 있음 → 미발행 (⚠️ 여부는 아이콘=수금완료 기준)
 *     - 확인필요/현금+카드 섞임 → 혼합
 *     - 발행 있음(그 외 방법도 처리됨) → 발행완료
 *     - 전부 카드 → 카드결제 / 전부 현금 → N입금
 * @param {object} stages - {계약금,중도금,잔금} 단계별 계산서 토큰
 * @param {object} row - 프로젝트 레코드(단계 금액 U/V/W 참조용)
 */
export function computeYSummary(stages, row) {
  const amt = {};
  const rawCol = {};
  BILL_STAGES.forEach((s) => {
    amt[s] = toNum(row && row[s]);
    rawCol[s] = String((row && row[BILL_STAGE_COL[s]]) == null ? '' : row[BILL_STAGE_COL[s]]).trim();
  });
  const vals = {};
  BILL_STAGES.forEach((s) => { vals[s] = normalizeToken(stages && stages[s]); });
  // covered('-') = 전체발행 한 장에 포함되어 발행된 단계 — **마지막 발행 앵커보다 앞 단계만**.
  //   뒤쪽 '-'는 빈 행(미발행). '-' 겹침(covered vs 빈칸) 존중 (2026-09-16 R4112 fix).
  let lastIdx = -1;
  BILL_STAGES.forEach((s, i) => {
    if (vals[s] === '발행' || rawCol[s] === '발행' || rawCol[s] === '일반') lastIdx = i;
  });
  const cov = {};
  BILL_STAGES.forEach((s, i) => { cov[s] = rawCol[s] === '-' && i < lastIdx; });
  // active = 입금됐거나(amt>0) 계산서 토큰이 있는 단계 (입금 전 계산서 선발행 케이스 포함)
  const active = BILL_STAGES.filter((s) => amt[s] > 0 || vals[s]);
  if (!active.length) return '미발행'; // 입금도 계산서도 없음 → 미발행(앵커 없음, ⚠️ 아님)
  const anchor = active[active.length - 1];
  const anchorInvoiced = vals[anchor] === '발행' || cov[anchor];
  // 수금완료 + 마지막 진행단계 발행 = 전체발행 완료로 간주 → 앞 단계 미발행 무시(계약금 미발행 알림 X).
  //   마지막 단계가 아직 미발행이면 전체발행 아님 → 아래 미발행 우선 로직으로 정당하게 노출(단계별 대응).
  if (isFullyCollected(row) && anchorInvoiced) return `${anchor} - 발행완료`;
  // 미발행(입금됐는데 계산서 없음) 우선 노출 — 마지막 미발행 단계 앵커 (알람). covered 제외.
  const uninv = active.filter((s) => amt[s] > 0 && (vals[s] === '' || vals[s] === '미발행') && !cov[s]);
  if (uninv.length) return `${uninv[uninv.length - 1]} - 미발행`;
  // 그 외: 마지막 진행단계(앵커)의 실제 상태 그대로 (covered 앵커 = 발행완료, 잔금 카드면 카드결제 등)
  const m = vals[anchor];
  const status = (m === '발행' || cov[anchor]) ? '발행완료' : m === '카드' ? '카드결제' : m === 'N입금' ? 'N입금' : '기타';
  return `${anchor} - ${status}`;
}

/**
 * 세금계산서 발행 금액 = '발행' 단계의 결제금액 합 (현금/카드는 세금계산서 아니라 제외).
 * @returns {number}
 */
export function computeInvoicedAmount(row) {
  let sum = 0;
  const lastIdx = lastIssuedIndex(row);
  BILL_STAGES.forEach((s, i) => {
    const raw = String((row && row[BILL_STAGE_COL[s]]) == null ? '' : row[BILL_STAGE_COL[s]]).trim();
    const amt = toNum(row && row[s]);
    // '발행' = 이 단계 세금계산서 발행. '-'+금액 = 전체발행 covered(단, 마지막 발행 앞 단계만).
    // 현금(N입금)·카드는 세금계산서 아니라 제외.
    if (normalizeToken(raw) === '발행' || (raw === '-' && amt > 0 && i < lastIdx)) sum += amt;
  });
  return sum;
}

/**
 * 세금계산서 발행 아이콘 툴팁용 금액 줄 — 사업자 + 공급가(VAT별도)/합계(VAT포함).
 * 실제 발행액 = 이 발행이 커버하는 금액. 통합발행이면 바로 앞의 연속된 '-'(covered)
 * 단계 금액까지 합산(예: 잔금 통합발행이 계약금·중도금 포함 → 총액).
 * 테이블(ProjectTable) 과 아코디언(ProjectRowAccordion) 공용 — 툴팁 파리티. @returns {string[]}
 */
export function billAmountLines(row, stage) {
  const lines = [];
  // 최근 발행이 여러 사업자로 나뉘었으면(고객 요청 사업자 분할) 시트 사업자명 대신 실제 발행 사업자들
  //   (G3991-YM: 시트 'SM CORPORATION' 인데 현재 계산서는 설린·SM 두 장, 2026-10-01).
  //   거래처(인테리어 업체) 경유 공사는 시트 사업자명=거래처, 계산서=실제 발주처 → 둘 다 표시
  //   (G4050-MJ: '오쿠드 (거래처 미공개스튜디오)'). 표기만 다른 같은 회사는 시트값 그대로.
  //   분할 발행은 사업자별 금액도 한 줄씩 (메모 '발행|선발행 X원' 줄에서, 금액 없는 옛 줄은 이름만).
  const parts = billIssuedParts(row, stage);
  const sheetBiz = String((row && row['사업자명']) || '').trim();
  const partner = parts.length && sheetBiz && !parts.some((p) => sameBizName(p.name, sheetBiz))
    ? sheetBiz : '';
  const amtOf = (s) => parseFloat((row && row[s]) || 0);
  const tokOf = (s) => String((row && row[`${s} 계산서`]) || '').trim();
  const idx = BILL_STAGES.indexOf(stage);
  let gross = amtOf(stage);
  // 앞 단계 합산(통합발행 커버분). '-'(명시적 covered)는 항상 포함, 금액 0(미발생) 단계는
  // 건너뛰어 체인 유지. 완납(미수금 0 = 1건 총액발행) 프로젝트는 앞의 '미발행'/빈 단계도
  // 이 발행에 포함된 것으로 합산(매니저 제보 2026-09-29 R4091-SJ: 계약금이 '-' 아닌 '미발행'
  // 이라 총액이 아닌 잔금만 표기되던 문제). 별도 발행/현금/카드 단계는 경계로 중단.
  const fully = isFullyCollected(row);
  for (let i = idx - 1; i >= 0; i--) {
    const t = tokOf(BILL_STAGES[i]);
    const a = amtOf(BILL_STAGES[i]);
    if (t === '-') { gross += a; continue; }              // 명시적 covered
    if (a === 0) continue;                                // 미발생 단계 → 건너뜀(체인 유지)
    if ((t === '미발행' || t === '') && fully) { gross += a; continue; }  // 완납 통합발행 포함
    break;                                                // 별도 발행/현금/카드 → 경계
  }
  if (parts.length > 1) {
    // 분할 발행 장별 입금 여부 — 입금 메모의 입금자 ↔ 발행 사업자 대조 (G3991-YM: 09/04 (주)설린 170.5만
    //   입금 = 설린 계산서와 이름·금액 일치 → 설린 입금, SM 대기). 입금자로 맞춘 장들의 합이 실제 입금액과
    //   같을 때만 표시 — 거래처가 대신 냈거나 이름이 안 맞으면 표시 안 함(오표기 방지).
    const paid = parts.map((p) => isPaidByBiz(row, p.name));
    const paidSum = parts.reduce((a, p, i) => a + (paid[i] ? p.amt : 0), 0);
    const showPaid = parts.every((p) => p.amt > 0) && paid.some(Boolean) && Math.abs(paidSum - gross) <= 1;
    lines.push(`분할 발행 ${parts.length}장${partner ? ` (거래처 ${partner})` : ''}`);
    parts.forEach((p, i) => lines.push(`· ${p.name}${p.amt > 0 ? ` ${p.amt.toLocaleString()}원` : ''}`
      + (showPaid ? (paid[i] ? ' — 입금 완료' : ' — 입금 대기') : '')));
  } else if (partner) {
    lines.push(`${parts[0].name} (거래처 ${partner})`);
  } else if (sheetBiz) {
    lines.push(sheetBiz);
  }
  // 선발행(입금 대기): 발행액 = 앞 단계 입금분(통합발행 covered, gross) + 이 단계 미입금분(메모).
  //   메모 'X원'은 자동기록 시 covered 를 뺀 순수분(R4080-JK: 638만 통합 = 입금 58만 + 대기 580만).
  //   더한 값이 총액2 를 넘으면 메모가 이미 전액이었던 것 → 메모 값을 전체 발행액으로.
  if (isPreIssued(row, stage)) {
    const pending = preIssuedAmount(row, stage);
    if (pending > 0) {
      const total2 = toNum(row && row['총액 2']);
      let issuedTotal = gross + pending;
      if (total2 > 0 && issuedTotal > total2 + 1) issuedTotal = pending;
      lines.push(`발행액${parts.length > 1 ? ' 합계' : ''} ${issuedTotal.toLocaleString()}원 (VAT 포함)`);
      if (gross > 0 && issuedTotal > gross) {
        lines.push(`입금 ${gross.toLocaleString()}원 · 입금 대기 ${(issuedTotal - gross).toLocaleString()}원`);
      }
    } else {
      lines.push(gross > 0
        ? `발행액 미상 — 입금 ${gross.toLocaleString()}원 외 계산서 메모 확인`
        : '발행액 미상 — 계산서 메모 확인');
    }
    return lines;
  }
  // 분할 발행(장별 금액 있음) — 결제칸 대신 장 합계 기준 (일부만 입금이면 '입금 · 입금 대기')
  if (parts.length > 1 && parts.every((p) => p.amt > 0)) {
    const issuedSum = parts.reduce((a, p) => a + p.amt, 0);
    lines.push(`발행액 합계 ${issuedSum.toLocaleString()}원 (VAT 포함)`);
    if (gross > 0 && issuedSum > gross + 1) {
      lines.push(`입금 ${gross.toLocaleString()}원 · 입금 대기 ${(issuedSum - gross).toLocaleString()}원`);
    }
    return lines;
  }
  // (ⓑ) 입금 0 단계는 실제 발행액을 결제칸으론 모름 → 총액2 추정 금지.
  //   gross>0(실입금 커버)면 결제칸 기준 금액, 아니면 계산서 메모에서 발행액(preIssuedAmount).
  if (gross > 0) {
    const supply = Math.round(gross / 1.1);     // 공급가 (VAT 별도)
    lines.push(`${stage} ${supply.toLocaleString()}원 (VAT 별도)`);
    lines.push(`합계 ${gross.toLocaleString()}원 (VAT 포함)`);
  } else {
    // 입금 0 이지만 수금 완료(옛 데이터 모양) — 메모에 금액이 있으면 참고로만 표시
    const issuedAmt = preIssuedAmount(row, stage);
    if (issuedAmt > 0) lines.push(`발행액 ${issuedAmt.toLocaleString()}원 (VAT 포함)`);
  }
  return lines;
}

/**
 * 선발행(입금 대기) — 그 단계에 세금계산서 '발행'이 있는데
 *   ① 그 단계 입금이 0 이고  ② 프로젝트에 미수금이 남아 있음(수금 미완료).
 * 상태 기준: 입금이 들어오면(금액>0) 자동으로 '발행완료'. 수금이 끝난 프로젝트(미수금 0·수금확인)는
 * 입금을 다른 단계 칸에 적은 옛 데이터 모양일 뿐이라 선발행 아님 (2026-10-01 정의, 실측: 입금0+발행
 * 61건 중 45건이 이 경우). 일부 입금(발행액 > 입금)은 구분하지 않음 → 발행완료.
 */
export function isPreIssued(row, stage) {
  if (!row || !BILL_STAGE_COL[stage]) return false;
  return normalizeToken(row[BILL_STAGE_COL[stage]]) === '발행'
    && toNum(row[stage]) === 0
    && !isFullyCollected(row);
}

/**
 * 입금 0 단계의 세금계산서 발행액(VAT 포함, 원). 출처 = 계산서_메모(Y 노트):
 *   ① 단계 표시 줄 'YYYY-MM-DD {단계} 선발행 X원 · 사업자명' (계산서 첨부 자동기록, 한 줄 = 한 장)
 *      → 그 단계의 **가장 최근 날짜 줄들의 합** = 현재 발행액. 앞 날짜 줄은 수정발행 전 이력으로 보존.
 *      (G3991-YM: 08-27 310만 한 장 → 09-02 고객 요청 사업자 분할 155만+155만 = 310만)
 *   ② 단계 표시 줄이 없으면 메모의 'X원' 이 정확히 하나 (SB 수기 메모 등, 단계 표시 줄은 제외)
 *   ③ 없거나 여럿이면 0 (어느 단계 금액인지 모호 → 오표기 방지)
 */
export function preIssuedAmount(row, stage) {
  const memo = String((row && row['계산서_메모']) || '');
  if (!memo.trim()) return 0;
  const TAG = /^\s*(?:(\d{4})[-./](\d{1,2})[-./](\d{1,2})\s+)?(계약금|중도금|잔금)\s*선발행\s*([\d,]+)\s*원/;
  const tagged = [];
  memo.split('\n').forEach((ln, i) => {
    const m = TAG.exec(ln);
    if (!m || m[4] !== stage) return;
    const date = m[1] ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : '';
    tagged.push({ date, amt: toNum(m[5]), i });
  });
  if (tagged.length) {
    const latest = tagged.reduce((a, t) => (t.date > a ? t.date : a), '');
    // 날짜 없는 옛 형식만 있으면 마지막 줄 하나
    if (!latest) return tagged[tagged.length - 1].amt;
    return tagged.filter((t) => t.date === latest).reduce((a, t) => a + t.amt, 0);
  }
  // 단계 표시 줄('YYYY-MM-DD 잔금 발행 X원 · …' = 일반 발행 장 금액, 2026-10-01~)은 폴백에서 제외
  const STAGE_LINE = /^\s*(?:\d{4}[-./]\d{1,2}[-./]\d{1,2}\s+)?(계약금|중도금|잔금)\s*선?발행/;
  const rest = memo.split('\n').filter((ln) => !STAGE_LINE.test(ln)).join('\n');
  const amts = rest.match(/[\d,]+\s*원/g) || [];
  return amts.length === 1 ? toNum(amts[0].replace(/원|\s/g, '')) : 0;
}

/**
 * 그 단계 가장 최근 발행일의 발행 사업자별 금액 — 계산서_메모 'YYYY-MM-DD {단계} 발행|선발행 X원 · 사업자명'
 * (2026-10-01~ 자동기록·소급). 줄 끝 '(수정발행: …)' 같은 설명 괄호는 뺌. 사업자 없는 줄은 무시.
 * 같은 사업자 여러 줄은 금액 합산. 금액 없는 일반 '발행' 줄은 amt 0.
 * @returns {{name: string, amt: number}[]} 메모 순서
 */
export function billIssuedParts(row, stage) {
  const memo = String((row && row['계산서_메모']) || '');
  if (!memo.trim()) return [];
  const LINE = /^\s*(\d{4})[-./](\d{1,2})[-./](\d{1,2})\s+(계약금|중도금|잔금)\s*선?발행\s*(?:([\d,]+)\s*원)?\s*·\s*(.+)$/;
  const items = [];
  memo.split('\n').forEach((ln) => {
    const m = LINE.exec(ln);
    if (!m || m[4] !== stage) return;
    const name = m[6].replace(/\s+\([^()]*:[^()]*\)\s*$/, '').trim();
    if (name) {
      items.push({ date: `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`,
        name, amt: m[5] ? toNum(m[5]) : 0 });
    }
  });
  if (!items.length) return [];
  const latest = items.reduce((a, t) => (t.date > a ? t.date : a), '');
  const parts = [];
  items.filter((t) => t.date === latest).forEach((t) => {
    const p = parts.find((x) => x.name === t.name);
    if (p) p.amt += t.amt; else parts.push({ name: t.name, amt: t.amt });
  });
  return parts;
}

/**
 * 같은 회사인지 — 법인 표기((주)·주식회사·㈜·(재) 등)·공백·대소문자 무시, 포함 관계 또는 한 글자 차이까지 같음
 * (G2894-YM 계산서 '엠제이디엔엠' ↔ 시트 '주식회사 엠제이디앤엠' 오탈자 표기).
 */
const normBizName = (s) => String(s || '').toLowerCase()
  .replace(/주식회사|유한회사|재단법인|사단법인|\((?:주|유|재|사)\)|[㈜㈔]/g, '')
  .replace(/[\s.,·()]/g, '');

/**
 * 입금 메모(계약금/중도금/잔금_메모)에 이 사업자 이름의 입금자가 있는지.
 * 메모 형식이 여럿('입금자: X', SMS '…\n입금 X원\n(주)설린\n계좌\n은행', '적요 X')이라 줄 단위로 비교.
 * 은행 SMS 는 입금자명을 잘라 보내므로('(주)삼양발브종') 4자 이상이면 앞부분 일치도 인정. 숫자 있는 줄은 제외.
 */
export function isPaidByBiz(row, name) {
  const n = normBizName(name);
  if (!n) return false;
  return BILL_STAGES.some((s) => String((row && row[`${s}_메모`]) || '').split('\n').some((ln) => {
    if (/\d/.test(ln)) return false;
    const l = normBizName(ln.replace(/^\s*(?:입금자\s*:|적요)\s*/, ''));
    if (!l) return false;
    return l === n || (l.length >= 4 && n.startsWith(l)) || (n.length >= 4 && l.startsWith(n));
  }));
}

export function sameBizName(a, b) {
  const x = normBizName(a), y = normBizName(b);
  if (!x || !y) return true;   // 비교 불가 → 같은 것으로 (괜히 두 이름 표시 안 함)
  if (x === y || x.includes(y) || y.includes(x)) return true;
  if (Math.abs(x.length - y.length) > 1 || Math.min(x.length, y.length) < 4) return false;
  // 편집거리 ≤ 1
  let i = 0, j = 0, diff = 0;
  while (i < x.length && j < y.length) {
    if (x[i] === y[j]) { i++; j++; continue; }
    if (++diff > 1) return false;
    if (x.length > y.length) i++; else if (y.length > x.length) j++; else { i++; j++; }
  }
  return diff + (x.length - i) + (y.length - j) <= 1;
}

/** 계산서(발행) 툴팁 첫 줄 — 선발행이면 입금 대기 표시 */
export function billHeadline(row, stage) {
  return isPreIssued(row, stage) ? '세금계산서 선발행 · 입금 대기' : '세금계산서 발행완료';
}

/**
 * 세금계산서 발행일 — 계산서_메모(Y 노트)에서 추출(툴팁 '발행일' 줄). 없으면 ''.
 * 발행일은 시트에 따로 남는 칸이 없어 메모가 유일한 출처 (2026-09-30 SB 요청).
 *  ① 단계 명시 줄 'YYYY-MM-DD {단계} 발행|선발행 …'(시스템 자동기록·슬랙 카드 소급) — 그 단계 최신 줄.
 *  ② SB 수기 줄 'YYYY-MM-DD … 발행'(단계 표기 없음): ①로 못 정한 발행 단계 수와 줄 수가
 *     같을 때만 단계 순서(계약금→잔금) ↔ 날짜 순서로 대응. 개수 다르면 모호 → 생략(오표기 방지).
 * @returns {string} 'YYYY-MM-DD' | ''
 */
export function billIssueDate(row, stage) {
  const memo = String((row && row['계산서_메모']) || '');
  if (!memo.trim()) return '';
  const DATE_LINE = /^\s*(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?!\d)(.*)$/;
  const tagged = [];
  const untagged = [];
  memo.split('\n').forEach((ln) => {
    const m = DATE_LINE.exec(ln);
    if (!m) return;
    const date = `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
    const st = /(계약금|중도금|잔금)\s*선?발행/.exec(m[4]);
    if (st) tagged.push({ stage: st[1], date });
    else if (/발행/.test(m[4])) untagged.push(date);
  });
  const own = tagged.filter((t) => t.stage === stage);
  if (own.length) return own[own.length - 1].date;
  const taggedStages = new Set(tagged.map((t) => t.stage));
  const remaining = BILL_STAGES.filter((s) => !taggedStages.has(s)
    && normalizeToken(String((row && row[BILL_STAGE_COL[s]]) ?? '').trim()) === '발행');
  const i = remaining.indexOf(stage);
  if (i < 0 || !untagged.length || untagged.length !== remaining.length) return '';
  return [...untagged].sort()[i];
}
