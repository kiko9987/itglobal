/**
 * 나눠 내는 공사(부담자 여러 명) — '계약금/중도금/잔금 부담자' 칸 해석 (2026-10-19 오남프라자).
 *
 * 파이썬 dashboard/services/payer_split.py 의 미러 — 규칙 바꾸면 같이 바꿀 것.
 * 칸 형식: '임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서' (구분 기호 '·', 첫 토큰=구분, 순서 자유)
 *  - 금액 = 부담액(부가세 포함 실제로 받을 돈). 사업자번호·이메일·결제 방식(계산서/현금/카드)은 선택.
 *  - 구분 기호가 없거나 이름·금액이 없으면 부담자 아님 → 옛 '입금자명' 값은 무시.
 */

export const PAYER_STAGES = ['계약금', '중도금', '잔금'];
export const PAYER_FIELD = { 계약금: '계약금 부담자', 중도금: '중도금 부담자', 잔금: '잔금 부담자' };
// 칸 배정 순서(줄 순서): 첫째 = 계약금 칸, 둘째 = 잔금 칸, 셋째 = 중도금 칸 — 화면 표시 순서도 이것
export const PAYER_ORDER = ['계약금', '잔금', '중도금'];
// 차수(2026-10-20) — 한 사람이 여러 번 나눠 낼 때 몇 번째 돈인지. 입금 메모 블록 안 '◇ 차수 중도금' 줄.
export const INSTALLMENTS = ['계약금', '중도금', '잔금', '일시불'];
export const PAYER_NOTE_PREFIX = '◆';
export const INSTALLMENT_PREFIX = '◇';
const INST_RE = /^\s*◇\s*차수\s*(계약금|중도금|잔금|일시불)/;

const SEP_RE = /\s*[·|｜]\s*/;
const BIZNO_RE = /^(\d{3})-?(\d{2})-?(\d{5})$/;
const AMT_RE = /^(?:₩\s*)?(\d[\d,]*)\s*원?$/;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const CASH_WORDS = new Set(['현금', '계산서x', '계산서없음', '계산서미발행', 'n입금', 'x']);
const METHOD_WORDS = { 계산서: '계산서', 세금계산서: '계산서', 카드: '카드', 미정: '미정' };
// 현금을 부가세 없이 받는 사람 표시(2026-10-21) — 칸 금액은 받을 돈 그대로, 부가세 별도 화면에서 1.1로 안 나눔
export const VAT_FREE_TOKEN = '부가세 없음';
const VAT_FREE_WORDS = new Set(['부가세없음', 'vat없음']);

const toNum = (v) => {
  const s = String(v == null ? '' : v).replace(/[^\d.-]/g, '');
  const n = parseFloat(s);
  return Number.isFinite(n) ? Math.trunc(n) : 0;
};

function methodOf(tok) {
  const t = String(tok || '').replace(/\s+/g, '').toLowerCase();
  if (CASH_WORDS.has(t)) return '현금';
  return METHOD_WORDS[t] || '';
}

/** 부담자 칸 → {role, name, bizNo, email, amount, method, vatFree, note} 또는 null */
export function parsePayerCell(text) {
  const s = String(text == null ? '' : text).trim();
  if (!s || s === '-' || !SEP_RE.test(s)) return null;
  const toks = s.split(SEP_RE).map((t) => t.trim()).filter(Boolean);
  if (toks.length < 3) return null;
  const role = toks[0];
  if (role.length > 12 || AMT_RE.test(role) || BIZNO_RE.test(role) || role.includes('@') || methodOf(role)) return null;
  let name = ''; let bizNo = ''; let email = ''; let method = ''; let amount = 0; let vatFree = false;
  const notes = [];
  toks.slice(1).forEach((t) => {
    const mBiz = t.match(BIZNO_RE);
    const mAmt = t.match(AMT_RE);
    if (!bizNo && mBiz) bizNo = `${mBiz[1]}-${mBiz[2]}-${mBiz[3]}`;
    else if (!email && EMAIL_RE.test(t)) email = t;
    else if (!amount && mAmt && toNum(mAmt[1]) >= 1000) amount = toNum(mAmt[1]);
    else if (!method && methodOf(t)) method = methodOf(t);
    else if (!vatFree && VAT_FREE_WORDS.has(t.replace(/\s+/g, '').toLowerCase())) vatFree = true;
    else if (!name) name = t;
    else notes.push(t);
  });
  if (!name || amount <= 0) return null;
  const m = method || '계산서';
  return { role, name, bizNo, email, amount, method: m, vatFree: vatFree && m === '현금', note: notes.join(' · ') };
}

/** 프로젝트 행 → {단계: 부담자(+received, remaining)} */
export function stagePayers(row) {
  const out = {};
  PAYER_STAGES.forEach((s) => {
    const p = parsePayerCell((row || {})[PAYER_FIELD[s]]);
    if (p) {
      const got = toNum((row || {})[s]);
      out[s] = { ...p, stage: s, received: got, remaining: p.amount - got };
    }
  });
  return out;
}

export function getStagePayer(row, stage) {
  return stagePayers(row)[stage] || null;
}

/** '계약금' → '계약금(임대인)' (부담자 없으면 그대로) */
export function stageLabel(row, stage) {
  const p = getStagePayer(row, stage);
  return p ? `${stage}(${p.role})` : stage;
}

/** '임대인 이희승 · 부담 6,887,100 · 남음 1,377,420' (툴팁 한 줄, 부담자 없으면 '') */
export function payerLine(row, stage) {
  const p = getStagePayer(row, stage);
  if (!p) return '';
  const tail = p.remaining <= 0 ? '완납' : `남음 ${p.remaining.toLocaleString()}`;
  const vf = p.vatFree ? ' · 현금·부가세 없음' : '';   // payer_split.payer_line 미러 (2026-10-21)
  return `${p.role} ${p.name} · 부담 ${p.amount.toLocaleString()}${vf} · ${tail}`;
}

export function isSplit(row) {
  return Object.keys(stagePayers(row)).length > 0;
}

/**
 * 화면에 보일 사업자명 — 나눠 내는 공사면 '임대인 최영희 / 임차인 기종원(위너스골프)', 아니면 시트 사업자명.
 * 시트 칸은 한 상호 기준(등록증·입금 대조)이라 그대로 (payer_split.biz_display 미러, 2026-10-28)
 */
export function bizDisplay(row) {
  const payers = stagePayers(row);
  const people = PAYER_STAGES.filter((s) => payers[s]).map((s) => payerTitle(payers[s]));
  return people.length ? people.join(' / ') : String((row || {})['사업자명'] ?? '').trim();
}

/** 부가세 별도 공사인가 — '부가세' 칸, 없으면 총액 2 > 총액 1 (payer_split.vat_flag 미러) */
export function vatFlag(row) {
  const r = row || {};
  const v = r['부가세'];
  if (v == null || v === '' || (typeof v === 'number' && Number.isNaN(v))) {
    const t1 = toNum(r['총액 1']); const t2 = toNum(r['총액 2']);
    return t1 > 0 && t2 > t1;
  }
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v === 1;
  return ['포함', 'true', 'yes', '1', 'vat 포함', 'y'].includes(String(v).trim().toLowerCase());
}

/** 부가세 별도 공사에서 부가세 없이 받는 몫 합('현금 · 부가세 없음', payer_split.vat_free_amount 미러) */
export function vatFreeAmount(row) {
  if (!vatFlag(row)) return 0;
  return Object.values(stagePayers(row)).filter((p) => p.vatFree).reduce((a, p) => a + p.amount, 0);
}

/**
 * 총액 2 — 부가세 절사 + 끝자리 1/9원 보정, 부가세 없는 몫은 그대로 (project_service._calculate_total2 미러).
 * 305호: calcTotal2(16,042,000, true, 6,000,000) = 6,000,000 + 11,046,200 = 17,046,200
 */
export function calcTotal2(total1, vatSep, vatFree = 0) {
  const t1 = Number(total1) || 0;
  const std = (x) => {
    if (!x) return 0;
    const base = x + Math.floor(x * 0.1);
    const d = base % 10;
    return base + (d === 1 ? -1 : d === 9 ? 1 : 0);
  };
  if (!vatSep) return t1;
  const vf = Math.min(Math.max(Number(vatFree) || 0, 0), t1);
  return vf + std(t1 - vf);
}

/** '임대인 이희승' (부담자 없으면 단계 이름) */
export function payerTitle(payer, stage = '') {
  return payer ? `${payer.role} ${payer.name}` : stage;
}

/** 입금 메모 블록의 차수 줄 '◇ 차수 중도금' (모르는 값이면 '') */
export function installmentLine(inst) {
  return INSTALLMENTS.includes(inst) ? `${INSTALLMENT_PREFIX} 차수 ${inst}` : '';
}

/** 메모 블록(여러 줄)에서 차수 '중도금' (없으면 '') */
export function parseInstallment(text) {
  for (const ln of String(text || '').split('\n')) {
    const m = INST_RE.exec(ln);
    if (m) return m[1];
  }
  return '';
}

/** 칸 메모에서 맨 위 '◆ 부담자' 줄(과 바로 아래 빈 줄)을 뺀 나머지 — 메모 편집·'메모 있음' 판단용 */
export function stripPayerLine(memo) {
  const body = String(memo || '').split('\n').filter((l) => !l.trim().startsWith(PAYER_NOTE_PREFIX));
  while (body.length && !body[0].trim()) body.shift();
  return body.join('\n').replace(/\s+$/, '');
}

/** 칸 메모의 맨 위 '◆' 줄을 line 하나로 맞춤(line='' 이면 지움) — payer_split.apply_note_line 미러 */
export function applyPayerLine(memo, line) {
  const text = stripPayerLine(memo);
  if (!line) return text;
  return text ? `${line}\n\n${text}` : line;
}

function billTok(t) {
  const s = String(t == null ? '' : t).trim();
  if (s === '' || s === '-') return '';
  return { 카드결제: '카드', 일반: '발행', 혼합: '기타', 확인필요: '기타' }[s] || s;
}

/**
 * 한 사람(칸)의 계산서 상태 — 그 칸에 발행한 금액 ↔ 부담액·받은 돈 (payer_split.payer_bill_state 미러).
 * state: issued(preissued=덜 받음) / partial(to_issue, need=⚠️) / needed(⚠️) / pending / cash / card / other / none
 */
export function payerBillState(payer, token, issuedAmt = 0, hasLines = false) {
  const amount = Number(payer.amount) || 0;
  const received = Number(payer.received) || 0;
  const remaining = amount - received;
  const tok = billTok(token);
  const base = { state: 'none', issued: 0, toIssue: 0, preissued: false, need: false, pending: 0 };
  if (tok === '카드') return { ...base, state: 'card' };
  if (tok === 'N입금') return { ...base, state: 'cash' };
  if (tok === '기타') return { ...base, state: 'other' };
  const iss = hasLines ? (Number(issuedAmt) || 0) : (tok === '발행' ? amount : 0);
  if (iss <= 0 && payer.method === '현금') return { ...base, state: 'cash' };
  if (iss <= 0 && payer.method === '카드') return { ...base, state: 'card' };
  if (iss >= amount - 1 && amount > 0) {
    return { ...base, state: 'issued', issued: iss, preissued: remaining > 1, pending: Math.max(0, iss - received) };
  }
  if (iss > 0) {
    return { ...base, state: 'partial', issued: iss, toIssue: amount - iss,
      need: received > iss + 1 && remaining <= 1, pending: Math.max(0, iss - received) };
  }
  if (received > 0) return { ...base, state: remaining <= 1 ? 'needed' : 'pending' };
  return base;
}

/** 나눠 내는 공사 → {단계: 상태 + payer}. activeLines = billStatus.activeInvoiceLines(계산서_메모) */
export function payerBillStates(row, activeLines) {
  const out = {};
  Object.entries(stagePayers(row)).forEach(([s, p]) => {
    const mine = (activeLines || []).filter((it) => it.stage === s && Number(it.amt) > 0);
    const st = payerBillState(p, (row || {})[`${s} 계산서`],
      mine.reduce((a, it) => a + Number(it.amt), 0), mine.length > 0);
    out[s] = { ...st, payer: p };
  });
  return out;
}

// 현금·카드로 내기로 했지만 아직 받은 돈이 없음 — 아이콘·요약에서 '움직임 없음'으로
const unpaidCashCard = (st) => (st.state === 'cash' || st.state === 'card') && !((Number(st.payer.received) || 0) > 0);

/** 나눠 내는 공사 계산서 요약(Y '{칸} - {상태}') — payer_split.split_y_summary 미러 */
export function splitYSummary(states) {
  const order = PAYER_ORDER.filter((s) => states[s]);
  const urgent = order.filter((s) => states[s].state === 'needed'
    || (states[s].state === 'partial' && states[s].need));
  if (urgent.length) return `${urgent[0]} - 미발행`;
  const uninvPaid = (st) => {
    const got = Number(st.payer.received) || 0;
    return (st.state === 'pending' && got > 0) || (st.state === 'partial' && got > st.issued + 1);
  };
  const quiet = order.filter((s) => uninvPaid(states[s]));
  if (quiet.length) return `${quiet[0]} - 미발행`;
  // 현금·카드로 내기로 한 사람은 아직 한 푼도 안 냈으면 움직임 없음 (payer_split.split_y_summary 미러, 10-21)
  const active = order.filter((s) => states[s].state !== 'none' && !unpaidCashCard(states[s]));
  if (!active.length) return '미발행';
  const s = active[active.length - 1];
  const label = { issued: '발행완료', partial: '발행완료', cash: 'N입금', card: '카드결제' }[states[s].state] || '기타';
  return `${s} - ${label}`;
}

/** 칸 메모 맨 윗줄 '◆ 임대인 이희승' — payer_split.payer_note_line 미러 */
export function payerNoteLine(payer) {
  return `${PAYER_NOTE_PREFIX} ${payer.role} ${payer.name}`;
}

/**
 * 사람 상태 → 계산서 아이콘 상태(billStatus 아이콘 값).
 * 일부만 발행(partial)은 이미 나간 계산서가 있으니 발행 아이콘(툴팁에 '남은 발행'), 그 사람 완납인데 덜 발행이면 ⚠️.
 */
export function payerBillIcon(st) {
  if (!st) return 'none';
  // 입금 전 현금·카드 사람에게 '현금 입금(계산서 불필요)' 아이콘이 뜨던 문제(10-21 305호) — 받은 돈이 있을 때만
  if (unpaidCashCard(st)) return 'none';
  if (st.state === 'partial') return st.need ? '미발행' : '발행';
  return { issued: '발행', needed: '미발행', pending: '발행예정', cash: 'N입금', card: '카드', other: '기타' }[st.state]
    || 'none';
}
