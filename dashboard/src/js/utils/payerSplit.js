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

const SEP_RE = /\s*[·|｜]\s*/;
const BIZNO_RE = /^(\d{3})-?(\d{2})-?(\d{5})$/;
const AMT_RE = /^(?:₩\s*)?(\d[\d,]*)\s*원?$/;
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const CASH_WORDS = new Set(['현금', '계산서x', '계산서없음', '계산서미발행', 'n입금', 'x']);
const METHOD_WORDS = { 계산서: '계산서', 세금계산서: '계산서', 카드: '카드', 미정: '미정' };

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

/** 부담자 칸 → {role, name, bizNo, email, amount, method, note} 또는 null */
export function parsePayerCell(text) {
  const s = String(text == null ? '' : text).trim();
  if (!s || s === '-' || !SEP_RE.test(s)) return null;
  const toks = s.split(SEP_RE).map((t) => t.trim()).filter(Boolean);
  if (toks.length < 3) return null;
  const role = toks[0];
  if (role.length > 12 || AMT_RE.test(role) || BIZNO_RE.test(role) || role.includes('@') || methodOf(role)) return null;
  let name = ''; let bizNo = ''; let email = ''; let method = ''; let amount = 0;
  const notes = [];
  toks.slice(1).forEach((t) => {
    const mBiz = t.match(BIZNO_RE);
    const mAmt = t.match(AMT_RE);
    if (!bizNo && mBiz) bizNo = `${mBiz[1]}-${mBiz[2]}-${mBiz[3]}`;
    else if (!email && EMAIL_RE.test(t)) email = t;
    else if (!amount && mAmt && toNum(mAmt[1]) >= 1000) amount = toNum(mAmt[1]);
    else if (!method && methodOf(t)) method = methodOf(t);
    else if (!name) name = t;
    else notes.push(t);
  });
  if (!name || amount <= 0) return null;
  return { role, name, bizNo, email, amount, method: method || '계산서', note: notes.join(' · ') };
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
  return `${p.role} ${p.name} · 부담 ${p.amount.toLocaleString()} · ${tail}`;
}
