// 이카운트 회계거래조회 읽기(2026-11-28) — 매일 밤 대조(services/ecount_recon.py)가 부르는 읽기 전용 도구.
//
// 이카운트는 조회 API 가 없어 헤드리스 Edge(CDP)로 화면을 읽는다. 아무것도 고치지 않는다.
//   1) 계정 없이 2025-01-01~ 전체 전표 → 매입전표 I 등(매입·지출) 읽기
//   2) 같은 기간에 계정 4049(제품매출)를 걸어 → 매출전표 읽기
// 쪽(ul.pagination)을 하나씩 넘기며 표의 tr/td 를 읽고, 쪽이 하나라도 빠지면 실패로 끝낸다(반쯤 읽은 자료로
// 대조하면 맞는 건이 '해결됨'으로 잘못 바뀐다).
//
// 로그인 조건(11-27 시험): .env 의 ECOUNT_WEB_COM_CODE·ECOUNT_WEB_USER_ID·ECOUNT_WEB_PASSWORD.
//   값은 Input.insertText, 버튼은 Input.dispatchMouseEvent 로 실제 누름 — JS 로 value 만 바꾸면 화면이 받지 않는다.
// 비밀번호는 출력·로그·결과 파일 어디에도 쓰지 않는다.
//
// 실행: node tools/ecount_reader/read_ecount.mjs --out <폴더> [--from 2025-01-01] [--only sales|buy] [--env <.env 경로>]
// 결과: <폴더>/ecount_read_<날짜시각>.json, 같은 이름 .log. 마지막 줄에 'RESULT <json 경로>' 출력.
// 종료 코드: 0 성공 · 1 설정 없음 · 2 로그인 실패 · 3 읽기 실패(쪽 빠짐 등) · 4 그 밖의 오류
import { spawn, execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..', '..');
const SLIP_HASH = '#menuType=MENUTREE_000001&menuSeq=MENUTREE_000098&groupSeq=MENUTREE_000016&prgId=E010701&depth=4';
const SALES_ACCOUNT = '4049';
const MAX_MS = 40 * 60 * 1000;          // 전체 상한 40분 — 넘으면 실패

// ── 인자 ──
const argv = process.argv.slice(2);
const arg = (k, d) => { const i = argv.indexOf(k); return i >= 0 && argv[i + 1] ? argv[i + 1] : d; };
const OUT = path.resolve(arg('--out', path.join(ROOT, 'instance', 'ecount_recon')));
const FROM = arg('--from', '2025-01-01');
const ONLY = arg('--only', 'all');
const ENV_FILE = arg('--env', process.env.ECOUNT_ENV_FILE || path.join(ROOT, '.env'));
if (!/^\d{4}-\d{2}-\d{2}$/.test(FROM)) { console.log('[실패] --from 은 YYYY-MM-DD'); process.exit(1); }
fs.mkdirSync(OUT, { recursive: true });

const stamp = (() => { const d = new Date(); const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`; })();
const JSON_OUT = path.join(OUT, `ecount_read_${stamp}.json`);
const LOGF = path.join(OUT, `ecount_read_${stamp}.log`);
fs.writeFileSync(LOGF, '');
// 긴 토큰(세션 값 등)은 가림 — 우리 파일 이름(ecount_read_…)은 그대로
const mask = s => String(s).replace(/[A-Za-z0-9+/=_-]{24,}/g, m => /ecount_read_\d{8}_\d{6}/.test(m) ? m : '[가림]');
const log = (...a) => {
  const line = new Date().toTimeString().slice(0, 8) + ' ' + a.map(x => typeof x === 'string' ? x : JSON.stringify(x)).join(' ');
  console.log(mask(line)); fs.appendFileSync(LOGF, mask(line) + '\n');
};

// ── .env ──
const env = {};
try {
  for (const line of fs.readFileSync(ENV_FILE, 'utf8').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*?)\s*$/);
    if (m) env[m[1]] = m[2].replace(/^(['"])(.*)\1$/, '$2');
  }
} catch { log('[실패] .env 를 읽지 못했어요:', ENV_FILE); process.exit(1); }
const COM = env.ECOUNT_WEB_COM_CODE || env.ECOUNT_COM_CODE, ID = env.ECOUNT_WEB_USER_ID, PW = env.ECOUNT_WEB_PASSWORD;
if (!COM || !ID || !PW) { log('[실패] .env 의 ECOUNT_WEB_COM_CODE·ECOUNT_WEB_USER_ID·ECOUNT_WEB_PASSWORD 중 빈 칸이 있어요.'); process.exit(1); }

// ── Edge ──
function edgeExe() {
  if (env.ECOUNT_EDGE_EXE && fs.existsSync(env.ECOUNT_EDGE_EXE)) return env.ECOUNT_EDGE_EXE;
  for (const base of ['C:\\Program Files (x86)\\Microsoft\\Edge\\Application', 'C:\\Program Files\\Microsoft\\Edge\\Application']) {
    if (!fs.existsSync(base)) continue;
    // 자동 업데이트 대기 중이면 Application\msedge.exe 가 바로 끝나므로 버전 폴더의 최신 것을 직접 쓴다
    const vers = fs.readdirSync(base).filter(d => /^\d+\./.test(d) && fs.existsSync(path.join(base, d, 'msedge.exe')))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    if (vers.length) return path.join(base, vers[vers.length - 1], 'msedge.exe');
  }
  return null;
}
const exe = edgeExe();
if (!exe) { log('[실패] Edge 를 찾지 못했어요.'); process.exit(1); }
const port = 9300 + Math.floor(Math.random() * 600);
const prof = fs.mkdtempSync(path.join(os.tmpdir(), 'ecount-read-'));
const edge = spawn(exe, ['--headless=new', '--disable-gpu', '--no-first-run', `--remote-debugging-port=${port}`,
  `--user-data-dir=${prof}`, '--window-size=1400,1000', 'about:blank'], { stdio: 'ignore' });
function cleanup() {
  try { execFileSync('taskkill', ['/PID', String(edge.pid), '/T', '/F'], { stdio: 'ignore' }); } catch {}
  for (let i = 0; i < 5; i++) { try { fs.rmSync(prof, { recursive: true, force: true }); break; } catch {} }
}
const sleep = ms => new Promise(r => setTimeout(r, ms));
const T0 = Date.now();
const overtime = () => Date.now() - T0 > MAX_MS;

// ── CDP ──
let ws, seq = 0;
const pending = new Map(), dialogs = [], probe = [];
function send(method, params = {}) {
  const id = ++seq;
  ws.send(JSON.stringify({ id, method, params }));
  return new Promise((res, rej) => {
    pending.set(id, { res, rej });
    setTimeout(() => { if (pending.has(id)) { pending.delete(id); rej(new Error(`${method} 응답 없음`)); } }, 60000);
  });
}
async function js(expr) {
  const r = await send('Runtime.evaluate', { expression: expr, awaitPromise: true, returnByValue: true });
  if (r.exceptionDetails) throw new Error('화면 스크립트 오류: ' + (r.exceptionDetails.exception?.description || r.exceptionDetails.text || '').slice(0, 200));
  return r.result && r.result.value;
}
async function shot(name) {
  try { const r = await send('Page.captureScreenshot', { format: 'png' }); fs.writeFileSync(path.join(OUT, name), Buffer.from(r.data, 'base64')); log('화면 그림:', path.join(OUT, name)); } catch {}
}
async function mouseClick(x, y, clickCount = 1) {
  await send('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y });
  await send('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount });
  await send('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', clickCount });
}
// 화면 안의 요소를 찾아(스크립트가 요소를 돌려주는 식) 가운데를 실제로 누른다. 못 찾으면 false
async function clickWhere(findExpr, clickCount = 1) {
  const rect = await js(`(() => { const el = (${findExpr}); if (!el) return null; el.scrollIntoView({ block: 'center' });
    const r = el.getBoundingClientRect(); return r.width && r.height ? { x: r.x + r.width / 2, y: r.y + r.height / 2 } : null; })()`);
  if (!rect) return false;
  await mouseClick(rect.x, rect.y, clickCount);
  return true;
}
async function key(k, code, vk) {
  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: k, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk });
  await send('Input.dispatchKeyEvent', { type: 'keyUp', key: k, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk });
}
async function waitFor(expr, ms, step = 300) {
  for (let t = 0; t < ms; t += step) { try { if (await js(expr)) return true; } catch {} await sleep(step); }
  return false;
}

// ── 화면 읽기(회계거래조회) ──
const VIS = 'e => e && e.offsetParent !== null';
// 이카운트 화면은 배열의 JSON 바꾸기를 고쳐 놓아(11-28 첫 실행: range 가 배열이 아니라 글자로 옴) 화면 안에서는
// JSON.stringify 를 쓰지 않고 구분 글자(\u0001 칸, \u0002 줄, \u0003 항목)로 이은 글자 하나를 받아 여기서 나눈다.
const STATE_JS = `(() => { const t = document.body.innerText;
  const rows = []; document.querySelectorAll('tr').forEach(tr => { const c = []; tr.querySelectorAll('td').forEach(td => c.push(td.innerText.trim().replace(/\\s+/g, ' ')));
    if (c.length >= 9 && /^\\d+$/.test(c[0])) rows.push(c.join('\\u0001')); });
  const m = t.match(/(\\d{4}\\/\\d{2}\\/\\d{2}) ~ (\\d{4}\\/\\d{2}\\/\\d{2})/) || ['', '', ''];
  const a = document.querySelector('ul.pagination li.active');
  return [document.title, location.hash, m[1], m[2], (t.match(/\\/ (\\d+)/) || [0, 0])[1], a ? a.textContent.trim() : '0',
    document.querySelector('ul.pagination') ? '1' : '', rows.join('\\u0002')].join('\\u0003'); })()`;
async function state() {
  const raw = await js(STATE_JS);
  if (typeof raw !== 'string') throw new Error('화면 상태를 읽지 못했어요.');
  const [title, hash, r0, r1, last, active, pager, rows] = raw.split('\u0003');
  return { title, hash, range: r0 ? [r0, r1] : [], last: +last || 0, active: +active || 0, pager: !!pager,
           rows: rows ? rows.split('\u0002').map(r => r.split('\u0001')) : [] };
}
const firstKey = st => (st.rows[0] || []).join('|');

// 쪽 이동 — ul.pagination 안의 a 만 누른다(밖의 a 를 누르면 ERP 첫 화면으로 튐, 11-26)
const GOTO_JS = p => `(() => { const u = document.querySelector('ul.pagination'); if (!u) return 'nopager';
  const L = [...u.querySelectorAll('li')]; const a = L.map(li => li.querySelector('a')).find(x => x && x.textContent.trim() === '${p}');
  if (a) { a.click(); return 'page'; }
  let idx = -1; L.forEach((li, i) => { if (/^\\d+$/.test(li.textContent.trim())) idx = i; });
  const nx = L[idx + 1] && L[idx + 1].querySelector('a'); if (!nx) return 'nonext'; nx.click(); return 'next'; })()`;

function today() { const d = new Date(); const p = n => String(n).padStart(2, '0'); return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())}`; }

async function openSlipList() {
  await js(`location.hash = ${JSON.stringify(SLIP_HASH)}; 1`);
  if (!await waitFor(`document.title === '회계거래조회' && !!document.querySelector('ul.pagination')`, 30000))
    throw new Error('회계거래조회 화면이 열리지 않았어요.');
  await sleep(1500);
}

async function openSearchForm() {
  // 이미 열려 있으면 그대로(F3 단추는 열고 닫기를 번갈아 함)
  if (await js(`[...document.querySelectorAll('input#day')].some(${VIS})`)) return;
  const ok = await clickWhere(`[...document.querySelectorAll('button,a,span')].find(e => (e.textContent || '').trim() === 'Search(F3)' && e.offsetParent !== null)`);
  if (!ok) throw new Error('검색 창(Search F3) 단추를 찾지 못했어요.');
  if (!await waitFor(`[...document.querySelectorAll('input#day')].some(${VIS})`, 8000)) throw new Error('검색 창이 열리지 않았어요.');
  await sleep(500);
}

// 연·월 고르기 — 단추를 눌러 뜬 목록에서 값 글자를 실제로 누른다
async function pick(kind, index, value) {
  const btn = `document.querySelectorAll('button[data-id=${kind}]')[${index}]`;
  if ((await js(`(${btn} || {}).textContent`) || '').trim() === value) return;
  if (!await clickWhere(btn)) throw new Error(`${kind} 단추를 찾지 못했어요.`);
  await sleep(600);
  const ok = await clickWhere(`(() => { const b = ${btn}; const id = b.getAttribute('data-own-layer-box-id');
    const box = (id && (document.getElementById(id) || document.querySelector('[data-layer-box-id="' + id + '"]'))) || document.body;
    const items = [...box.querySelectorAll('a,li,span,div')].filter(e => e.offsetParent !== null && e.children.length === 0 && e.textContent.trim() === '${value}');
    return items[items.length - 1] || null; })()`);
  if (!ok) throw new Error(`${kind} 목록에서 ${value} 를 찾지 못했어요.`);
  await sleep(500);
  const now = ((await js(`(${btn} || {}).textContent`)) || '').trim();
  if (now !== value) throw new Error(`${kind} 가 ${value} 로 바뀌지 않았어요(지금 ${now}).`);
}

async function setDay(index, d) {
  if (!await clickWhere(`document.querySelectorAll('input#day')[${index}]`, 3)) throw new Error('일 칸을 찾지 못했어요.');
  await js(`(() => { const e = document.querySelectorAll('input#day')[${index}]; e.select && e.select(); return 1; })()`);
  await send('Input.insertText', { text: d });
  await sleep(300);
  const v = await js(`document.querySelectorAll('input#day')[${index}].value`);
  if (String(v).padStart(2, '0') !== d) throw new Error(`일이 ${d} 로 들어가지 않았어요(지금 ${v}).`);
}

// 앞으로 갈 때(endFirst)는 끝 날짜부터 바꿔 잠깐이라도 시작 > 끝이 되지 않게
async function setRange(from, to, endFirst = false) {
  const [y1, m1, d1] = from.split('-'), [y2, m2, d2] = to.split('-');
  await openSearchForm();
  const start = async () => { await pick('year', 0, y1); await pick('month', 0, m1); await setDay(0, d1); };
  const end = async () => { await pick('year', 1, y2); await pick('month', 1, m2); await setDay(1, d2); };
  if (endFirst) { await end(); await start(); } else { await start(); await end(); }
}

// 분기마다 한 구간. 마지막 구간 끝은 다음 달 말일(미리 날짜로 넣은 전표까지)
function chunks() {
  const out = [], p = n => String(n).padStart(2, '0');
  const now = new Date(), endAll = new Date(now.getFullYear(), now.getMonth() + 2, 0);
  let [y, m] = FROM.split('-').map(Number), first = true;
  while (true) {
    const start = first ? FROM : `${y}-${p(m)}-01`;
    first = false;
    let ey = y, em = m + 2;
    if (em > 12) { em -= 12; ey++; }
    let end = new Date(ey, em, 0);
    if (end >= endAll || new Date(ey, em, 1) > now) end = endAll;
    out.push([start, `${end.getFullYear()}-${p(end.getMonth() + 1)}-${p(end.getDate())}`]);
    if (end >= endAll) break;
    m = em + 1; y = ey;
    if (m > 12) { m = 1; y++; }
  }
  return out;
}

async function setAccount(code) {
  if (!await clickWhere(`document.querySelector('input[placeholder="계정"]')`)) throw new Error('계정 칸을 찾지 못했어요.');
  await send('Input.insertText', { text: code });
  await sleep(1200);
  await key('Enter', 'Enter', 13);
  await sleep(1500);
  // 고른 계정은 칩으로 붙는다 — 제품매출 글자가 검색 창 안에 보이는지
  const chip = await js(`(() => { const i = document.querySelector('input[placeholder="계정"]'); let box = i; for (let k = 0; k < 6 && box; k++) box = box.parentElement;
    return box ? box.innerText.replace(/\\s+/g, ' ').slice(0, 200) : ''; })()`);
  if (!/제품매출|4049/.test(chip || '')) throw new Error('계정 4049(제품매출)가 걸리지 않았어요: ' + (chip || '').slice(0, 80));
  log('계정 걸림:', (chip.match(/[^ ]*제품매출[^ ]*/) || [code])[0]);
}

async function runSearch(from, to) {
  const before = firstKey(await state());
  const ok = await clickWhere(`[...document.querySelectorAll('button')].find(e => e.offsetParent !== null && /^검색\\(F8\\)$/.test(e.textContent.trim()))`);
  if (!ok) throw new Error('검색(F8) 단추를 찾지 못했어요.');
  const w0 = from.replace(/-/g, '/'), w1 = to.replace(/-/g, '/');
  let empty = 0;
  for (let t = 0; t < 60000; t += 500) {
    await sleep(500);
    const s = await state().catch(() => null);
    if (!s || s.range[0] !== w0 || s.range[1] !== w1) continue;
    if (s.pager && s.rows.length && firstKey(s) !== before) { await sleep(1500); return await state(); }
    if (!s.rows.length && ++empty >= 10) return s;          // 그 기간에 전표가 없음(5초 동안 빈 표)
  }
  const s = await state().catch(() => ({ range: [] }));
  throw new Error(`검색 결과 기간이 맞지 않아요(원하는 ${w0} ~ ${w1}, 지금 ${s.range.join(' ~ ') || '없음'}).`);
}

// 지금 검색 결과를 첫 쪽부터 끝까지 — 쪽마다 행 수를 남기고, 하나라도 못 읽으면 실패
async function readAll(label, from, to) {
  let s = await state();
  if (s.range[0] !== from.replace(/-/g, '/') || s.range[1] !== to.replace(/-/g, '/'))
    throw new Error(`${label}: 기간이 ${from} ~ ${to} 가 아니에요(${s.range.join(' ~ ')}).`);
  if (!s.rows.length) { log(`${label}: 전표 없음`); return { rows: [], pages: 0, page_size: 0, range: s.range }; }
  if (s.last >= 100) throw new Error(`${label}: 100쪽 한도에 닿았어요(잘렸을 수 있음) — 기간을 더 쪼개야 해요.`);
  if (s.active !== 1) {   // 첫 쪽으로
    await js(GOTO_JS(1)); await waitFor(`(document.querySelector('ul.pagination li.active') || {}).textContent === '1'`, 15000); await sleep(800);
    s = await state();
  }
  const last = s.last;
  if (!last) throw new Error(`${label}: 전체 쪽 수를 읽지 못했어요.`);
  const pages = { 1: s.rows };
  log(`${label}: 전체 ${last}쪽, 1쪽 ${s.rows.length}줄, 기간 ${s.range.join(' ~ ')}`);
  let p = 2, guard = 0;
  while (p <= last) {
    if (overtime()) throw new Error(`${label}: 시간 초과(${p}쪽에서 멈춤)`);
    if (++guard > last * 4 + 20) throw new Error(`${label}: 쪽 넘기기가 끝나지 않아요(${p}쪽).`);
    const firstBefore = firstKey(await state());
    const how = await js(GOTO_JS(p));
    if (how === 'nopager' || how === 'nonext') throw new Error(`${label}: ${p}쪽으로 갈 단추가 없어요(${how}).`);
    let moved = false;
    for (let t = 0; t < 20000; t += 300) {
      await sleep(300);
      const n = await state().catch(() => null);
      if (n && n.pager && firstKey(n) !== firstBefore && n.active) { moved = true; break; }
    }
    if (!moved) throw new Error(`${label}: ${p}쪽이 열리지 않았어요.`);
    await sleep(400);
    const n = await state();
    if (!/E010701/.test(n.hash)) throw new Error(`${label}: 회계거래조회 화면을 벗어났어요.`);
    if (how === 'next' && n.active !== p) {        // 다음 묶음으로 넘어가면 그 묶음 첫 쪽이 열림 — 원하는 쪽이 아니면 다시
      if (n.active < p) { pages[n.active] = n.rows; continue; }
      throw new Error(`${label}: ${p}쪽 대신 ${n.active}쪽이 열렸어요.`);
    }
    if (n.active !== p) throw new Error(`${label}: ${p}쪽 대신 ${n.active}쪽이 열렸어요.`);
    pages[p] = n.rows;
    if (p % 10 === 0) log(`${label}: ${p}/${last}쪽`);
    p++;
  }
  // 빠진 쪽·덜 찬 쪽 확인(마지막 쪽만 덜 찰 수 있음)
  const size = pages[1].length;
  for (let i = 1; i <= last; i++) {
    if (!pages[i]) throw new Error(`${label}: ${i}쪽을 못 읽었어요.`);
    if (i < last && pages[i].length !== size) throw new Error(`${label}: ${i}쪽이 ${pages[i].length}줄이에요(다른 쪽은 ${size}줄).`);
  }
  const num = v => Number(String(v || '0').replace(/[^0-9-]/g, '')) || 0;
  const rows = [];
  for (let i = 1; i <= last; i++) for (const c of pages[i]) {
    rows.push({ d: c[1], site: c[2], code: c[3], typ: c[4], amt: num(c[5]), s: num(c[6]), v: num(c[7]), partner: c[8] || '', memo: c[9] || '' });
  }
  log(`${label}: ${rows.length}줄 읽음(${last}쪽, 쪽당 ${size}줄)`);
  return { rows, pages: last, page_size: size, range: s.range };
}

// ── 실행 ──
const result = { ok: false, started: new Date().toISOString(), from: FROM, only: ONLY, sales: null, buy: null, dialogs, probe, error: '' };
try {
  let target;
  for (let i = 0; i < 60 && !target; i++) {
    await sleep(250);
    try { target = (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find(t => t.type === 'page'); } catch {}
  }
  if (!target) throw Object.assign(new Error('Edge 디버깅 연결 실패'), { code: 4 });
  ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((r, j) => { ws.addEventListener('open', r); ws.addEventListener('error', j); });
  ws.addEventListener('message', ev => {
    const m = JSON.parse(ev.data);
    if (m.id && pending.has(m.id)) { const p = pending.get(m.id); pending.delete(m.id); m.error ? p.rej(new Error(m.error.message)) : p.res(m.result); }
    if (m.method === 'Page.javascriptDialogOpening') {
      dialogs.push(m.params.message);
      log('[창]', m.params.type, JSON.stringify(m.params.message));
      send('Page.handleJavaScriptDialog', { accept: true }).catch(() => {});
    }
    // 화면 안 JSON API 모양만 남김(값은 안 남김) — 나중에 화면 대신 API 로 읽을지 판단용
    if (m.method === 'Network.requestWillBeSent' && /SelectAccountSearchListAction/.test(m.params.request.url) && probe.length < 4) {
      let keys = [];
      try { const b = JSON.parse(m.params.request.postData || '{}'); const walk = (o, pre) => { for (const k of Object.keys(o || {})) { keys.push(pre + k); if (o[k] && typeof o[k] === 'object' && !Array.isArray(o[k]) && pre.split('.').length < 3) walk(o[k], pre + k + '.'); } }; walk(b, ''); } catch {}
      probe.push({ url: m.params.request.url.replace(/\?.*$/, ''), method: m.params.request.method, keys: keys.slice(0, 80) });
    }
  });
  await send('Page.enable'); await send('Runtime.enable'); await send('Network.enable');

  // 로그인
  log('로그인 화면 여는 중');
  await send('Page.navigate', { url: 'https://login.ecount.com/' });
  if (!await waitFor(`!!document.querySelector('#passwd')`, 20000)) throw Object.assign(new Error('로그인 화면이 열리지 않았어요.'), { code: 2 });
  await sleep(1500);
  for (const [sel, v] of [['#com_code', COM], ['#id', ID], ['#passwd', PW]]) {
    await js(`(() => { const e = document.querySelector('${sel}'); e.focus(); e.select && e.select(); e.value = ''; return 1; })()`);
    await send('Input.insertText', { text: v });
    await sleep(200);
  }
  if (!await clickWhere(`document.querySelector('#save')`)) throw Object.assign(new Error('로그인 단추를 찾지 못했어요.'), { code: 2 });
  log('로그인 단추 누름');
  let url = '';
  for (let i = 0; i < 60; i++) {
    await sleep(500);
    url = await js('location.href').catch(() => '');
    if (/\/view\/erp/.test(url)) break;
  }
  if (!/\/view\/erp/.test(url)) {
    await shot(`ecount_read_${stamp}_login_fail.png`);
    const info = await js(`JSON.stringify({ err: (document.querySelector('#error_msg') || {}).value || '', text: document.body.innerText.slice(0, 200) })`).catch(() => '');
    throw Object.assign(new Error('로그인 뒤 ERP 화면으로 넘어가지 않았어요. ' + info), { code: 2 });
  }
  log('로그인 됨:', (url.match(/^https:\/\/[^/]+/) || [''])[0]);

  await openSlipList();
  const parts = chunks();
  log('읽을 구간:', parts.map(x => x.join('~')).join(', '));
  async function readChunks(label, keep) {
    const rows = [], info = [];
    for (const [i, [from, to]] of parts.entries()) {
      await setRange(from, to, i > 0);
      await runSearch(from, to);
      const r = await readAll(`${label} ${from}~${to}`, from, to);
      rows.push(...r.rows.filter(keep));
      info.push({ from, to, pages: r.pages, rows: r.rows.length });
    }
    return { rows, chunks: info, total_rows: info.reduce((a, x) => a + x.rows, 0) };
  }
  if (ONLY !== 'sales') {
    // 1) 계정 없이 — 매입·지출 전표(매출전표도 섞여 나오지만 매출은 2)에서 계정을 걸어 따로 읽음)
    await openSearchForm();
    if (await js(`(() => { const i = document.querySelector('input[placeholder="계정"]'); let b = i; for (let k = 0; k < 6 && b; k++) b = b.parentElement; return /제품매출/.test(b ? b.innerText : ''); })()`))
      throw new Error('계정 칸에 이미 계정이 걸려 있어요. 매입을 읽으려면 계정이 비어 있어야 해요.');
    result.buy = await readChunks('전체 전표', r => r.typ !== '매출전표 I');
    log(`매입·지출 전표 ${result.buy.rows.length}줄(전체 ${result.buy.total_rows}줄에서 매출전표 뺌)`);
  }
  if (ONLY !== 'buy') {
    await openSearchForm();
    await setAccount(SALES_ACCOUNT);
    result.sales = await readChunks('매출(4049)', () => true);
    const kinds = {};
    result.sales.rows.forEach(r => { kinds[r.typ] = (kinds[r.typ] || 0) + 1; });
    log(`매출 전표 ${result.sales.rows.length}줄, 종류:`, kinds);
  }
  result.ok = true;
} catch (e) {
  result.error = e.message;
  log('[실패]', e.message);
  await shot(`ecount_read_${stamp}_fail.png`);
  process.exitCode = e.code || 3;
} finally {
  result.finished = new Date().toISOString();
  try { ws && ws.close(); } catch {}
  cleanup();
  fs.writeFileSync(JSON_OUT + '.tmp', JSON.stringify(result));
  fs.renameSync(JSON_OUT + '.tmp', JSON_OUT);
  if (probe.length) log('화면 API 모양:', probe.map(p => `${p.method} ${p.url} 칸 ${p.keys.length}개`).join(' / '));
  log(result.ok ? '[성공]' : '[끝]', `매출 ${result.sales ? result.sales.rows.length : '-'}줄 · 매입 ${result.buy ? result.buy.rows.length : '-'}줄 · ${Math.round((Date.now() - T0) / 1000)}초`);
  console.log('RESULT ' + JSON_OUT);
}
