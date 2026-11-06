/**
 * 프로젝트 코드 복사 아이콘 (2026-11-03 직원 의견) — 목록 표 코드 칸·펼친 화면 요약줄.
 * 2026-11-06 현장 주소에도 같은 아이콘 (직원 의견) — label 로 '주소 복사'.
 * 항상 흐리게 보이게(마우스 올려야 보이면 있는지 모르고 태블릿에선 못 씀). 누르면 복사 + 잠깐 체크 표시.
 * 클릭은 document capture 단계에서 받아 행 펼침(아코디언 토글)보다 먼저 막는다 — 계산서 아이콘과 같은 방식.
 */
import { escHtml } from './escapeHtml.js';

const COPY_SVG = '<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" '
  + 'stroke-width="1.5" aria-hidden="true"><rect x="5.5" y="5.5" width="8" height="8" rx="1.5"/>'
  + '<path d="M10.5 3.5v-.5a1.5 1.5 0 0 0-1.5-1.5H4A1.5 1.5 0 0 0 2.5 3v5A1.5 1.5 0 0 0 4 9.5h.5"/></svg>';
const CHECK_SVG = '<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" '
  + 'stroke-width="2" aria-hidden="true"><path d="M3 8.5l3.2 3L13 4.5"/></svg>';

export function codeCopyButtonHtml(code, label = '코드 복사') {
  const c = String(code ?? '').trim();
  if (!c) return '';
  const l = escHtml(label);
  return `<button type="button" class="code-copy-btn" data-copy-code="${escHtml(c)}" data-copy-label="${l}" `
    + `title="${l}" aria-label="${l}">${COPY_SVG}</button>`;
}

async function writeClipboard(text) {
  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch (_) { /* 아래 방식으로 */ }
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.setAttribute('readonly', '');
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.select();
  let ok = false;
  try { ok = document.execCommand('copy'); } catch (_) { ok = false; }
  ta.remove();
  return ok;
}

// 마우스를 올리면 '코드 복사'·'주소 복사' 안내 — 목록 표는 표 툴팁 초기화가 이미 붙이지만 펼친 화면(아코디언)은
//   기본 title 만 있어 늦게 뜨거나 안 보였다(2026-11-06). 어디서든 같은 툴팁을 처음 올릴 때 붙인다.
function hoverTip(btn, create = false) {
  const T = window.bootstrap && window.bootstrap.Tooltip;
  if (!T) return null;
  const cur = T.getInstance(btn);
  if (cur || !create) return cur;
  return new T(btn, { trigger: 'hover', placement: 'top', container: 'body', delay: { show: 200, hide: 100 } });
}

function onMouseOver(e) {
  const btn = e.target.closest && e.target.closest('.code-copy-btn');
  if (!btn || hoverTip(btn)) return;
  const tip = hoverTip(btn, true);
  if (tip) tip.show();   // 이번 mouseenter 는 이미 지나갔을 수 있어 직접 띄움 (떠날 땐 툴팁이 알아서 숨김)
}

// 누른 뒤 '복사됨' 말풍선 — 툴팁 글자를 바꾸면 표·펼친 화면마다 다르게 보여 따로 띄운다
function flashBubble(btn, text, ok) {
  const tip = hoverTip(btn);
  if (tip) tip.hide();
  const r = btn.getBoundingClientRect();
  const b = document.createElement('div');
  b.className = ok ? 'code-copy-bubble' : 'code-copy-bubble fail';
  b.textContent = text;
  b.style.left = `${r.left + r.width / 2 + window.scrollX}px`;
  b.style.top = `${r.top + window.scrollY - 6}px`;
  document.body.appendChild(b);
  setTimeout(() => b.remove(), 1200);
}

function onClick(e) {
  const btn = e.target.closest && e.target.closest('.code-copy-btn');
  if (!btn) return;
  e.preventDefault();
  e.stopPropagation();
  const code = btn.dataset.copyCode || '';
  writeClipboard(code).then((ok) => {
    btn.innerHTML = ok ? CHECK_SVG : COPY_SVG;
    btn.classList.toggle('copied', ok);
    flashBubble(btn, ok ? '복사됨' : '복사 실패 — 다시 눌러 주세요', ok);
    clearTimeout(btn._copyTimer);
    btn._copyTimer = setTimeout(() => {
      btn.innerHTML = COPY_SVG;
      btn.classList.remove('copied');
    }, 1200);
  });
}

let installed = false;
export function installCodeCopy() {
  if (installed) return;
  installed = true;
  document.addEventListener('click', onClick, true);
  document.addEventListener('mouseover', onMouseOver, true);
}
