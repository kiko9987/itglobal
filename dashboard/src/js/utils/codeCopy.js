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

function onClick(e) {
  const btn = e.target.closest && e.target.closest('.code-copy-btn');
  if (!btn) return;
  e.preventDefault();
  e.stopPropagation();
  const code = btn.dataset.copyCode || '';
  writeClipboard(code).then((ok) => {
    btn.innerHTML = ok ? CHECK_SVG : COPY_SVG;
    btn.classList.toggle('copied', ok);
    btn.title = ok ? '복사됨' : '복사 실패 — 다시 눌러 주세요';
    clearTimeout(btn._copyTimer);
    btn._copyTimer = setTimeout(() => {
      btn.innerHTML = COPY_SVG;
      btn.classList.remove('copied');
      btn.title = btn.dataset.copyLabel || '코드 복사';
    }, 1200);
  });
}

let installed = false;
export function installCodeCopy() {
  if (installed) return;
  installed = true;
  document.addEventListener('click', onClick, true);
}
