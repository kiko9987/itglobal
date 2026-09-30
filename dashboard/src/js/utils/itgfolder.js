/**
 * itgfolder:// — Drive 폴더를 각 직원 PC 탐색기로 여는 커스텀 프로토콜 (프로젝트·리드 페이지 공용)
 *
 * 클라이언트에 install-itg-folder.bat(→ C:\ITG\open-itg-folder.vbs + HKCU 등록) 설치 필요.
 * VBS 가 explorer.exe "G:\.shortcut-targets-by-id\{폴더ID}" 실행. 서버 subprocess 방식은 불가
 * (NSSM LocalSystem 세션 격리) — 되돌리지 말 것.
 */

export const FOLDER_ID_RE = /^[a-zA-Z0-9_-]{20,}$/;

function escAttr(s) {
  return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 폴더 ID → 탐색기 열기 링크 (ID 패턴 아니면 '') */
export function renderItgfolderLink(folderId, { text } = {}) {
  const id = String(folderId || '').trim();
  if (!FOLDER_ID_RE.test(id)) return '';
  const safe = escAttr(id);
  return `<a href="itgfolder://${safe}" class="text-decoration-none itgfolder-link" data-folder-id="${safe}"
            style="color: #0d6efd;" title="탐색기에서 열기 (프로토콜 필요)">${text ? escAttr(text) : safe}</a>`;
}

/**
 * itgfolder:// 링크 클릭 감지 (2026-07-08). 클릭 후 1.5초 안에 탐색기로 포커스가 안 넘어가면
 * 프로토콜 미설치로 보고 설치 안내 + 폴더 ID 클립보드 복사. 한 번 성공한 브라우저는 다시 안내 안 함.
 * 페이지당 한 번만 등록 (전역 플래그).
 */
export function bindItgfolderProtocolHandler() {
  if (window._itgfolderHandlerBound) return;
  window._itgfolderHandlerBound = true;

  document.addEventListener('click', (e) => {
    const link = e.target.closest('.itgfolder-link');
    if (!link) return;
    const folderId = link.dataset.folderId;
    if (!folderId) return;

    if (localStorage.getItem('itg_folder_protocol_ok') === '1') return;

    // 프로토콜이 실행되면 탐색기(외부 앱)로 focus 이동 → window blur. 1.5초 안에 blur 없으면 미설치 유력.
    let focusLost = false;
    const onBlur = () => { focusLost = true; };
    window.addEventListener('blur', onBlur, { once: true });

    setTimeout(() => {
      window.removeEventListener('blur', onBlur);
      if (focusLost) {
        localStorage.setItem('itg_folder_protocol_ok', '1');
        return;
      }
      const proceed = confirm(
        '폴더가 열리지 않았나요?\n\n' +
        '탐색기 프로토콜 미설치일 수 있습니다.\n' +
        '설치 방법:\n' +
        '  1) 회사에서 배포한 install-itg-folder.bat 파일을 실행\n' +
        '  2) 브라우저를 완전히 종료 후 재실행\n' +
        '  3) 다시 폴더 링크 클릭\n\n' +
        '설치 가이드가 필요하면 관리자(kiko@itg-aircon.com)에게 문의하세요.\n\n' +
        '확인을 누르면 폴더 ID를 클립보드에 복사합니다 (Drive에서 직접 열기용).'
      );
      if (proceed && navigator.clipboard) {
        navigator.clipboard.writeText(folderId).catch(() => {});
      }
    }, 1500);
  });
}
