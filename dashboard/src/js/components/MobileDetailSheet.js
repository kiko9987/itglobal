/**
 * 휴대폰 상세 시트(화면 가득) — 프로젝트·고객 리드 화면 공용 (2026-11-16 휴대폰 리뉴얼).
 *
 * - 기존 펼침(아코디언) 코드는 '표 행 다음 줄'에 붙는다. 그래서 시트 안에 빈 표 행(anchorRow)을 두고
 *   각 화면이 펼침을 그 행에 붙인다(펼침 코드·편집·저장은 PC 와 같은 코드 그대로).
 * - 뒤로 가기 단추·휴대폰 뒤로 가기 = 닫기(history 한 칸), 열면 제목에 포커스·Tab 가두기, 닫으면 보던 카드로.
 * - 닫기 요청은 onRequestClose(fromPop) 로 화면에 넘긴다(고치던 내용 확인 등은 화면 쪽 책임).
 */
export default class MobileDetailSheet {
  constructor({ id, onRequestClose } = {}) {
    this.onRequestClose = onRequestClose || (() => this.hide());
    this.isOpen = false;
    this.closingFromPop = false;
    this.returnFocus = null;

    const sheet = document.createElement('div');
    sheet.className = 'pmm-sheet';
    sheet.id = id;
    sheet.setAttribute('role', 'dialog');
    sheet.setAttribute('aria-modal', 'true');
    sheet.setAttribute('aria-labelledby', `${id}Title`);
    sheet.hidden = true;
    sheet.innerHTML = `
      <div class="pmm-sheet-bar">
        <button type="button" class="pmm-sheet-back" aria-label="목록으로 돌아가기">
          <i class="fas fa-arrow-left" aria-hidden="true"></i>
        </button>
        <h2 class="pmm-sheet-title" id="${id}Title" tabindex="-1"></h2>
      </div>
      <div class="pmm-sheet-scroll">
        <div class="pmm-sheet-quick"></div>
        <table class="pmm-sheet-table" role="presentation">
          <tbody><tr class="pmm-sheet-anchor"><td></td></tr></tbody>
        </table>
      </div>`;
    document.body.appendChild(sheet);

    this.el = sheet;
    this.title = sheet.querySelector('.pmm-sheet-title');
    this.quick = sheet.querySelector('.pmm-sheet-quick');
    this.scroll = sheet.querySelector('.pmm-sheet-scroll');
    this.anchorRow = sheet.querySelector('.pmm-sheet-anchor');

    sheet.querySelector('.pmm-sheet-back').addEventListener('click', () => this.onRequestClose(false));
    window.addEventListener('popstate', () => {
      if (this.isOpen) this.onRequestClose(true);
    });
    // 포커스 가두기: 열려 있는 동안 Tab 으로 뒤 화면에 가지 않게. 시트 위에 뜬 창(메모·사유 입력 등
    //   부트스트랩 창, 몸통으로 옮겨 띄운 체크 메뉴)은 그대로 둔다. 뒤 화면 inert 는 그 창들이 본문 안에 있어 못 씀
    //   빠져나갈 때: 마지막 칸에서 앞으로 → 뒤로 단추, 첫 칸(뒤로 단추)에서 거꾸로 → 시트 마지막 칸
    const backBtn = sheet.querySelector('.pmm-sheet-back');
    let last = null;
    document.addEventListener('focusin', (e) => {
      if (!this.isOpen) return;
      if (sheet.contains(e.target)) { last = e.target; return; }
      if (e.target.closest?.('.modal, .dropdown-menu, [data-portal], .ts-dropdown, .tooltip, .popover')) return;
      if (last === backBtn) {
        const items = [...sheet.querySelectorAll('a[href], button:not([disabled]), input:not([disabled]), select, textarea, [tabindex="0"]')]
          .filter((el) => el.offsetParent !== null);
        (items[items.length - 1] || this.title).focus({ preventScroll: true });
      } else {
        backBtn.focus({ preventScroll: true });
      }
    });
  }

  /** 시트 열기(이미 열려 있으면 제목·바로 하기만 바꿈). returnFocus = 닫을 때 포커스를 돌려줄 요소 */
  show({ title, quickHtml = '', returnFocus = null } = {}) {
    this.title.textContent = title || '';
    this.quick.innerHTML = quickHtml;
    this.scroll.scrollTop = 0;
    if (returnFocus) this.returnFocus = returnFocus;
    if (this.isOpen) return;
    this.isOpen = true;
    this.el.hidden = false;
    document.documentElement.classList.add('pmm-sheet-open');
    try {
      history.pushState({ ...(history.state || {}), pmmSheet: title || '1' }, '');
    } catch (_) { /* 기록을 못 남겨도 닫기 단추로 닫힘 */ }
    requestAnimationFrame(() => this.title.focus({ preventScroll: true }));
  }

  setTitle(title) {
    this.title.textContent = title || '';
  }

  /** 시트 숨기기. fromPop=true 면 이미 history 가 한 칸 돌아간 상태(뒤로 가기) */
  hide({ fromPop = this.closingFromPop } = {}) {
    if (!this.isOpen) return;
    this.isOpen = false;
    this.el.hidden = true;
    document.documentElement.classList.remove('pmm-sheet-open');
    if (!fromPop && history.state && history.state.pmmSheet) {
      try { history.back(); } catch (_) { /* noop */ }
    }
    const back = this.returnFocus;
    this.returnFocus = null;
    if (back && document.contains(back) && back.offsetParent !== null) back.focus({ preventScroll: true });
  }

  /** 닫기 요청을 막았을 때(확인 창에서 '취소') 뒤로 가기로 빠진 history 한 칸을 되돌린다 */
  restoreHistory() {
    try {
      history.pushState({ ...(history.state || {}), pmmSheet: this.title.textContent || '1' }, '');
    } catch (_) { /* noop */ }
  }
}
