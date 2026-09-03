/**
 * InvoiceLicense — 사업자등록증 업로드/열람 + 세금계산서 발행 요청 (경량 컨트롤러).
 *
 * PM 아코디언 문서 카드의 인라인 버튼이 window.invoiceLicense 를 참조한다.
 *  - ② 등록증: 상태 lazy 조회 → 업로드(FormData)·열람(새 탭 Drive)
 *  - ③ 계산서: 모달 → POST /api/invoice/request → 슬랙 #계산서_관리 카드 발송(백엔드 공용 코어)
 *
 * 백엔드 경로는 모두 '/api/invoice/...' (보안 미들웨어 CSRF 통과 위해 /api/ 접두사 필수).
 */
import logger from '../utils/logger.js';

function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** 속성 선택자용 값 escape (프로젝트 코드는 단순하지만 안전하게). */
function attr(v) {
  return String(v == null ? '' : v).replace(/["\\]/g, '\\$&');
}

export default class InvoiceLicense {
  constructor() {
    this._modal = null;
    this._status = {};   // code → {exists, required, view_url}
    window.invoiceLicense = this;
  }

  // ───────────────────────────── 공통 헬퍼 ─────────────────────────────
  _findProject(code) {
    const c = String(code || '').trim();
    try {
      const sm = window.projectListApp && window.projectListApp.stateManager;
      if (sm && typeof sm.findProject === 'function') {
        const p = sm.findProject(c);
        if (p) return p;
      }
    } catch (_) { /* noop */ }
    if (Array.isArray(window.projectsData)) {
      return window.projectsData.find(x => String(x['프로젝트 코드'] || '').trim() === c) || null;
    }
    return null;
  }

  _toast(message, type = 'info') {
    import('./Toast.js').then(({ default: Toast }) => {
      new Toast().show(message, type);
    }).catch(() => logger.debug(`[알림] ${message}`));
  }

  _modalHost() {
    let host = document.getElementById('invoiceModalHost');
    if (!host) {
      host = document.createElement('div');
      host.id = 'invoiceModalHost';
      document.body.appendChild(host);
    }
    return host;
  }

  // ─────────────────────── ② 등록증 상태 lazy 조회 ───────────────────────
  /** 아코디언 body 렌더 직후 호출 — 등록증 상태 조회.
   * ⚠️ 이 훅은 renderAccordionContent 내부(아코디언이 아직 document 에 삽입되기 *전*)에서
   * 동기 호출된다. 따라서 DOM 존재 여부로 가드하면 안 된다(항상 null → 조기 return 버그).
   * refreshStatus 의 _applyStatus 는 fetch 완료(비동기) 시점에 실행되며, 그때는 이미 삽입돼
   * 있어 안전하게 뱃지를 갱신한다. */
  onAccordionRendered(code) {
    const c = String(code || '').trim();
    if (!c) return;
    this.refreshStatus(c);
  }

  async refreshStatus(code) {
    const c = String(code || '').trim();
    if (!c) return;
    try {
      const resp = await fetch(`/api/invoice/license/status/${encodeURIComponent(c)}`, {
        credentials: 'same-origin',
      });
      const json = await resp.json().catch(() => ({}));
      if (!resp.ok || json.success === false || !json.data) {
        this._applyStatusFail(c);
        return;
      }
      const d = json.data;
      this._status[c] = d;
      this._applyStatus(c, d);
    } catch (e) {
      logger.debug('[InvoiceLicense] 등록증 상태 조회 실패:', e);
      this._applyStatusFail(c);
    }
  }

  /** 상태 조회 실패 시 — '확인 중…' 에 멈추지 않도록 명확한 실패 뱃지로 전환. */
  _applyStatusFail(code) {
    const badge = document.querySelector(`[data-il-status="${attr(code)}"]`);
    if (badge) {
      badge.innerHTML = '';
      badge.textContent = '상태 확인 실패 (새로고침)';
      badge.style.background = '#fde7e9';
      badge.style.color = '#b42318';
      badge.style.padding = '3px 10px';   // 직전이 링크(padding 0)였을 수 있어 pill 복원
    }
  }

  _applyStatus(code, d) {
    const c = String(code || '').trim();
    const holder = document.querySelector(`[data-il-status="${attr(c)}"]`);
    if (holder) {
      if (d.exists && d.view_url) {
        // 등록증 있음 → 문서 폴더 값처럼 '링크'로 (클릭 시 새 탭 Drive 열람). 별도 열람 버튼 불필요.
        holder.style.background = 'transparent';
        holder.style.padding = '0';
        holder.style.color = '';
        holder.innerHTML =
          `<a href="${esc(d.view_url)}" target="_blank" rel="noopener"
              class="text-decoration-none" style="color:#0d6efd; font-weight:600;"
              title="사업자등록증 열람 (새 탭)"><i class="fas fa-file-invoice me-1"></i>사업자등록증 열기</a>`;
      } else if (d.source === 'partner') {
        // 등록증 파일은 없지만 거래처 탭에 상호 있음 → 계산서 발행 가능 (긍정 상태, 열람 링크 없음)
        holder.innerHTML = '';
        holder.textContent = '발행 가능 (거래처)';
        holder.style.background = '#e3f4e8';
        holder.style.color = '#1c7c3d';
        holder.style.padding = '3px 10px';
        holder.title = '거래처 탭에 사업자 정보가 있어 계산서 발행 가능 (등록증 파일 없음 — 업로드하면 열람·첨부 가능)';
      } else {
        // 없음 / 불필요(거래처) → pill 뱃지
        let label, bg, fg;
        if (d.exists) { label = '등록증 있음'; bg = '#e3f4e8'; fg = '#1c7c3d'; }  // view_url 없을 때 fallback
        else if (!d.required) { label = '등록증 불필요 (거래처)'; bg = '#eef0f2'; fg = '#6b7280'; }
        else { label = '등록증 없음'; bg = '#fdeede'; fg = '#b45309'; }
        holder.innerHTML = '';
        holder.textContent = label;
        holder.style.background = bg;
        holder.style.color = fg;
        holder.style.padding = '3px 10px';
      }
    }
    // 편집 모드 컨트롤(업로드/삭제) — 있으면 [삭제], 없으면 [업로드]. 컨테이너 표시는 편집 진입 시.
    const edit = document.querySelector(`[data-il-edit="${attr(c)}"]`);
    if (edit) {
      edit.innerHTML = d.exists
        ? `<button type="button" class="btn btn-outline-danger btn-sm construction-action-btn"
                   title="사업자등록증 삭제 (휴지통 이동)"
                   onclick="window.invoiceLicense && window.invoiceLicense.deleteLicense('${attr(c)}')">
             <i class="fas fa-trash"></i><span>삭제</span>
           </button>`
        : `<button type="button" class="btn btn-outline-secondary btn-sm construction-action-btn"
                   title="사업자등록증 업로드 (이미지·PDF)"
                   onclick="window.invoiceLicense && window.invoiceLicense.pickFile('${attr(c)}')">
             <i class="fas fa-upload"></i><span>등록증 업로드</span>
           </button>`;
    }
    const invBtn = document.querySelector(`[data-il-invoice="${attr(c)}"]`);
    if (invBtn) {
      const blocked = !!(d.required && !d.exists);
      // 비활성화하지 않는다 — disabled면 pointer-events:none 이라 클릭·hover 둘 다 죽는다.
      // 눌러서 '등록증 필요' 안내 모달(openInvoiceModal 가드)을 띄우는 편이 발견성이 좋다.
      invBtn.disabled = false;
      invBtn.classList.remove('disabled');
      invBtn.title = blocked ? '사업자등록증 업로드 후 요청 가능 — 눌러서 안내 확인' : '세금계산서 발행 요청';
    }
  }

  // ─────────────────────── ② 업로드 / 삭제 ───────────────────────
  pickFile(code) {
    const input = document.querySelector(`[data-il-file="${attr(code)}"]`);
    if (input) input.click();
  }

  onFilePicked(code, inputEl) {
    const file = inputEl && inputEl.files && inputEl.files[0];
    if (inputEl) inputEl.value = '';   // 같은 파일 재선택 시에도 change 발생하도록 초기화
    if (!file) return;
    // 서버 업로드 상한(프로덕션 MAX_CONTENT_LENGTH 8MB) — 클라이언트 사전 차단.
    const MAX = 8 * 1024 * 1024;
    if (file.size > MAX) {
      this._toast(`파일이 너무 큽니다 (${(file.size / 1024 / 1024).toFixed(1)}MB) — 8MB 이하로 올려주세요.`, 'error');
      return;
    }
    this._upload(code, file, false);
  }

  async _upload(code, file, force) {
    const c = String(code || '').trim();
    const badge = document.querySelector(`[data-il-status="${attr(c)}"]`);
    if (badge) {
      badge.innerHTML = '';
      badge.textContent = '업로드 중…';
      badge.style.background = '#eef0f2'; badge.style.color = '#6b7280'; badge.style.padding = '3px 10px';
    }

    let res;
    try {
      const fd = new FormData();
      fd.append('code', c);
      fd.append('file', file, file.name);
      if (force) fd.append('force', '1');
      const resp = await fetch('/api/invoice/license/upload', {
        method: 'POST', credentials: 'same-origin', body: fd,
      });
      res = await resp.json().catch(() => ({}));
      if (!resp.ok || res.success === false) {
        const e = res && res.error;
        const msg = (e && (typeof e === 'object' ? e.message : e)) || `업로드 실패 (HTTP ${resp.status})`;
        this._toast(msg, 'error');
        await this.refreshStatus(c);
        return;
      }
    } catch (e) {
      logger.debug('[InvoiceLicense] 업로드 오류:', e);
      this._toast('네트워크 오류로 업로드에 실패했습니다.', 'error');
      await this.refreshStatus(c);
      return;
    }

    const d = res.data || {};
    if (d.warned && !d.saved) {
      // OCR 게이트 경고 — 확인 후 강제 저장.
      const ok = window.confirm(`${d.warning}\n\n그래도 이 파일을 사업자등록증으로 등록할까요?`);
      if (ok) { await this._upload(c, file, true); }
      else { await this.refreshStatus(c); }
      return;
    }
    this._toast('사업자등록증을 저장했습니다.', 'success');
    await this.refreshStatus(c);
  }

  async deleteLicense(code) {
    const c = String(code || '').trim();
    if (!c) return;
    if (!window.confirm('사업자등록증을 삭제할까요?\n(Drive 휴지통으로 이동 — 필요 시 복구 가능)')) return;
    const r = await this._postRaw('/api/invoice/license/delete', { code: c });
    if (!r.ok) { this._toast(r.message || '삭제에 실패했습니다.', 'error'); return; }
    this._toast('사업자등록증을 삭제했습니다 (휴지통).', 'success');
    await this.refreshStatus(c);
  }

  // ─────────────────────── ③ 세금계산서 요청 ───────────────────────
  openInvoiceModal(code) {
    const c = String(code || '').trim();
    const d = this._status[c] || {};
    if (d.required && !d.exists) {
      this._showAlert('사업자등록증 필요',
        '이 프로젝트는 세금계산서 발행 전에 사업자등록증이 필요합니다.\n먼저 [등록증 업로드]로 파일을 올려주세요.');
      return;
    }
    const p = this._findProject(c) || {};
    // 시트 빈값 placeholder '-' 는 빈칸으로 (슬랙 _build_invoice_button_value 와 동일).
    const clean = (v) => { const s = String(v == null ? '' : v).trim(); return s === '-' ? '' : s; };
    const biz = clean(p['사업자명']) || clean(p['사업자']);
    const addr = clean(p['현장 주소']) || clean(p['현장주소']);
    // 총액 1 이 float('9900000.0')로 와도 '.0'의 0이 붙어 10배 되지 않게 정수부만 추출.
    // (digit-join 은 "9900000.0"→"99000000" 10배 버그. 슬랙 _build_invoice_button_value 의
    //  int(float()) 와 동일한 방어. 콤마 문자열·숫자·소수 문자열 모두 안전.)
    const amtNum = Number(String(p['총액 1'] ?? '').replace(/,/g, ''));
    const amtDigits = (Number.isFinite(amtNum) && amtNum > 0) ? String(Math.trunc(amtNum)) : '';
    const amtDisp = amtDigits ? Number(amtDigits).toLocaleString('ko-KR') : '';
    // 이메일: 서버가 해석한 값(거래처 탭 이메일 우선 > 발주처, 슬랙과 동일) → 없으면 발주처 fallback.
    const email = clean(d.email) || clean(p['발주처 이메일']);
    const vatRaw = String(p['부가세'] || '').trim();
    // 시트 부가세 truthy = 'VAT 별도'(sep). 슬랙 _build_invoice_button_value 와 동일 규칙.
    const vatSep = /^(true|y|yes|1|별도|vat\s*별도)$/i.test(vatRaw);

    const body = `
      <div class="mb-2"><label class="form-label">사업자명</label>
        <input id="ilBiz" type="text" class="form-control" value="${esc(biz)}" placeholder="예: (주)아이티글로벌"></div>
      <div class="mb-2"><label class="form-label">현장 주소</label>
        <input id="ilAddr" type="text" class="form-control" value="${esc(addr)}" placeholder="현장 주소"></div>
      <div class="row">
        <div class="col mb-2"><label class="form-label">발행 금액 <span class="text-muted small">(원)</span></label>
          <input id="ilAmt" type="text" inputmode="numeric" class="form-control" value="${esc(amtDisp)}" placeholder="예: 9,600,000"></div>
        <div class="col mb-2"><label class="form-label">부가세</label>
          <select id="ilVat" class="form-select">
            <option value="sep" ${vatSep ? 'selected' : ''}>VAT 별도</option>
            <option value="incl" ${vatSep ? '' : 'selected'}>VAT 포함</option>
          </select></div>
      </div>
      <div class="mb-2"><label class="form-label">이메일 <span class="text-muted small">(계산서 수신)</span></label>
        <input id="ilEmail" type="text" class="form-control" value="${esc(email)}" placeholder="example@company.com"></div>
      <div class="mb-1"><label class="form-label">요청사항 <span class="text-muted small">(선택)</span></label>
        <textarea id="ilMemo" class="form-control" rows="2" placeholder="수정발행·특이사항 등"></textarea></div>
    `;

    this._showModal(`세금계산서 발행 요청 — ${c}`, body, async (el) => {
      const amt = (el.querySelector('#ilAmt').value || '').replace(/[^\d]/g, '');
      const payload = {
        code: c,
        biz: (el.querySelector('#ilBiz').value || '').trim(),
        addr: (el.querySelector('#ilAddr').value || '').trim(),
        amt,
        vat: el.querySelector('#ilVat').value || 'sep',
        email: (el.querySelector('#ilEmail').value || '').trim(),
        memo: (el.querySelector('#ilMemo').value || '').trim(),
      };
      const err = await this._postRaw('/api/invoice/request', payload);
      if (!err.ok) return err.message;
      this._toast('세금계산서 요청을 #계산서_관리로 발송했습니다.', 'success');
      return null;
    }, { icon: 'fa-file-invoice-dollar', submitLabel: '요청 발송' });
  }

  // POST(JSON) → {ok, code, message}. APIResponse error.{code,message} 파싱.
  async _postRaw(url, payload) {
    try {
      const resp = await fetch(url, {
        method: 'POST', credentials: 'same-origin',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload || {}),
      });
      const json = await resp.json().catch(() => ({}));
      const ok = resp.ok && json.success !== false;
      const e = json.error;
      const isObj = e && typeof e === 'object';
      const message = (isObj ? e.message : e) || json.message || (ok ? '' : `요청 실패 (HTTP ${resp.status})`);
      return { ok, code: (isObj ? e.code : '') || '', message };
    } catch (_) {
      return { ok: false, code: 'NETWORK', message: '네트워크 오류로 실패했습니다.' };
    }
  }

  // ─────────────────────── 모달 (A/S 와 동일 룩) ───────────────────────
  _showModal(title, bodyHtml, onSubmit, opts = {}) {
    const host = this._modalHost();
    const icon = opts.icon || 'fa-file-invoice';
    const submitLabel = opts.submitLabel || '확인';
    host.innerHTML = `
      <div class="modal fade" id="ilActionModal" tabindex="-1">
        <div class="modal-dialog modal-dialog-centered" style="max-width: 520px;"><div class="modal-content">
          <div class="modal-header" style="background-color:#fafbfc; border-bottom:1px solid var(--gray-200); padding:1.1rem 1.25rem;">
            <h5 class="modal-title" style="font-weight:600; color:var(--gray-900); display:flex; align-items:center; gap:0.5rem;">
              <i class="fas ${icon}" style="color:#17a2b8;"></i>${esc(title)}</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
          <div class="modal-body">${bodyHtml}<div id="ilModalAlert" class="text-danger small mt-2"></div></div>
          <div class="modal-footer" style="background-color:var(--gray-50); border-top:1px solid var(--gray-200); padding:1rem 1.25rem;">
            <button type="button" class="btn as-btn-cancel" data-bs-dismiss="modal">취소</button>
            <button type="button" class="btn as-btn-submit" id="ilModalSubmit">${esc(submitLabel)}</button>
          </div>
        </div></div>
      </div>`;
    const el = host.querySelector('#ilActionModal');
    this._modal = new bootstrap.Modal(el);
    const submitBtn = el.querySelector('#ilModalSubmit');
    submitBtn.addEventListener('click', async () => {
      submitBtn.disabled = true;
      const err = await onSubmit(el);
      submitBtn.disabled = false;
      if (err) { el.querySelector('#ilModalAlert').textContent = err; return; }
      this._modal.hide();
    });
    this._modal.show();
  }

  _showAlert(title, message, opts = {}) {
    const host = this._modalHost();
    const icon = opts.icon || 'fa-triangle-exclamation';
    const color = opts.iconColor || '#e0a800';
    host.innerHTML = `
      <div class="modal fade" id="ilActionModal" tabindex="-1">
        <div class="modal-dialog modal-dialog-centered" style="max-width: 460px;"><div class="modal-content">
          <div class="modal-header" style="background-color:#fafbfc; border-bottom:1px solid var(--gray-200); padding:1.1rem 1.25rem;">
            <h5 class="modal-title" style="font-weight:600; color:var(--gray-900); display:flex; align-items:center; gap:0.5rem;">
              <i class="fas ${icon}" style="color:${color};"></i>${esc(title)}</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
          <div class="modal-body" style="padding:1.25rem;">
            <div style="display:flex; gap:0.75rem; align-items:flex-start;">
              <i class="fas ${icon}" style="color:${color}; font-size:1.6rem; line-height:1; margin-top:0.1rem;"></i>
              <div style="font-size:var(--font-size-sm); line-height:1.55; color:var(--gray-800);">${esc(message).replace(/\n/g, '<br>')}</div>
            </div>
          </div>
          <div class="modal-footer" style="background-color:var(--gray-50); border-top:1px solid var(--gray-200); padding:0.85rem 1.25rem;">
            <button type="button" class="btn as-btn-submit" data-bs-dismiss="modal">확인</button>
          </div>
        </div></div>
      </div>`;
    const el = host.querySelector('#ilActionModal');
    this._modal = new bootstrap.Modal(el);
    this._modal.show();
  }
}
