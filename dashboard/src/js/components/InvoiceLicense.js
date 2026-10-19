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
import { activeInvoiceLines, overIssuedAmount, sameBizName } from '../utils/billStatus.js';
import { stagePayers } from '../utils/payerSplit.js';

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
    this._issuedP = {};  // code → Promise(이미발행 요약) — 아코디언 렌더 때 미리 조회(모달 즉시 오픈)
    this._payersP = {};  // code → Promise(부담자 창 초기값) — 같은 방식(2026-10-20 '부담자 설정 느림' 제보)
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
    // 계산서 요청 모달용 이미발행 요약도 미리 (버튼 클릭 시 조회 대기로 모달이 늦게 뜨던 문제)
    this._issuedP[c] = this._fetchIssued(c);
    // [부담자 설정] 버튼이 보이는 사람(편집자 이상)만 — 창 초기값 미리
    if (['editor', 'admin', 'super_admin'].includes(window.userRole || 'viewer')) {
      this._payersP[c] = this._fetchPayers(c);
    }
  }

  /** 부담자 창 초기값 {payers, total2, stage_values}. 실패·행 없음이면 null. */
  async _fetchPayers(code) {
    try {
      const resp = await fetch(`/api/invoice/payers/${encodeURIComponent(code)}`, { credentials: 'same-origin' });
      const json = await resp.json().catch(() => ({}));
      return (resp.ok && json.success !== false && json.data) ? json.data : null;
    } catch (e) {
      logger.debug('[InvoiceLicense] 부담자 조회 실패:', e);
      return null;
    }
  }

  /** 이미 발행된 계산서 요약(슬랙 모달과 같은 서버 계산). 실패 시 null → 모달은 기존 동작. */
  async _fetchIssued(code) {
    try {
      const resp = await fetch(`/api/invoice/issued/${encodeURIComponent(code)}`, { credentials: 'same-origin' });
      const json = await resp.json().catch(() => ({}));
      return (resp.ok && json.success !== false && json.data) ? json.data : null;
    } catch (e) {
      logger.debug('[InvoiceLicense] 이미발행 조회 실패:', e);
      return null;
    }
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
              class="text-decoration-none" style="color:#0d6efd;"
              title="사업자등록증 열람 (새 탭)">사업자등록증 열기</a>`;
      } else if (d.source === 'partner') {
        // 등록증 파일은 없지만 거래처 탭에 상호 있음 → 계산서 발행 가능 (긍정 상태, 열람 링크 없음)
        holder.innerHTML = '';
        holder.textContent = '첨부되지 않음 (기존 거래처 정보 재활용)';
        holder.style.background = '#e3f4e8';
        holder.style.color = '#1c7c3d';
        holder.style.padding = '3px 10px';
        holder.title = '거래처 탭에 사업자 정보가 있어 계산서 발행 가능 (등록증 파일 없음 — 업로드하면 열람·첨부 가능)';
      } else {
        // 없음 / 불필요(거래처) → pill 뱃지
        let label, bg, fg;
        if (d.exists) { label = '등록증 있음'; bg = '#e3f4e8'; fg = '#1c7c3d'; }  // view_url 없을 때 fallback
        else if (!d.required) { label = '등록증 불필요 (거래처)'; bg = '#eef0f2'; fg = '#6b7280'; }
        else { label = '첨부되지 않음'; bg = '#fdeede'; fg = '#b45309'; }
        holder.innerHTML = '';
        holder.textContent = label;
        holder.style.background = bg;
        holder.style.color = fg;
        holder.style.padding = '3px 10px';
      }
    }
    // 편집 모드 컨트롤(업로드/삭제) — 내가 올린 파일(source==='own')일 때만 [삭제], 그 외 [업로드].
    // reuse(다른 폴더 재활용)·partner(거래처 이력, 파일 없음)는 이 프로젝트에 삭제할 파일이 없음.
    const edit = document.querySelector(`[data-il-edit="${attr(c)}"]`);
    if (edit) {
      edit.innerHTML = (d.source === 'own')
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
  async openInvoiceModal(code) {
    const c = String(code || '').trim();
    const d = this._status[c] || {};
    if (d.required && !d.exists) {
      this._showAlert('사업자등록증 필요',
        '이 프로젝트는 세금계산서 발행 전에 사업자등록증이 필요합니다.\n먼저 [등록증 업로드]로 파일을 올려주세요.');
      return;
    }
    // 이미 발행된 계산서 → 슬랙 모달과 동일하게 안내 + 발행 금액 '잔여 공급가' 프리필 (전액 재요청 방지)
    //   아코디언 렌더 때 미리 받아둔 결과 사용 → 보통 즉시. 없으면 이때 조회.
    const iss = await (this._issuedP[c] || (this._issuedP[c] = this._fetchIssued(c)));
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
    const remaining = iss && iss.remaining_supply != null ? Number(iss.remaining_supply) : null;
    const amtDisp = remaining != null ? remaining.toLocaleString('ko-KR')
      : (amtDigits ? Number(amtDigits).toLocaleString('ko-KR') : '');
    // 이메일: 서버가 해석한 값(거래처 탭 이메일 우선 > 발주처, 슬랙과 동일) → 없으면 발주처 fallback.
    const email = clean(d.email) || clean(p['발주처 이메일']);
    const vatRaw = String(p['부가세'] || '').trim();
    // 시트 부가세 truthy = 'VAT 별도'(sep). 슬랙 _build_invoice_button_value 와 동일 규칙.
    const vatSep = /^(true|y|yes|1|별도|vat\s*별도)$/i.test(vatRaw);

    // 계산서 발행 단계 — 단일 선택. 기본값 없음(placeholder) → 요청자가 직접 선택.
    //   드롭다운은 귀속 '단계'만 표시(입금액 미표시 — 그 금액으로 한정된 것처럼 보이지
    //   않게). 세금계산서는 입금 전 선발행 가능 → 금액 없는 단계도 선택 허용(슬랙과 동일),
    //   선발행이면 아래 힌트로 조용히 안내. 발행 금액은 '발행 금액' 필드에서 자유 입력.
    const stageInfo = ['계약금', '중도금', '잔금'].map((s) => {
      const amt = Number(String(p[s] ?? '').replace(/,/g, '')) || 0;
      const tok = String(p[`${s} 계산서`] ?? '').trim();
      const issued = tok === '발행' || tok === '일반';
      return { s, amt, issued };
    });
    // 나눠 내는 공사(2026-10-19): 단계 = 부담자 칸 → 목록에 부담자 이름, 고르면 그 사람 정보로 채움
    const payers = stagePayers(p);
    const payerPrefill = (d && d.payers) || {};
    const stageOptions = ['<option value="" selected disabled>단계 선택</option>'].concat(
      stageInfo.map((x) => {
        const issuedTxt = x.issued ? ' · 발행됨' : '';
        const py = payers[x.s];
        const who = py ? ` · ${py.role} ${py.name}` : '';
        return `<option value="${x.s}">${esc(`${x.s}${who}${issuedTxt}`)}</option>`;
      }),
    ).join('');

    // 수정발행 — 대체할 기존 계산서 (2026-10-01, 슬랙 모달과 동일). 장 금액이 기록된 현재 유효
    //   계산서만 후보. 체크한 장은 첨부 완료 시 메모에 '취소' 줄로 남아 현재 계산서에서 빠짐.
    //   후보는 서버 목록 우선(요청 시 서버 대조와 같은 기준), 조회 실패면 화면 데이터로 계산.
    const active = iss ? (iss.active || [])
      : activeInvoiceLines(p['계산서_메모']).filter((it) => it.amt > 0).slice(0, 10);
    // 총액 초과 경고는 서버 안내 문구에 포함 → 서버 조회 실패 때만 화면에서 계산
    const over = iss ? 0 : overIssuedAmount(p);
    const issuedHtml = iss && iss.warn ? `
      <div class="mb-2 small" style="background:#f8f9fa; border:1px solid var(--gray-200); border-radius:6px; padding:0.5rem 0.65rem; white-space:pre-line;">${esc(iss.warn)}</div>` : '';
    const replaceHtml = active.length ? `
      <div class="mb-2"><label class="form-label">기존 발행 계산서 <span class="text-muted small">(수정발행 시 선택)</span></label>
        ${active.map((it, i) => `
          <div class="form-check">
            <input class="form-check-input" type="checkbox" id="ilRep${i}" data-rep="${i}">
            <label class="form-check-label small" for="ilRep${i}">${esc(`${it.date ? it.date.slice(5) : '날짜없음'} ${it.stage} ${it.amt.toLocaleString()}원${it.biz ? ` · ${it.biz}` : ''}`)}</label>
          </div>`).join('')}
        <div class="small text-muted mt-1">🔁 기존 계산서 수정발행이 필요한 경우에만 체크</div>
        ${over ? `<div class="small mt-1" style="color:#d9480f;">⚠️ 발행 계산서 합계 ${over.toLocaleString()}원이 총액을 넘습니다 — 이전 수정발행의 대체 표시가 빠졌는지 확인하세요.</div>` : ''}
      </div>` : '';

    // 부담자 설정이 안 된 나눠 내는 공사 의심 — 이미 다른 사업자로 발행된 계산서가 있으면 안내 (2026-10-19).
    //   과거 기록상 이런 공사는 드물어(2건) 엉뚱하게 뜨는 일이 거의 없다.
    const otherBiz = !Object.keys(payers).length && biz
      && active.some((it) => it.biz && !sameBizName(it.biz, biz));
    const payerHint = otherBiz ? `
      <div class="small mb-2" style="color:#0a7ea4;">👥 이 공사에 다른 사업자로 발행된 계산서가 있습니다.
        여러 사람이 나눠 내는 공사라면 <a href="#" data-pl-open>부담자 설정</a>을 먼저 해 주세요.</div>` : '';

    const body = `
      <div class="mb-2"><label class="form-label">사업자명</label>
        <input id="ilBiz" type="text" class="form-control" value="${esc(biz)}" placeholder="예: (주)아이티글로벌"></div>
      <div class="mb-2"><label class="form-label">현장 주소</label>
        <input id="ilAddr" type="text" class="form-control" value="${esc(addr)}" placeholder="현장 주소"></div>
      ${payerHint}
      ${issuedHtml}
      ${replaceHtml}
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
      <div class="mb-2"><label class="form-label">발행 단계 <span class="text-muted small">(금액=총액이면 나머지 자동 '-')</span></label>
        <select id="ilStage" class="form-select">${stageOptions}</select>
        <div id="ilStageHint" class="small mt-1" style="display:none; color:#0a7ea4;"></div></div>
      <div class="mb-1"><label class="form-label">요청사항 <span class="text-muted small">(선택)</span></label>
        <textarea id="ilMemo" class="form-control" rows="2" placeholder="청구/영수 발행·특이사항 등"></textarea></div>
    `;

    this._showModal(`세금계산서 발행 요청 — ${c}`, body, async (el) => {
      const amt = (el.querySelector('#ilAmt').value || '').replace(/[^\d]/g, '');
      const stage = (el.querySelector('#ilStage') || {}).value || '';
      if (!stage) return '발행 단계를 선택해주세요.';
      const payload = {
        code: c,
        biz: (el.querySelector('#ilBiz').value || '').trim(),
        addr: (el.querySelector('#ilAddr').value || '').trim(),
        amt,
        vat: el.querySelector('#ilVat').value || 'sep',
        email: (el.querySelector('#ilEmail').value || '').trim(),
        memo: (el.querySelector('#ilMemo').value || '').trim(),
        stages: stage,
        replace: [...el.querySelectorAll('input[data-rep]:checked')]
          .map((cb) => active[Number(cb.dataset.rep)]).filter(Boolean),
      };
      let err = await this._postRaw('/api/invoice/request', payload);
      // 금액 사전 점검 경고(부가세 포함 금액 의심·총액 초과) — 확인하면 그대로 다시 보냄 (2026-10-15)
      if (!err.ok && err.code === 'INVOICE_AMOUNT_CHECK') {
        if (!window.confirm(`${err.message}\n\n이 금액 그대로 요청할까요?`)) return err.message;
        err = await this._postRaw('/api/invoice/request', { ...payload, confirm_amount: true });
      }
      if (!err.ok) return err.message;
      delete this._issuedP[c];   // 다음에 열 땐 새로 조회
      this._toast('세금계산서 요청을 #계산서_관리로 발송했습니다.', 'success');
      return null;
    }, {
      icon: 'fa-file-invoice-dollar', submitLabel: '요청 발송',
      // 선발행 안내(최소 가이드): 입금 없는 단계를 고르면 조용히 힌트만 — 차단하지 않음.
      onRender: (el) => {
        const plOpen = el.querySelector('[data-pl-open]');
        if (plOpen) {
          plOpen.addEventListener('click', (ev) => {
            ev.preventDefault();
            el.addEventListener('hidden.bs.modal', () => this.openPayerModal(c), { once: true });
            this._modal.hide();
          });
        }
        const sel = el.querySelector('#ilStage');
        const hint = el.querySelector('#ilStageHint');
        if (!sel || !hint) return;
        const amtByStage = Object.fromEntries(stageInfo.map((x) => [x.s, x.amt]));
        const issuedByStage = Object.fromEntries(stageInfo.map((x) => [x.s, x.issued]));
        sel.addEventListener('change', () => {
          const paid = amtByStage[sel.value] || 0;
          const notes = [];
          // 나눠 내는 공사: 부담자 칸을 고르면 사업자명·이메일·발행 금액(부담액)을 그 사람 것으로
          const py = payers[sel.value];
          if (py) {
            const pre = payerPrefill[sel.value] || {};
            const bizEl = el.querySelector('#ilBiz');
            const emailEl = el.querySelector('#ilEmail');
            const amtEl = el.querySelector('#ilAmt');
            const vatEl = el.querySelector('#ilVat');
            if (bizEl) bizEl.value = pre.name || py.name;
            if (emailEl) emailEl.value = pre.email || py.email || '';
            if (issuedByStage[sel.value]) {
              notes.push(`🧾 ${py.role} ${py.name} 계산서는 이미 발행됐습니다. 수정발행이면 위에서 기존 계산서를 체크하세요.`);
            } else if (amtEl && vatEl) {
              vatEl.value = 'sep';
              amtEl.value = Math.round(py.amount / 1.1).toLocaleString('ko-KR');
              notes.push(`👥 ${py.role} ${py.name} 부담액 ${py.amount.toLocaleString()}원(부가세 포함) 기준으로 채웠습니다.`);
            }
            if (py.method === '현금') notes.push('💵 이 부담자는 계산서 없이 현금으로 내기로 돼 있습니다. 맞는지 확인하세요.');
          }
          if (sel.value && paid <= 0) {
            notes.push('🕒 선발행 — 이 단계는 아직 입금 내역이 없습니다. 발행 금액을 확인하세요.');
          }
          hint.textContent = notes.join('\n');
          hint.style.whiteSpace = 'pre-line';
          hint.style.display = notes.length ? '' : 'none';
        });
      },
    });
  }

  // ─────────────────────── ④ 부담자 설정 (나눠 내는 공사) ───────────────────────
  /** 한 공사를 여러 사람(임대인·임차인 등)이 나눠 내는 경우의 부담자 설정 — 매니저도 사용 (2026-10-19).
   *  사람은 고르기·숫자만 입력하고, 서버가 칸(첫째=계약금·둘째=잔금·셋째=중도금)을 정해 '○○ 부담자' 칸과
   *  칸 메모 맨 윗줄('◆ 임대인 이희승')을 형식대로 기록한다. */
  async openPayerModal(code, opts = {}) {
    const c = String(code || '').trim();
    // 아코디언을 펼칠 때 미리 받아 둔 값 → 클릭 즉시 열림. 저장 때 서버가 시트 최신값으로 다시 검사한다.
    let info = opts.waitForRow ? null : await (this._payersP[c] || (this._payersP[c] = this._fetchPayers(c)));
    if (!info) delete this._payersP[c];   // 실패한 조회는 재사용하지 않음
    // 새 공사 등록 직후(waitForRow)는 시트 기록이 백그라운드라 행이 아직 없을 수 있음 → 최대 30초 재시도
    const tries = info ? 0 : (opts.waitForRow ? 15 : 1);
    for (let t = 0; t < tries && !info; t += 1) {
      if (t) await new Promise((res) => setTimeout(res, 2000));
      info = await this._fetchPayers(c);
    }
    if (!info) {
      this._toast('부담자 정보를 불러오지 못했습니다. 새로고침 후 다시 시도하세요.', 'error');
      return;
    }
    const total2 = Number(info.total2 || 0);
    const sv = info.stage_values || {};
    const fmt = (n) => (Number(n) || 0).toLocaleString('ko-KR');
    const STAGES3 = ['계약금', '중도금', '잔금'];
    const init = (info.payers && info.payers.length) ? info.payers
      : [{ role: '임대인', method: '계산서' }, { role: '임차인', method: '계산서' }];
    const rowHtml = (p) => `
      <div class="pl-row mb-2" style="border:1px solid var(--gray-200); border-radius:8px; padding:.6rem .7rem;">
        <div class="d-flex gap-2 mb-2">
          <input class="form-control form-control-sm" data-f="role" list="plRoleList" value="${esc(p.role || '')}"
                 placeholder="구분" style="max-width:96px;">
          <input class="form-control form-control-sm" data-f="name" value="${esc(p.name || '')}"
                 placeholder="사업자명 (거래처 탭에서 검색)" autocomplete="off">
          <button type="button" class="btn btn-sm btn-outline-danger" data-pl-del title="이 줄 빼기">
            <i class="fas fa-xmark"></i></button>
        </div>
        <div class="d-flex gap-2 flex-wrap">
          <input class="form-control form-control-sm" data-f="amount" inputmode="numeric"
                 value="${p.amount ? fmt(p.amount) : ''}" placeholder="부담액 (부가세 포함)" style="max-width:150px;">
          <select class="form-select form-select-sm" data-f="method" style="max-width:96px;">
            ${['계산서', '현금', '카드'].map((m) => `<option ${p.method === m ? 'selected' : ''}>${m}</option>`).join('')}
          </select>
          <select class="form-select form-select-sm" data-f="stage" style="max-width:112px;" title="이 사람 돈이 쌓일 칸">
            <option value="">칸 자동</option>
            ${STAGES3.map((s) => `<option value="${s}" ${p.stage === s ? 'selected' : ''}>${s} 칸</option>`).join('')}
          </select>
          <input class="form-control form-control-sm" data-f="biz_no" value="${esc(p.biz_no || '')}"
                 placeholder="사업자번호" style="max-width:130px;">
          <input class="form-control form-control-sm" data-f="email" value="${esc(p.email || '')}"
                 placeholder="계산서 받을 이메일" style="min-width:150px; flex:1;">
        </div>
        <div class="pl-suggest small" style="display:none; margin-top:.35rem;"></div>
      </div>`;
    const deposits = STAGES3.map((s) => `${s} ${fmt(sv[s])}`).join(' · ');
    const body = `
      <div class="small text-muted mb-2">한 공사를 여러 사람이 나눠 내는 경우에만 설정합니다.
        칸은 첫째 = 계약금, 둘째 = 잔금, 셋째 = 중도금으로 정해지고, 각자 낸 돈은 자기 칸에 쌓입니다.</div>
      <div class="small mb-2" style="background:#f8f9fa; border:1px solid var(--gray-200); border-radius:6px; padding:.4rem .6rem;">
        지금까지 입금 — ${deposits}원 · 공사 총액 ${fmt(total2)}원</div>
      <datalist id="plRoleList"><option value="임대인"></option><option value="임차인"></option></datalist>
      <div id="plRows">${init.map(rowHtml).join('')}</div>
      <button type="button" class="btn btn-sm btn-outline-secondary" id="plAdd"><i class="fas fa-plus"></i> 부담자 추가</button>
      <div id="plSum" class="small mt-2"></div>
      <div class="small text-muted mt-1">모두 빼고 저장하면 나눠 내기가 해제됩니다.</div>`;

    const collect = (el) => [...el.querySelectorAll('.pl-row')].map((r) => {
      const v = (f) => ((r.querySelector(`[data-f="${f}"]`) || {}).value || '').trim();
      return { role: v('role'), name: v('name'), amount: v('amount').replace(/[^\d]/g, ''),
        method: v('method'), stage: v('stage'), biz_no: v('biz_no'), email: v('email') };
    }).filter((p) => p.name || p.amount);

    const refreshSum = (el) => {
      const sum = collect(el).reduce((a, p) => a + (Number(p.amount) || 0), 0);
      const box = el.querySelector('#plSum');
      if (!box) return;
      const same = !total2 || Math.abs(sum - total2) <= 1;
      box.textContent = `부담액 합계 ${fmt(sum)}원 / 공사 총액 ${fmt(total2)}원${same ? '' : ` (차이 ${fmt(total2 - sum)}원)`}`;
      box.style.color = same ? '#1c7c3d' : '#d9480f';
    };

    const bindRow = (el, r) => {
      const del = r.querySelector('[data-pl-del]');
      if (del) del.addEventListener('click', () => { r.remove(); refreshSum(el); });
      const amt = r.querySelector('[data-f="amount"]');
      if (amt) {
        amt.addEventListener('input', () => refreshSum(el));
        amt.addEventListener('blur', () => {
          const n = Number(String(amt.value).replace(/[^\d]/g, ''));
          amt.value = n ? fmt(n) : '';
        });
      }
      // 거래처 탭 자동완성 — 고르면 사업자명·사업자번호·이메일 채움
      const name = r.querySelector('[data-f="name"]');
      const sug = r.querySelector('.pl-suggest');
      let timer = null;
      if (name && sug) {
        name.addEventListener('input', () => {
          clearTimeout(timer);
          const q = name.value.trim();
          if (q.length < 2) { sug.style.display = 'none'; return; }
          timer = setTimeout(async () => {
            try {
              const resp = await fetch(`/api/invoice/partner-search?q=${encodeURIComponent(q)}`, { credentials: 'same-origin' });
              const json = await resp.json().catch(() => ({}));
              const items = (json.data && json.data.items) || [];
              if (!items.length) { sug.style.display = 'none'; return; }
              sug.innerHTML = items.map((it, k) => `<a href="#" data-k="${k}" class="d-block text-decoration-none"
                  style="padding:2px 4px;">${esc(it.name)} <span class="text-muted">${esc(it.bno || '')}${it.ceo ? ` · ${esc(it.ceo)}` : ''}</span></a>`).join('');
              sug.style.display = '';
              sug.querySelectorAll('a[data-k]').forEach((a) => a.addEventListener('click', (ev) => {
                ev.preventDefault();
                const it = items[Number(a.dataset.k)] || {};
                name.value = it.name || name.value;
                const bn = r.querySelector('[data-f="biz_no"]');
                const em = r.querySelector('[data-f="email"]');
                if (bn && it.bno) bn.value = it.bno;
                if (em && it.email) em.value = it.email;
                sug.style.display = 'none';
              }));
            } catch (_) { sug.style.display = 'none'; }
          }, 250);
        });
      }
    };

    this._showModal(`부담자 설정 — ${c}`, body, async (el) => {
      const payers = collect(el);
      let r = await this._postRaw(`/api/invoice/payers/${encodeURIComponent(c)}`, { payers });
      if (!r.ok && r.code === 'PAYER_TOTAL_CHECK') {
        if (!window.confirm(`${r.message}\n\n그대로 저장할까요?`)) return r.message;
        r = await this._postRaw(`/api/invoice/payers/${encodeURIComponent(c)}`, { payers, confirm_total: true });
      }
      if (!r.ok) return r.message;
      this._payersP[c] = this._fetchPayers(c);   // 저장한 값으로 다음 열기 준비
      const fields = (r.data && r.data.fields) || {};
      try {
        if (window.projectRowAccordion && Object.keys(fields).length) {
          window.projectRowAccordion.applyExternalFields(c, fields);
        }
      } catch (_) { /* noop */ }
      this._toast(payers.length ? '부담자를 저장했습니다.' : '나눠 내기를 해제했습니다.', 'success');
      return null;
    }, {
      icon: 'fa-people-arrows', submitLabel: '저장', maxWidth: 640,
      onRender: (el) => {
        el.querySelectorAll('.pl-row').forEach((r) => bindRow(el, r));
        refreshSum(el);
        const add = el.querySelector('#plAdd');
        if (add) {
          add.addEventListener('click', () => {
            const rows = el.querySelector('#plRows');
            if (!rows || rows.querySelectorAll('.pl-row').length >= 3) {
              this._toast('부담자는 3명까지 정할 수 있습니다.', 'warning');
              return;
            }
            rows.insertAdjacentHTML('beforeend', rowHtml({ method: '계산서' }));
            bindRow(el, rows.lastElementChild);
          });
        }
      },
    });
  }

  // POST(JSON) → {ok, code, message, data}. APIResponse error.{code,message} 파싱.
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
      return { ok, code: (isObj ? e.code : '') || '', message, data: json.data };
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
        <div class="modal-dialog modal-dialog-centered" style="max-width: ${Number(opts.maxWidth) || 520}px;"><div class="modal-content">
          <div class="modal-header" style="background-color:#fafbfc; border-bottom:1px solid var(--gray-200); padding:1.1rem 1.25rem;">
            <h5 class="modal-title" style="font-weight:600; color:var(--gray-900); display:flex; align-items:center; gap:0.5rem;">
              <i class="fas ${icon}" style="color:#17a2b8;"></i>${esc(title)}</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
          <div class="modal-body">${bodyHtml}<div id="ilModalAlert" class="text-danger small mt-2" style="white-space:pre-line;"></div></div>
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
    if (typeof opts.onRender === 'function') { try { opts.onRender(el); } catch (_) {} }
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
