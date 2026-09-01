/**
 * AsView — A/S 데이터 조인맵 + 액션 모달 (경량 모듈).
 *
 * 테이블 전환·컬럼·필터는 ProjectTable(수금 모드와 동일 구조)이 담당한다.
 * 이 모듈은 (1) /as/api/list 를 프로젝트 코드→A/S 로 조인한 byCode 맵 제공,
 * (2) 접수/완료/요청/수동요청 Bootstrap 모달 + POST(슬랙 동기화) 만 담당.
 * ProjectTable 의 A/S 컬럼 render 와 accordion 'A/S 요청' 버튼이 window.asView 를 참조.
 */
import logger from '../utils/logger.js';

function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function isOpen(status) {
  const s = String(status || '').trim();
  return s === '요청됨' || s === '접수 완료';
}

export default class AsView {
  constructor() {
    this.byCode = {};     // 프로젝트 코드 → 대표 A/S 건 (열린 것 우선, 그다음 최신)
    this.byNo = {};       // AS 번호 → A/S 건
    this._modal = null;
    window.asView = this; // ProjectTable render / accordion 버튼에서 참조
  }

  /** /as/api/list → byCode/byNo 조인맵 구성 */
  async loadMap() {
    try {
      const resp = await fetch('/as/api/list', { credentials: 'same-origin' });
      const json = await resp.json();
      const data = json.data || json;
      const items = (data && data.items) || [];
      const byCode = {};
      const byNo = {};
      // AS 번호 오름차순 → 뒤(최신)가 덮어씀. 단 열린 건을 완료 건보다 우선.
      const num = (it) => { const m = /^AS-(\d+)/.exec(String(it.No || '')); return m ? parseInt(m[1], 10) : 0; };
      items.slice().sort((a, b) => num(a) - num(b)).forEach((it) => {
        byNo[it.No] = it;
        const code = String(it['프로젝트 코드'] || '').trim();
        if (!code) return;
        const cur = byCode[code];
        if (!cur) { byCode[code] = it; return; }
        const iOpen = isOpen(it['진행 상태']);
        const cOpen = isOpen(cur['진행 상태']);
        if (iOpen && !cOpen) byCode[code] = it;       // 열린 건 우선
        else if (iOpen === cOpen) byCode[code] = it;   // 동급이면 최신(뒤) 우선
      });
      this.byCode = byCode;
      this.byNo = byNo;
      logger.debug(`[AsView] A/S 조인맵 로드: ${items.length}건 → ${Object.keys(byCode).length}개 프로젝트`);
    } catch (err) {
      logger.error('[AsView] A/S 목록 로드 실패:', err);
      this.byCode = {};
      this.byNo = {};
    }
    // 열려있는 아코디언의 A/S 액션 버튼을 최신 상태로 갱신 (요청→접수→완료 전환 반영)
    this._refreshAccordionButtons();
  }

  /** 렌더된 모든 A/S 액션 슬롯을 현재 byCode 기준으로 다시 그림 */
  _refreshAccordionButtons() {
    try {
      const acc = window.projectRowAccordion;
      if (!acc || typeof acc.generateAsRequestButton !== 'function') return;
      document.querySelectorAll('.as-action-slot[data-project-code]').forEach((slot) => {
        slot.innerHTML = acc.generateAsRequestButton(slot.getAttribute('data-project-code'));
      });
    } catch (err) {
      logger.debug('[AsView] 아코디언 A/S 버튼 갱신 skip:', err);
    }
  }

  /** 액션 후 조인맵 재로드 + A/S 모드면 필터(has-A/S) 재적용 + 컬럼 재그림 */
  async refresh() {
    await this.loadMap();
    const inst = window.__projectTableInstance;
    if (inst && inst._asModeActive) {
      if (window.modernFilters && window.modernFilters.applyFilters) {
        window.modernFilters.applyFilters(null, true);
      } else if (inst.table) {
        try { inst.table.rows().invalidate().draw(false); } catch (_) { inst.table.draw(false); }
      }
    }
  }

  // ── Bootstrap 모달 ────────────────────────────────────────────
  _modalHost() {
    let host = document.getElementById('asModalHost');
    if (!host) {
      host = document.createElement('div');
      host.id = 'asModalHost';
      document.body.appendChild(host);
    }
    return host;
  }

  _showModal(title, bodyHtml, onSubmit, opts = {}) {
    const host = this._modalHost();
    const icon = opts.icon || 'fa-screwdriver-wrench';   // A/S 기본 아이콘
    const submitLabel = opts.submitLabel || '확인';
    host.innerHTML = `
      <div class="modal fade" id="asActionModal" tabindex="-1">
        <div class="modal-dialog modal-dialog-centered" style="max-width: 520px;"><div class="modal-content">
          <div class="modal-header" style="background-color:#fafbfc; border-bottom:1px solid var(--gray-200); padding:1.1rem 1.25rem;">
            <h5 class="modal-title" style="font-weight:600; color:var(--gray-900); display:flex; align-items:center; gap:0.5rem;">
              <i class="fas ${icon}" style="color:#17a2b8;"></i>${esc(title)}</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
          <div class="modal-body">${bodyHtml}<div id="asModalAlert" class="text-danger small mt-2"></div></div>
          <div class="modal-footer" style="background-color:var(--gray-50); border-top:1px solid var(--gray-200); padding:1rem 1.25rem;">
            <button type="button" class="btn as-btn-cancel" data-bs-dismiss="modal">취소</button>
            <button type="button" class="btn as-btn-submit" id="asModalSubmit">${esc(submitLabel)}</button>
          </div>
        </div></div>
      </div>`;
    const el = host.querySelector('#asActionModal');
    this._modal = new bootstrap.Modal(el);
    const submitBtn = el.querySelector('#asModalSubmit');
    submitBtn.addEventListener('click', async () => {
      submitBtn.disabled = true;
      const err = await onSubmit(el);
      submitBtn.disabled = false;
      if (err) { el.querySelector('#asModalAlert').textContent = err; return; }
      this._modal.hide();
      this.refresh();
    });
    this._modal.show();
  }

  /** 단순 안내/차단용 모달 (입력 없음, 확인 버튼 하나) — 페이지 토스트보다 눈에 띔 */
  _showAlert(title, message, opts = {}) {
    const host = this._modalHost();
    const icon = opts.icon || 'fa-triangle-exclamation';
    const color = opts.iconColor || '#e0a800';
    host.innerHTML = `
      <div class="modal fade" id="asActionModal" tabindex="-1">
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
    const el = host.querySelector('#asActionModal');
    this._modal = new bootstrap.Modal(el);
    this._modal.show();
  }

  /** POST → {ok, code, message, details}. APIResponse 의 error.{code,message,details} 구조를 정확히 파싱. */
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
      return { ok, code: (isObj ? e.code : '') || '', message, details: (isObj ? e.details : null) || null };
    } catch (_) {
      return { ok: false, code: 'NETWORK', message: '네트워크 오류로 실패했습니다.', details: null };
    }
  }

  async _post(url, payload) {
    const r = await this._postRaw(url, payload);
    return r.ok ? null : r.message;
  }

  openAccept(asNo) {
    this._showModal(`A/S 접수 — ${asNo}`, `
      <div class="mb-2"><label class="form-label">방문자 유형</label>
        <select id="asVisitorType" class="form-select">
          <option value="" selected disabled>선택</option>
          <option value="서비스 기사">서비스 기사</option>
          <option value="내부">내부 (아이티)</option>
          <option value="외주">외주 (시공자)</option>
        </select></div>
      <div class="mb-2"><label class="form-label">방문자 이름 <span class="text-muted small">(내부/외주 필수)</span></label>
        <input id="asVisitorName" type="text" class="form-control" placeholder="예: 강민석"></div>
      <div class="row"><div class="col mb-2"><label class="form-label">방문 예정일</label>
        <input id="asVisitStart" type="date" class="form-control" onclick="try{this.showPicker()}catch(e){}"></div>
        <div class="col mb-2"><label class="form-label">종료일 <span class="text-muted small">(선택·범위)</span></label>
        <input id="asVisitEnd" type="date" class="form-control" onclick="try{this.showPicker()}catch(e){}"></div></div>
      <div class="mb-1"><label class="form-label">접수 메모 <span class="text-muted small">(선택)</span></label>
        <textarea id="asAcceptMemo" class="form-control" rows="2" placeholder="접수번호·특이사항 등 — 메모/이력에 시간·이니셜과 함께 기록됩니다"></textarea></div>
    `, async (el) => {
      const visitor_type = el.querySelector('#asVisitorType').value;
      if (!visitor_type) return '방문자 유형을 선택해주세요.';
      const visitor_name = el.querySelector('#asVisitorName').value.trim();
      if ((visitor_type === '내부' || visitor_type === '외주') && !visitor_name) return '내부/외주는 방문자 이름이 필수입니다.';
      if (!el.querySelector('#asVisitStart').value) return '방문 예정일을 선택해주세요.';
      return this._post(`/as/api/accept/${encodeURIComponent(asNo)}`, {
        visitor_type,
        visitor_name,
        visit_date_start: el.querySelector('#asVisitStart').value,
        visit_date_end: el.querySelector('#asVisitEnd').value,
        memo: el.querySelector('#asAcceptMemo').value.trim(),
      });
    }, { icon: 'fa-clipboard-check', submitLabel: '접수' });
    // 방문자 유형 = 서비스 기사 → 이름칸 비활성 + '서비스 기사' 자동. 내부/외주 → 활성(직접 입력).
    const host = document.getElementById('asModalHost');
    const sel = host && host.querySelector('#asVisitorType');
    const name = host && host.querySelector('#asVisitorName');
    if (sel && name) {
      const sync = () => {
        if (sel.value === '서비스 기사') {
          name.value = '서비스 기사';
          name.disabled = true;
        } else {
          if (name.value === '서비스 기사') name.value = '';
          name.disabled = false;
        }
      };
      sel.addEventListener('change', sync);
      sync();  // 초기 상태(기본 '선택')에도 반영
    }
  }

  openComplete(asNo) {
    this._showModal(`A/S 조치 완료 — ${asNo}`, `
      <div class="mb-2"><label class="form-label">조치 내용</label>
        <textarea id="asResolution" class="form-control" rows="4" placeholder="조치 결과를 입력하세요"></textarea></div>
    `, async (el) => {
      const resolution = el.querySelector('#asResolution').value.trim();
      if (!resolution) return '조치 내용을 입력해주세요.';
      return this._post(`/as/api/complete/${encodeURIComponent(asNo)}`, { resolution });
    }, { icon: 'fa-flag-checkered', submitLabel: '조치 완료' });
  }

  openManualRequest() {
    this._showModal('수동 A/S 요청 (코드 없는 옛 공사)', `
      <div class="mb-2"><label class="form-label">현장 주소 *</label>
        <input id="asmAddress" type="text" class="form-control" placeholder="현장명/주소"></div>
      <div class="mb-2"><label class="form-label">공사 내용</label>
        <input id="asmWork" type="text" class="form-control"></div>
      <div class="mb-2"><label class="form-label">담당자 (영업)</label>
        <input id="asmManager" type="text" class="form-control" placeholder="담당자 이름"></div>
      <div class="mb-2"><label class="form-label">요청 내용 *</label>
        <textarea id="asmContent" class="form-control" rows="3"></textarea></div>
    `, async (el) => {
      const address = el.querySelector('#asmAddress').value.trim();
      const request_content = el.querySelector('#asmContent').value.trim();
      if (!address) return '현장 주소는 필수입니다.';
      if (!request_content) return '요청 내용은 필수입니다.';
      return this._post('/as/api/request', {
        request_content,
        manual: { address, work_content: el.querySelector('#asmWork').value.trim(), manager_name: el.querySelector('#asmManager').value.trim() },
      });
    }, { icon: 'fa-screwdriver-wrench', submitLabel: '요청 등록' });
  }

  /** 프로젝트 코드로 로컬 로드된 프로젝트 데이터 조회 (stateManager → projectsData 순) */
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

  /** 슬랙 A/S 카드와 동일한 공사 정보 카드 (읽기 전용) */
  _projectInfoCard(p) {
    const v = (k) => {
      const x = p[k];
      return (x === undefined || x === null || String(x).trim() === '') ? '-' : String(x);
    };
    // 공사 금액 = 총액 2 (VAT 포함/별도 표기는 부가세 필드에서)
    const amtRaw = p['총액 2'] || p['총액2'] || p['총액'] || p['총액 1'] || '';
    const amtNum = parseFloat(String(amtRaw).replace(/[^\d.-]/g, ''));
    let amount = amtNum ? `${amtNum.toLocaleString('ko-KR')}원` : '-';
    const vatRaw = String(p['부가세'] || '');
    if (amtNum && /별도/.test(vatRaw)) amount += ' (VAT 별도)';
    else if (amtNum && /포함/.test(vatRaw)) amount += ' (VAT 포함)';

    const rows = [
      ['📥 유입 구분', v('유입 구분')],
      ['🏢 사업자명', v('사업자')],
      ['📍 현장 주소', v('현장 주소')],
      ['👤 발주처 담당자', v('발주처 담당자')],
      ['📞 발주처 연락처', v('발주처 연락처')],
      ['✉️ 발주처 이메일', v('발주처 이메일')],
      ['📋 공사 내용', v('공사 내용')],
      ['🛠️ 도급 구분', v('도급 구분')],
      ['👷 시공자', v('시공자')],
      ['💲 공사 금액', amount],
      ['📅 공사 시작', v('공사 시작')],
      ['📅 공사 종료', v('공사 종료')],
    ];
    const items = rows.map(([label, val]) =>
      `<div class="d-flex" style="gap:0.5rem; padding:0.12rem 0;">
        <div style="min-width:104px; color:var(--text-muted,#6c757d); white-space:nowrap;">${esc(label)}</div>
        <div style="flex:1; word-break:break-word;">${esc(val)}</div>
      </div>`).join('');
    return `<div class="border rounded p-2 mb-3" style="background:var(--surface-secondary,#f8f9fa); font-size:var(--font-size-sm,0.875rem);">${items}</div>`;
  }

  /** 프로젝트 행 accordion 'A/S 요청' 버튼 (코드 있는 요청) — 공사 정보 카드 + 요청 내용.
   *  버튼 단계에서 진행 중(요청됨/접수완료) A/S 가 있으면 폼을 열지 않고 즉시 차단한다. */
  openProjectRequest(projectCode) {
    // 1) 버튼 단계 하드 블록 — 진행 중(요청됨/접수완료) A/S 가 있으면 폼을 열지 않음.
    //    byCode 는 페이지 로드 시 사전 적재돼 있어 즉시 판정(추가 요청 없음 → 지연 X).
    //    (미적재/경합 시엔 폼을 열되 서버 /api/request 가드가 최종 차단)
    const cur = this.byCode && this.byCode[String(projectCode || '').trim()];
    const st = cur ? String(cur['진행 상태'] || '').trim() : '';
    if (st === '요청됨' || st === '접수 완료') {
      this._showAlert('A/S 요청 불가',
        `이미 진행 중인 A/S 가 있습니다 (${cur['No']} · ${st}).\n기존 A/S 를 완료한 뒤 다시 요청해주세요.`,
        { icon: 'fa-triangle-exclamation', iconColor: '#e0a800' });
      return;
    }
    // 2) 진행 중 A/S 없음 → 요청 폼
    const p = this._findProject(projectCode);
    const info = p
      ? this._projectInfoCard(p)
      : `<div class="small text-muted mb-3">프로젝트 <b>${esc(projectCode)}</b></div>`;
    this._showModal(`A/S 요청 — ${projectCode}`, `
      ${info}
      <div class="mb-1"><label class="form-label">요청 내용 <span class="text-danger">*</span></label>
        <textarea id="aspContent" class="form-control" rows="3" placeholder="예) 실외기 소음 발생 / 냉방 약함 / 누수 등 A/S 요청 사유를 적어주세요"></textarea></div>
    `, async (el) => {
      const request_content = el.querySelector('#aspContent').value.trim();
      if (!request_content) return '요청 내용은 필수입니다.';
      const err = await this._post('/as/api/request', { project_code: projectCode, request_content });
      if (!err && window.showPageAlert) window.showPageAlert(`A/S 요청 등록 완료 (${projectCode})`, 'success');
      return err;  // AS_ALREADY_OPEN(레이스) 시 서버 메시지 그대로 표시
    }, { icon: 'fa-screwdriver-wrench', submitLabel: '요청 등록' });
  }
}
