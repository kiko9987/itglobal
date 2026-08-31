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

  _showModal(title, bodyHtml, onSubmit) {
    const host = this._modalHost();
    host.innerHTML = `
      <div class="modal fade" id="asActionModal" tabindex="-1">
        <div class="modal-dialog"><div class="modal-content">
          <div class="modal-header"><h5 class="modal-title">${esc(title)}</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal"></button></div>
          <div class="modal-body">${bodyHtml}<div id="asModalAlert" class="text-danger small mt-2"></div></div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">취소</button>
            <button type="button" class="btn btn-primary" id="asModalSubmit">확인</button>
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

  async _post(url, payload) {
    try {
      const resp = await fetch(url, {
        method: 'POST', credentials: 'same-origin',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload || {}),
      });
      const json = await resp.json().catch(() => ({}));
      if (!resp.ok || json.success === false) {
        return json.message || json.error || `요청 실패 (HTTP ${resp.status})`;
      }
      return null;
    } catch (e) {
      return '네트워크 오류로 실패했습니다.';
    }
  }

  openAccept(asNo) {
    this._showModal(`A/S 접수 — ${asNo}`, `
      <div class="mb-2"><label class="form-label">방문자 유형</label>
        <select id="asVisitorType" class="form-select">
          <option value="서비스 기사">서비스 기사</option>
          <option value="내부">내부 (아이티)</option>
          <option value="외주">외주 (시공자)</option>
        </select></div>
      <div class="mb-2"><label class="form-label">방문자 이름 <span class="text-muted small">(내부/외주 필수)</span></label>
        <input id="asVisitorName" type="text" class="form-control" placeholder="예: 강민석"></div>
      <div class="row"><div class="col mb-2"><label class="form-label">방문 예정일</label>
        <input id="asVisitStart" type="date" class="form-control"></div>
        <div class="col mb-2"><label class="form-label">종료일 <span class="text-muted small">(선택·범위)</span></label>
        <input id="asVisitEnd" type="date" class="form-control"></div></div>
    `, async (el) => this._post(`/as/api/accept/${encodeURIComponent(asNo)}`, {
      visitor_type: el.querySelector('#asVisitorType').value,
      visitor_name: el.querySelector('#asVisitorName').value.trim(),
      visit_date_start: el.querySelector('#asVisitStart').value,
      visit_date_end: el.querySelector('#asVisitEnd').value,
    }));
  }

  openComplete(asNo) {
    this._showModal(`A/S 처리 완료 — ${asNo}`, `
      <div class="mb-2"><label class="form-label">처리 내용</label>
        <textarea id="asResolution" class="form-control" rows="4" placeholder="처리 결과를 입력하세요"></textarea></div>
    `, async (el) => {
      const resolution = el.querySelector('#asResolution').value.trim();
      if (!resolution) return '처리 내용을 입력해주세요.';
      return this._post(`/as/api/complete/${encodeURIComponent(asNo)}`, { resolution });
    });
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
    });
  }

  /** 프로젝트 행 accordion 'A/S 요청' 버튼 (코드 있는 요청) */
  openProjectRequest(projectCode) {
    this._showModal(`A/S 요청 — ${projectCode}`, `
      <div class="mb-2 small text-muted">프로젝트 <b>${esc(projectCode)}</b> 에 대한 A/S 요청을 등록합니다.</div>
      <div class="mb-2"><label class="form-label">요청 내용 *</label>
        <textarea id="aspContent" class="form-control" rows="4" placeholder="A/S 요청 사유"></textarea></div>
    `, async (el) => {
      const request_content = el.querySelector('#aspContent').value.trim();
      if (!request_content) return '요청 내용은 필수입니다.';
      const err = await this._post('/as/api/request', { project_code: projectCode, request_content });
      if (!err && window.showPageAlert) window.showPageAlert(`A/S 요청 등록 완료 (${projectCode})`, 'success');
      return err;
    });
  }
}
