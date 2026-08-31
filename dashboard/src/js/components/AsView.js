/**
 * AsView — PM 대시보드 A/S 관리 모드 (자체 완결형).
 *
 * 토글 ON 시 프로젝트 뷰(.filter-section-sticky/.table-section/.mobile-card-container)를
 * 숨기고 #asView(A/S 전용 테이블)를 그 자리에 표시. 프로젝트 DataTable 은 건드리지 않음.
 * 왼쪽은 연결 프로젝트 컬럼, 오른쪽은 A/S 컬럼. 요청/접수/완료는 슬랙 카드·DM 과 동기화(백엔드).
 * 백엔드: /as/api/list · /as/api/request · /as/api/accept/<no> · /as/api/complete/<no>
 */
import logger from '../utils/logger.js';

const STATUS_REQUESTED = '요청됨';
const STATUS_ACCEPTED = '접수 완료';
const STATUS_COMPLETED = '처리 완료';

const STATUS_META = {
  [STATUS_REQUESTED]: { label: '🔔 요청됨', cls: 'bg-warning text-dark' },
  [STATUS_ACCEPTED]: { label: '📥 접수완료', cls: 'bg-info text-dark' },
  [STATUS_COMPLETED]: { label: '✅ 처리완료', cls: 'bg-success' },
};

const PROJECT_VIEW_SELECTORS = ['.filter-section-sticky', '.table-section', '.mobile-card-container'];

function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export default class AsView {
  constructor() {
    this.container = document.getElementById('asView');
    this.toggleBtn = document.getElementById('asModeToggle');
    this.active = false;
    this.items = [];
    this.filter = { status: '', manager: '', q: '' };
    this._bound = false;
    this._modal = null; // bootstrap.Modal 인스턴스
    if (!this.container || !this.toggleBtn) {
      logger.warn('[AsView] #asView / #asModeToggle 없음 — 초기화 skip');
      return;
    }
    this._bindToggle();
  }

  _bindToggle() {
    this.toggleBtn.addEventListener('click', () => this.toggle());
  }

  toggle() {
    this.active = !this.active;
    this._applyMode();
    this.toggleBtn.classList.toggle('active', this.active);
    this.toggleBtn.setAttribute('aria-pressed', String(this.active));
    if (this.active) {
      this.container.style.display = 'block';
      this._ensureShell();
      this.load();
    } else {
      this.container.style.display = 'none';
    }
  }

  _applyMode() {
    // 클래스 기반 숨김 — 원래 display(모바일 컨테이너 기본 none 등)를 보존하며 복원.
    PROJECT_VIEW_SELECTORS.forEach((sel) => {
      document.querySelectorAll(sel).forEach((el) => {
        el.classList.toggle('as-hidden', this.active);
      });
    });
  }

  /** 컨테이너 골격(필터바 + 테이블 + 모달)을 1회 구성 */
  _ensureShell() {
    if (this._bound) return;
    this.container.innerHTML = `
      <div class="filter-section" style="margin-bottom:12px;">
        <div class="d-flex flex-wrap align-items-end gap-2">
          <div>
            <label class="form-label mb-1">상태</label>
            <select id="asFilterStatus" class="form-select form-select-sm">
              <option value="">전체</option>
              <option value="미완료">미완료(요청+접수)</option>
              <option value="${STATUS_REQUESTED}">요청됨</option>
              <option value="${STATUS_ACCEPTED}">접수 완료</option>
              <option value="${STATUS_COMPLETED}">처리 완료</option>
            </select>
          </div>
          <div>
            <label class="form-label mb-1">담당자</label>
            <select id="asFilterManager" class="form-select form-select-sm"><option value="">전체</option></select>
          </div>
          <div style="flex:1;min-width:180px;">
            <label class="form-label mb-1">검색 (코드/주소/요청)</label>
            <input id="asFilterSearch" type="text" class="form-control form-control-sm" placeholder="검색어">
          </div>
          <div class="ms-auto d-flex gap-2">
            <button id="asManualRequestBtn" class="btn btn-soft-primary btn-sm">
              <i class="fas fa-plus me-1"></i>수동 A/S 요청
            </button>
            <button id="asReloadBtn" class="btn btn-outline-secondary btn-sm" title="새로고침">
              <i class="fas fa-sync"></i>
            </button>
          </div>
        </div>
      </div>
      <div id="asTableWrap" class="table-responsive"></div>
      <div id="asModalHost"></div>
    `;
    // 이벤트 위임
    this.container.querySelector('#asFilterStatus').addEventListener('change', (e) => { this.filter.status = e.target.value; this.render(); });
    this.container.querySelector('#asFilterManager').addEventListener('change', (e) => { this.filter.manager = e.target.value; this.render(); });
    this.container.querySelector('#asFilterSearch').addEventListener('input', (e) => { this.filter.q = e.target.value.trim(); this.render(); });
    this.container.querySelector('#asManualRequestBtn').addEventListener('click', () => this._openManualRequest());
    this.container.querySelector('#asReloadBtn').addEventListener('click', () => this.load());
    this.container.querySelector('#asTableWrap').addEventListener('click', (e) => this._onTableClick(e));
    this._bound = true;
  }

  async load() {
    const wrap = this.container.querySelector('#asTableWrap');
    if (wrap) wrap.innerHTML = '<div class="text-center text-muted py-4">불러오는 중…</div>';
    try {
      const resp = await fetch('/as/api/list', { credentials: 'same-origin' });
      const json = await resp.json();
      const data = json.data || json;
      this.items = (data && data.items) || [];
      this._populateManagers();
      this.render();
    } catch (err) {
      logger.error('[AsView] 목록 로드 실패:', err);
      if (wrap) wrap.innerHTML = '<div class="text-center text-danger py-4">A/S 목록을 불러오지 못했습니다.</div>';
    }
  }

  _populateManagers() {
    const sel = this.container.querySelector('#asFilterManager');
    if (!sel) return;
    const cur = this.filter.manager;
    const names = [...new Set(this.items.map((i) => (i['담당자'] || '').trim()).filter(Boolean))].sort();
    sel.innerHTML = '<option value="">전체</option>' + names.map((n) => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
    sel.value = cur;
  }

  _filtered() {
    const { status, manager, q } = this.filter;
    return this.items.filter((it) => {
      const st = (it['진행 상태'] || '').trim();
      if (status === '미완료') { if (st === STATUS_COMPLETED) return false; }
      else if (status) { if (st !== status) return false; }
      if (manager && (it['담당자'] || '').trim() !== manager) return false;
      if (q) {
        const hay = `${it['프로젝트 코드'] || ''} ${it['현장주소'] || ''} ${it['요청 내용'] || ''} ${it['사업자명'] || ''}`.toLowerCase();
        if (!hay.includes(q.toLowerCase())) return false;
      }
      return true;
    });
  }

  render() {
    const wrap = this.container.querySelector('#asTableWrap');
    if (!wrap) return;
    const rows = this._filtered();
    if (!rows.length) {
      wrap.innerHTML = '<div class="text-center text-muted py-4">해당하는 A/S 건이 없습니다.</div>';
      return;
    }
    const body = rows.map((it) => {
      const st = (it['진행 상태'] || '').trim();
      const meta = STATUS_META[st] || { label: esc(st), cls: 'bg-secondary' };
      const code = (it['프로젝트 코드'] || '').trim();
      let action = '';
      if (st === STATUS_REQUESTED) action = `<button class="btn btn-primary btn-sm" data-as-action="accept" data-as-no="${esc(it.No)}">접수</button>`;
      else if (st === STATUS_ACCEPTED) action = `<button class="btn btn-success btn-sm" data-as-action="complete" data-as-no="${esc(it.No)}">완료</button>`;
      return `<tr>
        <td class="text-nowrap">${code ? esc(code) : '<span class="text-muted">-</span>'}<div class="small text-muted">${esc(it.No)}</div></td>
        <td>${esc(it['담당자'] || '-')}</td>
        <td>${esc(it['유입 구분'] || '-')}</td>
        <td>${esc(it['사업자명'] || '-')}</td>
        <td>${esc(it['현장주소'] || '-')}</td>
        <td class="text-nowrap">${esc(it['접수 일자'] || '-')}</td>
        <td class="text-nowrap">${esc(it['방문 예정일'] || '-')}</td>
        <td style="max-width:240px;white-space:normal;">${esc(it['요청 내용'] || '-')}</td>
        <td>${esc(it['방문 예정자'] || '-')}</td>
        <td><span class="badge ${meta.cls}">${meta.label}</span></td>
        <td style="max-width:220px;white-space:normal;">${esc(it['처리 내용'] || '-')} ${action}</td>
      </tr>`;
    }).join('');
    wrap.innerHTML = `
      <table class="table table-sm table-hover align-middle">
        <thead><tr class="text-nowrap">
          <th>프로젝트 코드</th><th>담당자</th><th>유입 구분</th><th>사업자명</th><th>현장 주소</th>
          <th>A/S 접수일</th><th>방문 예정일</th><th>요청 내용</th><th>방문 예정자</th><th>상태</th><th>처리 내용</th>
        </tr></thead>
        <tbody>${body}</tbody>
      </table>
      <div class="small text-muted">총 ${rows.length}건</div>`;
  }

  _onTableClick(e) {
    const btn = e.target.closest('[data-as-action]');
    if (!btn) return;
    const asNo = btn.getAttribute('data-as-no');
    const action = btn.getAttribute('data-as-action');
    if (action === 'accept') this._openAccept(asNo);
    else if (action === 'complete') this._openComplete(asNo);
  }

  // ── 모달 (Bootstrap) ──────────────────────────────────────────
  _showModal(title, bodyHtml, onSubmit) {
    const host = this.container.querySelector('#asModalHost');
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
    el.querySelector('#asModalSubmit').addEventListener('click', async () => {
      const err = await onSubmit(el);
      if (err) { el.querySelector('#asModalAlert').textContent = err; return; }
      this._modal.hide();
      this.load();
    });
    this._modal.show();
  }

  async _post(url, payload) {
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
  }

  _openAccept(asNo) {
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
    `, async (el) => {
      const payload = {
        visitor_type: el.querySelector('#asVisitorType').value,
        visitor_name: el.querySelector('#asVisitorName').value.trim(),
        visit_date_start: el.querySelector('#asVisitStart').value,
        visit_date_end: el.querySelector('#asVisitEnd').value,
      };
      return this._post(`/as/api/accept/${encodeURIComponent(asNo)}`, payload);
    });
  }

  _openComplete(asNo) {
    this._showModal(`A/S 처리 완료 — ${asNo}`, `
      <div class="mb-2"><label class="form-label">처리 내용</label>
        <textarea id="asResolution" class="form-control" rows="4" placeholder="처리 결과를 입력하세요"></textarea></div>
    `, async (el) => {
      const resolution = el.querySelector('#asResolution').value.trim();
      if (!resolution) return '처리 내용을 입력해주세요.';
      return this._post(`/as/api/complete/${encodeURIComponent(asNo)}`, { resolution });
    });
  }

  _openManualRequest() {
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

  /** 프로젝트 행 accordion 의 'A/S 요청' 버튼에서 호출 (코드 있는 요청) */
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
