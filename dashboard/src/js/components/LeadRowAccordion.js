/**
 * LeadRowAccordion — 고객 리드 행 아코디언 (프로젝트 ProjectRowAccordion 과 같은 골조)
 *
 *   tbody tr 클릭 → 바로 아래 tr.accordion-row 삽입 (한 번에 하나, 슬라이드 애니메이션, ESC 닫기)
 *   .accordion-shell > .row-details > 타이틀(헤더 필드 + 액션 버튼) + 정보 카드 4개
 *     고객 정보 / 방문·진행 / 문의 원문(J열) / 상담 이력(K열 회차 타임라인)
 *
 * 1단계: 보기 전용. 편집 모드·공사 확정 버튼은 project-actions-container 에 이어서 붙인다.
 */
import logger from '../utils/logger.js';
import {
  splitOwners, parseConsultEntries, consultFull, splitVisitRange, leadStatusClass, esc,
  lastContact, agoLabel,
} from '../utils/leadFormat.js';
import { managerBadges, platformBadge, selfVisitLabel } from './LeadTable.js';

const val = (r, k) => {
  const v = String(r?.[k] ?? '').trim();
  return v === '-' ? '' : v;
};

export default class LeadRowAccordion {
  constructor() {
    this.dataTable = null;
    this.currentLead = null;
    this.isOpen = false;
    this.container = document.createElement('div');
    this.container.className = 'project-accordion-container lead-accordion-container';
    this.eventsBound = false;
    this.slackLinkCache = new Map();   // 리드 No → {inquiry, visit} (페이지 머무는 동안)
  }

  attachToTable(tableElement, dataTable) {
    this.dataTable = dataTable;
    if (this.eventsBound) return;

    tableElement.addEventListener('click', (e) => {
      const tr = e.target.closest('tbody tr');
      if (!tr || tr.classList.contains('accordion-row')) return;
      if (e.target.closest('.action-buttons, a, button')) return;   // 링크·버튼은 토글 제외
      const data = this.dataTable.row(tr).data();
      if (data) this.toggleAccordion(tr, data);
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && this.isOpen && !document.querySelector('.modal.show')) this.closeAccordion();
    });

    this.eventsBound = true;
  }

  currentLeadNo() {
    return this.isOpen ? val(this.currentLead, '리드 No') : '';
  }

  toggleAccordion(tr, lead) {
    if (this.isOpen && val(this.currentLead, '리드 No') === val(lead, '리드 No')) {
      this.closeAccordion();
      return;
    }
    if (this.isOpen) this.detachQuietly();
    this.openAccordion(tr, lead);
  }

  openAccordion(tr, lead, { animate = true } = {}) {
    try {
      this.currentLead = lead;
      this.container.innerHTML = this.renderContent(lead);

      tr.nextElementSibling?.classList.contains('accordion-row') && tr.nextElementSibling.remove();
      const accRow = document.createElement('tr');
      accRow.className = 'accordion-row';
      accRow.innerHTML = '<td colspan="100%"><div class="accordion-content"></div></td>';
      accRow.querySelector('.accordion-content').appendChild(this.container);
      tr.insertAdjacentElement('afterend', accRow);

      this.container.classList.remove('accordion-slide-down', 'accordion-slide-up');
      this.container.classList.add('show');
      if (animate) {
        this.container.style.display = 'none';
        void this.container.offsetHeight;   // 강제 리플로우 (프로젝트 아코디언과 동일)
        this.container.style.display = 'block';
        this.container.classList.add('accordion-slide-down');
      } else {
        this.container.style.display = 'block';
      }

      document.querySelectorAll('#leadsTable tbody tr.table-active').forEach((r) => r.classList.remove('table-active'));
      tr.classList.add('table-active');
      this.isOpen = true;
      this.loadSlackLinks(val(lead, '리드 No'));
    } catch (err) {
      logger.error('[LeadAccordion] 렌더링 오류:', err, lead);
      this.detachQuietly();
      window.showPageAlert?.('리드 상세를 열 수 없습니다', 'error');
    }
  }

  closeAccordion() {
    if (!this.isOpen) return;
    this.isOpen = false;
    this.currentLead = null;
    this.container.classList.add('accordion-slide-up');
    let done = false;
    const cleanup = () => {
      if (done) return;
      done = true;
      document.querySelectorAll('#leadsTable .accordion-row').forEach((r) => r.remove());
      this.container.classList.remove('show', 'accordion-slide-up');
    };
    this.container.addEventListener('animationend', cleanup, { once: true });
    setTimeout(cleanup, 1000);   // animationend 누락 대비
    document.querySelectorAll('#leadsTable tbody tr.table-active').forEach((r) => r.classList.remove('table-active'));
  }

  /** 애니메이션 없이 즉시 제거 (데이터 갱신·다른 행 전환 시) */
  detachQuietly() {
    document.querySelectorAll('#leadsTable .accordion-row').forEach((r) => r.remove());
    document.querySelectorAll('#leadsTable tbody tr.table-active').forEach((r) => r.classList.remove('table-active'));
    this.container.classList.remove('show', 'accordion-slide-down', 'accordion-slide-up');
    this.isOpen = false;
    this.currentLead = null;
  }

  /** 현재 페이지에 그 리드 행이 있으면 다시 연다 (데이터 갱신 후 유지용) */
  reopen(leadNo) {
    if (!this.dataTable || !leadNo) return false;
    const row = this.dataTable.row(`#row-${leadNo}`);
    const node = row.node();
    if (!node) return false;
    this.openAccordion(node, row.data(), { animate: false });
    return true;
  }

  // ── 렌더링 ────────────────────────────────────────────────
  renderContent(lead) {
    const no = val(lead, '리드 No');
    return `
      <div class="accordion-shell lead-accordion-shell" data-lead-no="${esc(no)}">
        <div class="row-details">
          <div class="card border-0 shadow-sm">
            <div class="card-body p-4">
              <div class="row mb-3">
                <div class="col-12">
                  <div class="project-title-section" data-lead-no="${esc(no)}">
                    <div class="project-title-flex">
                      <div class="project-title-info">${this.renderHeader(lead)}</div>
                      <div class="project-actions-container"></div>
                    </div>
                  </div>
                </div>
              </div>
              <div class="all-info-container">
                <div class="row">
                  <div class="col-xl-3 col-lg-6 col-md-6 info-card-column">${this.renderCustomerCard(lead)}</div>
                  <div class="col-xl-3 col-lg-6 col-md-6 info-card-column">${this.renderVisitCard(lead)}</div>
                  <div class="col-xl-3 col-lg-6 col-md-6 info-card-column">${this.renderInquiryCard(lead)}</div>
                  <div class="col-xl-3 col-lg-6 col-md-6 info-card-column">${this.renderHistoryCard(lead)}</div>
                </div>
              </div>
              <!-- 하단 (프로젝트의 문서 폴더·사업자등록증·수금 특이사항 줄과 같은 legacy-card) -->
              <div class="mt-3">
                <div class="row">
                  <div class="col-md-6">${this.renderFolderSection(lead)}</div>
                  <div class="col-md-6">${this.renderSlackSection(lead)}</div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>`;
  }

  renderHeader(lead) {
    const no = val(lead, '리드 No');
    const consult = splitOwners(lead['온라인 상담자']).join(', ') || '-';
    const sales = splitOwners(lead['영업 담당자']).join(', ') || '미배정';
    const addr = val(lead, '방문 주소') || '주소 정보 없음';
    const platform = val(lead, '플랫폼');
    const when = consultFull(lead['상담 시간']);
    // 연결 프로젝트는 헤더에 (카드에서 빼서 카드 3줄 유지)
    const codes = val(lead, '_project_code').split(',').map((c) => c.trim()).filter(Boolean);
    const projectLinks = codes.length
      ? `<div class="project-content-info lead-header-projects"><i class="fas fa-link me-2"></i>${codes.map((c) =>
        `<a class="badge lead-project-link" href="/projects?search=${encodeURIComponent(c)}" target="_blank"
            rel="noopener" title="연결된 프로젝트 열기">${esc(c)}</a>`).join(' ')}</div>`
      : '';
    return `
      <div class="project-code-badge"><i class="fas fa-tag me-2"></i>${esc(no)}</div>
      <div class="project-manager-info"><i class="fas fa-user me-2"></i>상담 ${esc(consult)} · 영업 ${esc(sales)}</div>
      <div class="project-address-info" title="${esc(addr)}"><i class="fas fa-map-marker-alt me-2"></i>${esc(addr)}</div>
      <div class="project-content-info"><i class="fas fa-inbox me-2"></i>${esc([platform, when].filter(Boolean).join(' · ') || '-')}</div>
      ${projectLinks}`;
  }

  item(label, html, { wide = false } = {}) {
    return `
      <div class="compact-item${wide ? ' lead-item-wide' : ''}">
        <small>${label}</small>
        <div class="editable-value">${html || '-'}</div>
      </div>`;
  }

  renderCustomerCard(lead) {
    return `
      <div class="info-card compact-card">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h6 class="text-primary"><i class="fas fa-user me-1"></i>고객 정보</h6>
        </div>
        <div class="card-grid">
          ${this.item('고객명', esc(val(lead, '고객명')))}
          ${this.item('연락처', esc(val(lead, '고객 연락처')))}
          ${this.item('플랫폼', platformBadge(val(lead, '플랫폼')))}
          ${this.item('키워드', esc(val(lead, '키워드')))}
          ${this.item('이메일', esc(val(lead, '이메일')), { wide: true })}
        </div>
      </div>`;
  }

  renderVisitCard(lead) {
    const status = val(lead, '상태');
    const rawVisit = String(lead['방문 예정일'] || '').replace(/^'/, '').trim();
    const { start } = splitVisitRange(rawVisit);
    const visitHtml = rawVisit && rawVisit !== '-'
      ? `${esc(rawVisit)}${start ? '' : ' <i class="fas fa-exclamation-triangle text-warning" title="날짜 형식 확인 필요"></i>'}`
      : '';
    const lc = lastContact(lead);
    const lastHtml = lc.iso
      ? `${esc(lc.label)}${lc.ini ? ` ${esc(lc.ini)}` : ''}${lc.fromIntake ? ' 접수' : ''} · <span class="lead-sub">${agoLabel(lc.iso)}</span>`
      : '';
    return `
      <div class="info-card compact-card">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h6 class="text-success"><i class="fas fa-calendar-check me-1"></i>방문·진행</h6>
        </div>
        <div class="card-grid">
          ${this.item('상태', status ? `<span class="badge lead-status ${leadStatusClass(status)}">${esc(status)}</span>` : '')}
          ${this.item('방문 예정일', visitHtml)}
          ${this.item('상담자', managerBadges(lead['온라인 상담자']))}
          ${this.item('영업 담당', managerBadges(lead['영업 담당자']))}
          ${this.item('본인 방문', selfVisitLabel(val(lead, '본인 방문 여부')) || esc(val(lead, '본인 방문 여부')))}
          ${this.item('최근 연락', lastHtml)}
        </div>
      </div>`;
  }

  renderFolderSection(lead) {
    const folder = val(lead, '_folder_id');
    return `
      <div class="legacy-card mt-3 document-card ${folder ? '' : 'document-card-empty'}">
        <div class="legacy-card-row">
          <div class="legacy-card-main">
            <span class="legacy-card-label"><i class="fab fa-google-drive me-2" style="color: #4285f4;"></i>사진 폴더</span>
            <div class="editable-value">${folder
              ? `<a href="https://drive.google.com/drive/folders/${esc(folder)}" target="_blank" rel="noopener">방문 사진 폴더 열기</a>`
              : '<span class="lead-empty">방문 사진 없음</span>'}</div>
          </div>
        </div>
      </div>`;
  }

  renderSlackSection(lead) {
    return `
      <div class="legacy-card mt-3">
        <div class="legacy-card-row">
          <div class="legacy-card-main">
            <span class="legacy-card-label"><i class="fab fa-slack me-2" style="color: #4a154b;"></i>슬랙 바로가기</span>
            <div class="editable-value lead-slack-links" data-lead-no="${esc(val(lead, '리드 No'))}">
              <span class="lead-empty">불러오는 중…</span>
            </div>
          </div>
        </div>
      </div>`;
  }

  /** 슬랙 카드 permalink lazy 조회 → 하단 칸 채움 (리드별 캐시) */
  async loadSlackLinks(leadNo) {
    if (!leadNo) return;
    let links = this.slackLinkCache.get(leadNo);
    if (!links) {
      try {
        const res = await fetch(`/leads/api/${encodeURIComponent(leadNo)}/slack-links`, {
          credentials: 'same-origin', headers: { Accept: 'application/json' },
        });
        const json = res.ok ? await res.json() : null;
        links = json?.data || { inquiry: '', visit: '' };
        this.slackLinkCache.set(leadNo, links);
      } catch (_) {
        links = { inquiry: '', visit: '' };
      }
    }
    const box = this.container.querySelector(`.lead-slack-links[data-lead-no="${CSS.escape(leadNo)}"]`);
    if (!box) return;   // 그 사이 다른 리드로 전환됨
    const parts = [];
    if (links.inquiry) parts.push(`<a href="${esc(links.inquiry)}" target="_blank" rel="noopener"><i class="fas fa-inbox me-1"></i>문의 카드</a>`);
    if (links.visit) parts.push(`<a href="${esc(links.visit)}" target="_blank" rel="noopener"><i class="fas fa-map-marker-alt me-1"></i>방문 카드</a>`);
    box.innerHTML = parts.length
      ? parts.join('<span class="text-muted mx-2">·</span>')
      : '<span class="lead-empty">연결된 슬랙 카드 없음</span>';
  }

  renderInquiryCard(lead) {
    const text = val(lead, '문의 내용');
    return `
      <div class="info-card compact-card">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h6 class="text-warning"><i class="fas fa-inbox me-1"></i>문의 내용</h6>
        </div>
        <div class="lead-text-scroll">${text ? esc(text) : '<span class="lead-empty">문의 내용 없음</span>'}</div>
      </div>`;
  }

  renderHistoryCard(lead) {
    const entries = parseConsultEntries(lead['상담 내용']);
    const body = entries.length
      ? entries.map((e, i) => ({ ...e, n: i + 1 })).reverse().map((e) => `
          <div class="lead-history-entry">
            <div class="lead-history-head">
              <span class="lead-history-n">${e.n}차</span>
              ${e.time ? `<span>${esc(e.time)}</span>` : ''}
              ${e.ini ? `<span>${esc(e.ini)}</span>` : ''}
              ${e.status ? `<span class="badge lead-status ${leadStatusClass(e.status)}">${esc(e.status)}</span>` : ''}
            </div>
            <div class="lead-history-body">${e.content ? esc(e.content) : '<span class="lead-empty">내용 없음</span>'}</div>
          </div>`).join('')
      : '<span class="lead-empty">상담 기록 없음</span>';
    return `
      <div class="info-card compact-card">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h6 class="text-info"><i class="fas fa-comments me-1"></i>상담 이력${entries.length ? ` <small class="text-muted">(${entries.length}회)</small>` : ''}</h6>
        </div>
        <div class="lead-text-scroll lead-history">${body}</div>
      </div>`;
  }
}
