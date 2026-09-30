/**
 * LeadTable — 고객 리드 메인 테이블 (읽기 전용)
 *
 * 프로젝트 페이지 ProjectTable 과 같은 골조:
 *   - DataTables(npm v2) 설정·언어·dom 동일, 이름 붙은 열(name)로 모드별 열 전환
 *   - '개씩 보기' 줄에 [내 리드만 보기] [기타 포함] ··· [방문 모드] 토글 주입
 *   - 행 클릭 → LeadRowAccordion (편집은 아코디언 편집 모드에서만)
 * 모드: 기본(접수·상담 위주) / 방문(방문 예정일·본인 방문·영업 담당 위주). sessionStorage 유지.
 */
import DataTable from 'datatables.net';
import 'datatables.net-bs5';
import logger from '../utils/logger.js';
import UnifiedBadgeSystem from './UnifiedBadgeSystem.js';
import { FOLDER_ID_RE } from '../utils/itgfolder.js';
import {
  consultSortKey, consultFull, splitVisitRange, splitOwners,
  parseConsultEntries, leadStatusClass, localIso, daysBetween, esc,
} from '../utils/leadFormat.js';

// 플랫폼·담당자 배지는 프로젝트 페이지와 같은 배지 체계(유입 구분·담당자 색) 재사용
const badges = new UnifiedBadgeSystem();

/** 이름 목록 → 담당자 배지들 (공동방문 "권태훈,강정권" 은 이름마다 배지) */
export function managerBadges(raw) {
  return splitOwners(raw).map((n) => badges.createManagerBadge(n)).join(' ');
}

export function platformBadge(p) {
  return p ? badges.createCompanyBadge(p) : '';
}

/** 취소된 리드(방문 취소·공사 취소) — 프로젝트 테이블의 '공사 취소' 효과를 같은 클래스로 적용 */
export const CANCELLED_STATUSES = ['방문 취소', '공사 취소'];
export function isCancelledLead(r) {
  return CANCELLED_STATUSES.includes(String(r?.['상태'] ?? '').trim());
}

/** 본인 방문 여부(O열) 짧은 표시 — 테이블·아코디언 공용 */
export function selfVisitLabel(v) {
  const s = String(v || '').trim();
  if (s === '본인 방문 필수') return '<span class="badge lead-self-required">본인 필수</span>';
  if (s === '아무나 방문 가능') return '<span class="lead-sub">아무나</span>';
  return '';
}

const MODE_KEY = 'itg_lead_visit_mode';

const COLUMN_NAMES = [
  'leadNo', 'consultTime', 'visitDate', 'platform', 'name', 'phone', 'address',
  'status', 'consultant', 'sales', 'selfVisit', 'progress',
];

// 모드별로 숨길 열 (나머지는 표시)
const HIDDEN_BY_MODE = {
  base: ['visitDate', 'selfVisit'],
  visit: ['consultTime', 'consultant'],
};

const val = (r, k) => {
  const v = String(r[k] ?? '').trim();
  return v === '-' ? '' : v;
};

export default class LeadTable {
  constructor({ tableSelector = '#leadsTable', filters, accordion } = {}) {
    this.tableElement = document.querySelector(tableSelector);
    this.filters = filters;
    this.accordion = accordion;
    this.table = null;
    this.visitMode = false;
    this._firstDrawResolve = null;
    this.firstDraw = new Promise((res) => { this._firstDrawResolve = res; });
  }

  init(data) {
    if (!this.tableElement) {
      logger.error('[LeadTable] 테이블 요소 없음');
      return;
    }
    this.visitMode = this.restoreMode();
    this.initializeDataTable(data);
    this.accordion?.attachToTable(this.tableElement, this.table);
  }

  // ── DataTables ───────────────────────────────────────────
  initializeDataTable(data) {
    const today = localIso(new Date());

    this.table = new DataTable(this.tableElement, {
      data,
      rowId: (r) => `row-${val(r, '리드 No')}`,
      responsive: false,
      autoWidth: false,
      pageLength: 15,
      lengthMenu: [[15, 25, 50, 100], ['15', '25', '50', '100']],
      order: [[this.visitMode ? 2 : 1, 'desc']],
      orderMulti: false,
      searching: false,
      stateSave: false,
      dom: '<"top"l>rt<"bottom d-flex justify-content-between"<"info-left"i><"paging-right"p>><"clear">',
      language: {
        lengthMenu: '_MENU_개씩 보기',
        info: '_START_~_END_ / 전체 _TOTAL_개 (페이지 _PAGE_ / _PAGES_)',
        infoEmpty: '0개',
        infoFiltered: '',
        paginate: { first: '처음', last: '마지막', next: '다음', previous: '이전' },
        emptyTable: '데이터가 없습니다',
        zeroRecords: '조건에 맞는 리드가 없습니다',
      },
      columns: [
        {
          name: 'leadNo', data: null, className: 'lcol-no',
          render: (d, t, r) => {
            const no = val(r, '리드 No');
            if (t === 'sort' || t === 'type') return Number((no.match(/\d+/) || ['0'])[0]);
            return t === 'display' ? `<span class="lead-no-badge">${esc(no)}</span>` : no;
          },
        },
        {
          name: 'consultTime', data: null, className: 'lcol-time',
          render: (d, t, r) => {
            const raw = r['상담 시간'];
            if (t === 'sort' || t === 'type') return consultSortKey(raw);
            return t === 'display' ? esc(consultFull(raw) || '-') : raw;
          },
        },
        {
          name: 'visitDate', data: null, className: 'lcol-visit',
          render: (d, t, r) => {
            const raw = String(r['방문 예정일'] || '').replace(/^'/, '').trim();
            const { start } = splitVisitRange(raw);
            if (t === 'sort' || t === 'type') return start;
            if (t !== 'display') return raw;
            if (!start) {
              return raw && raw !== '-'
                ? `<span title="날짜 형식 확인 필요">${esc(raw)}</span>`
                : '<span class="lead-empty">미정</span>';
            }
            // D-n 은 줄바꿈 없이 같은 줄에 (행 높이 유지)
            const dd = daysBetween(today, start);
            const tag = dd === 0 ? ' <span class="lead-dday lead-dday-today">오늘</span>'
              : dd > 0 ? ` <span class="lead-dday">D-${dd}</span>` : '';
            return `${esc(raw)}${tag}`;
          },
        },
        {
          name: 'platform', data: null, className: 'lcol-platform',
          render: (d, t, r) => {
            const p = val(r, '플랫폼');
            return t === 'display' ? (platformBadge(p) || '-') : p;
          },
        },
        {
          name: 'name', data: null, className: 'lcol-name',
          render: (d, t, r) => {
            const n = val(r, '고객명');
            if (t !== 'display') return n;
            // 긴 상호명이 두 줄로 꺾여 행 높이가 늘지 않게 한 줄 + 말줄임 (전체는 툴팁)
            return n ? `<span class="lead-ellipsis" title="${esc(n)}">${esc(n)}</span>` : '-';
          },
        },
        {
          name: 'phone', data: null, className: 'lcol-phone',
          render: (d, t, r) => (t === 'display' ? esc(val(r, '고객 연락처') || '-') : val(r, '고객 연락처')),
        },
        {
          name: 'address', data: null, className: 'lcol-address',
          render: (d, t, r) => {
            const a = val(r, '방문 주소');
            if (t !== 'display') return a;
            return a
              ? `<span class="lead-ellipsis" title="${esc(a)}">${esc(a)}</span>`
              : '<span class="text-danger"><i class="fas fa-map-marker-alt me-1"></i>주소 없음</span>';
          },
        },
        {
          name: 'status', data: null, className: 'lcol-status text-center',
          render: (d, t, r) => {
            const s = val(r, '상태');
            return t === 'display' ? `<span class="badge lead-status ${leadStatusClass(s)}">${esc(s || '-')}</span>` : s;
          },
        },
        {
          name: 'consultant', data: null, className: 'lcol-consultant',
          render: (d, t, r) => {
            const v = splitOwners(r['온라인 상담자']).join(', ');
            return t === 'display' ? (managerBadges(r['온라인 상담자']) || '-') : v;
          },
        },
        {
          name: 'sales', data: null, className: 'lcol-sales',
          render: (d, t, r) => {
            const v = splitOwners(r['영업 담당자']).join(', ');
            if (t !== 'display') return v;
            if (v) return managerBadges(r['영업 담당자']);
            return val(r, '상태') === '방문 예약' ? '<span class="lead-unassigned">미배정</span>' : '-';
          },
        },
        {
          name: 'selfVisit', data: null, className: 'lcol-self text-center',
          render: (d, t, r) => {
            const v = val(r, '본인 방문 여부');
            if (t !== 'display') return v;
            return selfVisitLabel(v) || '-';
          },
        },
        {
          name: 'progress', data: null, className: 'lcol-progress', orderable: false,
          render: (d, t, r) => (t === 'display' ? this.renderProgress(r) : ''),
        },
      ],
      columnDefs: [
        { targets: HIDDEN_BY_MODE[this.visitMode ? 'visit' : 'base'].map((n) => `${n}:name`), visible: false },
      ],
      createdRow: (row, r) => {
        const s = val(r, '상태');
        // 상태 칸 식별 — 취소 행에서도 상태 배지는 취소선·흐림 없이 (프로젝트 status-column-cell 과 동일)
        row.querySelector('td.lcol-status')?.classList.add('status-column-cell');
        if (isCancelledLead(r)) row.classList.add('project-cancelled-row');   // 프로젝트 '공사 취소' 행 효과
        else if (s === '문의 드랍') row.classList.add('lead-row-closed');
      },
      initComplete: () => this.injectLengthBarControls(),
      drawCallback: () => {
        if (this._firstDrawResolve) {
          this._firstDrawResolve();
          this._firstDrawResolve = null;
        }
      },
    });

    logger.debug('[LeadTable] DataTables 초기화 완료');
  }

  // 진행: 사진 폴더 · 연결 프로젝트 · 상담 회차 (링크 클릭은 아코디언 토글 제외 = .action-buttons)
  renderProgress(r) {
    const parts = [];
    const folder = val(r, '_folder_id');
    if (folder && FOLDER_ID_RE.test(folder)) {
      // itgfolder:// = 탐색기에서 열기 (프로젝트 문서 폴더와 같은 방식, 미설치 감지는 itgfolder-link 클래스로)
      parts.push(`<a class="lead-progress-icon itgfolder-link" href="itgfolder://${esc(folder)}"
        data-folder-id="${esc(folder)}" title="방문 사진 폴더 (탐색기에서 열기)"><i class="fas fa-folder"></i></a>`);
    }
    const codes = val(r, '_project_code');
    if (codes) {
      codes.split(',').map((c) => c.trim()).filter(Boolean).forEach((c) => {
        parts.push(`<a class="badge lead-project-link" href="/projects?search=${encodeURIComponent(c)}"
          target="_blank" rel="noopener" title="연결된 프로젝트 열기"><i class="fas fa-link me-1"></i>${esc(c)}</a>`);
      });
    }
    const n = parseConsultEntries(r['상담 내용']).length;
    if (n) parts.push(`<span class="lead-progress-icon" title="상담 ${n}회"><i class="far fa-comment-dots"></i>${n}</span>`);
    return parts.length ? `<span class="action-buttons lead-progress">${parts.join('')}</span>` : '';
  }

  // ── '개씩 보기' 줄 토글 (ProjectTable.addMyProjectsFilter 와 같은 위치·마크업) ──
  injectLengthBarControls() {
    const lengthContainer = this.tableElement.closest('.table-wrapper')?.querySelector('.dt-length');
    if (!lengthContainer || lengthContainer.querySelector('.my-projects-filter')) return;
    Object.assign(lengthContainer.style, { display: 'flex', alignItems: 'center', gap: '15px' });

    const check = (id, text, title, checked) => {
      const box = document.createElement('div');
      box.className = 'my-projects-filter';
      const input = document.createElement('input');
      input.type = 'checkbox';
      input.id = id;
      input.title = title;
      input.checked = !!checked;
      const label = document.createElement('label');
      label.setAttribute('for', id);
      label.textContent = text;
      box.append(input, label);
      lengthContainer.appendChild(box);
      return input;
    };

    const t = this.filters?.toggles || {};
    const my = check('myLeadsOnly', '내 리드만 보기', '내가 상담자이거나 영업 담당인 리드만', t.myLeadsOnly);
    my.addEventListener('change', (e) => this.filters?.setToggle('myLeadsOnly', e.target.checked));
    if (!(window.userDisplayName || '').trim()) {
      my.disabled = true;
      my.title = '로그인 사용자 이름을 확인할 수 없습니다';
    }

    const etc = check('includeEtc', '기타 포함', '사후관리·A/S 등 기타(ETC) 방문도 표시', t.includeEtc);
    etc.addEventListener('change', (e) => this.filters?.setToggle('includeEtc', e.target.checked));

    const sw = document.createElement('div');
    sw.className = 'form-check form-switch';
    Object.assign(sw.style, { marginLeft: 'auto', marginRight: '20px', display: 'flex', alignItems: 'center' });
    sw.innerHTML = `
      <input type="checkbox" id="visitModeToggle" class="form-check-input" style="cursor:pointer;margin-top:0;">
      <label class="form-check-label" for="visitModeToggle"
             style="cursor:pointer;user-select:none;margin-left:8px;font-weight:500;">방문 모드</label>`;
    lengthContainer.appendChild(sw);
    const visitSwitch = sw.querySelector('input');
    visitSwitch.checked = this.visitMode;
    visitSwitch.addEventListener('change', (e) => this.setVisitMode(e.target.checked));

    // 프리셋 로드 시 토글 UI 동기화
    document.addEventListener('leadTogglesRestored', (e) => {
      const tg = e.detail || {};
      my.checked = !!tg.myLeadsOnly;
      etc.checked = !!tg.includeEtc;
      if (!!tg.visitMode !== this.visitMode) {
        visitSwitch.checked = !!tg.visitMode;
        this.setVisitMode(!!tg.visitMode, { skipFilter: true });
      }
    });
  }

  // ── 모드 전환 (수금관리 모드처럼 이름 열 표시 전환 + 필터 재적용) ─────
  setVisitMode(on, { skipFilter = false } = {}) {
    this.visitMode = !!on;
    try { sessionStorage.setItem(MODE_KEY, this.visitMode ? '1' : '0'); } catch (_) { /* noop */ }

    this.accordion?.closeAccordion?.();
    const hide = HIDDEN_BY_MODE[this.visitMode ? 'visit' : 'base'];
    COLUMN_NAMES.forEach((name) => this.table.column(`${name}:name`).visible(!hide.includes(name), false));
    this.table.order([this.table.column(this.visitMode ? 'visitDate:name' : 'consultTime:name').index(), 'desc']);
    this.table.columns.adjust();

    document.dispatchEvent(new CustomEvent('leadModeChanged', { detail: { visitMode: this.visitMode } }));
    if (this.filters) {
      this.filters.toggles.visitMode = this.visitMode;
      if (!skipFilter) this.filters.applyFilters(null, true);
    } else {
      this.table.draw(false);
    }
  }

  restoreMode() {
    try {
      const v = sessionStorage.getItem(MODE_KEY);
      const on = v === '1';
      if (this.filters) this.filters.toggles.visitMode = on;
      return on;
    } catch (_) {
      return false;
    }
  }

  // ── 데이터 갱신 (열린 아코디언 유지) ─────────────────────
  //   keepPage: 새로고침(같은 필터)이면 보던 페이지 유지, 필터 변경이면 1페이지로
  updateData(rows, { keepPage = false } = {}) {
    if (!this.table) return;
    const openNo = this.accordion?.currentLeadNo?.() || '';
    this.accordion?.detachQuietly?.();
    this.table.clear();
    this.table.rows.add(rows);
    this.table.draw(!keepPage);
    if (openNo) this.accordion?.reopen?.(openNo);
  }

  /** 리드 번호로 행을 찾아 해당 페이지로 이동 후 아코디언 열기 (캘린더 → 테이블 이동용) */
  revealLead(leadNo) {
    const idx = this.table.rows({ order: 'current', search: 'applied' }).indexes().toArray();
    const pos = idx.findIndex((i) => val(this.table.row(i).data(), '리드 No') === leadNo);
    if (pos < 0) return false;
    this.table.page(Math.floor(pos / this.table.page.len())).draw(false);
    this.accordion?.reopen?.(leadNo);
    return true;
  }
}
