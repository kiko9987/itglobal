/**
 * 고객 리드 관리 페이지 — 구글 시트 스타일 인라인 편집
 *
 * - 15열 한 줄에 모두 표시 (가로 스크롤)
 * - 상태/플랫폼/영업 담당자: 클릭 한 번에 드롭다운 즉시 변경
 * - 텍스트/날짜 필드: 더블클릭으로 인라인 편집 (Enter 저장, Esc 취소)
 * - 변경 시 시트에 즉시 반영
 */

import DataTable from 'datatables.net';
import 'datatables.net-bs5';
import 'datatables.net-bs5/css/dataTables.bootstrap5.min.css';
import '../../css/pages/leads.css';
import logger from '../utils/logger.js';
import ModernLeadsFilters from '../components/ModernLeadsFilters.js';

import { Calendar } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import koLocale from '@fullcalendar/core/locales/ko';

// ─────────────────────────────────────────────────────────────
// 전역
// ─────────────────────────────────────────────────────────────
let leadsTable = null;
let calendar = null;
let leadsData = [];
let currentView = 'table';
let leadsFilters = null;

// 한글 컬럼명 ↔ 프런트 필드명 매핑 (서버 PUT 시 사용)
const KO = {
    leadNo:       '리드 No',
    consultTime:  '상담 시간',
    platform:     '플랫폼',
    status:       '상태',
    visitDate:    '방문 예정일',
    phone:        '고객 연락처',
    email:        '이메일',
    customerName: '고객명',
    address:      '방문 주소',
    content:      '문의 내용',   // 인입 원본 (옛 '상담 내용')
    feedback:     '상담 내용',   // 매니저 처리 결과 (옛 '피드백')
    keyword:      '키워드',
    onlineOwner:  '온라인 상담자',
    salesOwner:   '영업 담당자',
    lastContact:  '마지막 연락일',
};

// 상태별 배지 색상 (Bootstrap 5)
const STATUS_COLORS = {
    '상담 대기': 'secondary',
    '유선 상담': 'secondary',
    '부재중':   'warning',
    '방문 예약': 'primary',
    '방문 대기': 'info',
    '방문 완료': 'primary',
    '방문 취소': 'warning',
    '견적 제출': 'info',
    '문의 드랍': 'dark',
    '공사 확정': 'success',
    '공사 취소': 'danger',
    '공사 드랍': 'danger',
};

const STATUS_HEX = {
    '상담 대기': '#adb5bd', '유선 상담': '#6c757d', '부재중': '#fd7e14',
    '방문 예약': '#0d6efd', '방문 대기': '#0dcaf0', '방문 완료': '#0d6efd',
    '방문 취소': '#ffc107', '견적 제출': '#0dcaf0', '문의 드랍': '#343a40',
    '공사 확정': '#198754', '공사 취소': '#dc3545', '공사 드랍': '#dc3545',
};

const STATUS_OPTIONS = [
    '상담 대기', '유선 상담', '부재중',
    '방문 예약', '방문 대기', '방문 완료', '방문 취소',
    '견적 제출', '문의 드랍',
    '공사 확정', '공사 취소', '공사 드랍',
];

// 옛 필드명 alias 호환을 위한 안전 액세서
const F = {
    leadNo:       (r) => r['리드 No'] || '',
    consultTime:  (r) => r['상담 시간'] || '',
    platform:     (r) => r['플랫폼'] || r['거래처'] || '',
    status:       (r) => r['상태'] || '',
    visitDate:    (r) => r['방문 예정일'] || '',
    phone:        (r) => r['고객 연락처'] || r['연락처'] || '',
    email:        (r) => r['이메일'] || '',
    customerName: (r) => r['고객명'] || '',
    address:      (r) => r['방문 주소'] || '',
    content:      (r) => r['문의 내용'] || r['상담 내용'] || '',   // 새 J열 우선, 옛 이름 fallback
    keyword:      (r) => r['키워드'] || '',
    onlineOwner:  (r) => r['온라인 상담자'] || '',
    salesOwner:   (r) => r['영업 담당자'] || r['담당자'] || '',
    lastContact:  (r) => r['마지막 연락일'] || '',
    feedback:     (r) => r['상담 내용'] || r['피드백'] || r['비고'] || '',   // 새 K열 우선, 옛 이름 fallback
};

// 멀티라인 편집이 어울리는 필드
const MULTILINE_FIELDS = new Set(['address', 'content', 'feedback']);

// CSRF 토큰 (메타태그에서 읽어 모든 mutation 요청 헤더에 첨부)
function getCSRFToken() {
    const meta = document.querySelector('meta[name="csrf-token"]');
    return meta ? meta.getAttribute('content') : '';
}

// 공통 mutation 헤더 (JSON + CSRF)
function mutationHeaders() {
    return {
        'Content-Type': 'application/json',
        'X-CSRF-Token': getCSRFToken(),
        'X-Requested-With': 'XMLHttpRequest',
    };
}

// HTML 이스케이프
function esc(s) {
    if (s === null || s === undefined) return '';
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

// ─────────────────────────────────────────────────────────────
// 초기화
// ─────────────────────────────────────────────────────────────
async function initLeadsPage() {
    logger.debug('[LEADS] 리드 관리 페이지 초기화 시작');
    try {
        await loadLeadsData();
        initializeDataTable();
        initializeCalendar();

        leadsFilters = new ModernLeadsFilters();
        await leadsFilters.init();

        leadsFilters.onFilterChange((filteredData) => {
            if (!leadsTable) return;
            leadsTable.clear();
            leadsTable.rows.add(filteredData);
            leadsTable.draw(false);
        });

        leadsFilters.applyFilters(leadsData, true);
        populateModalDropdowns();

        logger.debug('[LEADS] 리드 관리 페이지 초기화 완료');
    } catch (error) {
        logger.error('[LEADS] 페이지 초기화 실패:', error);
        const detail = error && error.message ? `: ${error.message}` : '';
        showAlert(`페이지 로드 중 오류가 발생했습니다${detail}`, 'error');
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initLeadsPage);
} else {
    initLeadsPage();
}

// ─────────────────────────────────────────────────────────────
// 데이터 로드
// ─────────────────────────────────────────────────────────────
async function loadLeadsData(forceRefresh = false) {
    try {
        const url = forceRefresh ? '/leads/api/list?force_refresh=true' : '/leads/api/list';
        const response = await fetch(url, {
            credentials: 'same-origin',
            headers: { Accept: 'application/json', 'X-Requested-With': 'XMLHttpRequest' },
        });

        if (response.status === 401 || response.status === 440) {
            const redirectUrl = new URL('/login', window.location.origin);
            redirectUrl.searchParams.set('session_expired', 'true');
            window.location.href = redirectUrl.toString();
            return [];
        }
        const ct = response.headers.get('content-type');
        if (!ct || !ct.includes('application/json')) {
            const redirectUrl = new URL('/login', window.location.origin);
            redirectUrl.searchParams.set('session_expired', 'true');
            window.location.href = redirectUrl.toString();
            return [];
        }
        if (!response.ok) throw new Error(`HTTP ${response.status}`);

        const result = await response.json();
        if (!result.success) throw new Error(result.message || '리드 데이터 로드 실패');

        leadsData = result.data?.leads || [];
        logger.debug(`[LEADS] 리드 데이터 로드 완료: ${leadsData.length}건`);
        return leadsData;
    } catch (error) {
        logger.error('[LEADS] 리드 데이터 로드 실패:', error);
        throw error;
    }
}

// ─────────────────────────────────────────────────────────────
// 인라인 셀 렌더러
// ─────────────────────────────────────────────────────────────
function renderSelect(field, getOptions, opts = {}) {
    return (d, t, r) => {
        const value = F[field](r);
        const options = getOptions();
        const list = options
            .map((o) => `<option value="${esc(o)}"${o === value ? ' selected' : ''}>${esc(o)}</option>`)
            .join('');
        const colorClass = field === 'status' && value
            ? `border-${STATUS_COLORS[value] || 'secondary'}`
            : '';
        return `<select class="js-inline-select ${colorClass}"
                        data-lead-no="${esc(F.leadNo(r))}"
                        data-field="${field}">
                    <option value=""${!value ? ' selected' : ''}>—</option>
                    ${list}
                </select>`;
    };
}

function renderText(field, opts = {}) {
    return (d, t, r) => {
        const value = F[field](r);
        const hasValue = value && value !== '-';
        const displayClass = hasValue ? '' : 'is-empty';
        return `<span class="js-inline-text ${displayClass}"
                      data-lead-no="${esc(F.leadNo(r))}"
                      data-field="${field}"
                      title="${esc(value || '더블클릭으로 편집')}">${esc(value) || '비어있음'}</span>`;
    };
}

function renderDate(field) {
    return (d, t, r) => {
        const value = F[field](r);
        const hasValue = value && value !== '-';
        const displayClass = hasValue ? '' : 'is-empty';
        return `<span class="js-inline-date ${displayClass}"
                      data-lead-no="${esc(F.leadNo(r))}"
                      data-field="${field}"
                      title="${esc(value || '더블클릭으로 편집 (YYYY-MM-DD)')}">${esc(value) || '비어있음'}</span>`;
    };
}

function renderStatusBadge(field) {
    // 상태는 select 대신 badge로 표시하되 클릭하면 select로 전환
    return (d, t, r) => {
        const value = F.status(r);
        const color = STATUS_COLORS[value] || 'secondary';
        const list = STATUS_OPTIONS
            .map((o) => `<option value="${esc(o)}"${o === value ? ' selected' : ''}>${esc(o)}</option>`)
            .join('');
        return `<select class="js-inline-select status-select"
                        data-lead-no="${esc(F.leadNo(r))}"
                        data-field="status"
                        style="color: ${STATUS_HEX[value] || '#6c757d'}; font-weight: 600;">
                    <option value=""${!value ? ' selected' : ''}>—</option>
                    ${list}
                </select>`;
    };
}

// 동적 옵션 추출
function getPlatformOptions() {
    const fromData = [...new Set(leadsData.map(F.platform).filter(Boolean))].sort();
    return fromData.length ? fromData : ['홈페이지', '전화', '채널톡', '당근', '카카오톡', '메일'];
}
function getSalesOwnerOptions() {
    const fromData = [...new Set(leadsData.map(F.salesOwner).filter(Boolean).filter((v) => v !== '-'))].sort();
    return fromData;
}

// ─────────────────────────────────────────────────────────────
// DataTables 초기화 (15열)
// ─────────────────────────────────────────────────────────────
function initializeDataTable() {
    if (leadsTable) leadsTable.destroy();

    leadsTable = new DataTable('#leadsTable', {
        data: leadsData,
        columns: [
            { className: 'col-time',      data: null, render: (d, t, r) => esc(F.consultTime(r)) || '-' },
            { className: 'col-platform',  data: null, render: renderSelect('platform', getPlatformOptions) },
            { className: 'col-status',    data: null, render: renderStatusBadge('status') },
            { className: 'col-visitdate', data: null, render: renderDate('visitDate') },
            { className: 'col-name',      data: null, render: renderText('customerName') },
            { className: 'col-phone',     data: null, render: renderText('phone') },
            { className: 'col-email',     data: null, render: renderText('email') },
            { className: 'col-address',   data: null, render: renderText('address') },
            { className: 'col-content',   data: null, render: renderText('content') },
            { className: 'col-keyword',   data: null, render: renderText('keyword') },
            { className: 'col-online',    data: null, render: renderText('onlineOwner') },
            { className: 'col-sales',     data: null, render: renderSelect('salesOwner', getSalesOwnerOptions) },
            { className: 'col-feedback',  data: null, render: renderText('feedback') },
        ],
        order: [[0, 'desc']],   // 상담 시간 내림차순 (첫 번째 컬럼)
        pageLength: 15,
        lengthMenu: [[15, 25, 50, 100], ['15', '25', '50', '100']],
        pagingType: 'full_numbers',
        searching: false,
        scrollX: false,         // 외부 wrapper에서 가로 스크롤 처리 (CSS로)
        autoWidth: false,
        dom: '<"top"l>rt<"bottom d-flex justify-content-between"<"info-left"i><"paging-right"p>><"clear">',
        language: {
            lengthMenu: '_MENU_개씩 보기',
            info: '_START_~_END_ / 전체 _TOTAL_개 (페이지 _PAGE_ / _PAGES_)',
            infoEmpty: '0개',
            infoFiltered: '',
            paginate: { first: '처음', last: '마지막', next: '다음', previous: '이전' },
            emptyTable: '데이터가 없습니다',
        },
        drawCallback: function () {
            const tableEl = document.getElementById('leadsTable');
            if (tableEl) tableEl.style.visibility = 'visible';

            const filterResultCount = document.getElementById('filterResultCount');
            if (filterResultCount) {
                let api = null;
                try { api = this.api ? this.api() : null; } catch (_) {}
                if (!api && leadsTable) api = leadsTable;
                if (api) {
                    const visible = api.rows({ filter: 'applied' }).nodes().length;
                    filterResultCount.textContent = `${visible}개 리드 표시`;
                }
            }
        },
    });

    logger.debug('[LEADS] DataTables 초기화 완료');
}

// ─────────────────────────────────────────────────────────────
// 인라인 편집: select 변경 / 텍스트·날짜 더블클릭
// ─────────────────────────────────────────────────────────────
function setupInlineEditDelegation() {
    // 1) select 즉시 변경 (상태/플랫폼/영업 담당자)
    document.addEventListener('change', async (e) => {
        const sel = e.target.closest('.js-inline-select');
        if (!sel) return;
        const leadNo = sel.dataset.leadNo;
        const field = sel.dataset.field;
        const newValue = sel.value;
        const ko = KO[field];
        if (!leadNo || !ko) return;

        // 로컬 데이터 우선 갱신 + DOM 효과
        await updateField(leadNo, ko, newValue, sel.closest('td'));
        // 상태 색상 즉시 반영
        if (field === 'status') sel.style.color = STATUS_HEX[newValue] || '#6c757d';
    });

    // select 클릭 시 곧바로 드롭다운 펼침(클릭 한 번)
    document.addEventListener('mousedown', (e) => {
        const sel = e.target.closest('.js-inline-select');
        if (sel) e.stopPropagation();   // DataTables 정렬 등에 방해받지 않게
    });

    // 2) 텍스트·날짜 더블클릭 → 편집
    document.addEventListener('dblclick', (e) => {
        const textSpan = e.target.closest('.js-inline-text');
        if (textSpan) {
            startInlineEdit(textSpan, 'text');
            return;
        }
        const dateSpan = e.target.closest('.js-inline-date');
        if (dateSpan) {
            startInlineEdit(dateSpan, 'date');
        }
    });
}

function startInlineEdit(span, type) {
    const td = span.closest('td');
    const leadNo = span.dataset.leadNo;
    const field = span.dataset.field;
    const ko = KO[field];
    if (!leadNo || !ko) return;

    // 원래 값 (placeholder '비어있음'은 빈값으로 처리)
    const lead = leadsData.find((l) => F.leadNo(l) === leadNo);
    const originalValue = lead ? (F[field](lead) || '') : '';

    // input/textarea 생성
    let input;
    if (type === 'text' && MULTILINE_FIELDS.has(field)) {
        input = document.createElement('textarea');
        input.rows = 3;
    } else {
        input = document.createElement('input');
        input.type = type === 'date' ? 'date' : 'text';
    }
    input.className = 'js-inline-input';
    input.value = originalValue;

    // 더블클릭 텍스트 선택 해제
    window.getSelection().removeAllRanges();

    span.style.display = 'none';
    td.appendChild(input);
    input.focus();
    if (input.select) input.select();

    let finished = false;

    const cleanup = (commit) => {
        if (finished) return;
        finished = true;

        if (commit && input.value !== originalValue) {
            updateField(leadNo, ko, input.value, td).then(() => {
                // updateField가 leadsTable.draw()를 호출해 셀이 재렌더링됨
            });
        } else {
            // 취소: input 제거, span 복원
            input.remove();
            span.style.display = '';
        }
    };

    input.addEventListener('blur', () => cleanup(true));
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            if (input.tagName === 'TEXTAREA' && !e.ctrlKey && !e.metaKey) return;  // textarea는 Ctrl+Enter로 저장
            e.preventDefault();
            cleanup(true);
        } else if (e.key === 'Escape') {
            e.preventDefault();
            cleanup(false);
        }
    });
}

async function updateField(leadNo, koField, value, cellEl = null) {
    if (cellEl) cellEl.classList.add('js-saving');
    try {
        const response = await fetch(`/leads/api/update/${leadNo}`, {
            method: 'PUT',
            headers: mutationHeaders(),
            credentials: 'same-origin',
            body: JSON.stringify({ [koField]: value }),
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const result = await response.json();
        if (!result.success) throw new Error(result.message);

        // 로컬 데이터 갱신
        const lead = leadsData.find((l) => F.leadNo(l) === leadNo);
        if (lead) lead[koField] = value;

        // 테이블 부분 갱신 (정렬·페이지 유지)
        if (leadsTable) {
            leadsTable.row((idx, data) => F.leadNo(data) === leadNo).invalidate().draw(false);
        }

        // 성공 플래시
        setTimeout(() => {
            const refreshed = document.querySelector(`[data-lead-no="${CSS.escape(leadNo)}"][data-field]`);
            const flashCell = refreshed ? refreshed.closest('td') : cellEl;
            if (flashCell) {
                flashCell.classList.add('save-indicator-flash');
                setTimeout(() => flashCell.classList.remove('save-indicator-flash'), 600);
            }
        }, 50);
    } catch (error) {
        logger.error('[LEADS] 인라인 수정 실패:', error);
        showAlert(`${koField} 저장 실패: ${error.message || ''}`, 'error');
        // 실패 시 테이블 재렌더링하여 원복
        if (leadsTable) leadsTable.draw(false);
    } finally {
        if (cellEl) cellEl.classList.remove('js-saving');
    }
}

// ─────────────────────────────────────────────────────────────
// FullCalendar
// ─────────────────────────────────────────────────────────────
function initializeCalendar() {
    const calendarEl = document.getElementById('calendar');
    if (!calendarEl) return;

    calendar = new Calendar(calendarEl, {
        plugins: [dayGridPlugin, timeGridPlugin, interactionPlugin],
        initialView: 'dayGridMonth',
        locale: koLocale,
        headerToolbar: {
            left: 'prev,next today', center: 'title',
            right: 'dayGridMonth,timeGridWeek,timeGridDay',
        },
        editable: true,
        droppable: true,
        eventDrop: handleEventDrop,
        eventClick: handleEventClick,
        events: [],
    });
}

function loadCalendarEvents() {
    const events = leadsData
        .filter((lead) => F.visitDate(lead))
        .map((lead) => ({
            id: F.leadNo(lead),
            title: `${F.customerName(lead)} - ${F.status(lead)}`,
            start: F.visitDate(lead),
            backgroundColor: STATUS_HEX[F.status(lead)] || '#6c757d',
        }));
    calendar.removeAllEvents();
    calendar.addEventSource(events);
}

async function handleEventDrop(info) {
    const leadNo = info.event.id;
    const newDate = info.event.startStr;
    try {
        const response = await fetch(`/leads/api/update/${leadNo}`, {
            method: 'PUT',
            headers: mutationHeaders(),
            credentials: 'same-origin',
            body: JSON.stringify({ '방문 예정일': newDate }),
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const result = await response.json();
        if (!result.success) throw new Error(result.message);

        showAlert('방문 예정일이 변경되었습니다', 'success');
        const lead = leadsData.find((l) => F.leadNo(l) === leadNo);
        if (lead) lead['방문 예정일'] = newDate;
        if (leadsTable) leadsTable.draw(false);
    } catch (error) {
        logger.error('[LEADS] 방문 예정일 변경 실패:', error);
        showAlert('방문 예정일 변경 중 오류가 발생했습니다', 'error');
        info.revert();
    }
}

function handleEventClick(info) {
    // 캘린더에서 클릭하면 편집 모달 열기
    openEditLeadModal(info.event.id);
}

function switchToTableView() {
    currentView = 'table';
    document.getElementById('tableView').style.display = 'block';
    document.getElementById('calendarView').style.display = 'none';
    document.getElementById('filterSection').style.display = 'block';
    document.getElementById('tableViewBtn').classList.add('active');
    document.getElementById('calendarViewBtn').classList.remove('active');
}

function switchToCalendarView() {
    currentView = 'calendar';
    document.getElementById('tableView').style.display = 'none';
    document.getElementById('calendarView').style.display = 'block';
    document.getElementById('filterSection').style.display = 'none';
    document.getElementById('tableViewBtn').classList.remove('active');
    document.getElementById('calendarViewBtn').classList.add('active');
    calendar.render();
    loadCalendarEvents();
}

// ─────────────────────────────────────────────────────────────
// 모달 드롭다운
// ─────────────────────────────────────────────────────────────
function populateModalDropdowns() {
    const platforms = getPlatformOptions();
    const salesOwners = getSalesOwnerOptions();

    fillSelect('newLeadPlatform', platforms, '선택하세요');
    fillSelect('newLeadSalesOwner', salesOwners, '선택하세요');
    fillSelect('editLeadPlatform', platforms, '');
    fillSelect('editLeadSalesOwner', salesOwners, '');
}

function fillSelect(id, options, placeholder) {
    const el = document.getElementById(id);
    if (!el) return;
    const current = el.value;
    el.innerHTML = '';
    if (placeholder !== null && placeholder !== undefined) {
        const opt = document.createElement('option');
        opt.value = '';
        opt.textContent = placeholder;
        el.appendChild(opt);
    }
    options.forEach((o) => {
        const opt = document.createElement('option');
        opt.value = o;
        opt.textContent = o;
        el.appendChild(opt);
    });
    if (current && options.includes(current)) el.value = current;
}

// ─────────────────────────────────────────────────────────────
// 신규 리드 등록 / 편집 모달
// ─────────────────────────────────────────────────────────────
function openNewLeadModal() {
    document.getElementById('newLeadForm').reset();
    new bootstrap.Modal(document.getElementById('newLeadModal')).show();
}

async function saveNewLead() {
    try {
        const leadData = {
            '플랫폼': document.getElementById('newLeadPlatform').value,
            '상태': document.getElementById('newLeadStatus').value || '유선 상담',
            '영업 담당자': document.getElementById('newLeadSalesOwner').value,
            '상담 시간': document.getElementById('newLeadConsultTime').value,
            '온라인 상담자': document.getElementById('newLeadOnlineOwner').value,
            '고객명': document.getElementById('newLeadCustomerName').value,
            '고객 연락처': document.getElementById('newLeadPhone').value,
            '이메일': document.getElementById('newLeadEmail').value,
            '방문 주소': document.getElementById('newLeadAddress').value,
            '방문 예정일': document.getElementById('newLeadVisitDate').value,
            '마지막 연락일': document.getElementById('newLeadLastContact').value,
            '문의 내용': document.getElementById('newLeadContent').value,   // 인입 원본
            '키워드': document.getElementById('newLeadKeyword').value,
            '상담 내용': document.getElementById('newLeadFeedback').value,   // 매니저 처리 결과 (옛 피드백)
        };

        if (!leadData['플랫폼'] || !leadData['고객명'] || !leadData['고객 연락처']) {
            showAlert('필수 항목(플랫폼/고객명/고객 연락처)을 모두 입력해주세요', 'warning');
            return;
        }

        const response = await fetch('/leads/api/create', {
            method: 'POST',
            headers: mutationHeaders(),
            credentials: 'same-origin',
            body: JSON.stringify(leadData),
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const result = await response.json();
        if (!result.success) throw new Error(result.message);

        showAlert(`리드 ${result.data?.lead_no || ''}가 생성되었습니다`, 'success');
        bootstrap.Modal.getInstance(document.getElementById('newLeadModal')).hide();
        await loadLeadsData(true);
        leadsFilters.applyFilters(leadsData, true);
        if (currentView === 'calendar') loadCalendarEvents();
    } catch (error) {
        logger.error('[LEADS] 새 리드 생성 실패:', error);
        showAlert('리드 생성 중 오류가 발생했습니다', 'error');
    }
}

function openEditLeadModal(leadNo) {
    const lead = leadsData.find((l) => F.leadNo(l) === leadNo);
    if (!lead) {
        showAlert('리드를 찾을 수 없습니다', 'error');
        return;
    }

    document.getElementById('editLeadNo').value = F.leadNo(lead);
    document.getElementById('editLeadNoDisplay').value = F.leadNo(lead);
    document.getElementById('editLeadStatus').value = F.status(lead) || '유선 상담';
    document.getElementById('editLeadPlatform').value = F.platform(lead) || '';
    document.getElementById('editLeadSalesOwner').value = F.salesOwner(lead) || '';
    document.getElementById('editLeadConsultTime').value = F.consultTime(lead);
    document.getElementById('editLeadOnlineOwner').value = F.onlineOwner(lead);
    document.getElementById('editLeadCustomerName').value = F.customerName(lead);
    document.getElementById('editLeadPhone').value = F.phone(lead);
    document.getElementById('editLeadEmail').value = F.email(lead);
    document.getElementById('editLeadAddress').value = F.address(lead);
    document.getElementById('editLeadVisitDate').value = F.visitDate(lead);
    document.getElementById('editLeadLastContact').value = F.lastContact(lead);
    document.getElementById('editLeadContent').value = F.content(lead);
    document.getElementById('editLeadKeyword').value = F.keyword(lead);
    document.getElementById('editLeadFeedback').value = F.feedback(lead);

    new bootstrap.Modal(document.getElementById('editLeadModal')).show();
}

async function saveEditLead() {
    try {
        const leadNo = document.getElementById('editLeadNo').value;
        const updateData = {
            '상태': document.getElementById('editLeadStatus').value,
            '플랫폼': document.getElementById('editLeadPlatform').value,
            '영업 담당자': document.getElementById('editLeadSalesOwner').value,
            '상담 시간': document.getElementById('editLeadConsultTime').value,
            '온라인 상담자': document.getElementById('editLeadOnlineOwner').value,
            '고객명': document.getElementById('editLeadCustomerName').value,
            '고객 연락처': document.getElementById('editLeadPhone').value,
            '이메일': document.getElementById('editLeadEmail').value,
            '방문 주소': document.getElementById('editLeadAddress').value,
            '방문 예정일': document.getElementById('editLeadVisitDate').value,
            '마지막 연락일': document.getElementById('editLeadLastContact').value,
            '문의 내용': document.getElementById('editLeadContent').value,
            '키워드': document.getElementById('editLeadKeyword').value,
            '상담 내용': document.getElementById('editLeadFeedback').value,
        };

        const response = await fetch(`/leads/api/update/${leadNo}`, {
            method: 'PUT',
            headers: mutationHeaders(),
            credentials: 'same-origin',
            body: JSON.stringify(updateData),
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const result = await response.json();
        if (!result.success) throw new Error(result.message);

        showAlert('리드가 수정되었습니다', 'success');
        bootstrap.Modal.getInstance(document.getElementById('editLeadModal')).hide();

        await loadLeadsData(true);
        leadsFilters.applyFilters(leadsData, true);
        if (currentView === 'calendar') loadCalendarEvents();
    } catch (error) {
        logger.error('[LEADS] 리드 수정 실패:', error);
        showAlert('리드 수정 중 오류가 발생했습니다', 'error');
    }
}

// ─────────────────────────────────────────────────────────────
// 알림
// ─────────────────────────────────────────────────────────────
function showAlert(message, type = 'info') {
    const alertContainer = document.getElementById('headerAlertContainer');
    if (!alertContainer) return;
    const typeMap = { success: 'alert-success', error: 'alert-danger', warning: 'alert-warning', info: 'alert-info' };
    const klass = typeMap[type] || 'alert-info';
    alertContainer.innerHTML = `
        <div class="alert ${klass} alert-dismissible fade show mb-0 py-1 px-3" role="alert" style="font-size: 0.9rem;">
            ${esc(message)}
            <button type="button" class="btn-close btn-close-sm" data-bs-dismiss="alert"></button>
        </div>
    `;
    setTimeout(() => { alertContainer.innerHTML = ''; }, 3000);
}

// ─────────────────────────────────────────────────────────────
// 이벤트 리스너
// ─────────────────────────────────────────────────────────────
function setupEventListeners() {
    document.getElementById('tableViewBtn')?.addEventListener('click', switchToTableView);
    document.getElementById('calendarViewBtn')?.addEventListener('click', switchToCalendarView);
    document.getElementById('newLeadBtn')?.addEventListener('click', openNewLeadModal);
    document.getElementById('saveNewLeadBtn')?.addEventListener('click', saveNewLead);
    document.getElementById('saveEditLeadBtn')?.addEventListener('click', saveEditLead);

    setupInlineEditDelegation();

    logger.debug('[LEADS] 이벤트 리스너 등록 완료');
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(setupEventListeners, 100));
} else {
    setTimeout(setupEventListeners, 100);
}

export default { loadLeadsData, openEditLeadModal, switchToTableView, switchToCalendarView };

// 하위 호환
window.openEditLeadModal = openEditLeadModal;
