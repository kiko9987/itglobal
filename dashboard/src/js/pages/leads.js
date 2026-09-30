/**
 * 고객 리드 관리 페이지 — 조회·관리 전용 (리드 등록은 슬랙이 단일 창구, 2026-09-30 결정)
 *
 * 골조는 프로젝트 페이지와 동일: 필터 → 읽기 전용 메인 테이블 → 행 클릭 아코디언.
 *   LeadTable(테이블·모드) / LeadRowAccordion(상세) / ModernLeadsFilters(필터·프리셋)
 * 헤더 [테이블 | 캘린더] 보기 전환 (캘린더는 현재 필터 결과의 방문 예정일을 표시).
 */
import 'datatables.net-bs5/css/dataTables.bootstrap5.min.css';
import '../../css/pages/leads.css';
import logger from '../utils/logger.js';
import ModernLeadsFilters from '../components/ModernLeadsFilters.js';
import LeadTable from '../components/LeadTable.js';
import LeadRowAccordion from '../components/LeadRowAccordion.js';
import { splitVisitRange, shiftIso, esc } from '../utils/leadFormat.js';

import { Calendar } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import koLocale from '@fullcalendar/core/locales/ko';

const STATUS_HEX = {
  '상담 대기': '#adb5bd', '유선 상담': '#6c757d', '부재중': '#fd7e14',
  '방문 예약': '#0d6efd', '방문 완료': '#6610f2', '방문 취소': '#ffc107',
  '견적 제출': '#0dcaf0', '문의 드랍': '#343a40',
  '공사 확정': '#198754', '공사 취소': '#dc3545',
};

const app = {
  data: [],
  filtered: [],
  filters: null,
  table: null,
  accordion: null,
  calendar: null,
  view: 'table',          // 'table' | 'calendar'
  lastLoadedAt: 0,
};

// ── 헤더 알림 (프로젝트 페이지 showPageAlert 와 같은 동작) ────────
function showPageAlert(message, type = 'info', duration = 3000) {
  const box = document.getElementById('headerAlertContainer');
  if (!box) return;
  const cls = { success: 'alert-success', error: 'alert-danger', warning: 'alert-warning', info: 'alert-info' }[type] || 'alert-info';
  box.innerHTML = `
    <div class="alert ${cls} alert-dismissible fade show mb-0 py-1 px-3" role="alert" style="font-size: 0.9rem;">
      ${esc(message)}
      <button type="button" class="btn-close btn-close-sm" data-bs-dismiss="alert"></button>
    </div>`;
  clearTimeout(showPageAlert._t);
  showPageAlert._t = setTimeout(() => { box.innerHTML = ''; }, duration);
}
window.showPageAlert = showPageAlert;

// ── 데이터 로드 ────────────────────────────────────────────
async function fetchLeads(force = false) {
  // 기타(ETC)는 서버에서 함께 받아 필터의 [기타 포함] 토글로 거른다
  const url = `/leads/api/list?include_etc=true${force ? '&force_refresh=true' : ''}`;
  const res = await fetch(url, {
    credentials: 'same-origin',
    headers: { Accept: 'application/json', 'X-Requested-With': 'XMLHttpRequest' },
  });
  const ct = res.headers.get('content-type') || '';
  if (res.status === 401 || res.status === 440 || !ct.includes('application/json')) {
    const u = new URL('/login', window.location.origin);
    u.searchParams.set('session_expired', 'true');
    window.location.href = u.toString();
    return null;
  }
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  if (!json.success) throw new Error(json.message || '리드 데이터 로드 실패');
  return json.data?.leads || [];
}

function hideLoadingOverlay() {
  const overlay = document.getElementById('tableLoadingOverlay');
  if (overlay) overlay.style.display = 'none';
  const wrapper = document.querySelector('.table-section .table-wrapper');
  if (wrapper) {
    wrapper.style.visibility = 'visible';
    wrapper.style.opacity = '1';
  }
}

async function refreshData(force = true) {
  if (app.accordion?.isOpen && !force) return;
  try {
    const btn = document.getElementById('fullRefreshBtn');
    btn?.classList.add('disabled');
    const rows = await fetchLeads(force);
    if (!rows) return;
    app.data = rows;
    app.lastLoadedAt = Date.now();
    app.filters.applyFilters(rows, true);
    if (force) showPageAlert('최신 데이터를 불러왔습니다', 'success');
  } catch (err) {
    logger.error('[LEADS] 새로고침 실패:', err);
    showPageAlert(`새로고침 실패: ${err.message || ''}`, 'error');
  } finally {
    document.getElementById('fullRefreshBtn')?.classList.remove('disabled');
  }
}
window.fullRefresh = () => refreshData(true);

// ── 캘린더 (보기 전용 — 일정 클릭 시 테이블의 그 리드로 이동) ─────────
function ensureCalendar() {
  if (app.calendar) return app.calendar;
  const el = document.getElementById('calendar');
  if (!el) return null;
  app.calendar = new Calendar(el, {
    plugins: [dayGridPlugin, timeGridPlugin],
    initialView: 'dayGridMonth',
    locale: koLocale,
    headerToolbar: { left: 'prev,next today', center: 'title', right: 'dayGridMonth,timeGridWeek' },
    editable: false,
    dayMaxEvents: 5,   // 하루 20건 가까이 몰려 칸이 길어지는 것 방지 → '+N개' 접기
    eventClick: (info) => {
      setView('table');
      if (!app.table.revealLead(info.event.id)) showPageAlert('현재 필터에서 이 리드를 찾을 수 없습니다', 'warning');
    },
    events: [],
  });
  return app.calendar;
}

function loadCalendarEvents() {
  const cal = ensureCalendar();
  if (!cal) return;
  const events = [];
  app.filtered.forEach((lead) => {
    const { start, end } = splitVisitRange(lead['방문 예정일']);
    if (!start) return;
    const color = STATUS_HEX[String(lead['상태'] || '').trim()] || '#6c757d';
    events.push({
      id: String(lead['리드 No'] || ''),
      title: `${lead['고객명'] || '(이름 없음)'} · ${lead['상태'] || '-'}`,
      start,
      end: end ? shiftIso(end, 1) : undefined,   // FullCalendar end 는 배타적
      allDay: true,
      backgroundColor: color,
      borderColor: color,
    });
  });
  cal.removeAllEvents();
  cal.addEventSource(events);
}

function setView(view) {
  app.view = view === 'calendar' ? 'calendar' : 'table';
  const isCal = app.view === 'calendar';
  document.getElementById('tableSection').style.display = isCal ? 'none' : '';
  document.getElementById('calendarView').style.display = isCal ? '' : 'none';
  document.getElementById('tableViewBtn')?.classList.toggle('active', !isCal);
  document.getElementById('calendarViewBtn')?.classList.toggle('active', isCal);
  if (isCal) {
    app.accordion?.detachQuietly();
    ensureCalendar()?.render();
    loadCalendarEvents();
  }
}

// ── 초기화 ────────────────────────────────────────────────
async function init() {
  try {
    app.filters = new ModernLeadsFilters();
    app.filters.init();
    app.accordion = new LeadRowAccordion();

    const rows = await fetchLeads(false);
    if (!rows) return;
    app.data = rows;
    app.lastLoadedAt = Date.now();

    app.table = new LeadTable({ filters: app.filters, accordion: app.accordion });
    app.filters.onFilterChange((filtered, meta = {}) => {
      app.filtered = filtered;
      app.table.updateData(filtered, { keepPage: !!meta.dataRefresh });
      if (app.view === 'calendar') loadCalendarEvents();
    });

    // 첫 필터 결과로 테이블 생성 (populate + 세션 복원 필터 적용)
    app.table.init([]);
    app.filters.applyFilters(rows, true);
    await app.table.firstDraw;
    hideLoadingOverlay();

    document.getElementById('tableViewBtn')?.addEventListener('click', () => setView('table'));
    document.getElementById('calendarViewBtn')?.addEventListener('click', () => setView('calendar'));

    // 탭 복귀 시 5분 넘었으면 조용히 갱신 (아코디언 열려 있으면 보류 — 프로젝트 페이지와 동일)
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible' && Date.now() - app.lastLoadedAt > 5 * 60 * 1000
          && !app.accordion.isOpen) {
        refreshData(false);
      }
    });

    window.leadsApp = app;
    logger.debug('[LEADS] 초기화 완료');
  } catch (err) {
    logger.error('[LEADS] 초기화 실패:', err);
    hideLoadingOverlay();
    showPageAlert(`페이지 로드 중 오류가 발생했습니다${err?.message ? `: ${err.message}` : ''}`, 'error', 8000);
  }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();
