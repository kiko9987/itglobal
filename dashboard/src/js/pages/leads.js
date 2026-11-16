/**
 * 고객 리드 관리 페이지 — 조회·관리 전용 (리드 등록은 슬랙이 단일 창구, 2026-09-30 결정)
 *
 * 골조는 프로젝트 페이지와 동일: 필터 → 읽기 전용 메인 테이블 → 행 클릭 아코디언.
 *   LeadTable(테이블·모드) / LeadRowAccordion(상세) / ModernLeadsFilters(필터·프리셋)
 * 헤더 [테이블 | 캘린더] 보기 전환 (캘린더는 현재 필터 결과의 방문 예정일을 표시).
 */
import '../../css/pages/leads.css';
import logger from '../utils/logger.js';
import ModernLeadsFilters from '../components/ModernLeadsFilters.js';
import LeadTable from '../components/LeadTable.js';
import LeadRowAccordion from '../components/LeadRowAccordion.js';
import MobileChrome from '../components/MobileChrome.js';
import LeadMobileView from '../components/LeadMobileView.js';
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

// 흰 글자가 안 읽히는 밝은 상태 색(캘린더 일정 글자를 진하게) — PC 점검 2차 11-16
const LIGHT_STATUS_BG = new Set(['#adb5bd', '#fd7e14', '#ffc107', '#0dcaf0']);

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
  // 오류는 직원이 닫을 때까지 남긴다 — 3~8초 뒤 사라져 무엇이 잘못됐는지 못 읽고 빈 표만 남았음(PC 점검 2차 11-16)
  if (type === 'error') duration = 0;
  const box = document.getElementById('headerAlertContainer');
  if (!box) return;
  const cls = { success: 'alert-success', error: 'alert-danger', warning: 'alert-warning', info: 'alert-info' }[type] || 'alert-info';
  box.innerHTML = `
    <div class="alert ${cls} alert-dismissible fade show mb-0 py-1 px-3" role="alert" style="font-size: 0.9rem;">
      ${esc(message)}
      <button type="button" class="btn-close btn-close-sm" data-bs-dismiss="alert" aria-label="알림 닫기"></button>
    </div>`;
  clearTimeout(showPageAlert._t);
  if (duration > 0) showPageAlert._t = setTimeout(() => { box.innerHTML = ''; }, duration);
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
    if (!rows) return;   // 로그인 화면으로 넘어가는 중 — 성공 알림을 띄우지 않음
    app.data = rows;
    app.lastLoadedAt = Date.now();
    // 첫 불러오기가 실패해 표가 아직 없으면 표 만들기부터 다시 — 전에는 표도 없이
    //   "최신 데이터를 불러왔습니다"만 떠 빈 화면이 그대로였음(최종 점검 11-16)
    if (!app.table) await setupTable(rows);
    else app.filters.applyFilters(rows, true);
    if (force) showPageAlert('최신 데이터를 불러왔습니다', 'success');
  } catch (err) {
    logger.error('[LEADS] 새로고침 실패:', err);
    showPageAlert('리드를 다시 불러오지 못했습니다. 잠시 뒤 [새로고침]을 눌러 주세요.', 'error');
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
    const name = String(lead['고객명'] || '').trim();
    events.push({
      id: String(lead['리드 No'] || ''),
      title: `${name && name !== '-' ? name : '(이름 없음)'} · ${lead['상태'] || '-'}`,
      textColor: LIGHT_STATUS_BG.has(color) ? '#1f2937' : '#ffffff',
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
  app.mobile?.render();   // 휴대폰: 카드 목록 ↔ 방문 일정 목록 (넓은 화면에선 아무것도 안 함)
}

// ── 표 만들기 (첫 불러오기, 또는 첫 불러오기 실패 뒤 [새로고침]) ─────────────
async function setupTable(rows) {
  // 필터 연결·보기 버튼·탭 복귀 갱신은 한 번만 — 다시 불러도 겹쳐 붙지 않게
  if (!app.wired) {
    app.filters.onFilterChange((filtered, meta = {}) => {
      app.filtered = filtered;
      app.table?.updateData(filtered, { keepPage: !!meta.dataRefresh });
      app.mobile?.update(filtered);
      app.chrome?.syncFilterCount?.();
      if (app.view === 'calendar') loadCalendarEvents();
    });

    document.getElementById('tableViewBtn')?.addEventListener('click', () => setView('table'));
    document.getElementById('calendarViewBtn')?.addEventListener('click', () => setView('calendar'));

    // 탭 복귀 시 5분 넘었으면 조용히 갱신 (아코디언 열려 있으면 보류 — 프로젝트 페이지와 동일)
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible' && Date.now() - app.lastLoadedAt > 5 * 60 * 1000
          && !app.accordion?.isOpen) {
        refreshData(false);
      }
    });
    app.wired = true;
  }

  // 첫 필터 결과로 테이블 생성 (populate + 세션 복원 필터 적용)
  const table = new LeadTable({ filters: app.filters, accordion: app.accordion });
  app.table = table;
  try {
    table.init([]);
    app.filters.applyFilters(rows, true);
    await table.firstDraw;
  } catch (err) {
    app.table = null;   // 반쯤 만든 표는 버림 — 다음 [새로고침]에서 처음부터 다시
    throw err;
  }
  hideLoadingOverlay();
  window.leadsApp = app;
}

// ── 초기화 ────────────────────────────────────────────────
async function init() {
  try {
    app.filters = new ModernLeadsFilters();
    app.filters.init();
    app.accordion = new LeadRowAccordion();
    // 휴대폰 폭에서만: 머리 ⋯ 메뉴·필터 접기·카드 목록·화면 가득 상세 (넓은 화면은 CSS 로 그대로, 2026-11-16)
    try {
      app.chrome = new MobileChrome({
        pageClass: 'pmm-leads',
        menu: [{ id: 'fullRefreshBtn', icon: 'fa-sync-alt', label: '새로고침' }],
      });
      app.mobile = new LeadMobileView({ accordion: app.accordion, getView: () => app.view });
    } catch (err) {
      logger.error('[LEADS] 휴대폰 화면 준비 실패:', err);
    }

    const rows = await fetchLeads(false);
    if (!rows) return;
    app.data = rows;
    app.lastLoadedAt = Date.now();

    await setupTable(rows);
    logger.debug('[LEADS] 초기화 완료');
  } catch (err) {
    logger.error('[LEADS] 초기화 실패:', err);
    hideLoadingOverlay();
    showPageAlert('리드 목록을 불러오지 못했습니다. 잠시 뒤 [새로고침]을 눌러 주세요.', 'error');
  }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();
