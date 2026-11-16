/**
 * 고객 리드 휴대폰 화면 — 카드 목록·화면 가득 상세·방문 일정 목록 (2026-11-16 휴대폰 리뉴얼 2단계).
 *
 * - 휴대폰 폭(767.98px 이하)에서만 켜진다. 넓은 화면은 DataTables 표·FullCalendar 그대로(CSS 로 숨기기만).
 * - 목록은 페이지 필터(ModernLeadsFilters)가 거른 결과를 그대로 받는다(leads.js 가 update 를 부름).
 * - 카드를 누르면 기존 펼침(LeadRowAccordion, 보기 전용)을 공용 시트(MobileDetailSheet) 안에서 연다.
 * - [캘린더] 보기: 휴대폰에선 달력 칸이 좁아 읽을 수 없어 방문 예정일을 날짜별로 묶은 목록을 대신 보여 준다.
 */
import MobileDetailSheet from './MobileDetailSheet.js';
import {
  splitOwners, splitVisitRange, consultFull, localIso, daysBetween, esc,
} from '../utils/leadFormat.js';

const MOBILE_MQ = '(max-width: 767.98px)';
const PAGE_SIZE = 30;
const WEEKDAY = ['일', '월', '화', '수', '목', '금', '토'];

// 상태 점 색(글자는 늘 진한 색 — 대비는 점이 아니라 글자로 맞춤). leads.js STATUS_HEX 와 같은 계열
const STATUS_DOT = {
  '상담 대기': '#adb5bd', '유선 상담': '#6c757d', '부재중': '#fd7e14',
  '방문 예약': '#0d6efd', '방문 완료': '#6610f2', '방문 취소': '#ffc107',
  '견적 제출': '#0dcaf0', '문의 드랍': '#343a40',
  '공사 확정': '#198754', '공사 취소': '#dc3545',
};
const CANCELLED = ['방문 취소', '공사 취소', '문의 드랍'];

const val = (r, k) => {
  const v = String(r?.[k] ?? '').trim();
  return v === '-' ? '' : v;
};

function telHref(raw) {
  const digits = String(raw || '').replace(/[^0-9+]/g, '');
  return digits.length >= 8 ? `tel:${digits}` : '';
}

/** 2026-11-18 → 11.18(화) */
function dayLabel(iso) {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return '';
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return `${m[2]}.${m[3]}(${WEEKDAY[d.getDay()]})`;
}

function ddayTag(today, start) {
  const dd = daysBetween(today, start);
  if (dd === 0) return '<span class="pmm-dday pmm-dday-today">오늘</span>';
  if (dd === 1) return '<span class="pmm-dday">내일</span>';
  if (dd > 1) return `<span class="pmm-dday">D-${dd}</span>`;
  return '';
}

function statusHtml(status) {
  const dot = STATUS_DOT[status] || '#adb5bd';
  return `<span class="pmm-status"><span class="pmm-dot" style="background:${dot}" aria-hidden="true"></span>${esc(status || '상태 없음')}</span>`;
}


/** 날짜 칸('2026-11-17', "'2026-11-17", '2026-11-17 ~ 11-20')에서 첫 날짜만 */
function firstDate(raw) {
  const m = String(raw || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : '';
}

/** '가까운 순': 오늘 이후(가까운 날 먼저) → 지난 날(최근 먼저) → 날짜 없음 */
function nearFirst(a, b, today) {
  if (!a || !b) return (a ? 0 : 1) - (b ? 0 : 1);
  const fa = a >= today, fb = b >= today;
  if (fa !== fb) return fa ? -1 : 1;
  if (fa) return a < b ? -1 : a > b ? 1 : 0;
  return a > b ? -1 : a < b ? 1 : 0;
}

function sortStore(key, value) {
  try {
    if (value === undefined) return localStorage.getItem(key) || '';
    localStorage.setItem(key, value);
  } catch (_) { /* 기억 못 해도 정렬은 됨 */ }
  return value || '';
}

const LEAD_SORTS = [
  ['', '최신순'],
  ['visit', '방문일 가까운 순'],
];

export default class LeadMobileView {
  constructor({ accordion, getView } = {}) {
    this.acc = accordion;
    this.getView = getView || (() => 'table');
    this.mq = window.matchMedia(MOBILE_MQ);
    this.rows = [];
    this.shown = PAGE_SIZE;
    this.showPast = false;
    this.pastShown = PAGE_SIZE;
    this.lastNo = null;

    this.createList();
    this.createAgenda();
    this.sheet = new MobileDetailSheet({ id: 'pmmLeadSheet', onRequestClose: (fromPop) => this.requestClose(fromPop) });
    this.patchAccordion();
    this.bindEvents();
    this.sync();
  }

  get isActive() {
    return this.mq.matches;
  }

  createList() {
    const section = document.getElementById('tableSection');
    const el = document.createElement('div');
    el.className = 'pmm-list pmm-lead-list';
    el.style.display = 'none';
    el.innerHTML = `
      <h2 class="visually-hidden">고객 리드 목록</h2>
      <div class="pmm-sort">
        <label for="pmmSortLeads">정렬</label>
        <select id="pmmSortLeads" class="pmm-sort-select">
          ${LEAD_SORTS.map(([k, t]) => `<option value="${k}">${t}</option>`).join('')}
        </select>
      </div>
      <div class="pmm-cards"></div>
      <p class="pmm-empty" hidden>조건에 맞는 리드가 없어요. 검색어나 필터를 바꿔 보세요.</p>
      <button type="button" class="pmm-more-btn" hidden></button>`;
    // 표 칸(#tableSection) 안에 둬서 [캘린더] 보기로 바꾸면 같이 숨게
    (section || document.body).appendChild(el);
    this.listBox = el;
    this.listEl = el.querySelector('.pmm-cards');
    this.emptyEl = el.querySelector('.pmm-empty');
    this.moreBtn = el.querySelector('.pmm-more-btn');
    this.sortKey = sortStore('pmm.sort.leads');
    if (!LEAD_SORTS.some(([k]) => k === this.sortKey)) this.sortKey = '';
    this.sortEl = el.querySelector('.pmm-sort-select');
    this.sortEl.value = this.sortKey;
    this.sortEl.addEventListener('change', () => {
      this.sortKey = sortStore('pmm.sort.leads', this.sortEl.value);
      this.shown = PAGE_SIZE;
      this.render();
    });
  }

  /** 고른 정렬로 줄 세움(같으면 리드 No 최신순) */
  sortedRows() {
    if (this.sortKey !== 'visit') return this.rows;
    const today = new Date().toISOString().slice(0, 10);
    return this.rows.map((r, i) => [r, i, firstDate(r['방문 예정일'])])
      .sort((a, b) => nearFirst(a[2], b[2], today) || a[1] - b[1]).map((x) => x[0]);
  }

  createAgenda() {
    const cal = document.getElementById('calendarView');
    const el = document.createElement('div');
    el.className = 'pmm-agenda';
    el.style.display = 'none';
    // 위쪽 줄: [내 리드만](표 위 '내 리드만 보기'와 같은 값 — 일정 보기에선 그 줄이 표와 함께 숨으므로) · [지난 방문]
    el.innerHTML = `
      <div class="pmm-agenda-bar">
        <button type="button" class="pmm-chip pmm-agenda-mine" aria-pressed="false">
          <i class="fas fa-user-check" aria-hidden="true"></i>내 리드만</button>
        <button type="button" class="pmm-chip pmm-agenda-past" aria-pressed="false"></button>
      </div>
      <p class="pmm-agenda-note" tabindex="-1"></p>
      <div class="pmm-agenda-body"></div>
      <button type="button" class="pmm-more-btn pmm-agenda-more" hidden></button>`;
    // 달력보다 앞에(휴대폰에선 달력을 CSS 로 숨김)
    if (cal) cal.insertBefore(el, cal.firstChild);
    this.agendaBox = el;
    this.agendaNote = el.querySelector('.pmm-agenda-note');
    this.agendaBody = el.querySelector('.pmm-agenda-body');
    this.pastBtn = el.querySelector('.pmm-agenda-past');
    this.mineBtn = el.querySelector('.pmm-agenda-mine');
    this.agendaMore = el.querySelector('.pmm-agenda-more');
  }

  bindEvents() {
    const onMq = () => this.sync();
    if (this.mq.addEventListener) this.mq.addEventListener('change', onMq);
    else this.mq.addListener(onMq);

    const onCard = (e) => {
      if (e.target.closest('a')) return;   // 카드 안 [전화]는 그대로 전화
      const card = e.target.closest('.pmm-card');
      if (card) this.openDetail(card.dataset.no, card);
    };
    const onKey = (e) => {
      if ((e.key !== 'Enter' && e.key !== ' ') || e.target.closest('a')) return;
      const card = e.target.closest('.pmm-card');
      if (!card) return;
      e.preventDefault();
      this.openDetail(card.dataset.no, card);
    };
    [this.listEl, this.agendaBody].forEach((el) => {
      el.addEventListener('click', onCard);
      el.addEventListener('keydown', onKey);
    });
    this.moreBtn.addEventListener('click', () => {
      this.shown += PAGE_SIZE;
      this.renderList();
    });
    this.pastBtn.addEventListener('click', () => {
      this.showPast = !this.showPast;
      this.pastShown = PAGE_SIZE;
      this.renderAgenda();
      // 바꾼 뒤 맨 위 안내 줄로(아래쪽 단추에 포커스가 남아 화면이 한참 아래에 있던 것)
      this.agendaBox.scrollIntoView({ block: 'start' });
      this.agendaNote.focus({ preventScroll: true });
    });
    this.agendaMore.addEventListener('click', () => {
      this.pastShown += PAGE_SIZE;
      this.renderAgenda();
    });
    this.mineBtn.addEventListener('click', () => {
      document.getElementById('myLeadsOnly')?.click();   // 필터가 다시 돌며 update → render
    });
  }

  /** 휴대폰 ↔ 넓은 화면 전환 */
  sync() {
    const on = this.isActive;
    this.listBox.style.display = on ? '' : 'none';
    this.agendaBox.style.display = on ? '' : 'none';
    this.relabelViewButtons(on);
    if (on) this.render();
    else if (this.sheet.isOpen) this.requestClose();
  }

  /** 휴대폰에선 [테이블|캘린더] → [목록|일정] (표·달력 대신 카드·날짜별 목록을 보여 주므로) */
  relabelViewButtons(on) {
    [['tableViewBtn', 'fa-list', '목록'], ['calendarViewBtn', 'fa-calendar-day', '일정']].forEach(([id, icon, text]) => {
      const b = document.getElementById(id);
      if (!b || (!on && b.dataset.pcHtml === undefined)) return;   // 넓은 화면은 손대지 않음
      if (b.dataset.pcHtml === undefined) b.dataset.pcHtml = b.innerHTML;
      b.innerHTML = on ? `<i class="fas ${icon} me-1" aria-hidden="true"></i>${text}` : b.dataset.pcHtml;
    });
  }

  /** 필터 결과·보기 전환 때 leads.js 가 부른다 */
  update(rows) {
    // PC 표와 같은 순서: 리드 No 최신순, ETC-(번호 없는 기타)는 맨 뒤
    const key = (r) => {
      const m = val(r, '리드 No').match(/^L-(\d+)$/);
      return m ? Number(m[1]) : -1;
    };
    this.rows = (Array.isArray(rows) ? [...rows] : []).sort((a, b) => key(b) - key(a));
    this.shown = PAGE_SIZE;
    if (this.isActive) this.render();
  }

  render() {
    if (!this.isActive) return;
    const cal = this.getView() === 'calendar';
    // 휴대폰에서만: 보기 단추에 눌림 상태(화면 낭독기용)
    document.getElementById('tableViewBtn')?.setAttribute('aria-pressed', cal ? 'false' : 'true');
    document.getElementById('calendarViewBtn')?.setAttribute('aria-pressed', cal ? 'true' : 'false');
    if (cal) this.renderAgenda();
    else this.renderList();
  }

  renderList() {
    const rows = this.sortedRows();
    const list = rows.slice(0, this.shown);
    this.listEl.innerHTML = list.map((r) => this.cardHtml(r)).join('');
    this.emptyEl.hidden = rows.length > 0;
    const rest = rows.length - list.length;
    this.moreBtn.hidden = rest <= 0;
    if (rest > 0) this.moreBtn.textContent = `더 보기 (${rest.toLocaleString('ko-KR')}건 남음)`;
  }

  /** 방문 일정: 오늘부터 앞으로(가까운 날 먼저), [지난 방문 보기]를 누르면 지난 30일(최근 먼저) */
  renderAgenda() {
    const today = localIso(new Date());
    const withDate = [];
    this.rows.forEach((r) => {
      const { start } = splitVisitRange(r['방문 예정일']);
      if (start) withDate.push({ r, start });
    });
    const upcoming = withDate.filter((x) => x.start >= today).sort((a, b) => a.start.localeCompare(b.start));
    const past = withDate.filter((x) => x.start < today && daysBetween(x.start, today) <= 30)
      .sort((a, b) => b.start.localeCompare(a.start));
    const items = this.showPast ? past.slice(0, this.pastShown) : upcoming;

    const groups = [];
    items.forEach((x) => {
      const g = groups[groups.length - 1];
      if (g && g.day === x.start) g.items.push(x.r);
      else groups.push({ day: x.start, items: [x.r] });
    });
    this.agendaNote.textContent = this.showPast
      ? `지난 30일 방문 ${past.length.toLocaleString('ko-KR')}건 (최근 먼저)`
      : `오늘부터 방문 예정 ${upcoming.length.toLocaleString('ko-KR')}건 · 지금 걸린 필터 기준`;
    this.agendaBody.innerHTML = groups.length
      ? groups.map((g) => `
          <section class="pmm-agenda-day">
            <h3 class="pmm-agenda-head">${esc(dayLabel(g.day))} ${ddayTag(today, g.day)}
              <span class="pmm-agenda-n">${g.items.length}건</span></h3>
            <div class="pmm-cards">${g.items.map((r) => this.cardHtml(r, { agenda: true })).join('')}</div>
          </section>`).join('')
      : `<p class="pmm-empty">${this.showPast ? '지난 30일 방문이 없어요.' : '앞으로 잡힌 방문이 없어요.'}</p>`;
    this.pastBtn.innerHTML = '<i class="fas fa-history" aria-hidden="true"></i>'
      + (this.showPast ? '앞으로 방문' : `지난 방문 ${past.length.toLocaleString('ko-KR')}건`);
    this.pastBtn.setAttribute('aria-pressed', this.showPast ? 'true' : 'false');
    const rest = this.showPast ? past.length - items.length : 0;
    this.agendaMore.hidden = rest <= 0;
    if (rest > 0) this.agendaMore.textContent = `더 보기 (${rest.toLocaleString('ko-KR')}건 남음)`;
    const mine = document.getElementById('myLeadsOnly');
    this.mineBtn.hidden = !mine || mine.disabled;
    this.mineBtn.setAttribute('aria-pressed', mine?.checked ? 'true' : 'false');
  }

  cardHtml(r, { agenda = false } = {}) {
    const no = val(r, '리드 No');
    const status = val(r, '상태');
    const name = val(r, '고객명');
    const platform = val(r, '플랫폼');
    const addr = val(r, '방문 주소');
    const tel = telHref(r['고객 연락처']);
    const consult = splitOwners(r['온라인 상담자']).join(', ');
    const sales = splitOwners(r['영업 담당자']).join(', ');
    const rawVisit = String(r['방문 예정일'] || '').replace(/^'/, '').trim();
    const { start } = splitVisitRange(rawVisit);
    const today = localIso(new Date());

    let when;
    // 방문일: 하루면 '11.17(화)', 기간이면 시트 값 그대로
    const visitText = /~/.test(rawVisit) ? rawVisit : dayLabel(start);
    if (start && !agenda) when = `방문 ${esc(visitText)} ${ddayTag(today, start)}`;
    else if (!start) when = `상담 ${esc(consultFull(r['상담 시간']) || '-')}`;
    else when = '';   // 일정 목록은 날짜 머리에 이미 있음

    const people = [consult && `상담 ${consult}`, sales ? `영업 ${sales}` : (status === '방문 예약' ? '영업 미배정' : '')]
      .filter(Boolean).join(' · ');

    return `
      <div class="pmm-card pmm-lead-card${CANCELLED.includes(status) ? ' is-cancelled' : ''}" role="button" tabindex="0"
           data-no="${esc(no)}" aria-label="${esc(`${no} ${status} ${name}`)} 상세 보기">
        <div class="pmm-card-top">
          ${statusHtml(status)}
          <span class="pmm-code">${esc(no)}</span>
        </div>
        <div class="pmm-addr">${esc(name || '이름 없음')}${platform ? ` <span class="pmm-platform">${esc(platform)}</span>` : ''}</div>
        <div class="pmm-partner">${esc(addr || '방문 주소 없음')}</div>
        ${when ? `<div class="pmm-meta pmm-lead-meta">${when}</div>` : ''}
        ${people ? `<div class="pmm-meta pmm-lead-meta">${esc(people)}</div>` : ''}
        ${tel ? `<a class="pmm-card-call" href="${esc(tel)}" aria-label="${esc(name || '고객')}에게 전화">
          <i class="fas fa-phone" aria-hidden="true"></i></a>` : ''}
      </div>`;
  }

  findLead(no) {
    return this.rows.find((r) => val(r, '리드 No') === no)
      || (window.leadsApp?.data || []).find((r) => val(r, '리드 No') === no);
  }

  /** 펼침이 표 행 대신 시트 안 빈 행에 붙게 — 휴대폰일 때만. 인스턴스 메서드만 감싸 PC 동작은 그대로 */
  patchAccordion() {
    const acc = this.acc;
    if (!acc || acc.__pmmPatched) return;
    acc.__pmmPatched = true;
    const sheetRows = () => this.sheet.el.querySelectorAll('.accordion-row');

    const open = acc.openAccordion.bind(acc);
    acc.openAccordion = (tr, lead, opts = {}) => {
      if (this.isActive && lead) {
        sheetRows().forEach((r) => r.remove());
        this.showSheet(lead);
        return open(this.sheet.anchorRow, lead, { ...opts, animate: false });
      }
      return open(tr, lead, opts);
    };

    const close = acc.closeAccordion.bind(acc);
    acc.closeAccordion = (...args) => {
      const inSheet = this.sheet.isOpen;
      const r = close(...args);
      if (inSheet) {
        // 원래 닫기 정리는 애니메이션 끝 또는 1초 뒤 — 그 안에 다른 카드를 열면 새 상세의 'show'까지 지워짐
        //   (프로젝트 화면 11-15 점검과 같은 함정) → 지금 바로 정리
        acc.container.dispatchEvent(new Event('animationend'));
        sheetRows().forEach((row) => row.remove());
        this.sheet.hide();
      }
      return r;
    };

    // 필터·새로고침으로 표가 다시 그려질 때 펼침을 떼었다가(reopen) 다시 붙인다. 시트가 열려 있으면 떼지 않음
    //   (그 리드가 표의 지금 쪽에 없으면 reopen 이 못 찾아 시트가 빈 화면이 되므로)
    const detach = acc.detachQuietly.bind(acc);
    acc.detachQuietly = (...args) => {
      if (this.sheet.isOpen) return undefined;
      return detach(...args);
    };
  }

  openDetail(no, card = null) {
    const lead = this.findLead(no);
    if (!lead || !this.acc) return;
    this.lastNo = no;
    this.returnFocus = card;
    this.acc.openAccordion(this.sheet.anchorRow, lead);
  }

  showSheet(lead) {
    this.sheet.show({ title: val(lead, '리드 No'), quickHtml: this.quickHtml(lead), returnFocus: this.returnFocus });
    this.returnFocus = null;
  }

  /** 시트 위: 요약(상태·방문 예정일) + 바로 하기(고객 전화, 방문 주소 지도) */
  quickHtml(lead) {
    const status = val(lead, '상태');
    const name = val(lead, '고객명');
    const tel = telHref(lead['고객 연락처']);
    const addr = val(lead, '방문 주소');
    const rawVisit = String(lead['방문 예정일'] || '').replace(/^'/, '').trim();
    const { start } = splitVisitRange(rawVisit);
    const parts = [`
      <div class="pmm-sheet-sum">
        ${statusHtml(status)}
        <span class="pmm-sum-money">${esc(name || '이름 없음')}</span>
        ${start ? `<span class="pmm-sum-money">방문 <b>${esc(rawVisit)}</b> ${ddayTag(localIso(new Date()), start)}</span>` : ''}
      </div>`];
    if (tel) {
      parts.push(`<a class="pmm-quick-btn" href="${esc(tel)}">
        <i class="fas fa-phone" aria-hidden="true"></i>고객 전화</a>`);
    }
    if (addr) {
      const url = `https://map.kakao.com/link/search/${encodeURIComponent(addr)}`;
      parts.push(`<a class="pmm-quick-btn" href="${esc(url)}" target="_blank" rel="noopener">
        <i class="fas fa-map-marker-alt" aria-hidden="true"></i>지도에서 보기</a>`);
    }
    return parts.join('');
  }

  /** 닫기(뒤로 단추·휴대폰 뒤로 가기). 리드 상세는 보기 전용이라 확인 없이 닫음 */
  requestClose(fromPop = false) {
    this.sheet.closingFromPop = fromPop;
    if (this.acc?.isOpen) this.acc.closeAccordion();
    if (this.sheet.isOpen) {
      this.sheet.el.querySelectorAll('.accordion-row').forEach((r) => r.remove());
      this.sheet.hide();
    }
    this.sheet.closingFromPop = false;
  }
}
