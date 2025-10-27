/**
 * 고객 리드 관리 페이지 메인 JavaScript
 * DataTables, 필터, CRUD, FullCalendar 통합
 */

import DataTable from 'datatables.net';
import 'datatables.net-bs5';
import 'datatables.net-bs5/css/dataTables.bootstrap5.min.css';
import '../../css/pages/leads.css';
import logger from '../utils/logger.js';
import ModernLeadsFilters from '../components/ModernLeadsFilters.js';

// FullCalendar imports
import { Calendar } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import koLocale from '@fullcalendar/core/locales/ko';

// 전역 변수
let leadsTable = null;
let calendar = null;
let leadsData = [];
let currentView = 'table'; // 'table' or 'calendar'
let leadsFilters = null; // 필터 컴포넌트

/**
 * 페이지 초기화 함수
 */
async function initLeadsPage() {
    logger.debug('[LEADS] 리드 관리 페이지 초기화 시작');

    try {
        // 리드 데이터 로드
        await loadLeadsData();

        // DataTables 초기화
        initializeDataTable();

        // FullCalendar 초기화 (나중에 표시)
        initializeCalendar();

        // 필터 컴포넌트 초기화 (프로젝트 관리 패턴)
        leadsFilters = new ModernLeadsFilters();
        await leadsFilters.init();

        // 필터 콜백 등록 (테이블 업데이트)
        leadsFilters.onFilterChange((filteredData) => {
            if (leadsTable) {
                leadsTable.clear().rows.add(filteredData).draw();
                logger.debug(`[LEADS] 테이블 업데이트: ${filteredData.length}개 행`);
            }
        });

        // 필터에 데이터 전달 (필터 옵션 생성 + 초기 필터 적용)
        leadsFilters.applyFilters(leadsData, true);

        // 새 리드 등록 모달 드롭다운 초기화
        populateModalDropdowns();

        logger.debug('[LEADS] 리드 관리 페이지 초기화 완료');

    } catch (error) {
        logger.error('[LEADS] 페이지 초기화 실패:', error);
        showAlert('페이지 로드 중 오류가 발생했습니다', 'error');
    }
}

/**
 * 페이지 초기화 - DOM 준비 상태 체크
 * 동적 스크립트 로드 시에도 작동하도록 readyState 체크
 */
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initLeadsPage);
} else {
    // DOM이 이미 준비된 경우 즉시 실행
    initLeadsPage();
}

/**
 * 리드 데이터 로드
 */
async function loadLeadsData(forceRefresh = false) {
    try {
        const url = forceRefresh
            ? '/leads/api/list?force_refresh=true'
            : '/leads/api/list';

        logger.debug(`[LEADS] API 요청 시작: ${url}`);

        const response = await fetch(url, {
            credentials: 'same-origin',
            headers: {
                'Accept': 'application/json',
                'X-Requested-With': 'XMLHttpRequest'
            }
        });

        logger.debug(`[LEADS] API 응답 상태: ${response.status} ${response.statusText}`);

        if (response.status === 401 || response.status === 440) {
            logger.warn('[LEADS] 세션 만료 응답 수신 - 로그인 페이지로 리다이렉트');
            leadsData = [];
            const redirectUrl = new URL('/login', window.location.origin);
            redirectUrl.searchParams.set('session_expired', 'true');
            window.location.href = redirectUrl.toString();
            return leadsData;
        }

        // Content-Type 체크 (세션 만료로 HTML 로그인 페이지 반환 시 감지)
        const contentType = response.headers.get('content-type');
        if (!contentType || !contentType.includes('application/json')) {
            logger.warn('[LEADS] 세션 만료 감지 - 로그인 페이지로 리다이렉트');
            leadsData = [];
            const redirectUrl = new URL('/login', window.location.origin);
            redirectUrl.searchParams.set('session_expired', 'true');
            window.location.href = redirectUrl.toString();
            return leadsData;
        }

        if (!response.ok) {
            const errorText = await response.text();
            logger.error(`[LEADS] HTTP 오류 응답: ${errorText.substring(0, 200)}`);
            throw new Error(`HTTP ${response.status}`);
        }

        const result = await response.json();
        logger.debug(`[LEADS] JSON 파싱 완료. success: ${result.success}, data 존재: ${!!result.data}`);

        if (!result.success) {
            throw new Error(result.message || '리드 데이터 로드 실패');
        }

        leadsData = result.data?.leads || [];
        logger.debug(`[LEADS] 리드 데이터 로드 완료: ${leadsData.length}건`);

        return leadsData;

    } catch (error) {
        logger.error('[LEADS] 리드 데이터 로드 실패:', error);
        throw error;
    }
}

/**
 * DataTables 초기화
 */
function initializeDataTable() {
    if (leadsTable) {
        leadsTable.destroy();
    }

    leadsTable = new DataTable('#leadsTable', {
        data: leadsData,
        columns: [
            { data: '리드 No', width: '8%' },
            { data: '거래처', width: '8%' },
            {
                data: '상태',
                width: '8%',
                render: (data) => {
                    const statusColors = {
                        '유선 상담': 'secondary',
                        '방문 예약': 'primary',
                        '견적 제출': 'info',
                        '방문 취소': 'warning',
                        '공사 확정': 'success',
                        '공사 드랍': 'danger'
                    };
                    const color = statusColors[data] || 'secondary';
                    return `<span class="badge bg-${color}">${data || '-'}</span>`;
                }
            },
            { data: '방문 예정일', width: '10%' },
            { data: '담당자', width: '8%' },
            { data: '연락처', width: '12%' },
            { data: '고객명', width: '10%' },
            { data: '방문 주소', width: '18%' },
            // 상담 내용은 테이블에서 제외 (편집 모달에서만 표시)
            { data: '마지막 연락일', width: '10%', defaultContent: '-' },
            {
                data: null,
                width: '8%',
                orderable: false,
                render: (data, type, row) => {
                    return `
                        <div class="btn-group btn-group-sm">
                            <button class="btn btn-outline-primary" onclick="openEditLeadModal('${row['리드 No']}')" title="편집">
                                <i class="fas fa-edit"></i>
                            </button>
                        </div>
                    `;
                }
            }
        ],
        order: [[3, 'desc']], // 방문 예정일 내림차순
        pageLength: 15,
        lengthMenu: [[15, 25, 50, 100], ['15', '25', '50', '100']],
        pagingType: 'full_numbers', // 페이지 숫자 표시 (처음/이전/1/2/3/다음/마지막)
        searching: false, // 테이블 내장 검색 비활성화 (상단 필터 사용)
        dom: '<"top"l>rt<"bottom d-flex justify-content-between"<"info-left"i><"paging-right"p>><"clear">',
        language: {
            lengthMenu: "_MENU_개씩 보기",
            info: "_START_~_END_ / 전체 _TOTAL_개 (페이지 _PAGE_ / _PAGES_)",
            infoEmpty: "0개",
            infoFiltered: "",
            paginate: {
                first: "처음",
                last: "마지막",
                next: "다음",
                previous: "이전"
            },
            emptyTable: "데이터가 없습니다"
        },
        responsive: true,
        drawCallback: function() {
            // 테이블 렌더링 완료 후 표시
            $('#leadsTable').css('visibility', 'visible');

            // 필터 결과 카운트 초기화 (프로젝트 필터 스타일과 동일)
            const filterResultCount = document.getElementById('filterResultCount');
            if (filterResultCount) {
                const visibleRows = leadsTable.rows({ filter: 'applied' }).nodes().length;
                filterResultCount.textContent = `${visibleRows}개 리드 표시`;
            }
        }
    });

    logger.debug('[LEADS] DataTables 초기화 완료');
}

/**
 * FullCalendar 초기화
 */
function initializeCalendar() {
    const calendarEl = document.getElementById('calendar');

    if (!calendarEl) {
        logger.warn('[LEADS] 캘린더 요소를 찾을 수 없습니다');
        return;
    }

    calendar = new Calendar(calendarEl, {
        plugins: [dayGridPlugin, timeGridPlugin, interactionPlugin],
        initialView: 'dayGridMonth',
        locale: koLocale,
        headerToolbar: {
            left: 'prev,next today',
            center: 'title',
            right: 'dayGridMonth,timeGridWeek,timeGridDay'
        },
        editable: true,
        droppable: true,
        eventDrop: handleEventDrop,
        eventClick: handleEventClick,
        events: []
    });

    logger.debug('[LEADS] FullCalendar 초기화 완료');
}

/**
 * 캘린더 이벤트 로드
 */
function loadCalendarEvents() {
    const events = leadsData
        .filter(lead => lead['방문 예정일'])
        .map(lead => {
            const statusColors = {
                '유선 상담': '#6c757d',
                '방문 예약': '#0d6efd',
                '견적 제출': '#0dcaf0',
                '방문 취소': '#ffc107',
                '공사 확정': '#198754',
                '공사 드랍': '#dc3545'
            };

            return {
                id: lead['리드 No'],
                title: `${lead['고객명']} - ${lead['상태']}`,
                start: lead['방문 예정일'],
                backgroundColor: statusColors[lead['상태']] || '#6c757d',
                extendedProps: {
                    leadNo: lead['리드 No'],
                    customerName: lead['고객명'],
                    phone: lead['연락처'],
                    address: lead['방문 주소'],
                    status: lead['상태']
                }
            };
        });

    calendar.removeAllEvents();
    calendar.addEventSource(events);

    logger.debug(`[LEADS] 캘린더 이벤트 로드 완료: ${events.length}건`);
}

/**
 * 캘린더 이벤트 드래그 핸들러
 */
async function handleEventDrop(info) {
    const leadNo = info.event.id;
    // startStr 사용: FullCalendar가 이미 로컬 날짜로 계산해둔 값 (UTC 변환 문제 방지)
    const newDate = info.event.startStr;

    try {
        const response = await fetch(`/leads/api/update/${leadNo}`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            credentials: 'same-origin',
            body: JSON.stringify({
                '방문 예정일': newDate
            })
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const result = await response.json();

        if (!result.success) {
            throw new Error(result.message);
        }

        logger.debug(`[LEADS] 방문 예정일 변경 완료: ${leadNo} → ${newDate}`);
        showAlert('방문 예정일이 변경되었습니다', 'success');

        // 데이터 새로고침
        await loadLeadsData(true);

    } catch (error) {
        logger.error('[LEADS] 방문 예정일 변경 실패:', error);
        showAlert('방문 예정일 변경 중 오류가 발생했습니다', 'error');
        info.revert();
    }
}

/**
 * 캘린더 이벤트 클릭 핸들러
 */
function handleEventClick(info) {
    const leadNo = info.event.id;
    openEditLeadModal(leadNo);
}

/**
 * 뷰 전환: 테이블
 */
function switchToTableView() {
    currentView = 'table';

    document.getElementById('tableView').style.display = 'block';
    document.getElementById('calendarView').style.display = 'none';
    document.getElementById('filterSection').style.display = 'block';

    document.getElementById('tableViewBtn').classList.add('active');
    document.getElementById('calendarViewBtn').classList.remove('active');

    logger.debug('[LEADS] 테이블 뷰로 전환');
}

/**
 * 뷰 전환: 캘린더
 */
function switchToCalendarView() {
    currentView = 'calendar';

    document.getElementById('tableView').style.display = 'none';
    document.getElementById('calendarView').style.display = 'block';
    document.getElementById('filterSection').style.display = 'none';

    document.getElementById('tableViewBtn').classList.remove('active');
    document.getElementById('calendarViewBtn').classList.add('active');

    // 캘린더 렌더링
    calendar.render();
    loadCalendarEvents();

    logger.debug('[LEADS] 캘린더 뷰로 전환');
}

/**
 * 모달 드롭다운 채우기
 */
function populateModalDropdowns() {
    // 거래처 & 담당자 목록은 프로젝트 관리와 동일하게 가져오기
    // 임시로 하드코딩 (실제로는 API에서 가져와야 함)
    const companies = ['글로벌', '서울지점', '경기지점', '인천지점'];
    const owners = ['김철수', '이영희', '박민수', '정수진'];

    // 새 리드 모달
    const newLeadCompany = document.getElementById('newLeadCompany');
    const newLeadOwner = document.getElementById('newLeadOwner');

    if (newLeadCompany) {
        companies.forEach(company => {
            const option = document.createElement('option');
            option.value = company;
            option.textContent = company;
            newLeadCompany.appendChild(option.cloneNode(true));
        });
    }

    if (newLeadOwner) {
        owners.forEach(owner => {
            const option = document.createElement('option');
            option.value = owner;
            option.textContent = owner;
            newLeadOwner.appendChild(option.cloneNode(true));
        });
    }

    // 편집 모달
    const editLeadCompany = document.getElementById('editLeadCompany');
    const editLeadOwner = document.getElementById('editLeadOwner');

    if (editLeadCompany) {
        companies.forEach(company => {
            const option = document.createElement('option');
            option.value = company;
            option.textContent = company;
            editLeadCompany.appendChild(option.cloneNode(true));
        });
    }

    if (editLeadOwner) {
        owners.forEach(owner => {
            const option = document.createElement('option');
            option.value = owner;
            option.textContent = owner;
            editLeadOwner.appendChild(option.cloneNode(true));
        });
    }

    logger.debug('[LEADS] 모달 드롭다운 초기화 완료');
}

/**
 * 새 리드 등록 모달 열기
 */
function openNewLeadModal() {
    // 폼 초기화
    document.getElementById('newLeadForm').reset();

    // 모달 표시
    const modal = new bootstrap.Modal(document.getElementById('newLeadModal'));
    modal.show();

    logger.debug('[LEADS] 새 리드 등록 모달 열기');
}

/**
 * 새 리드 저장
 */
async function saveNewLead() {
    try {
        // 폼 데이터 수집
        const leadData = {
            '거래처': document.getElementById('newLeadCompany').value,
            '담당자': document.getElementById('newLeadOwner').value,
            '고객명': document.getElementById('newLeadCustomerName').value,
            '연락처': document.getElementById('newLeadPhone').value,
            '방문 주소': document.getElementById('newLeadAddress').value,
            '방문 예정일': document.getElementById('newLeadVisitDate').value,
            '마지막 연락일': document.getElementById('newLeadLastContact').value,
            '상담 내용': document.getElementById('newLeadContent').value,
            '비고': document.getElementById('newLeadNote').value
        };

        // 필수 필드 검증
        if (!leadData['거래처'] || !leadData['고객명'] || !leadData['연락처']) {
            showAlert('필수 항목을 모두 입력해주세요', 'warning');
            return;
        }

        // API 호출
        const response = await fetch('/leads/api/create', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            credentials: 'same-origin',
            body: JSON.stringify(leadData)
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const result = await response.json();

        if (!result.success) {
            throw new Error(result.message);
        }

        logger.debug(`[LEADS] 새 리드 생성 완료: ${result.lead_no}`);
        showAlert(`리드 ${result.lead_no}가 생성되었습니다`, 'success');

        // 모달 닫기
        bootstrap.Modal.getInstance(document.getElementById('newLeadModal')).hide();

        // 데이터 새로고침
        await loadLeadsData(true);
        leadsTable.clear().rows.add(leadsData).draw();

        if (currentView === 'calendar') {
            loadCalendarEvents();
        }

    } catch (error) {
        logger.error('[LEADS] 새 리드 생성 실패:', error);
        showAlert('리드 생성 중 오류가 발생했습니다', 'error');
    }
}

/**
 * 리드 편집 모달 열기
 */
function openEditLeadModal(leadNo) {
    const lead = leadsData.find(l => l['리드 No'] === leadNo);

    if (!lead) {
        showAlert('리드를 찾을 수 없습니다', 'error');
        return;
    }

    // 폼 채우기
    document.getElementById('editLeadNo').value = lead['리드 No'];
    document.getElementById('editLeadNoDisplay').value = lead['리드 No'];
    document.getElementById('editLeadStatus').value = lead['상태'] || '유선 상담';
    document.getElementById('editLeadCompany').value = lead['거래처'] || '';
    document.getElementById('editLeadOwner').value = lead['담당자'] || '';
    document.getElementById('editLeadCustomerName').value = lead['고객명'] || '';
    document.getElementById('editLeadPhone').value = lead['연락처'] || '';
    document.getElementById('editLeadAddress').value = lead['방문 주소'] || '';
    document.getElementById('editLeadVisitDate').value = lead['방문 예정일'] || '';
    document.getElementById('editLeadLastContact').value = lead['마지막 연락일'] || '';
    document.getElementById('editLeadContent').value = lead['상담 내용'] || '';
    document.getElementById('editLeadNote').value = lead['비고'] || '';

    // 모달 표시
    const modal = new bootstrap.Modal(document.getElementById('editLeadModal'));
    modal.show();

    logger.debug(`[LEADS] 리드 편집 모달 열기: ${leadNo}`);
}

/**
 * 리드 편집 저장
 */
async function saveEditLead() {
    try {
        const leadNo = document.getElementById('editLeadNo').value;

        // 수정할 데이터
        const updateData = {
            '상태': document.getElementById('editLeadStatus').value,
            '거래처': document.getElementById('editLeadCompany').value,
            '담당자': document.getElementById('editLeadOwner').value,
            '고객명': document.getElementById('editLeadCustomerName').value,
            '연락처': document.getElementById('editLeadPhone').value,
            '방문 주소': document.getElementById('editLeadAddress').value,
            '방문 예정일': document.getElementById('editLeadVisitDate').value,
            '마지막 연락일': document.getElementById('editLeadLastContact').value,
            '상담 내용': document.getElementById('editLeadContent').value,
            '비고': document.getElementById('editLeadNote').value
        };

        // API 호출
        const response = await fetch(`/leads/api/update/${leadNo}`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            credentials: 'same-origin',
            body: JSON.stringify(updateData)
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const result = await response.json();

        if (!result.success) {
            throw new Error(result.message);
        }

        logger.debug(`[LEADS] 리드 수정 완료: ${leadNo}`);
        showAlert('리드가 수정되었습니다', 'success');

        // 모달 닫기
        bootstrap.Modal.getInstance(document.getElementById('editLeadModal')).hide();

        // 데이터 새로고침
        await loadLeadsData(true);
        leadsTable.clear().rows.add(leadsData).draw();

        if (currentView === 'calendar') {
            loadCalendarEvents();
        }

    } catch (error) {
        logger.error('[LEADS] 리드 수정 실패:', error);
        showAlert('리드 수정 중 오류가 발생했습니다', 'error');
    }
}

/**
 * 알림 메시지 표시
 */
function showAlert(message, type = 'info') {
    const alertContainer = document.getElementById('headerAlertContainer');

    const alertTypes = {
        'success': 'alert-success',
        'error': 'alert-danger',
        'warning': 'alert-warning',
        'info': 'alert-info'
    };

    const alertClass = alertTypes[type] || 'alert-info';

    alertContainer.innerHTML = `
        <div class="alert ${alertClass} alert-dismissible fade show mb-0 py-1 px-3" role="alert" style="font-size: 0.9rem;">
            ${message}
            <button type="button" class="btn-close btn-close-sm" data-bs-dismiss="alert"></button>
        </div>
    `;

    // 3초 후 자동 제거
    setTimeout(() => {
        alertContainer.innerHTML = '';
    }, 3000);
}

/**
 * 이벤트 리스너 등록 (DOM 로드 후)
 * 참고: 필터 관련 이벤트는 ModernLeadsFilters 컴포넌트에서 처리
 */
function setupEventListeners() {
    // 뷰 전환 버튼
    const tableViewBtn = document.getElementById('tableViewBtn');
    const calendarViewBtn = document.getElementById('calendarViewBtn');
    if (tableViewBtn) tableViewBtn.addEventListener('click', switchToTableView);
    if (calendarViewBtn) calendarViewBtn.addEventListener('click', switchToCalendarView);

    // 새 리드 등록 버튼
    const newLeadBtn = document.getElementById('newLeadBtn');
    if (newLeadBtn) newLeadBtn.addEventListener('click', openNewLeadModal);

    // 새 리드 저장 버튼
    const saveNewLeadBtn = document.getElementById('saveNewLeadBtn');
    if (saveNewLeadBtn) saveNewLeadBtn.addEventListener('click', saveNewLead);

    // 리드 편집 저장 버튼
    const saveEditLeadBtn = document.getElementById('saveEditLeadBtn');
    if (saveEditLeadBtn) saveEditLeadBtn.addEventListener('click', saveEditLead);

    logger.debug('[LEADS] 이벤트 리스너 등록 완료');
}

// 페이지 초기화 시 이벤트 리스너 등록
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        // DOM 로드 후 이벤트 리스너 등록은 initLeadsPage가 완료된 후에
        setTimeout(setupEventListeners, 100);
    });
} else {
    // DOM이 이미 로드된 경우
    setTimeout(setupEventListeners, 100);
}

export default {
    loadLeadsData,
    openEditLeadModal,
    switchToTableView,
    switchToCalendarView
};

// inline onclick 핸들러를 위해 전역으로 노출
window.openEditLeadModal = openEditLeadModal;
