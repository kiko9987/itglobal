/**
 * DataTables 모듈 가져오기
 */
import DataTable from 'datatables.net';
import 'datatables.net-bs5';

// DataTables CSS 스타일 import (Bootstrap 5 테마)
import 'datatables.net-bs5/css/dataTables.bootstrap5.css';

// 🆕 ModeManager import
import { getGlobalModeManager } from '../utils/globalModeManager.js';
import { TABLE_MODE, ACCORDION_MODE } from '../constants/ViewModes.js';

import logger from '../utils/logger.js';

/** A/S 컬럼 렌더용 최소 HTML 이스케이프 */
function _asEsc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * 날짜 포맷팅 헬퍼 함수
 */
function formatDate(data) {
  if (!data) return '';

  try {
    // 다양한 날짜 형식 처리
    let date;
    if (typeof data === 'string') {
      // "2025-09-12" 형식이면 그대로 반환
      if (/^\d{4}-\d{2}-\d{2}$/.test(data)) {
        return data;
      }
      date = new Date(data);
    } else if (data instanceof Date) {
      date = data;
    } else {
      return '';
    }

    if (isNaN(date.getTime())) return '';

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');

    return `${year}-${month}-${day}`;
  } catch (error) {
    logger.warn('날짜 포맷팅 오류:', error);
    return '';
  }
}

/**
 * HTML 특수문자 이스케이프 (XSS 방지)
 */
function escapeHTML(str) {
  if (!str) return '';
  return str.replace(/[&<>"']/g, function(match) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[match];
  });
}

/**
 * 금액 필드 렌더링 헬퍼 (메모 상태 + 계산서 아이콘 포함)
 * @param {number|string} data - 금액 데이터
 * @param {Object} row - 행 데이터
 * @param {string} memoFieldName - 메모 필드명 (예: '계약금_메모')
 * @param {string} stage - 단계명 (예: '계약금', '중도금', '잔금')
 * @returns {string} 렌더링된 HTML
 */
function renderPaymentFieldWithMemo(data, row, memoFieldName, stage) {
  const amount = parseFloat(data || 0);
  const memo = row[memoFieldName] || '';

  const amountText = amount > 0 ?
    `${amount.toLocaleString()}원` :
    '-';

  // 계산서 정보 파싱
  const billStatus = parseBillStatusForTable(row['계산서']);
  const billIcon = billStatus.stages[stage] ? getBillStatusIconForTable(billStatus.stages[stage]) : '';

  // 3단계 시각화: 금액 없음 / 메모 없음(경고) / 메모 있음(정상)
  if (amount > 0) {
    const memoStatus = checkMemoStatus(memo);

    if (memoStatus.isEmpty) {
      // 금액 있음 + 메모 없음 → 빈 아이콘 + 초록색 (금액이 있으니 중요)
      const memoIcon = `
        <span class="memo-value-wrapper" data-bs-toggle="tooltip" data-bs-title="메모를 작성해주세요" aria-label="메모를 작성해주세요">
          <span class="memo-tooltip-trigger no-memo">
            <i class="far fa-sticky-note fa-lg text-success"></i>
          </span>
        </span>`;
      return `<div class="payment-field-container">
        <span class="payment-amount">${amountText}</span>
        <span class="payment-icons">${memoIcon}${wrapBillIcon(billIcon, stage)}</span>
      </div>`;
    } else {
      // 금액 있음 + 메모 있음 → 채워진 아이콘 + 초록색 (아이콘에만 툴팁)
      const memoIcon = `
        <span class="memo-value-wrapper" data-bs-toggle="tooltip" data-bs-title="${escapeHTML(memo)}" aria-label="메모 보기">
          <span class="memo-tooltip-trigger has-memo">
            <i class="fas fa-sticky-note fa-lg text-success"></i>
          </span>
        </span>`;
      return `<div class="payment-field-container">
        <span class="payment-amount">${amountText}</span>
        <span class="payment-icons">${memoIcon}${wrapBillIcon(billIcon, stage)}</span>
      </div>`;
    }
  }

  return '<span class="text-muted">-</span>';
}

function wrapBillIcon(billIcon, stage) {
  if (!billIcon) return '';

  const titleMatch = billIcon.match(/title="([^"]*)"/);
  const tooltipText = titleMatch ? titleMatch[1] : `${stage} 단계 계산서`;
  const sanitizedIcon = billIcon
    .replace(/title="([^"]*)"/, '')
    .replace(/ms-1/g, '')
    .replace(/\s{2,}/g, ' ');

  return `
    <span class="memo-value-wrapper" data-bs-toggle="tooltip" data-bs-title="${escapeHTML(tooltipText)}" aria-label="${escapeHTML(tooltipText)}">
      <span class="memo-tooltip-trigger bill-status">
        ${sanitizedIcon}
      </span>
    </span>`;
}

/**
 * 메모 상태 확인 (빈 메모 vs 실제 메모)
 * @param {string} memo - 메모 내용
 * @returns {object} { isEmpty: boolean, isTemplate: boolean }
 */
function checkMemoStatus(memo) {
  if (!memo || !memo.trim()) {
    return { isEmpty: true, isTemplate: false };
  }

  // 자동 생성 템플릿 체크
  const template = '입금일: \n입금자:';
  const trimmedMemo = memo.trim();

  if (trimmedMemo === template || trimmedMemo === '입금일:\n입금자:') {
    return { isEmpty: true, isTemplate: true };
  }

  return { isEmpty: false, isTemplate: false };
}

/**
 * 계산서 정보 파싱 (테이블용)
 * @param {string} billValue - 계산서 값 (예: "일반-계약금, N입금-중도금")
 * @returns {object} { stages: { '계약금': '일반', '중도금': 'N입금', ... } }
 */
function parseBillStatusForTable(billValue) {
  const result = {
    stages: {}  // { '계약금': '일반', '중도금': 'N입금', ... }
  };

  if (!billValue || billValue === '-' || billValue === '미발행') {
    return result;
  }

  // Split by comma and parse each item
  const items = billValue.split(',').map(s => s.trim());

  items.forEach(item => {
    const parts = item.split('-');
    if (parts.length === 2) {
      const category = parts[0].trim();
      const stage = parts[1].trim();
      result.stages[stage] = category;
    }
  });

  return result;
}

/**
 * 계산서 카테고리에 대한 아이콘 반환 (테이블용)
 * @param {string} category - 카테고리명 (일반, N입금, 카드)
 * @returns {string} 아이콘 HTML
 */
function getBillStatusIconForTable(category) {
  const iconMap = {
    '일반': '<i class="fas fa-receipt fa-lg ms-1 text-primary" title="세금계산서 발행"></i>',
    'N입금': '<i class="fas fa-sack-dollar fa-lg ms-1 text-danger" title="현금거래 (계산서 미발행)"></i>',
    '카드': '<i class="fas fa-credit-card fa-lg ms-1 text-info" title="카드결제"></i>'
  };

  return iconMap[category] || '';
}

/**
 * Project Table 컴포넌트
 * DataTables를 사용한 프로젝트 목록 테이블
 */
export default class ProjectTable {
  constructor() {
    this.table = null;
    this.tableElement = null;
    this.accordion = null;
    this.statusBadge = null;
    this.performanceOptimizer = null;
    this.isLargeDataset = false;
    this.largeDatasetThreshold = 500; // 500개 이상이면 대용량으로 간주
    this.pendingBatchTimeout = null;
    this.currentBatchToken = 0;

    // 🆕 ModeManager 초기화
    this.modeManager = getGlobalModeManager();

    this.isReceivablesTransitioning = false; // 수금 모드 전환 중 플래그
    this.skeletonTimeout = null; // 스켈레톤 타임아웃
    this.skeletonOverlay = null; // 스켈레톤 오버레이 요소
    this.originalWrapperPosition = null; // wrapper position 복원용

    // 🆕 테이블 모드 변경 이벤트 리스너 (메모리 누수 방지: bound handler 저장)
    this.boundHandleTableModeChange = this.handleTableModeChange.bind(this);
    this.modeManager.on('tableModeChanged', this.boundHandleTableModeChange);
  }

  /**
   * 🆕 테이블 모드 변경 이벤트 핸들러
   * @param {object} data - { oldMode, newMode, currentCombination }
   */
  handleTableModeChange(data) {
    logger.debug(`[ProjectTable] 테이블 모드 변경 감지: ${data.oldMode} → ${data.newMode}`);

    // 컬럼 visibility 업데이트는 setupReceivablesToggle에서 처리
    // A/S 모드 이탈 시(수금 토글 등으로) A/S 컬럼/상태 정리 — 상호배타 (2026-08-31)
    if (data.newMode !== TABLE_MODE.AS && this._asModeActive) {
      this._asModeActive = false;
      const t = document.getElementById('asModeToggle');
      if (t) { t.classList.remove('active'); t.setAttribute('aria-pressed', 'false'); }
      try { sessionStorage.removeItem('itg_as_mode'); } catch (_) { /* noop */ }
      if (this.table) { this._applyAsColumns(false); this.table.columns.adjust(); }
    }
  }

  /**
   * 전역 툴팁 초기화 함수
   * Bootstrap 5 표준 방식: data-bs-toggle="tooltip" 속성 사용
   * 레거시 호환: title 속성도 지원 (데이터 완성도, 주소, 공사내용 등)
   */
  initializeTooltips(container = null) {
    const targetContainer = container || this.tableElement || document;

    // 1. data-bs-toggle="tooltip" 속성 초기화 (메모 툴팁용)
    const existingBootstrapTooltips = targetContainer.querySelectorAll('[data-bs-toggle="tooltip"]');
    existingBootstrapTooltips.forEach(el => {
      const instance = window.bootstrap?.Tooltip.getInstance(el);
      if (instance) {
        instance.dispose();
      }
    });

    const bootstrapTooltipElements = targetContainer.querySelectorAll('[data-bs-toggle="tooltip"]');
    bootstrapTooltipElements.forEach(el => {
      if (window.bootstrap && window.bootstrap.Tooltip) {
        new window.bootstrap.Tooltip(el, {
          container: 'body',  // 툴팁이 테이블 바깥에 표시되도록 함
          trigger: 'hover',
          placement: 'top',
          html: false,
          delay: { show: 200, hide: 100 }  // show 지연: 200ms, hide 지연: 100ms
        });
      }
    });

    // 2. 일반 title 속성 초기화 (레거시 호환: 데이터 완성도, 주소, 공사내용 등)
    const existingTitleTooltips = targetContainer.querySelectorAll('[title]:not([data-bs-toggle="tooltip"])');
    existingTitleTooltips.forEach(el => {
      if (el._tooltip) {
        el._tooltip.dispose();
        el._tooltip = null;
      }
    });

    const titleTooltipElements = targetContainer.querySelectorAll('[title]:not([data-bs-toggle="tooltip"])');
    titleTooltipElements.forEach(el => {
      if (window.bootstrap && window.bootstrap.Tooltip) {
        el._tooltip = new window.bootstrap.Tooltip(el, {
          trigger: 'hover',
          placement: 'top',
          html: false,
          delay: { show: 200, hide: 100 }  // show 지연: 200ms, hide 지연: 100ms
        });
      }
    });

    const totalCount = bootstrapTooltipElements.length + titleTooltipElements.length;
    if (totalCount > 0) {
      logger.debug(`[ProjectTable] 툴팁 ${totalCount}개 초기화 완료 (Bootstrap: ${bootstrapTooltipElements.length}, Title: ${titleTooltipElements.length})`);
    }
  }

  /**
   * 테이블 초기화
   */
  async init() {
    this.tableElement = document.getElementById('projectsTable'); // 올바른 ID 사용
    if (!this.tableElement) {
      logger.error('프로젝트 테이블 엘리먼트를 찾을 수 없습니다.');
      return;
    }

    // CRITICAL FIX: Force remove any existing contain styles on initialization
    this.removeContainStyles();

    await this.initializeBadgeSystem();
    await this.initializePerformanceOptimizer();
    await this.initializeDataTable(); // async/await 적용
    await this.initializeAccordion();
    this.setupReceivablesToggle();
    this.setupAsToggle();
  }

  /**
   * UnifiedBadgeSystem 초기화 (모듈형 뱃지 시스템)
   */
  async initializeBadgeSystem() {
    try {
      const { default: UnifiedBadgeSystem } = await import('./UnifiedBadgeSystem.js');
      this.badgeSystem = new UnifiedBadgeSystem();
      // UnifiedBadgeSystem 초기화 완료
    } catch (error) {
      logger.error('[ERROR] UnifiedBadgeSystem 초기화 실패:', error);
    }
  }


  /**
   * PerformanceOptimizer 초기화
   */
  async initializePerformanceOptimizer() {
    try {
      const { default: PerformanceOptimizer } = await import('./PerformanceOptimizer.js');
      this.performanceOptimizer = new PerformanceOptimizer({
        enableVirtualScrolling: true,
        enableMemoization: true,
        enableLazyLoading: true,
        virtualScrollItemHeight: 60,
        memoizationCacheSize: 200,
        performanceMonitoring: true
      });
      // PerformanceOptimizer 초기화 완료
    } catch (error) {
      logger.error('[ERROR] PerformanceOptimizer 초기화 실패:', error);
    }
  }

  /**
   * 아코디언 초기화
   */
  async initializeAccordion() {
    // Accordion 초기화

    // 이미 초기화되어 있으면 스킵
    if (this.accordion) {
      logger.warn('[ProjectTable] 아코디언이 이미 초기화되어 있습니다. 중복 초기화 방지');
      return;
    }

    try {
      // ProjectRowAccordion 컴포넌트 동적 로드
      const { default: ProjectRowAccordion } = await import('./ProjectRowAccordion.js');

      this.accordion = new ProjectRowAccordion();
      this.accordion.init();
      this.accordion.attachToTable(this.tableElement, this.table);

      // 전역 접근을 위해 윈도우 객체에 등록
      window.projectRowAccordion = this.accordion;


    } catch (error) {
      logger.error('아코디언 초기화 실패:', error);
    }
  }


  /**
   * DataTable 초기화
   */
  async initializeDataTable() {
    // DataTables 초기화 시작

    // 프로젝트 코드 정렬을 위한 커스텀 타입 등록 (숫자만으로 정렬)
    DataTable.ext.type.order['project-code-pre'] = function(data) {
      // HTML 태그가 있는 경우 텍스트만 추출
      const div = document.createElement('div');
      const safeData = (data ?? '').toString();
      div.innerHTML = safeData;
      const textOnly = div.textContent || div.innerText || safeData;

      // 프로젝트 코드에서 숫자만 추출 (G, P, R 접두사 무시하고 숫자만 비교)
      const match = textOnly.match(/([GPR])(\d+)/);
      if (match) {
        const number = parseInt(match[2], 10);
        return number; // 숫자만으로 정렬 (접두사 구분 없음)
      }
      return 9999999; // 매치되지 않는 경우 뒤쪽으로
    };

    // 프로젝트 코드 타입 감지
    DataTable.ext.type.detect.unshift(function (data) {
      const div = document.createElement('div');
      const safeData = (data ?? '').toString();
      div.innerHTML = safeData;
      const textOnly = div.textContent || div.innerText || safeData;
      return /[GPR]\d+/.test(textOnly) ? 'project-code' : null;
    });

    // 기존 DataTable이 있으면 제거
    if (DataTable.isDataTable(this.tableElement)) {
      new DataTable(this.tableElement).destroy();
    }

    this.table = new DataTable(this.tableElement, {
      responsive: false, // 아코디언과 충돌 방지
      autoWidth: false, // 자동 너비 계산 비활성화
      pageLength: 15,
      lengthMenu: [[15, 25, 50, 100], ['15', '25', '50', '100']],
      order: [[0, 'desc']], // 프로젝트 코드 내림차순 정렬 (기본)
      ordering: true, // 정렬 기능 활성화
      orderMulti: false, // 여러 열을 동시에 정렬하지 않음
      searching: false, // 테이블 내장 검색 기능 비활성화 (상단 필터 사용)
      // 2026-07-08 stateSave 비활성화. 매니저가 이전 세션에서 실수로 다른 컬럼을 클릭해
      // 정렬 상태가 유지되면 새로고침 후 뒤죽박됨 순서로 보이고, 필터도 안 켜져 있는데
      // 왜 이 순서인지 판단이 어려운 UX 사고 유발. 매 방문 시 기본 정렬(코드 내림차순)로
      // 시작하는 것이 매니저 20명 운영에 안전.
      stateSave: false,
      stateDuration: -1,

      // DOM 구조 설정 - 겹침 문제 근본 해결, 한 줄에 나란히 배치
      dom: '<"top"l>rt<"bottom d-flex justify-content-between"<"info-left"i><"paging-right"p>><"clear">',
      language: {
        lengthMenu: "_MENU_개씩 보기", // 레거시 스타일과 동일
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
      columnDefs: [
        {
          targets: '_all',
          orderSequence: ['asc', 'desc'] // 미정렬 상태 제거
        }
      ],
      columns: [
        {
          name: 'projectCode',
          data: '프로젝트 코드',
          type: 'project-code',
          defaultContent: '',
          render: function(data) {
            if (this.badgeSystem) {
              return this.badgeSystem.createProjectBadge(data);
            }
            return data || '';
          }.bind(this)
        },
        {
          name: 'manager',
          data: '담당자',
          defaultContent: '',
          render: function(data) {
            if (this.badgeSystem) {
              return this.badgeSystem.createManagerBadge(data);
            }
            return data || '';
          }.bind(this)
        },
        {
          name: 'company',
          data: '유입 구분',
          defaultContent: '',
          render: function(data) {
            if (this.badgeSystem) {
              return this.badgeSystem.createCompanyBadge(data);
            }
            return data || '';
          }.bind(this)
        },
        {
          // 사업자명 (E열, 신규) - 연한 회색 라벨
          name: 'businessName',
          data: '사업자명',
          defaultContent: '',
          render: function(data) {
            if (data && data.trim() !== '') {
              // HTML 이스케이프 (간단)
              const safe = String(data).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
              return `<span class="badge business-name-badge" style="background-color: #e9ecef; color: #495057; font-weight: 500; padding: 0.3em 0.55em; font-size: 0.8125rem; white-space: normal; text-align: left; max-width: 100%;">${safe}</span>`;
            } else {
              return '<span class="text-muted">-</span>';
            }
          }
        },
        {
          name: 'address',
          data: '현장 주소',
          defaultContent: '',
          render: function(data) {
            if (data && data.trim() !== '') {
              return data;
            } else {
              return '<i class="fas fa-exclamation-circle text-danger me-2" title="데이터가 비어 있습니다. 현장 주소를 입력해주세요."></i><span class="text-muted">주소 정보 없음</span>';
            }
          }
        },
        // 공사 내용 컬럼은 제거됨 (행 펼침 아코디언에서 표시)
        {
          name: 'startDate',
          data: '공사 시작',
          defaultContent: '',
          render: function(data) {
            if (data) {
              return formatDate(data);
            } else {
              return '<i class="fas fa-exclamation-triangle" title="데이터가 비어 있습니다. 공사 시작일을 입력해주세요." style="font-size: 1.2rem; color: #f39c12;"></i>';
            }
          }
        },
        {
          name: 'endDate',
          data: '공사 종료',
          defaultContent: '',
          render: function(data) {
            if (data) {
              return formatDate(data);
            } else {
              return '<i class="fas fa-exclamation-triangle" title="데이터가 비어 있습니다. 공사 종료일을 입력해주세요." style="font-size: 1.2rem; color: #f39c12;"></i>';
            }
          }
        },
        {
          name: 'totalAmount',
          data: null,
          className: 'amount-cell',
          render: function(data, type, row) {
            const totalAmount = row['총액 2'] || row['총액2'] || row['S'] || row['총액'] || 0;
            const numValue = parseFloat(totalAmount) || 0;

            if (numValue > 0) {
              return `<span>${numValue.toLocaleString()}원</span>`;
            } else {
              return '<i class="fas fa-exclamation-triangle" title="데이터가 비어 있습니다. 총액을 입력해주세요." style="font-size: 1.2rem; color: #f39c12;"></i>';
            }
          }
        },
        {
          name: 'outstandingNormal',
          data: null,
          className: 'text-center amount-cell',
          visible: true,  // 일반 모드에서 항상 표시
          render: function(data, type, row) {
            const outstandingData = row['미수금'] || 0;
            const outstandingAmount = parseFloat(outstandingData) || 0;
            const totalAmount = parseFloat(row['총액 2'] || row['총액2'] || row['S'] || row['총액'] || 0);

            if (outstandingAmount >= 0) {
              // 총액이 없는데 미수금이 0원인 경우 - 경고 아이콘 표시
              if (totalAmount === 0 && outstandingAmount === 0) {
                return '<i class="fas fa-exclamation-triangle" title="데이터가 비어 있습니다. 총액을 입력해주세요." style="font-size: 1.2rem; color: #f39c12;"></i>';
              }

              // 공사 완료된 경우 (미수금 0원) - 일반 모드에서는 '-' 표시
              if (outstandingAmount === 0) {
                return '<span class="text-muted">-</span>';
              }

              // 미수금이 있는 경우만 금액 표시 (빨간색)
              return `<span class="amount-outstanding" style="color: #dc3545;">${outstandingAmount.toLocaleString()}원</span>`;
            } else {
              return '<span class="text-muted">-</span>';
            }
          }
        },
        {
          name: 'status',
          data: null,
          render: (data, type, row) => {
            // 동적 상태 계산
            const status = this.calculateDynamicStatus(row);

            // 새로운 뱃지 시스템 사용
            if (this.badgeSystem) {
              return this.badgeSystem.createStatusBadge(status);
            }

            // 폴백: 기본 뱃지
            return `<span class="badge status-waiting">${status}</span>`;
          }
        },
        {
          name: 'dataCompleteness',
          data: null,
          className: 'text-center',
          render: (data, type, row) => {
            // 공사 취소 여부 확인 (띄어쓰기 무시)
            const collectionNotes = row['수금 관련 특이사항'] || row['AG'] || '';
            if (collectionNotes && /공사\s*취소/.test(collectionNotes)) {
              return '<i class="fas fa-check-circle" title="공사 취소" style="font-size: 1.2rem; color: #23923c;"></i>';
            }

            const completeness = this.checkDataCompleteness(row);

            if (completeness.missingCount === 0) {
              return '<i class="fas fa-check-circle" title="모든 필수 데이터가 입력되었습니다" style="font-size: 1.2rem; color: #23923c;"></i>';
            } else {
              const missingList = completeness.missingFields.join(', ');
              return `<i class="fas fa-exclamation-triangle" title="누락된 필드: ${missingList}" style="font-size: 1.2rem; color: #f39c12;"></i>`;
            }
          }
        },
        {
          name: 'downPayment',
          data: '계약금',
          className: 'text-center amount-cell',
          visible: false,
          defaultContent: '',
          render: function(data, type, row, meta) {
            return renderPaymentFieldWithMemo(data, row, '계약금_메모', '계약금');
          }
        },
        {
          name: 'midPayment',
          data: '중도금',
          className: 'text-center amount-cell',
          visible: false,
          render: function(data, type, row) {
            return renderPaymentFieldWithMemo(data, row, '중도금_메모', '중도금');
          }
        },
        {
          name: 'finalPayment',
          data: '잔금',
          className: 'text-center amount-cell',
          visible: false,
          render: function(data, type, row) {
            return renderPaymentFieldWithMemo(data, row, '잔금_메모', '잔금');
          }
        },
        {
          name: 'outstandingReceivable',
          data: null,
          className: 'text-center amount-cell',
          visible: false,
          render: function(data, type, row) {
            const outstandingData = row['미수금'] || 0;
            const outstandingAmount = parseFloat(outstandingData) || 0;
            const totalAmount = parseFloat(row['총액 2'] || row['총액2'] || row['S'] || row['총액'] || 0);

            if (outstandingAmount >= 0) {
              // 총액이 없는데 미수금이 0원인 경우 - 경고 아이콘 표시
              if (totalAmount === 0 && outstandingAmount === 0) {
                return '<i class="fas fa-exclamation-triangle" title="데이터가 비어 있습니다. 총액을 입력해주세요." style="font-size: 1.2rem; color: #f39c12;"></i>';
              }

              // 공사 완료된 경우 (미수금 0원) - 부모 테이블에서는 '-' 표시
              if (outstandingAmount === 0) {
                return '<span class="text-muted">-</span>';
              }

              // 미수금이 있는 경우만 금액 표시 (볼드체 + 빨간색)
              return `<span class="amount-outstanding" style="color: #dc3545;">${outstandingAmount.toLocaleString()}원</span>`;
            } else {
              return '<span class="text-muted">-</span>';
            }
          }
        },
        {
          name: 'receivableNotes',
          data: null,
          className: 'text-center',
          visible: false,
          render: (data, type, row) => {
            const notes = row['수금 관련 특이사항'] || row['수금특이사항'] || row['수금 특이사항'] || '';
            if (!notes) return '';
            if (type === 'display') {
              return `<span class="d-inline-block" style="word-break: break-word; white-space: normal;" title="${notes}">${notes}</span>`;
            }
            return notes;
          }
        },
        // ── A/S 모드 컬럼 (A/S 토글 시 표시). data 는 window.asView.byCode[코드] 조인.
        //    행은 프로젝트 그대로라 아코디언/편집/필터 전부 정상 동작. (2026-08-31)
        {
          name: 'asAcceptDate', data: null, className: 'text-center text-nowrap', visible: false,
          render: (d, t, row) => {
            const a = window.asView && window.asView.byCode && window.asView.byCode[row['프로젝트 코드']];
            return (a && a['접수 일자']) ? _asEsc(a['접수 일자']) : '<span class="text-muted">-</span>';
          }
        },
        {
          name: 'asVisitDate', data: null, className: 'text-center text-nowrap', visible: false,
          render: (d, t, row) => {
            const a = window.asView && window.asView.byCode && window.asView.byCode[row['프로젝트 코드']];
            return (a && a['방문 예정일']) ? _asEsc(a['방문 예정일']) : '<span class="text-muted">-</span>';
          }
        },
        {
          name: 'asRequest', data: null, className: 'text-left', visible: false,
          render: (d, t, row) => {
            const a = window.asView && window.asView.byCode && window.asView.byCode[row['프로젝트 코드']];
            const v = a ? (a['요청 내용'] || '') : '';
            return v ? `<span style="white-space:normal;" title="${_asEsc(v)}">${_asEsc(v)}</span>` : '<span class="text-muted">-</span>';
          }
        },
        {
          name: 'asVisitor', data: null, className: 'text-center', visible: false,
          render: (d, t, row) => {
            const a = window.asView && window.asView.byCode && window.asView.byCode[row['프로젝트 코드']];
            return (a && a['방문 예정자']) ? _asEsc(a['방문 예정자']) : '<span class="text-muted">-</span>';
          }
        },
        {
          name: 'asStatus', data: null, className: 'text-center', visible: false,
          render: (d, t, row) => {
            const a = window.asView && window.asView.byCode && window.asView.byCode[row['프로젝트 코드']];
            if (!a) return '<span class="text-muted">-</span>';
            const st = String(a['진행 상태'] || '').trim();
            const map = { '요청됨': ['🔔 요청됨', 'bg-warning text-dark'], '접수 완료': ['📥 접수완료', 'bg-info text-dark'], '처리 완료': ['✅ 처리완료', 'bg-success'] };
            const m = map[st] || [st, 'bg-secondary'];
            return `<span class="badge ${m[1]}">${_asEsc(m[0])}</span>`;
          }
        },
        {
          name: 'asResolution', data: null, className: 'text-left', visible: false,
          render: (d, t, row) => {
            const a = window.asView && window.asView.byCode && window.asView.byCode[row['프로젝트 코드']];
            if (!a) return '<span class="text-muted">-</span>';
            const st = String(a['진행 상태'] || '').trim();
            const asNo = _asEsc(a['No']);
            let btn = '';
            if (st === '요청됨') btn = ` <button class="btn btn-primary btn-sm py-0 px-2" onclick="event.stopPropagation(); window.asView&&window.asView.openAccept('${asNo}')">접수</button>`;
            else if (st === '접수 완료') btn = ` <button class="btn btn-success btn-sm py-0 px-2" onclick="event.stopPropagation(); window.asView&&window.asView.openComplete('${asNo}')">완료</button>`;
            const res = a['처리 내용'] || '';
            return `${res ? `<span style="white-space:normal;" title="${_asEsc(res)}">${_asEsc(res)}</span>` : '<span class="text-muted">-</span>'}${btn}`;
          }
        }
      ],
      columnDefs: [
        {
          // 프로젝트 코드: 컴팩트 (8자 정도) + 줄바꿈 방지로 해상도 영향 없음
          targets: [0],
          width: '8%',
          className: 'text-left text-nowrap'
        },
        {
          // 담당자, 거래처: 넓이 축소 (비율)
          targets: [1, 2],
          width: '7%',
          className: 'text-left'
        },
        {
          // 사업자명: 가운데 정렬
          targets: [3],
          className: 'text-center'
        },
        {
          // 현장 주소: 왼쪽 정렬
          targets: [4],
          className: 'text-left'
        },
        {
          // 날짜: 넓이 축소 (비율)
          targets: [5, 6],
          width: '7%',
          className: 'text-center'
        },
        {
          // 금액, 미수금: 넓이 축소 (비율)
          targets: [7, 8],
          width: '8%',
          className: 'text-center'
        },
        {
          // 상태, 데이터 완성도: 넓이 축소 (비율)
          targets: [9, 10],
          width: '6%',
          className: 'text-center'
        }
        // visible 설정은 각 컬럼 정의에서 직접 설정됨 (중복 방지)
      ],
      createdRow: function(row, data, dataIndex) {
        const status = data['공사상태'] || '';
        const outstandingAmount = parseFloat(data['미수금'] || 0);

        // 상태에 따른 행 스타일링 (jQuery 제거)
        row.classList.remove('table-success', 'table-warning', 'table-danger', 'table-info', 'table-secondary');

        switch(status) {
          case '공사완료':
            if (outstandingAmount > 0) {
              row.classList.add('table-warning'); // 완료했지만 미수금 있음
            } else {
              row.classList.add('table-success'); // 완료 + 수금 완료
            }
            break;
          case '수금필요':
            row.classList.add('table-danger'); // 긴급 처리 필요
            break;
          case '확인필요':
            row.classList.add('table-info'); // 확인 필요
            break;
          case '공사진행':
            row.classList.add('table-light'); // 정상 진행
            break;
          case '공사대기':
            row.classList.add('table-secondary'); // 대기 상태
            break;
        }

        // 미수금이 높은 경우 추가 표시
        if (outstandingAmount > 5000000) { // 500만원 이상
          row.classList.add('border-danger', 'border-3');
        }

        // 취소된 공사 체크 및 빨간 취소선 적용 (띄어쓰기 무시)
        const collectionNotes = data['수금 관련 특이사항'] || data['AG'] || '';
        if (collectionNotes && /공사\s*취소/.test(collectionNotes)) {
          row.classList.add('project-cancelled-row');
        }

        // 상태 뱃지가 있는 셀에 식별 클래스 추가
        Array.from(row.cells).forEach(cell => {
          const badgeEl = cell.querySelector('.badge');
          if (badgeEl && /\bstatus-[a-z-]+\b/.test(badgeEl.className)) {
            cell.classList.add('status-column-cell');
          }
        });

        // Note: 툴팁 초기화는 drawCallback의 initializeTooltips()에서 처리됨
      },
      drawCallback: (settings) => {
        // 테이블이 그려진 후 전역 툴팁 초기화 함수 호출
        this.initializeTooltips();

        // 상태 뱃지 셀 식별 클래스 동기화
        const tableBody = settings?.nTable?.tBodies?.[0];
        if (tableBody) {
          tableBody.querySelectorAll('td').forEach(cell => {
            const badgeEl = cell.querySelector('.badge');
            if (!badgeEl) return;

            const badgeClasses = badgeEl.className || '';
            const isStatusBadge = /\bstatus-[a-z-]+\b/.test(badgeClasses);
            const badgeType = badgeEl.dataset?.badgeType || '';
            const isScoped = badgeType === 'status';

            if (isStatusBadge || isScoped) {
              cell.classList.add('status-column-cell');
            } else {
              cell.classList.remove('status-column-cell');
            }
          });
        }
      },
      initComplete: (settings, json) => {

        // 테이블 초기화 완료 후 정렬 이벤트 직접 제어
        const api = settings.oInstance.api();
        const tableElement = settings.nTable;

        // 테이블 보이기 (FOUC 방지)
        tableElement.style.visibility = 'visible';

        // "내 공사만 보기" 체크박스 추가
        this.addMyProjectsFilter(tableElement);

        // 모든 정렬 가능한 헤더에 클릭 이벤트 추가
        tableElement.querySelectorAll('thead th').forEach((th, index) => {
          if (th.classList.contains('sorting') || th.classList.contains('sorting_asc') || th.classList.contains('sorting_desc')) {
            th.addEventListener('click', function(e) {
              e.preventDefault();
              e.stopPropagation();

              // 현재 정렬 상태 확인
              const currentOrder = api.order()[0]; // [columnIndex, direction]
              const currentColumn = currentOrder ? currentOrder[0] : -1;
              const currentDirection = currentOrder ? currentOrder[1] : null;

              if (currentColumn === index) {
                // 같은 컬럼 클릭 시: asc ↔ desc 토글
                const newDirection = currentDirection === 'asc' ? 'desc' : 'asc';
                api.order([index, newDirection]).draw();
              } else {
                // 다른 컬럼 클릭 시: 항상 asc부터 시작
                api.order([index, 'asc']).draw();
              }
            });
          }
        });
      }
    });
  }

  /**
   * 테이블 데이터 업데이트
   */
  updateData(data) {
    if (!this.table) return;

    // 2026-07-08 재진입 방지: 이전 updateData가 아직 draw 완료 안 됐으면 이 요청 스킵.
    // 짧은 간격 연속 호출로 인한 draw 중복·아코디언 상태 유실·수금 모드 플래그 꼬임 방지.
    // 플래그는 draw.dt 콜백에서 clear (updateDataRegular / updateDataWithOptimization).
    if (this._updateInProgress) {
      logger.debug('[ProjectTable] updateData 재진입 감지 - 스킵');
      return;
    }
    this._updateInProgress = true;

    // 수금 모드 전환 중에는 배치 처리 건너뛰기 (빠른 draw 보장)
    if (this.isReceivablesTransitioning) {
      logger.debug('[ProjectTable] 수금 모드 전환 중 - 일반 업데이트 사용');
      this.updateDataRegular(data);
      return;
    }

    // 대용량 데이터셋 감지
    this.isLargeDataset = data && data.length > this.largeDatasetThreshold;

    if (this.isLargeDataset) {
      logger.debug(`[ROCKET] 대용량 데이터셋 감지: ${data.length}개 항목`);
      this.updateDataWithOptimization(data);
    } else {
      this.updateDataRegular(data);
    }
  }

  /**
   * 일반 데이터 업데이트
   */
  updateDataRegular(data) {
    this.cancelPendingBatch();

    // 아코디언 상태 저장 (currentProject가 객체이므로 프로젝트 코드 추출)
    const openAccordionCode = this.accordion?.currentProject?.['프로젝트 코드'] || null;
    const wasOpen = this.accordion?.isOpen || false;

    this.table.clear();

    if (data && data.length > 0) {
      this.table.rows.add(data);
    }

    // DataTables draw 완료 콜백 사용 (더 안전한 타이밍)
    this.table.one('draw.dt', () => {
      // ❌ 여기서는 스켈레톤을 제거하지 않음
      // 수금 모드 토글의 refreshData() 완료 후에만 제거됨

      // 수금 모드 상태 복원
      this.restoreReceivablesMode();
      // A/S 모드 상태 복원 (컬럼 visibility 재적용)
      this.restoreAsMode();

      // 아코디언 상태 복원
      if (wasOpen && openAccordionCode) {
        // draw 완료 직후 DOM 안정화 대기
        requestAnimationFrame(() => {
          const row = this.table.rows().nodes().toArray().find(r => {
            const rowData = this.table.row(r).data();
            return rowData && rowData['프로젝트 코드'] === openAccordionCode;
          });

          if (row) {
            this.accordion.openAccordion(row, this.table.row(row).data());
          } else {
            logger.warn(`[ProjectTable] 아코디언 복원 실패: 행을 찾을 수 없음 (${openAccordionCode})`);
          }
        });
      }

      // 재진입 방지 플래그 해제 (다음 updateData 호출 허용)
      this._updateInProgress = false;
    });

    this.table.draw();
  }

  /**
   * 최적화된 데이터 업데이트 (대용량)
   */
  updateDataWithOptimization(data) {
    if (!this.performanceOptimizer) {
      logger.warn('PerformanceOptimizer가 없습니다. 일반 모드로 전환합니다.');
      this.updateDataRegular(data);
      return;
    }
    this.cancelPendingBatch();

    // 아코디언 상태 저장 (currentProject가 객체이므로 프로젝트 코드 추출)
    const openAccordionCode = this.accordion?.currentProject?.['프로젝트 코드'] || null;
    const wasOpen = this.accordion?.isOpen || false;

    // 성능 측정 시작
    const startTime = performance.now();

    // 메모이제이션된 데이터 처리
    const processedData = this.performanceOptimizer.memoize(
      (rawData) => this.preprocessTableData(rawData),
      (rawData) => `table_data_${rawData.length}_${this.hashData(rawData)}`
    )(data);

    // 배치 업데이트 with draw 완료 콜백
    this.table.clear();
    this.batchUpdateRows(processedData, () => {
      // ❌ 여기서는 스켈레톤을 제거하지 않음
      // 수금 모드 토글의 refreshData() 완료 후에만 제거됨

      // 성능 측정 완료
      const duration = performance.now() - startTime;
      logger.debug(`⚡ 최적화된 업데이트 완료: ${duration.toFixed(2)}ms`);

      // 성능 리포트 (개발 모드에서만)
      if (this.performanceOptimizer.config.performanceMonitoring) {
        this.performanceOptimizer.logPerformanceStats();
      }

      // 수금 모드 상태 복원
      this.restoreReceivablesMode();
      // A/S 모드 상태 복원 (컬럼 visibility 재적용)
      this.restoreAsMode();

      // 아코디언 상태 복원
      if (wasOpen && openAccordionCode) {
        requestAnimationFrame(() => {
          const row = this.table.rows().nodes().toArray().find(r => {
            const rowData = this.table.row(r).data();
            return rowData && rowData['프로젝트 코드'] === openAccordionCode;
          });

          if (row) {
            this.accordion.openAccordion(row, this.table.row(row).data());
          } else {
            logger.warn(`[ProjectTable] 아코디언 복원 실패: 행을 찾을 수 없음 (${openAccordionCode})`);
          }
        });
      }
    });
  }

  cancelPendingBatch() {
    if (this.pendingBatchTimeout) {
      clearTimeout(this.pendingBatchTimeout);
      this.pendingBatchTimeout = null;
    }
    this.currentBatchToken += 1;
  }

  /**
   * 배치 행 업데이트
   */
  batchUpdateRows(data, onComplete = null) {
    const batchSize = 15;  // 모던 테이블 페이지 사이즈에 맞춤
    let index = 0;
    const batchToken = this.currentBatchToken;

    const runBatch = () => {
      if (batchToken !== this.currentBatchToken) {
        return;
      }

      if (index >= data.length) {
        this.pendingBatchTimeout = null;

        // draw 완료 콜백 등록 후 draw 호출
        // 재진입 방지 플래그도 함께 해제 (2026-07-08)
        this.table.one('draw.dt', () => {
          if (onComplete) {
            try { onComplete(); } catch (e) { logger.warn('[ProjectTable] onComplete 오류:', e); }
          }
          this._updateInProgress = false;
        });
        this.table.draw();
        return;
      }

      const batch = data.slice(index, index + batchSize);
      this.table.rows.add(batch);
      index += batchSize;

      this.pendingBatchTimeout = setTimeout(runBatch, 0);
    };

    this.pendingBatchTimeout = null;
    this.pendingBatchTimeout = setTimeout(runBatch, 0);
  }

  /**
   * 테이블 데이터 전처리
   */
  preprocessTableData(data) {
    return data.map(item => {
      // 데이터 정규화 및 캐싱 가능한 형태로 변환
      return {
        ...item,
        _processed: true,
        _timestamp: Date.now()
      };
    });
  }

  /**
   * 데이터 해시 생성 (간단한 체크섬)
   */
  hashData(data) {
    if (!data || data.length === 0) return '0';

    const sampleSize = Math.min(10, data.length);
    const sample = data.slice(0, sampleSize);
    const str = JSON.stringify(sample);

    let hash = 0;
    for (let i = 0; i < str.length; i++) {
      const char = str.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash;
    }
    return hash.toString();
  }


  /**
   * 필터 적용
   */
  applyFilter(searchTerm) {
    if (this.table) {
      this.table.search(searchTerm).draw();
    }
  }

  /**
   * 컴포넌트 정리
   */
  destroy() {
    // ModeManager 이벤트 리스너 제거 (메모리 누수 방지)
    if (this.boundHandleTableModeChange) {
      this.modeManager.off('tableModeChanged', this.boundHandleTableModeChange);
      this.boundHandleTableModeChange = null;
    }

    this.cancelPendingBatch();

    // 스켈레톤 정리
    this.hideSkeletonLoader();

    if (this.accordion) {
      this.accordion.destroy();
      this.accordion = null;
    }

    if (this.statusBadge) {
      this.statusBadge.destroy();
      this.statusBadge = null;
    }

    if (this.performanceOptimizer) {
      this.performanceOptimizer.destroy();
      this.performanceOptimizer = null;
    }

    if (this.table) {
      this.table.destroy();
      this.table = null;
    }
  }

  /**
   * 동적 상태 계산 (원본 getStatusBadge 로직 완전 복제)
   */
  calculateDynamicStatus(rowData) {
    // 공사 취소 여부 확인 (최우선, 띄어쓰기 무시)
    const collectionNotes = rowData['수금 관련 특이사항'] || rowData['AG'] || '';
    if (collectionNotes && /공사\s*취소/.test(collectionNotes)) {
      return '공사취소';
    }

    const startDateStr = rowData['공사 시작'] || '';
    const endDateStr = rowData['공사 종료'] || '';

    // 금액 정보
    const contractAmount = parseFloat(rowData['계약금'] || 0);
    const midAmount = parseFloat(rowData['중도금'] || 0);
    const finalAmount = parseFloat(rowData['잔금'] || 0);
    const outstandingAmount = parseFloat(rowData['미수금'] || 0);
    const totalAmount = parseFloat(rowData['총액 2'] || rowData['총액2'] || rowData['S'] || rowData['총액'] || 0);

    // 수금 확인 (Z열)
    const paymentConfirmed = rowData['수금 확인'] || rowData['Z'] || '';
    const isPaymentConfirmed = (paymentConfirmed === true || paymentConfirmed === 'TRUE' ||
                              paymentConfirmed === '✓' || paymentConfirmed === 'true' ||
                              paymentConfirmed === 'Y' || paymentConfirmed === 'y' ||
                              paymentConfirmed === '1' || paymentConfirmed === 1);

    // 총 입금액 계산
    const totalPaidAmount = contractAmount + midAmount + finalAmount;

    // 날짜 정보 처리
    let startDate = null, endDate = null, today = new Date();
    today.setHours(0, 0, 0, 0);

    if (startDateStr) {
        try {
            startDate = new Date(startDateStr);
            startDate.setHours(0, 0, 0, 0);
        } catch (e) {
            logger.warn('시작일 파싱 오류:', e);
        }
    }

    if (endDateStr) {
        try {
            endDate = new Date(endDateStr);
            endDate.setHours(0, 0, 0, 0);
        } catch (e) {
            logger.warn('종료일 파싱 오류:', e);
        }
    }

    // 1. 🔴 수금필요: 공사 종료일 지났는데 입금 완료 안됨 (최우선)
    if (endDate && today > endDate) {
        if ((totalAmount > 0 && totalPaidAmount < totalAmount) ||
            (totalAmount === 0 && outstandingAmount > 0)) {
            return '수금필요';
        }
    }

    // 2. 🟢 공사완료: 모든 입금 완료
    if ((totalAmount > 0 && totalPaidAmount === totalAmount && outstandingAmount === 0) ||
        (totalAmount === 0 && outstandingAmount === 0 && finalAmount > 0)) {
        return '공사완료';
    }

    // 3. 🟠 확인필요: 이상한 상황들 + 데이터 부족
    if ((totalAmount > 0 && totalPaidAmount === totalAmount && outstandingAmount > 0) ||
        (totalAmount > 0 && totalPaidAmount < totalAmount && contractAmount > 0 && midAmount > 0 && finalAmount > 0) ||
        (totalAmount === 0 && outstandingAmount > 0 && contractAmount > 0 && midAmount > 0 && finalAmount > 0) ||
        (totalPaidAmount === 0 && outstandingAmount === 0) || // 모든 금액 정보 없음
        (!startDateStr && !endDateStr)) { // 날짜 정보 없음
        return '확인필요';
    }

    // 4. ⚫ 공사대기: 오직 공사 시작 전만
    if (startDate && today < startDate) {
        return '공사대기';
    }

    // 5. 🔵 공사진행: 나머지 모든 경우 (입금이 있거나 공사가 시작됨)
    return '공사진행';
  }


  /**
   * 데이터 완성도 확인 (레거시 로직 복제)
   */
  checkDataCompleteness(row) {
    // 필수 필드 목록 (레거시와 동일)
    const requiredFields = [
      '사업자', '발주처 담당자', '발주처 연락처', '현장 주소', '공사 내용',
      '공사 시작', '공사 종료', '총액 1', '총액 2', '기계 분류', '브랜드',
      '견적서 및 계약서 폴더 경로'
    ];

    const missingFields = [];

    requiredFields.forEach(field => {
      const value = row[field];
      if (!value || (typeof value === 'string' && value.trim() === '')) {
        missingFields.push(field);
      }
    });

    return {
      totalFields: requiredFields.length,
      missingCount: missingFields.length,
      missingFields: missingFields,
      completionRate: Math.round(((requiredFields.length - missingFields.length) / requiredFields.length) * 100)
    };
  }

  /**
   * Initialize CSS Grid layout
   * @private
   */
  async initializeGridLayout() {
    const wrapper = this.tableElement.closest('.table-wrapper');
    if (wrapper) {
      wrapper.classList.add('grid-modern-table', 'gpu-accelerated');
      // wrapper.style.containerType = 'inline-size'; // Disabled to fix width alignment
    }
    // Grid layout 초기화 완료
  }

  /**
   * Initialize container queries support
   * @private
   */
  async initializeContainerQueries() {
    // Check for container queries support
    if (!CSS.supports('container-type', 'inline-size')) {
      logger.warn('[ProjectTable] Container queries not supported, falling back to media queries');
      return;
    }

    const section = this.tableElement.closest('.table-section');
    if (section) {
      // section.classList.add('container-query-enabled'); // Disabled to fix width alignment
    }
    // Container queries 설정 완료
  }

  /**
   * Apply modern grid layout to table row
   * @param {HTMLElement} row - Table row element
   * @param {Object} data - Row data
   * @param {number} dataIndex - Row index
   */
  applyGridLayout(row, data, dataIndex) {
    if (!row) return;

    // Add GPU acceleration and containment
    row.classList.add('gpu-accelerated', 'contain-layout');

    // Apply grid styling
    const cells = row.querySelectorAll('td');
    cells.forEach((cell, index) => {
      cell.style.gridColumn = index + 1;
      cell.classList.add('transition-fast');
    });
  }

  /**
   * CRITICAL FIX: Remove any existing contain styles that may cause width issues
   */
  removeContainStyles() {
    const wrappers = document.querySelectorAll('.table-wrapper, .table-virtual-container, .table-section');
    wrappers.forEach(wrapper => {
      wrapper.style.contain = '';
      wrapper.style.removeProperty('contain');
      wrapper.style.containerType = '';
      wrapper.style.removeProperty('container-type');
      wrapper.style.removeProperty('container-name');

      // 테이블 섹션의 padding도 제거하여 완전한 폭 정렬 달성
      if (wrapper.classList.contains('table-section')) {
        wrapper.style.padding = '0';
      }

    });
  }

  /**
   * Enable virtual scrolling for large datasets
   * @param {Array} data - Table data
   */
  enableVirtualScrolling(data) {
    if (!data || data.length < this.largeDatasetThreshold) {
      return;
    }

    const wrapper = this.tableElement.closest('.table-wrapper');
    if (wrapper) {
      wrapper.classList.add('table-virtual-container');
      wrapper.style.height = '400px';
      // wrapper.style.contain = 'strict'; // Disabled to fix width alignment

      // CRITICAL FIX: Force remove any existing contain styles
      wrapper.style.contain = '';
      wrapper.style.removeProperty('contain');
    }

  }

  /**
   * "내 공사만 보기" 체크박스를 "개씩 보기" 우측에 추가
   * @param {HTMLElement} tableElement - 테이블 엘리먼트
   */
  addMyProjectsFilter(tableElement) {
    // 기존 체크박스가 있으면 제거
    const existingFilter = document.querySelector('.my-projects-filter');
    if (existingFilter) {
      existingFilter.remove();
    }

    // .dt-length 엘리먼트 찾기
    const lengthContainer = tableElement.closest('.table-wrapper').querySelector('.dt-length');
    if (!lengthContainer) {
      logger.warn('[ProjectTable] .dt-length container not found');
      return;
    }

    // lengthContainer를 flex로 변경하여 우측 정렬 가능하게 함
    lengthContainer.style.display = 'flex';
    lengthContainer.style.alignItems = 'center';
    lengthContainer.style.gap = '15px';


    // 체크박스 컨테이너 생성
    const filterContainer = document.createElement('div');
    filterContainer.className = 'my-projects-filter';

    // 라벨 생성 (텍스트가 먼저 오도록)
    const label = document.createElement('label');
    label.textContent = '내 프로젝트만 보기';
    label.setAttribute('for', 'myProjectsOnly');

    // 체크박스 생성
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = 'myProjectsOnly';
    checkbox.title = '내 담당 프로젝트만 표시';

    // 체크박스 이벤트 리스너 추가 — 담당자 필터에 로그인 사용자 이름 세팅으로 위임
    // (2026-07-07 변경, ModernProjectFilters의 새 로직과 동일)
    checkbox.addEventListener('change', (e) => {
      if (window.modernFilters) {
        if (e.target.checked) {
          const userName = (window.userDisplayName || '').trim();
          if (userName) {
            window.modernFilters.filters.manager = userName;
            if (window.modernFilters.managerFilter) {
              window.modernFilters.managerFilter.value = userName;
            }
          }
        } else {
          delete window.modernFilters.filters.manager;
          if (window.modernFilters.managerFilter) {
            window.modernFilters.managerFilter.value = '';
          }
        }
        window.modernFilters.applyFilters(null, true);
      }
    });

    // 엘리먼트 순서: 체크박스 먼저, 라벨 나중에
    filterContainer.appendChild(checkbox);
    filterContainer.appendChild(label);

    // lengthContainer에 추가
    lengthContainer.appendChild(filterContainer);


    // 수금 관리 모드 토글 추가
    this.addReceivablesToggleToLengthContainer(lengthContainer);
  }

  /**
   * lengthContainer에 수금 관리 모드 토글 추가
   */
  addReceivablesToggleToLengthContainer(lengthContainer) {
    // 토글 컨테이너 생성 (Bootstrap form-switch)
    const toggleContainer = document.createElement('div');
    toggleContainer.className = 'form-check form-switch';
    toggleContainer.style.marginLeft = 'auto'; // 우측 끝으로 밀기
    toggleContainer.style.marginRight = '20px'; // 오른쪽 여백
    toggleContainer.style.display = 'flex';
    toggleContainer.style.alignItems = 'center';

    // 토글 스위치 생성
    const toggleSwitch = document.createElement('input');
    toggleSwitch.type = 'checkbox';
    toggleSwitch.id = 'receivablesToggle';
    toggleSwitch.className = 'form-check-input';
    toggleSwitch.style.cursor = 'pointer';
    toggleSwitch.style.marginTop = '0';
    toggleSwitch.style.verticalAlign = 'middle';

    // 라벨 생성
    const label = document.createElement('label');
    label.className = 'form-check-label';
    label.setAttribute('for', 'receivablesToggle');
    label.textContent = '수금 관리 모드';
    label.style.cursor = 'pointer';
    label.style.userSelect = 'none';
    label.style.marginLeft = '8px';
    label.style.fontWeight = '500';
    label.style.lineHeight = '1.5';
    label.style.verticalAlign = 'middle';
    label.style.display = 'inline-block';

    // 컨테이너에 추가
    toggleContainer.appendChild(toggleSwitch);
    toggleContainer.appendChild(label);

    // lengthContainer에 추가
    lengthContainer.appendChild(toggleContainer);

  }

  /**
   * 수금 관리 모드 토글 설정
   * 수금 모드: 공사내용(4), 공사시작(5), 상태(8), 데이터(9) 숨기고 → 계약금, 중도금, 잔금, 미수금, 수금특이사항 표시
   */
  setupReceivablesToggle() {
    const toggle = document.getElementById('receivablesToggle');
    if (!toggle) {
      logger.warn('[ProjectTable] 수금 관리 토글 버튼을 찾을 수 없습니다.');
      return;
    }

    // 초기값 복원: 우선순위 = URL query > sessionStorage > 기본(false)
    // URL query `?receivables=1` → 통계 페이지의 '수금관리' 링크로 진입 시 자동 활성.
    // sessionStorage → 새로고침(F5) 시 이전 모드 유지.
    try {
      const params = new URLSearchParams(window.location.search);
      const fromUrl = params.get('receivables');
      const fromStorage = sessionStorage.getItem('itg_receivables_mode');
      let initial = false;
      if (fromUrl === '1') {
        initial = true;
        // URL 파라미터 정리 (다음 새로고침엔 sessionStorage 만 참조)
        params.delete('receivables');
        const clean = window.location.pathname + (params.toString() ? '?' + params : '') + window.location.hash;
        window.history.replaceState({}, '', clean);
      } else if (fromStorage === '1') {
        initial = true;
      }
      if (initial) {
        toggle.checked = true;
        // change 이벤트 발생시켜 실제 모드 전환 실행 (컬럼 visibility, 필터 등)
        toggle.dispatchEvent(new Event('change', { bubbles: true }));
      }
    } catch (err) {
      logger.warn('[ProjectTable] 수금 모드 초기값 복원 실패:', err);
    }

    toggle.addEventListener('change', async (e) => {
      // 🔒 모드 전환 시 편집 중인 아코디언 처리
      if (this.accordion && this.accordion.modeManager.isEditMode()) {
        // 변경사항 확인 (메모 포함) - EditState 에러 시 자동 편집 종료
        const projectCode = this.accordion.currentProject?.['프로젝트 코드'];
        let allChanges = {};
        let hasChanges = false;

        try {
          allChanges = projectCode ? this.accordion.collectAllChanges(projectCode) : {};
          hasChanges = Object.keys(allChanges).length > 0;
        } catch (error) {
          logger.error('[모드 전환] collectAllChanges 오류:', error);
          // 🆕 자동으로 편집 모드 종료
          if (projectCode) {
            this.accordion.disableUnifiedEditMode(projectCode);
          }
          // 🆕 사용자에게 알림
          if (this.accordion && this.accordion.showMessage) {
            this.accordion.showMessage('편집 모드가 비정상적으로 종료되었습니다. 모드 전환을 진행합니다.', 'warning');
          }
          // 오류 발생 시 변경사항 없는 것으로 간주하고 모드 전환 진행
          allChanges = {};
          hasChanges = false;
        }

        if (hasChanges) {
          // 변경사항이 있으면 확인 다이얼로그 표시
          const currentProjectName = this.accordion.currentProject?.['프로젝트명'] || '현재 프로젝트';
          const confirmed = confirm(
            `"${currentProjectName}"에 저장하지 않은 변경사항이 있습니다.\n\n` +
            `모드를 전환하시겠습니까?\n\n` +
            `• 확인: 변경사항을 버리고 모드 전환\n` +
            `• 취소: 현재 프로젝트 계속 편집`
          );

          if (!confirmed) {
            // 사용자가 취소 선택 → 토글 상태 복원
            e.target.checked = this.modeManager.isReceivablesMode();
            return;
          }
        }

        // 편집 모드 정리 (잠금 해제 포함)
        logger.debug('[ProjectTable] 모드 전환: 편집 모드 정리');
        this.accordion.cleanupEditMode();
      }

      // 🆕 ModeManager를 통한 모드 상태 업데이트
      const newMode = e.target.checked ? TABLE_MODE.RECEIVABLES : TABLE_MODE.NORMAL;
      const modeChangeSuccess = this.modeManager.setTableMode(newMode);

      if (!modeChangeSuccess) {
        // 모드 변경 실패 시 토글 상태 복원 (일관성 위해 showSystemAlert 사용)
        e.target.checked = this.modeManager.isReceivablesMode();
        const failMsg = '모드 전환에 실패했습니다. 편집 모드를 종료한 후 다시 시도해주세요.';
        if (window.showSystemAlert) window.showSystemAlert(failMsg, 'warning');
        else alert(failMsg);
        return;
      }

      // sessionStorage 에 모드 저장 → F5 시 복원
      try {
        if (newMode === TABLE_MODE.RECEIVABLES) {
          sessionStorage.setItem('itg_receivables_mode', '1');
        } else {
          sessionStorage.removeItem('itg_receivables_mode');
        }
      } catch (_) { /* private mode 등 무시 */ }

      // 🔖 아코디언 상태 저장 (모드 전환 후 복원용)
      const openAccordionCode = this.accordion?.currentProject?.['프로젝트 코드'] || null;
      const wasOpen = this.accordion?.isOpen || false;

      // 전환 플래그 설정 및 스켈레톤 표시
      this.isReceivablesTransitioning = true;
      this.showSkeletonLoader();

      logger.debug(`[ProjectTable] 수금 모드 전환 시작 (${this.modeManager.isReceivablesMode() ? 'ON' : 'OFF'})`);

      // ✅ 1단계: 컬럼 visibility 먼저 변경 (draw 없이 - 깜빡임 방지)
      // (workContent 컬럼은 메인 테이블에서 제거됨 — 아코디언에서 표시)
      if (this.modeManager.isReceivablesMode()) {
        // 수금 관리 모드 ON
        this.table.column('startDate:name').visible(false, false);
        this.table.column('outstandingNormal:name').visible(false, false);
        this.table.column('status:name').visible(false, false);
        this.table.column('dataCompleteness:name').visible(false, false);

        this.table.column('downPayment:name').visible(true, false);
        this.table.column('midPayment:name').visible(true, false);
        this.table.column('finalPayment:name').visible(true, false);
        this.table.column('outstandingReceivable:name').visible(true, false);
        this.table.column('receivableNotes:name').visible(true, false);
      } else {
        // 일반 모드 복원
        this.table.column('startDate:name').visible(true, false);
        this.table.column('outstandingNormal:name').visible(true, false);
        this.table.column('status:name').visible(true, false);
        this.table.column('dataCompleteness:name').visible(true, false);

        this.table.column('downPayment:name').visible(false, false);
        this.table.column('midPayment:name').visible(false, false);
        this.table.column('finalPayment:name').visible(false, false);
        this.table.column('outstandingReceivable:name').visible(false, false);
        this.table.column('receivableNotes:name').visible(false, false);
      }

      this.table.columns.adjust();

      // ✅ 2단계: 필터 설정 (값만)
      const outstandingFilter = document.getElementById('outstandingFilter');
      if (outstandingFilter) {
        outstandingFilter.value = this.modeManager.isReceivablesMode() ? 'outstanding' : '';
      }

      if (window.modernFilters) {
        if (this.modeManager.isReceivablesMode()) {
          window.modernFilters.filters.outstanding = 'outstanding';
        } else {
          delete window.modernFilters.filters.outstanding;
        }
      }

      // ✅ 3단계: draw 완료 대기 - 필터 적용 후
      this.table.one('draw.dt', () => {
        // ✅ 4단계: 툴팁 초기화 (메모 아이콘 툴팁 표시)
        this.initializeTooltips();

        // ✅ 5단계: 아코디언 상태 복원 (모드 전환 시에도 열린 상태 유지)
        if (wasOpen && openAccordionCode) {
          requestAnimationFrame(() => {
            const row = this.table.rows().nodes().toArray().find(r => {
              const rowData = this.table.row(r).data();
              return rowData && rowData['프로젝트 코드'] === openAccordionCode;
            });

            if (row) {
              this.accordion.openAccordion(row, this.table.row(row).data());
              logger.debug(`[ProjectTable] 모드 전환 후 아코디언 복원: ${openAccordionCode}`);
            } else {
              logger.warn(`[ProjectTable] 아코디언 복원 실패: 행을 찾을 수 없음 (${openAccordionCode})`);
            }
          });
        }

        // ✅ 6단계: 스켈레톤 제거 (StateManager에 이미 최신 데이터 있음)
        // 더 이상 서버에서 데이터를 가져올 필요 없음 - 저장 시 이미 동기화됨
        this.hideSkeletonLoader();
        this.isReceivablesTransitioning = false;
        logger.debug('[ProjectTable] 수금 모드 전환 완료 (0-10ms, 서버 호출 없음)');
      });

      // ✅ 3단계: 필터 적용 (draw 트리거)
      if (window.modernFilters && window.modernFilters.applyFilters) {
        window.modernFilters.applyFilters(null, true);
      }
    });

  }

  /**
   * 스켈레톤 로더 표시 (Overlay 방식으로 안전하게 처리)
   */
  showSkeletonLoader() {
    if (!this.table || !this.tableElement) {
      logger.warn('[ProjectTable] DataTable이 초기화되지 않음');
      return;
    }

    // 기존 타임아웃 제거
    if (this.skeletonTimeout) {
      clearTimeout(this.skeletonTimeout);
      this.skeletonTimeout = null;
    }

    // 기존 오버레이 제거
    this.hideSkeletonLoader();

    try {
      // 테이블 래퍼 찾기
      const wrapper = this.tableElement.closest('.table-wrapper') || this.tableElement.parentElement;
      if (!wrapper) {
        logger.error('[ProjectTable] 테이블 래퍼를 찾을 수 없음');
        return;
      }

      // 오버레이 생성
      const overlay = document.createElement('div');
      overlay.className = 'skeleton-overlay';
      overlay.style.position = 'absolute';
      overlay.style.top = '0';
      overlay.style.left = '0';
      overlay.style.right = '0';
      overlay.style.bottom = '0';
      overlay.style.backgroundColor = 'rgba(255, 255, 255, 0.9)';
      overlay.style.display = 'flex';
      overlay.style.alignItems = 'center';
      overlay.style.justifyContent = 'center';
      overlay.style.zIndex = '1000';
      overlay.style.pointerEvents = 'auto';  // 클릭 차단 (전환 중 사용자 입력 방지)

      // 스켈레톤 컨텐츠
      const content = document.createElement('div');
      content.style.textAlign = 'center';

      const skeletonDiv = document.createElement('div');
      skeletonDiv.className = 'skeleton';
      skeletonDiv.style.height = '20px';
      skeletonDiv.style.width = '200px';
      skeletonDiv.style.margin = '0 auto 10px';

      const text = document.createElement('div');
      text.textContent = '모드 변경 중...';
      text.style.color = '#666';
      text.style.fontSize = '14px';
      text.style.marginTop = '10px';

      content.appendChild(skeletonDiv);
      content.appendChild(text);
      overlay.appendChild(content);

      // wrapper가 position: relative인지 확인 (원래 값 저장)
      const wrapperPosition = window.getComputedStyle(wrapper).position;
      if (wrapperPosition === 'static') {
        this.originalWrapperPosition = wrapperPosition; // 원래 값 저장
        wrapper.style.position = 'relative';
      }

      wrapper.appendChild(overlay);
      this.skeletonOverlay = overlay;

      // Fallback: 3초 후 자동 제거 (draw가 실행되지 않을 경우 대비)
      this.skeletonTimeout = setTimeout(() => {
        logger.warn('[ProjectTable] 스켈레톤 타임아웃: draw가 3초 내에 실행되지 않았습니다.');
        this.hideSkeletonLoader();
        this.isReceivablesTransitioning = false;
      }, 3000);

    } catch (error) {
      logger.error('[ProjectTable] 스켈레톤 표시 실패:', error);
      this.isReceivablesTransitioning = false;
    }
  }

  /**
   * 스켈레톤 로더 숨기기
   */
  hideSkeletonLoader() {
    // 타임아웃 정리
    if (this.skeletonTimeout) {
      clearTimeout(this.skeletonTimeout);
      this.skeletonTimeout = null;
    }

    // 오버레이 제거
    if (this.skeletonOverlay && this.skeletonOverlay.parentNode) {
      try {
        const wrapper = this.skeletonOverlay.parentNode;

        // wrapper position 복원
        if (this.originalWrapperPosition !== null) {
          wrapper.style.position = this.originalWrapperPosition;
          this.originalWrapperPosition = null;
        }

        this.skeletonOverlay.parentNode.removeChild(this.skeletonOverlay);
        this.skeletonOverlay = null;
      } catch (error) {
        logger.error('[ProjectTable] 스켈레톤 오버레이 제거 실패:', error);
      }
    }
  }

  /**
   * 미수금 필터 적용 (미수금 > 0)
   */
  applyOutstandingFilter() {
    const outstandingFilter = document.getElementById('outstandingFilter');
    if (outstandingFilter) {
      outstandingFilter.value = 'outstanding';
      // 필터 변경 이벤트 발생
      outstandingFilter.dispatchEvent(new Event('change'));
    }
  }

  /**
   * 미수금 필터 해제
   */
  clearOutstandingFilter() {
    const outstandingFilter = document.getElementById('outstandingFilter');
    if (outstandingFilter) {
      outstandingFilter.value = '';
      // 필터 변경 이벤트 발생
      outstandingFilter.dispatchEvent(new Event('change'));
    }
  }

  /**
   * 수금 모드 상태 복원
   * updateData() 후 테이블이 다시 그려질 때 수금 모드를 유지하기 위한 함수
   */
  restoreReceivablesMode() {
    if (!this.modeManager.isReceivablesMode() || !this.table) return;

    // 수금 모드 컬럼 visibility 복원 - name 기반으로
    // (workContent 컬럼은 메인 테이블에서 제거됨 — 아코디언에서 표시)
    this.table.column('startDate:name').visible(false, false);         // 공사시작
    this.table.column('outstandingNormal:name').visible(false, false); // 미수금(일반용)
    this.table.column('status:name').visible(false, false);            // 상태
    this.table.column('dataCompleteness:name').visible(false, false);  // 데이터

    this.table.column('downPayment:name').visible(true, false);         // 계약금
    this.table.column('midPayment:name').visible(true, false);          // 중도금
    this.table.column('finalPayment:name').visible(true, false);        // 잔금
    this.table.column('outstandingReceivable:name').visible(true, false); // 미수금(수금관리용)
    this.table.column('receivableNotes:name').visible(true, false);     // 수금특이사항

    this.table.columns.adjust();

    // 수금 모드 컬럼의 툴팁 재초기화 (visibility 변경으로 인해 drawCallback이 호출되지 않음)
    // requestAnimationFrame으로 DOM 업데이트 후 실행 보장
    requestAnimationFrame(() => {
      this.initializeTooltips();
      logger.debug('[ProjectTable] restoreReceivablesMode: 툴팁 재초기화 완료');
    });
  }

  /**
   * A/S 관리 모드 토글 설정 (수금 모드와 동일 구조 — 같은 테이블, 우측 컬럼만 A/S로 교체).
   * 행은 프로젝트 그대로라 아코디언/편집/필터 전부 정상. A/S 있는 프로젝트만 표시.
   */
  setupAsToggle() {
    const toggle = document.getElementById('asModeToggle');
    if (!toggle) {
      logger.warn('[ProjectTable] A/S 관리 토글 버튼을 찾을 수 없습니다.');
      return;
    }
    window.__projectTableInstance = this;

    // A/S 전용 DataTables 필터: A/S 모드일 때 A/S 있는 프로젝트만 (1회 등록)
    if (!ProjectTable._asFilterRegistered) {
      DataTable.ext.search.push((settings, dataArr, dataIndex, rowObj) => {
        const inst = window.__projectTableInstance;
        if (!inst || !inst._asModeActive) return true;
        if (inst.tableElement && settings.nTable !== inst.tableElement) return true;
        const code = rowObj && rowObj['프로젝트 코드'];
        return !!(window.asView && window.asView.byCode && window.asView.byCode[code]);
      });
      ProjectTable._asFilterRegistered = true;
    }

    toggle.addEventListener('click', async () => {
      const turningOn = !this.modeManager.isAsMode();
      await this._activateAsMode(toggle, turningOn);
    });

    // 새로고침(F5) 시 A/S 모드 복원
    try {
      if (sessionStorage.getItem('itg_as_mode') === '1') {
        setTimeout(() => this._activateAsMode(toggle, true), 0);
      }
    } catch (_) { /* noop */ }
  }

  /** A/S 모드 ON/OFF 실제 전환 */
  async _activateAsMode(toggle, on) {
    // 편집 중이면 정리 (수금 토글과 동일 정책)
    if (this.accordion && this.accordion.modeManager && this.accordion.modeManager.isEditMode()) {
      this.accordion.cleanupEditMode();
    }
    // 수금 모드와 상호배타 — A/S 켜면 수금 토글 시각적으로 끔 (컬럼은 _applyAsColumns가 정리)
    if (on) {
      const rt = document.getElementById('receivablesToggle');
      if (rt && rt.checked) { rt.checked = false; }
      try { sessionStorage.removeItem('itg_receivables_mode'); } catch (_) { /* noop */ }
    }

    const newMode = on ? TABLE_MODE.AS : TABLE_MODE.NORMAL;
    if (!this.modeManager.setTableMode(newMode)) {
      toggle.classList.toggle('active', this.modeManager.isAsMode());
      return;
    }
    this._asModeActive = on;
    toggle.classList.toggle('active', on);
    toggle.setAttribute('aria-pressed', String(on));
    try { on ? sessionStorage.setItem('itg_as_mode', '1') : sessionStorage.removeItem('itg_as_mode'); } catch (_) { /* noop */ }

    this.isReceivablesTransitioning = true;  // 빠른 draw 경로 재사용
    this.showSkeletonLoader();

    if (on) {
      try { if (window.asView && window.asView.loadMap) await window.asView.loadMap(); }
      catch (e) { logger.warn('[ProjectTable] A/S 맵 로드 실패:', e); }
    }
    this._applyAsColumns(on);
    this.table.columns.adjust();

    this.table.one('draw.dt', () => {
      this.initializeTooltips();
      this.hideSkeletonLoader();
      this.isReceivablesTransitioning = false;
    });
    if (window.modernFilters && window.modernFilters.applyFilters) {
      window.modernFilters.applyFilters(null, true);
    } else {
      this.table.draw();
    }
  }

  /** A/S 모드 컬럼 visibility 전환 (draw 없이) */
  _applyAsColumns(on) {
    if (!this.table) return;
    const asCols = ['asAcceptDate', 'asVisitDate', 'asRequest', 'asVisitor', 'asStatus', 'asResolution'];
    const normalRightCols = ['startDate', 'endDate', 'totalAmount', 'outstandingNormal', 'status', 'dataCompleteness'];
    const receivableCols = ['downPayment', 'midPayment', 'finalPayment', 'outstandingReceivable', 'receivableNotes'];
    if (on) {
      normalRightCols.forEach(n => this.table.column(`${n}:name`).visible(false, false));
      receivableCols.forEach(n => this.table.column(`${n}:name`).visible(false, false));
      asCols.forEach(n => this.table.column(`${n}:name`).visible(true, false));
    } else {
      asCols.forEach(n => this.table.column(`${n}:name`).visible(false, false));
      normalRightCols.forEach(n => this.table.column(`${n}:name`).visible(true, false));
      receivableCols.forEach(n => this.table.column(`${n}:name`).visible(false, false));
    }
  }

  /** updateData 후 A/S 모드 컬럼 visibility 재적용 (수금 restoreReceivablesMode 와 동일 역할) */
  restoreAsMode() {
    if (!this.modeManager.isAsMode() || !this.table) return;
    this._asModeActive = true;
    this._applyAsColumns(true);
    this.table.columns.adjust();
    requestAnimationFrame(() => this.initializeTooltips());
  }
}
