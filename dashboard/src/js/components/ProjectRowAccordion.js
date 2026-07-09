/**
 * Project Row Accordion 컴포넌트
 * 행 클릭시 5개 정보 카드를 표시하는 아코디언
 * 통합 편집 모드 지원 (권한별 편집 제어)
 */
import UnifiedBadgeSystem from './UnifiedBadgeSystem.js';
import { showFriendlyError, getProjectSaveError } from '../utils/userFriendlyErrors.js';
import { validateAndNormalizeDate, validateDateRange, attachDateValidator } from '../utils/dateValidator.js';
import { getSocketManager } from './SocketManager.js';
import EditState from '../utils/EditState.js';
import TomSelect from 'tom-select';
import 'tom-select/dist/css/tom-select.bootstrap5.css';
import AmountCalculator from '../utils/AmountCalculator.js';

// 🆕 ModeManager import
import { getGlobalModeManager } from '../utils/globalModeManager.js';
import { TABLE_MODE, ACCORDION_MODE } from '../constants/ViewModes.js';

// 🆕 전역 로거 import
import logger from '../utils/logger.js';

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

export default class ProjectRowAccordion {
  // 클래스 상수 (모든 인스턴스에서 공유)
  static CARD_TYPES = ['basic', 'construction', 'financial', 'collection', 'profit'];

  constructor() {
    this.accordionContainer = null;
    this.isOpen = false;
    this.currentProject = null;
    this.projectCards = {};
    this.fieldMemoButton = null; // 필드 메모 버튼 컴포넌트
    this.lastTriggeredField = null; // 자동 계산을 트리거한 필드 추적
    this.dataTable = null; // DataTable 인스턴스 저장
    this.unifiedBadgeSystem = new UnifiedBadgeSystem(); // 통합 뱃지 시스템

    // 🆕 탭 고유 ID 생성 (sessionStorage 사용, 탭마다 독립적)
    this.tabId = sessionStorage.getItem('tab_id');
    if (!this.tabId) {
      this.tabId = this.generateTabId();
      sessionStorage.setItem('tab_id', this.tabId);
      logger.debug(`[ProjectRowAccordion] 새 탭 ID 생성: ${this.tabId}`);
    } else {
      logger.debug(`[ProjectRowAccordion] 기존 탭 ID 사용: ${this.tabId}`);
    }

    // 🆕 ModeManager 초기화 (전역 인스턴스 사용)
    this.modeManager = getGlobalModeManager();

    this.eventListenersInitialized = false; // 이벤트 리스너 중복 방지 플래그
    this.tomSelectInstances = []; // Tom Select 인스턴스 저장
    this.pendingMemoChanges = {}; // 저장 대기 중인 메모 변경사항

    // itgfolder:// 프로토콜 실패 감지 (2026-07-08)
    this.bindItgfolderProtocolHandler();

    // 🆕 EditState: 단일 진실 소스 (Single Source of Truth)
    this.editState = null;

    // 🆕 아코디언 모드 변경 이벤트 리스너 (메모리 누수 방지: bound handler 저장)
    this.boundHandleAccordionModeChange = this.handleAccordionModeChange.bind(this);
    this.modeManager.on('accordionModeChanged', this.boundHandleAccordionModeChange);

    // 동시 저장 요청 큐잉 시스템
    this.saveQueue = [];  // 대기 중인 저장 요청들

    // Heartbeat 타이머 (잠금 자동 연장)
    this.heartbeatInterval = null;
    this.heartbeatFailureCount = 0;  // Heartbeat 연속 실패 횟수
    this.isSavingInProgress = false;  // 현재 저장 진행 중 여부

    // 전역 이벤트 리스너 참조 저장 (destroy 시 제거용)
    this.documentClickHandler = null;
    this.documentKeydownHandler = null;

    // beforeunload 핸들러 (편집 중 페이지 이탈 경고)
    this.beforeUnloadHandler = (e) => {
      if (this.modeManager.isEditMode()) {
        e.preventDefault();
        e.returnValue = ''; // Chrome에서 필요
        return ''; // 일부 브라우저에서 필요
      }
    };

    // 권한별 편집 가능 카드 정의
    this.CARD_EDIT_PERMISSIONS = {
      'basic': ['admin', 'editor'],          // 기본정보
      'construction': ['admin', 'editor'],   // 공사정보
      'financial': ['admin'],                // 금액정보 (admin만)
      'collection': ['admin'],               // 수금정보 (admin만)
      'profit': ['admin', 'editor'],         // 손익정보
      'document': ['admin', 'editor'],       // 문서폴더
      'notes': ['admin', 'editor']           // 수금특이사항
    };
  }

  /**
   * 탭 고유 ID 생성 (UUID v4)
   * @returns {string} UUID 문자열
   */
  generateTabId() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
      const r = Math.random() * 16 | 0;
      const v = c === 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }

  /**
   * 사용자 권한별 편집 가능한 카드 목록 반환
   */
  getEditableCards(userRole) {
    const editableCards = [];

    for (const [cardType, allowedRoles] of Object.entries(this.CARD_EDIT_PERMISSIONS)) {
      if (allowedRoles.includes(userRole)) {
        editableCards.push(cardType);
      }
    }

    return editableCards;
  }

  /**
   * 🆕 아코디언 모드 변경 이벤트 핸들러
   * @param {object} data - { oldMode, newMode, currentCombination }
   */
  handleAccordionModeChange(data) {
    logger.debug(`[ProjectRowAccordion] 아코디언 모드 변경 감지: ${data.oldMode} → ${data.newMode}`);

    if (data.newMode === ACCORDION_MODE.EDIT) {
      // 편집 모드 활성화 시 beforeunload 경고 활성화
      window.addEventListener('beforeunload', this.beforeUnloadHandler);
      logger.debug('[ProjectRowAccordion] beforeunload 경고 활성화');
    } else {
      // 뷰 모드로 전환 시 beforeunload 경고 비활성화
      window.removeEventListener('beforeunload', this.beforeUnloadHandler);
      logger.debug('[ProjectRowAccordion] beforeunload 경고 비활성화');
    }
  }

  /**
   * 사용자가 특정 카드를 편집할 수 있는지 확인
   */
  canEditCard(cardType, userRole) {
    // super_admin은 admin으로 처리
    const normalizedRole = userRole === 'super_admin' ? 'admin' : userRole;
    const allowedRoles = this.CARD_EDIT_PERMISSIONS[cardType] || [];
    return allowedRoles.includes(normalizedRole);
  }

  /**
   * 아코디언 초기화
   */
  async init() {
    this.createAccordionContainer();
    await this.initializeFieldMemoButton();
    this.setupFieldMemoEventListener(); // 메모 업데이트 이벤트 구독
    this.setupProjectLockEventListener(); // 프로젝트 잠금 변경 이벤트 구독
  }

  /**
   * FieldMemoButton 초기화
   */
  async initializeFieldMemoButton() {
    try {
      const { default: FieldMemoButton } = await import('./FieldMemoButton.js');
      this.fieldMemoButton = new FieldMemoButton();
      // FieldMemoButton 초기화 완료
    } catch (error) {
      logger.error('[ERROR] FieldMemoButton 초기화 실패:', error);
    }
  }

  /**
   * fieldMemoUpdated 이벤트 리스너 설정
   * 메모 저장 후 UI 즉시 갱신
   */
  setupFieldMemoEventListener() {
    document.addEventListener('fieldMemoUpdated', (e) => {
      const { projectCode, fieldName, memo } = e.detail;

      // 현재 열려있는 아코디언이 해당 프로젝트라면 데이터 갱신
      if (this.currentProject && this.currentProject['프로젝트 코드'] === projectCode) {
        const memoKey = `${fieldName}_메모`;
        this.currentProject[memoKey] = memo;

        // 🆕 편집 모드일 때는 pendingMemoChanges에 저장 (저장 버튼 클릭 시 한번에 저장)
        if (this.modeManager.isEditMode() && this.editState && this.editState.isActive) {
          this.pendingMemoChanges[memoKey] = memo;
          logger.debug(`[메모 추적] ${memoKey} 변경 감지:`, memo?.substring(0, 50) || '(빈 값)');
        }

        // 버튼 아이콘 즉시 업데이트
        const card = document.getElementById(`card-collection-${projectCode}`);
        if (card && this.fieldMemoButton) {
          const button = card.querySelector(`.field-memo-btn[data-field="${fieldName}"]`);
          if (button) {
            this.fieldMemoButton.updateButtonIcon(button, memo);
          }
        }

        logger.debug(`[ACCORDION] 메모 UI 갱신 완료: ${projectCode} / ${fieldName}`);
      }
    });
  }

  /**
   * 프로젝트 잠금 변경 이벤트 리스너 설정
   * WebSocket을 통해 실시간으로 잠금 상태 변경 감지
   */
  setupProjectLockEventListener() {
    try {
      const socketManager = getSocketManager();

      // 프로젝트 잠금 상태 변경 이벤트 구독
      socketManager.on('project_lock_changed', (data) => {
        logger.debug('🔔 [실시간] 프로젝트 잠금 상태 변경 이벤트 수신:', data);

        // 현재 열려있는 아코디언이 해당 프로젝트인지 확인
        if (this.currentProject && this.currentProject['프로젝트 코드'] === data.project_code) {
          logger.debug(`🔄 [실시간] 현재 프로젝트의 잠금 상태가 변경됨: ${data.project_code}`);

          // 편집 버튼 상태 즉시 업데이트
          this.updateEditButtonLockStatus(data.project_code);
        }
      });

      logger.debug('✅ [WebSocket] 프로젝트 잠금 이벤트 리스너 설정 완료');
    } catch (error) {
      logger.warn('⚠️ [WebSocket] SocketManager 초기화 실패 (WebSocket 기능 비활성):', error);
    }
  }

  /**
   * 아코디언 컨테이너 생성
   */
  createAccordionContainer() {
    this.accordionContainer = document.createElement('div');
    this.accordionContainer.className = 'project-accordion-container';
    // 초기 숨김 상태는 CSS에서 처리됨
  }

  /**
   * 테이블 행에 아코디언 기능 추가
   */
  attachToTable(tableElement, dataTable) {
    if (!tableElement) {
      logger.error('[ERROR] [아코디언] attachToTable: 테이블 요소가 없습니다');
      return;
    }

    if (!dataTable) {
      logger.error('[ERROR] [아코디언] attachToTable: DataTable 인스턴스가 없습니다');
      return;
    }

    // DataTable 인스턴스 저장
    this.dataTable = dataTable;

    // 이벤트 리스너가 이미 등록되어 있으면 스킵 (중복 방지)
    if (this.eventListenersInitialized) {      return;
    }

    // DataTable의 row 클릭 이벤트 바인딩 (아코디언 행 제외)
    tableElement.addEventListener('click', (e) => {
      // 이벤트 위임으로 tbody tr:not(.accordion-row) 체크
      const clickedRow = e.target.closest('tbody tr');
      if (!clickedRow || clickedRow.classList.contains('accordion-row')) return;

      // 액션 버튼 클릭 시 아코디언 토글 방지
      if (e.target.closest('.action-buttons')) {
        return;
      }

      const row = clickedRow;
      const rowData = this.dataTable.row(row).data();

      if (rowData) {
        // 먼저 행 선택 표시 (아코디언과 별개로)
        this.selectTableRow(row);

        this.toggleAccordion(row, rowData);
      } else {
      }
    });

    // 카드별 이벤트 바인딩 (편집/저장/취소 버튼 등)
    this.bindCardEvents();

    // ESC 키 및 외부 클릭 이벤트 바인딩
    this.bindEvents();

    // 이벤트 리스너 등록 완료 플래그 설정
    this.eventListenersInitialized = true;
  }

  /**
   * 테이블 행 선택 표시 (아코디언과 별개)
   */
  selectTableRow(tableRow) {
    // 이전 선택된 행들에서 table-active 제거 (아코디언이 열려있지 않은 행만)
    document.querySelectorAll('tbody tr.table-active').forEach(row => {
      // 현재 아코디언이 열린 행은 유지
      if (this.isOpen && row.querySelector('.accordion-container')) {
        return; // 아코디언 열린 행은 그대로 유지
      }
      row.classList.remove('table-active');
    });

    // 클릭한 행에 선택 표시
    tableRow.classList.add('table-active');
  }

  /**
   * 편집 모드 종료 및 잠금 해제
   * (아코디언 열림/닫힘 상태는 별도 처리)
   */
  cleanupEditMode() {
    if (this.modeManager.isEditMode() && this.currentProject) {
      const currentProjectCode = this.currentProject['프로젝트 코드'];
      logger.debug(`🧹 [편집 모드] 정리 - ${currentProjectCode} 잠금 해제`);

      // 편집 모드 종료 (잠금 해제 포함)
      this.disableUnifiedEditMode(currentProjectCode);
    }
  }

  /**
   * 아코디언 토글
   */
  toggleAccordion(tableRow, projectData) {
    const projectCode = projectData['프로젝트 코드'];

    // 같은 행 클릭시 닫기
    if (this.isOpen && this.currentProject && this.currentProject['프로젝트 코드'] === projectCode) {
      this.closeAccordion();
      return;
    }

    // 다른 행이 열려있다면 직접 전환 (닫기 없이)
    if (this.isOpen) {
      // 🆕 편집 중이면 변경사항 확인 경고
      if (this.modeManager.isEditMode() && this.editState && this.editState.isActive) {
        const changes = this.editState.collectChanges();

        if (Object.keys(changes).length > 0) {
          const currentProjectName = this.currentProject?.['프로젝트명'] || '현재 프로젝트';
          const confirmed = confirm(
            `"${currentProjectName}"에 저장하지 않은 변경사항이 있습니다.\n\n` +
            `다른 프로젝트로 이동하시겠습니까?\n\n` +
            `• 확인: 변경사항을 버리고 이동\n` +
            `• 취소: 현재 프로젝트 계속 편집`
          );
          if (!confirmed) {
            return; // 이동 취소
          }
        }
      }

      // 편집 모드 정리
      this.cleanupEditMode();

      // 기존 아코디언 제거만 하고 상태는 유지
      const existingAccordion = document.querySelector('.accordion-row');
      if (existingAccordion) {
        existingAccordion.remove();
      }
      // 기존 테이블 행 활성화 해제
      document.querySelectorAll('tbody tr').forEach(row => row.classList.remove('table-active'));
    }

    // 새 아코디언 열기 (isOpen 상태 연속성 유지)
    this.openAccordion(tableRow, projectData);
  }

  /**
   * 아코디언 열기
   */
  openAccordion(tableRow, projectData) {
    // DOM 검증: 유효한 테이블 행인지 확인
    if (!tableRow || !tableRow.parentElement) {
      logger.warn('[아코디언] 유효하지 않은 테이블 행, 아코디언을 열 수 없습니다');
      return;
    }

    // 데이터 검증: 필수 데이터가 있는지 확인
    if (!projectData || !projectData['프로젝트 코드']) {
      logger.error('[아코디언] 프로젝트 데이터가 유효하지 않습니다:', projectData);
      // 페이지 액션 결과 (아코디언 표 관련) → 페이지 헤더
      if (window.showPageAlert) window.showPageAlert('프로젝트 정보를 불러올 수 없습니다', 'error');
      else this.showToast('프로젝트 정보를 불러올 수 없습니다', 'error');
      return;
    }

    try {
      this.currentProject = projectData; // 전체 프로젝트 데이터 저장
      this.originalProjectCode = projectData['프로젝트 코드']; // 원본 프로젝트 코드 저장

      // 행 번호는 현재 프로젝트 코드에서 추출 (G2851-YG에서 2851 추출)
      const projectCode = projectData['프로젝트 코드'];
      const match = projectCode.match(/([GPR])(\d+)/);
      if (match) {
        this.currentRowNumber = parseInt(match[2], 10);
        // 행번호 추출 완료
      } else {
        logger.warn(`[행번호] 프로젝트 코드에서 행번호 추출 실패: ${projectCode}`);
      }

      const wasAlreadyOpen = this.isOpen; // 이미 열려있었는지 기록

      const isCancelledInitial = this.isProjectCancelled(projectData);

      // 아코디언 내용 생성
      this.renderAccordionContent(projectData, isCancelledInitial);

      this.toggleCancelledStyles(projectCode, isCancelledInitial);

      // 테이블 행 다음에 아코디언 삽입
      const nextRow = tableRow.nextElementSibling;
      if (nextRow && nextRow.classList.contains('accordion-row')) {
        nextRow.remove();
      }

      // 새 아코디언 행 생성
      const accordionRow = document.createElement('tr');
      accordionRow.className = 'accordion-row';
      accordionRow.innerHTML = `
        <td colspan="100%">
          <div class="accordion-content"></div>
        </td>
      `;

      accordionRow.querySelector('.accordion-content').appendChild(this.accordionContainer);
      tableRow.insertAdjacentElement('afterend', accordionRow);

      // 아코디언 표시 - CSS 클래스로 제어
      this.accordionContainer.classList.remove('accordion-slide-down', 'accordion-slide-up');
      this.accordionContainer.classList.add('show');

      // 애니메이션 효과 (취소 상태는 건너뜀)
      if (!isCancelledInitial) {
        this.accordionContainer.style.display = 'none';
        void this.accordionContainer.offsetHeight; // 강제 리플로우
        this.accordionContainer.style.display = 'block';
        this.accordionContainer.classList.add('accordion-slide-down');
      } else {
        this.accordionContainer.style.display = 'block';
      }

      // 테이블 행 활성화 표시
      document.querySelectorAll('tbody tr').forEach(row => row.classList.remove('table-active'));
      tableRow.classList.add('table-active');

      // 이미 열려있던 경우 즉시 상태 설정, 새로 여는 경우 DOM 완성 후 설정
      if (wasAlreadyOpen) {
        this.isOpen = true; // 즉시 설정하여 상태 연속성 유지
      }

      // 계산식 실행 및 데이터 검증 (브라우저 렌더링 완료 후)
      requestAnimationFrame(() => {
        // DOM 요소 준비 확인 후 계산 실행
        const projectCode = projectData['프로젝트 코드'];
        const financialCard = document.getElementById(`card-financial-${projectCode}`);

        if (financialCard) {
          this.calculateAmountFields(projectCode);
          this.calculateProfitFields(projectCode);
          // 검증 UI 완전 비활성화 - 일반 텍스트로 깔끔하게 표시
          // this.validateCalculatedFields(projectCode, projectData);

          // 기존 검증 아이콘 완전 제거
          this.removeAllValidationIcons();

          // 아코디언 내부의 모든 툴팁 초기화 (메모 툴팁 포함)
          this.initializeAccordionTooltips();

          // 프로젝트 잠금 상태 체크 및 버튼 업데이트
          this.updateEditButtonLockStatus(projectCode);

          // 새로 여는 경우에만 DOM 완성 후 isOpen 상태 설정
          if (!wasAlreadyOpen) {
            this.isOpen = true;
          }
        } else {
          // DOM이 준비되지 않은 경우 fallback으로 작은 지연 후 재시도
          setTimeout(() => {
            this.calculateAmountFields(projectCode);
            this.calculateProfitFields(projectCode);
            this.removeAllValidationIcons();
            this.initializeAccordionTooltips();

            // 프로젝트 잠금 상태 체크 및 버튼 업데이트
            this.updateEditButtonLockStatus(projectCode);

            if (!wasAlreadyOpen) {
              this.isOpen = true;
            }
          }, 50);
        }
      });

    } catch (error) {
      // 오류 처리: 렌더링 중 오류 발생 시 안전하게 복구
      logger.error('[아코디언] 렌더링 오류:', error);
      logger.error('[아코디언] 문제 데이터:', projectData);

      // 상태 복구
      this.isOpen = false;
      this.currentProject = null;

      // 사용자에게 알림 — 아코디언 표 관련 → 페이지 헤더
      if (window.showPageAlert) window.showPageAlert('아코디언을 열 수 없습니다. 데이터를 확인해주세요.', 'error');
      else this.showToast('아코디언을 열 수 없습니다. 데이터를 확인해주세요.', 'error');

      // 오류 발생 시 생성된 DOM 정리
      document.querySelectorAll('.accordion-row').forEach(row => row.remove());
      document.querySelectorAll('tbody tr').forEach(row => row.classList.remove('table-active'));
    }
  }

  /**
   * 아코디언 닫기
   */
  closeAccordion() {
    if (!this.isOpen) return;

    // 편집 모드 정리
    this.cleanupEditMode();

    // 애니메이션으로 닫기
    this.accordionContainer.classList.add('accordion-slide-up');

    // 정리 작업을 한 번만 실행하도록 플래그 관리
    let cleanupExecuted = false;

    // 애니메이션 완료 후 콜백
    const handleAnimationEnd = () => {
      if (cleanupExecuted) return; // 중복 실행 방지
      cleanupExecuted = true;

      document.querySelectorAll('.accordion-row').forEach(row => row.remove());
      this.accordionContainer.classList.remove('show');
      this.accordionContainer.classList.remove('accordion-slide-up');
      this.accordionContainer.removeEventListener('animationend', handleAnimationEnd);

      // pending된 테이블 업데이트가 있으면 실행
      if (this.pendingTableUpdate && this.dataTable) {
        this.dataTable.draw(false);
        this.pendingTableUpdate = false;
        logger.debug('✅ [DataTable] pending 업데이트 실행 완료 (아코디언 닫힘)');
      }
    };

    this.accordionContainer.addEventListener('animationend', handleAnimationEnd, { once: true });

    // 안전장치: animationend가 발생하지 않을 경우를 대비한 timeout
    // 브라우저가 바쁘거나 백그라운드 탭일 때 애니메이션 이벤트가 누락될 수 있음
    setTimeout(() => {
      if (!cleanupExecuted) {
        logger.warn('[아코디언] animationend 이벤트 누락, 강제 정리 실행');
        handleAnimationEnd();
      }
    }, 1000);

    // 테이블 행 활성화 해제
    document.querySelectorAll('tbody tr').forEach(row => row.classList.remove('table-active'));

    this.isOpen = false;
    this.currentProject = null;
  }

  /**
   * 아코디언 내용 렌더링
   */
  renderAccordionContent(projectData, isCancelled = false) {
    const projectCode = projectData['프로젝트 코드'];
    const shellClass = isCancelled ? 'accordion-shell project-cancelled' : 'accordion-shell';

    this.accordionContainer.innerHTML = `
      <div class="${shellClass}" data-project-code="${projectCode}">
        <div class="row-details">
          <div class="card border-0 shadow-sm">
            <div class="card-body p-4">

              <!-- 프로젝트 정보 타이틀 -->
              <div class="row mb-3">
                <div class="col-12">
                  <div class="project-title-section" data-project-code="${projectCode}">
                    <div class="project-title-flex">
                      <div class="project-title-info" id="header-${projectCode}">
                        ${this.generateHeaderFields(projectCode, projectData, false)}
                      </div>
                      <div class="project-actions-container">
                        ${this.generateUnifiedEditButtons(projectCode, projectData)}
                        ${this.generateCancelResumeButton(projectCode, projectData)}
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <!-- 5개 정보 카드를 가로로 배치 -->
              <div class="all-info-container">
                <div class="row">

                  <!-- 기본정보 카드 -->
                  <div class="col-xxl-2 col-xl-3 col-lg-4 col-md-6 info-card-column">
                    ${this.generateBasicInfoCard(projectCode, projectData)}
                  </div>

                  <!-- 공사정보 카드 -->
                  <div class="col-xxl-2 col-xl-3 col-lg-4 col-md-6 info-card-column">
                    ${this.generateConstructionInfoCard(projectCode, projectData)}
                  </div>

                  <!-- 금액정보 카드 -->
                  <div class="col-xxl-2 col-xl-3 col-lg-4 col-md-6 info-card-column">
                    ${this.generateFinancialInfoCard(projectCode, projectData)}
                  </div>

                  <!-- 수금정보 카드 -->
                  <div class="col-xxl-2 col-xl-3 col-lg-4 col-md-6 info-card-column">
                    ${this.generateCollectionInfoCard(projectCode, projectData)}
                  </div>

                  <!-- 손익정보 카드 -->
                  <div class="col-xxl-2 col-xl-3 col-lg-4 col-md-6 info-card-column">
                    ${this.generateProfitInfoCard(projectCode, projectData)}
                  </div>

                </div>
              </div>

              <!-- 하단 문서 정보 섹션 -->
              <div class="mt-3">
                <div class="row">
                  <div class="col-md-6">
                    ${this.generateDocumentSection(projectData)}
                  </div>
                  <div class="col-md-6">
                    ${this.generateCollectionNotesSection(projectData)}
                  </div>
                </div>
              </div>

            </div>
          </div>
        </div>
      </div>
    `;

    // 실시간 계산 이벤트 바인딩 (프로젝트별로 매번 새로 바인딩 필요)
    // 총액1 → 총액2, 결제금액 → 미수금 등의 계산 이벤트
    this.bindCalculationEvents(projectCode);
  }

  /**
   * 시공자 필드 포맷팅 (뷰모드): 3명 이상일 때 3번째부터 줄바꿈
   */
  formatConstructorField(constructorValue) {
    if (!constructorValue || constructorValue === '-') {
      return '-';
    }

    // 쉼표로 분리
    const names = constructorValue.split(',').map(n => n.trim()).filter(n => n);

    // 3명 이상이면 2명까지는 한 줄, 3번째부터 줄바꿈
    if (names.length >= 3) {
      const firstTwo = names.slice(0, 2).join(', ');
      const remaining = names.slice(2).join(', ');
      return `${firstTwo}, <br>${remaining}`;
    }

    // 2명 이하는 그대로
    return names.join(', ');
  }

  /**
   * 기본정보 카드 생성
   */
  generateBasicInfoCard(projectCode, rowData) {
    return `
      <div class="info-card compact-card" id="card-basic-${projectCode}">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h6 class="text-primary">
            <i class="fas fa-building me-1"></i>기본정보
          </h6>
          <div class="card-edit-buttons">
            ${this.generateEditButtons(projectCode, 'basic')}
          </div>
        </div>
        <div class="card-grid">
          <div class="compact-item">
            <small>사업자</small>
            <div class="editable-value" data-field="사업자">${rowData['사업자'] || '-'}</div>
          </div>
          <div class="compact-item">
            <small>현장담당자</small>
            <div class="editable-value" data-field="발주처 담당자">${rowData['발주처 담당자'] || '-'}</div>
          </div>
          <div class="compact-item">
            <small>도급구분</small>
            <div class="editable-value" data-field="도급 구분">${rowData['도급 구분'] || '-'}</div>
          </div>
          <div class="compact-item">
            <small>연락처</small>
            <div class="editable-value" data-field="발주처 연락처">${rowData['발주처 연락처'] || '-'}</div>
          </div>
          <div class="compact-item">
            <small>시공자</small>
            <div class="editable-value" data-field="시공자">${this.formatConstructorField(rowData['시공자'])}</div>
          </div>
          <div class="compact-item">
            <small>이메일</small>
            <div class="editable-value" data-field="발주처 이메일">${rowData['발주처 이메일'] || '-'}</div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * 공사정보 카드 생성
   */
  generateConstructionInfoCard(projectCode, rowData) {
    return `
      <div class="info-card compact-card" id="card-construction-${projectCode}">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h6 class="text-success">
            <i class="fas fa-hammer me-1"></i>공사정보
          </h6>
          <div class="card-edit-buttons">
            ${this.generateEditButtons(projectCode, 'construction')}
          </div>
        </div>
        <div class="card-grid">
          <div class="compact-item">
            <small>공사구분</small>
            <div class="editable-value" data-field="공사 구분">${rowData['공사 구분'] || '-'}</div>
          </div>
          <div class="compact-item">
            <small>공사확정</small>
            <div>${rowData['공사 확정'] ? this.formatDate(rowData['공사 확정']) : '-'}</div>
          </div>
          <div class="compact-item">
            <small>기계분류</small>
            <div class="editable-value" data-field="기계 분류">${rowData['기계 분류'] || '-'}</div>
          </div>
          <div class="compact-item">
            <small>공사시작</small>
            <div class="editable-value" data-field="공사 시작">${rowData['공사 시작'] ? this.formatDate(rowData['공사 시작']) : '-'}</div>
          </div>
          <div class="compact-item">
            <small>브랜드</small>
            <div class="editable-value" data-field="브랜드">${rowData['브랜드'] || '-'}</div>
          </div>
          <div class="compact-item">
            <small>공사종료</small>
            <div class="editable-value" data-field="공사 종료">${rowData['공사 종료'] ? this.formatDate(rowData['공사 종료']) : '-'}</div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * 금액정보 카드 생성
   */
  generateFinancialInfoCard(projectCode, rowData) {
    // 총액2는 API에서 받은 값을 그대로 사용 (재계산하지 않음)
    const totalAmount = rowData['총액 2'] || rowData['총액2'] || rowData['S'] || rowData['총액'] || 0;
    const total2Value = parseFloat(totalAmount) || 0;
    // 총액2 값 설정

    return `
      <div class="info-card compact-card" id="card-financial-${projectCode}" data-card-type="financial">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h6 class="text-warning">
            <i class="fas fa-won-sign me-1"></i>금액정보
          </h6>
          <div class="card-edit-buttons">
            ${this.generateEditButtons(projectCode, 'financial')}
          </div>
        </div>
        <div class="card-grid">
          <div class="compact-item">
            <small>총액1</small>
            <div class="editable-value" data-field="총액 1">${rowData['총액 1'] ? this.formatCurrency(rowData['총액 1']) : '-'}</div>
          </div>
          <div class="compact-item">
            <small>부가세</small>
            <div class="editable-value" data-field="부가세" data-original-value="${this.unifiedBadgeSystem.vatBadge.normalizeStatus(rowData['부가세'])}">${this.unifiedBadgeSystem.createBadge('vat', rowData['부가세'])}</div>
          </div>
          <div class="compact-item">
            <small>총액2</small>
            <div class="editable-value calculated-field" data-field="총액 2" data-original-value="${total2Value}">${total2Value > 0 ? this.formatCurrency(total2Value) : '-'}</div>
          </div>
          <div class="compact-item">
            <small>계산서</small>
            <div class="editable-value" data-field="계산서">${this.formatBillStatus(rowData)}</div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * 계산서 정보 파싱
   * 형식: "일반-계약금, N입금-중도금, 카드-잔금"
   */
  parseBillStatus(billValue) {
    const result = {
      stages: {}  // { '계약금': '일반', '중도금': 'N입금', ... }
    };

    if (!billValue || billValue === '-' || billValue === '미발행') {
      logger.debug('[parseBillStatus] 빈 값 또는 미발행:', billValue);
      return result;
    }

    logger.debug('[parseBillStatus] 파싱 시작:', billValue);

    // 콤마로 분리된 각 항목 처리
    const items = billValue.split(',').map(s => s.trim());

    items.forEach(item => {
      // "카테고리-단계" 형식 파싱
      const parts = item.split('-');
      if (parts.length === 2) {
        const category = parts[0].trim();
        const stage = parts[1].trim();
        result.stages[stage] = category;
        logger.debug(`[parseBillStatus] 매핑: ${stage} => ${category}`);
      } else {
        logger.warn(`[parseBillStatus] 잘못된 형식: "${item}"`);
      }
    });

    logger.debug('[parseBillStatus] 파싱 결과:', result);
    return result;
  }

  /**
   * 계산서 카테고리에 따른 아이콘 반환
   */
  getBillStatusIcon(category) {
    const iconMap = {
      '일반': { icon: 'fas fa-receipt text-primary', label: '세금계산서 발행' },
      'N입금': { icon: 'fas fa-sack-dollar text-danger', label: '현금거래 (계산서 미발행)' },
      '카드': { icon: 'fas fa-credit-card text-info', label: '카드결제' }
    };

    const info = iconMap[category];
    if (!info) return '';

    return ` <span class="memo-tooltip-trigger bill-status" data-bs-toggle="tooltip" data-bs-title="${info.label}" aria-label="${info.label}"><span class="bill-icon-spacer"></span><i class="${info.icon}"></i></span>`;
  }

  /**
   * 수금정보 카드 생성
   */
  generateCollectionInfoCard(projectCode, rowData) {
    // 메모 데이터 및 금액 추출
    const memos = {
      계약금: rowData['계약금_메모'] || null,
      중도금: rowData['중도금_메모'] || null,
      잔금: rowData['잔금_메모'] || null
    };

    const amounts = {
      계약금: rowData['계약금'] || 0,
      중도금: rowData['중도금'] || 0,
      잔금: rowData['잔금'] || 0
    };

    // 계산서 정보 파싱
    const billStatus = this.parseBillStatus(rowData['계산서']);

    // 메모 버튼 HTML 생성 (FieldMemoButton이 초기화되어 있을 때만)
    // UX 개선: 아코디언 내부에서는 항상 메모 버튼 표시 (금액 입력 중일 수 있음)
    const getMemoButton = (fieldName) => {
      if (!this.fieldMemoButton) return '';
      return this.fieldMemoButton.createButton(
        fieldName,
        projectCode,
        memos[fieldName],
        amounts[fieldName],
        true  // 아코디언 내부에서는 항상 편집 가능 (금액 입력 중일 수 있음)
      );
    };

    // 카드 생성 후 메모 버튼 이벤트 연결 (setTimeout으로 DOM 렌더링 후 실행)
    setTimeout(() => {
      const card = document.getElementById(`card-collection-${projectCode}`);
      if (card && this.fieldMemoButton) {
        // Getter 함수로 최신 메모 데이터 제공 (클로저 문제 해결)
        this.fieldMemoButton.attachButtonListeners(card, () => {
          return {
            '계약금_메모': this.currentProject?.['계약금_메모'] || null,
            '중도금_메모': this.currentProject?.['중도금_메모'] || null,
            '잔금_메모': this.currentProject?.['잔금_메모'] || null
          };
        });
      }
    }, 50);

    // 메모 필드 HTML 생성 (3단계 시각화 + 계산서 아이콘)
    // 1. 금액 없음 → 아이콘 없음
    // 2. 금액 있음 + 메모 없음 → 빈 아이콘 + 주황색 경고
    // 3. 금액 있음 + 메모 있음 → 채워진 아이콘 + 초록색
    // 4. 계산서 발행 → 해당 단계에 계산서 아이콘 표시
    const createMemoField = (fieldName, memoKey) => {
      const amount = AmountCalculator.safeParseCurrency(rowData[fieldName] || 0);
      const memo = memos[memoKey];
      const formattedAmount = this.formatCurrency(amount);

      let valueMarkup;

      if (amount > 0) {
        const memoStatus = checkMemoStatus(memo);
        const tooltipText = memoStatus.isEmpty ? '메모를 작성해주세요' : this.escapeHTML(memo);
        const iconType = memoStatus.isEmpty ? 'far' : 'fas';
        const stateClass = memoStatus.isEmpty ? 'no-memo' : 'has-memo';
        const iconColorClass = 'text-success';  // 금액이 있으면 항상 초록색

        valueMarkup = `
          <span class="memo-value-wrapper">
            <span class="memo-amount-text">${formattedAmount}</span>
            <span class="memo-tooltip-trigger ${stateClass}" data-bs-toggle="tooltip" data-bs-title="${tooltipText}" aria-label="${memoStatus.isEmpty ? '메모를 작성해주세요' : '메모 보기'}">
              <i class="${iconType} fa-sticky-note ${iconColorClass}"></i>
            </span>
          </span>
        `;
      } else {
        valueMarkup = formattedAmount;
      }

      return `
        <div class="editable-value" data-field="${fieldName}">
          ${valueMarkup}
        </div>
        ${getMemoButton(memoKey)}
      `;
    };

    return `
      <div class="info-card compact-card" id="card-collection-${projectCode}" data-card-type="collection">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h6 class="text-danger">
            <i class="fas fa-credit-card me-1"></i>수금정보
          </h6>
          <div class="card-edit-buttons">
            ${this.generateEditButtons(projectCode, 'collection')}
          </div>
        </div>
        <div class="card-grid">
          <div class="compact-item">
            <small>계약금</small>
            <div class="editable-value-wrapper">
              ${createMemoField('계약금', '계약금')}
            </div>
          </div>
          <div class="compact-item">
            <small>미수금</small>
            <div class="editable-value calculated-field" data-field="미수금" data-original-value="${rowData['미수금'] || 0}">${(() => {
              const outstanding = rowData['미수금'] || 0;
              const numValue = parseFloat(outstanding) || 0;
              const colorClass = numValue === 0 ? 'text-success fw-semibold' : 'text-danger fw-semibold';
              const displayValue = this.formatCurrency(numValue);
              return `<span class="${colorClass}">${displayValue}</span>`;
            })()}</div>
          </div>
          <div class="compact-item">
            <small>중도금</small>
            <div class="editable-value-wrapper">
              ${createMemoField('중도금', '중도금')}
            </div>
          </div>
          <div class="compact-item">
            <small>수금확인</small>
            <div class="editable-value" data-field="수금 확인" data-original-value="${this.unifiedBadgeSystem.collectionBadge.normalizeStatus(rowData['수금 확인'])}">${this.formatCollectionStatus(rowData['수금 확인'])}</div>
          </div>
          <div class="compact-item">
            <small>잔금</small>
            <div class="editable-value-wrapper">
              ${createMemoField('잔금', '잔금')}
            </div>
          </div>
          <div class="compact-item">
            <small>수금날짜</small>
            <div class="editable-value" data-field="수금 날짜">${rowData['수금 날짜'] ? this.formatDate(rowData['수금 날짜']) : '-'}</div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * 손익정보 카드 생성
   */
  generateProfitInfoCard(projectCode, rowData) {
    return `
      <div class="info-card compact-card" id="card-profit-${projectCode}">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h6 class="text-info">
            <i class="fas fa-chart-line me-1"></i>손익정보
          </h6>
          <div class="card-edit-buttons">
            ${this.generateEditButtons(projectCode, 'profit')}
          </div>
        </div>
        <div class="card-grid">
          <div class="compact-item">
            <small>제품대</small>
            <div class="editable-value" data-field="제품대">${rowData['제품대'] ? this.formatCurrency(rowData['제품대']) : '-'}</div>
          </div>
          <div class="compact-item">
            <small>도급비</small>
            <div class="editable-value" data-field="도급비">${rowData['도급비'] ? this.formatCurrency(rowData['도급비']) : '-'}</div>
          </div>
          <div class="compact-item">
            <small>자재비</small>
            <div class="editable-value" data-field="자재비">${rowData['자재비'] ? this.formatCurrency(rowData['자재비']) : '-'}</div>
          </div>
          <div class="compact-item">
            <small>기타비</small>
            <div class="editable-value" data-field="기타비">${rowData['기타비'] ? this.formatCurrency(rowData['기타비']) : '-'}</div>
          </div>
          <div class="compact-item">
            <small>마진율</small>
            <div class="editable-value calculated-field" data-field="마진율" data-original-value="${rowData['마진율'] || 0}">${(() => {
              const marginRate = rowData['마진율'];
              // API에서 마진율 값이 있으면 사용, 없으면 계산
              if (marginRate !== undefined && marginRate !== null && marginRate !== '') {
                const numValue = parseFloat(marginRate) || 0;
                const colorClass = numValue > 0 ? 'text-success fw-semibold' :
                                   numValue < 0 ? 'text-danger fw-semibold' : 'text-muted';
                return `<span class="${colorClass}">${numValue.toFixed(1)}%</span>`;
              }
              // Fallback: 클라이언트에서 계산
              return this.calculateMarginRate(rowData);
            })()}</div>
          </div>
          <div class="compact-item">
            <small>순익</small>
            <div class="editable-value calculated-field" data-field="순익" data-original-value="${rowData['순익'] || 0}">${(() => {
              const netProfit = rowData['순익'];
              if (netProfit === undefined || netProfit === null) {
                return '-';
              }
              const numValue = parseFloat(netProfit) || 0;
              const colorClass = numValue > 0 ? 'text-success fw-semibold' :
                                numValue < 0 ? 'text-danger fw-semibold' : 'text-muted';
              const displayValue = numValue >= 0 ? this.formatCurrency(numValue) : `-${this.formatCurrency(Math.abs(numValue))}`;
              return `<span class="${colorClass}">${displayValue}</span>`;
            })()}</div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * 통합 편집 버튼 생성 (아코디언 헤더용)
   */
  generateUnifiedEditButtons(projectCode, rowData) {
    const userRole = window.userRole || 'viewer';

    // viewer 권한은 편집 버튼 자체를 숨김
    if (userRole === 'viewer') {
      return '';
    }

    // 공사 취소 상태일 때는 편집 버튼 숨김
    const collectionNotes = rowData?.['수금 관련 특이사항'] || '';
    if (collectionNotes && /공사\s*취소/.test(collectionNotes)) {
      return '';
    }

    return `
      <div class="unified-edit-buttons me-2" data-project-code="${projectCode}">
        <button class="btn btn-outline-secondary btn-sm unified-edit-btn"
                data-project-code="${projectCode}"
                title="편집">
          <i class="fas fa-edit"></i><span>편집</span>
        </button>
        <button class="btn btn-outline-success btn-sm unified-save-btn d-none"
                data-project-code="${projectCode}"
                title="저장">
          <i class="fas fa-check"></i><span>저장</span>
        </button>
        <button class="btn btn-outline-danger btn-sm unified-cancel-btn d-none"
                data-project-code="${projectCode}"
                title="취소">
          <i class="fas fa-times"></i><span>취소</span>
        </button>
      </div>
    `;
  }

  /**
   * 카드별 편집 버튼 생성 (통합 편집 모드로 전환하여 제거)
   */
  generateEditButtons(projectCode, cardType) {
    // 통합 편집 모드로 전환하므로 카드별 버튼은 비활성화
    return '';
  }

  /**
   * 취소/재개 버튼 생성
   */
  generateCancelResumeButton(projectCode, rowData) {
    const collectionNotes = rowData['수금 관련 특이사항'] || '';
    const isCancelled = collectionNotes && /공사\s*취소/.test(collectionNotes);

    if (isCancelled) {
      return `
        <button type="button" class="btn btn-outline-secondary btn-sm construction-action-btn construction-resume-btn resume-construction-btn"
                data-project-code="${projectCode}"
                title="공사 재개">
          <i class="fas fa-play"></i><span>공사 재개</span>
        </button>
      `;
    } else {
      return `
        <button type="button" class="btn btn-outline-secondary btn-sm construction-action-btn construction-cancel-btn cancel-construction-btn"
                data-project-code="${projectCode}"
                title="공사 취소">
          <i class="fas fa-times"></i><span>공사 취소</span>
        </button>
      `;
    }
  }

  /**
   * 문서 섹션 생성 (통합 편집 모드 적용)
   */
  /**
   * 폴더 링크 HTML 렌더링.
   * - 폴더 ID 패턴이면 itgfolder:// 커스텀 프로토콜 (클라이언트 Explorer 직접 열기)
   * - 그 외 경로면 기존 API 경유
   */
  renderFolderLink(localPath, projectCode) {
    if (!localPath) return '폴더 경로가 설정되지 않았습니다.';
    const isFolderId = /^[a-zA-Z0-9_-]{20,}$/.test(localPath);
    if (isFolderId) {
      // 2026-07-08 프로토콜 실패 감지: 클릭 후 페이지가 여전히 활성 상태면 프로토콜 미설치로 판단, 설치 안내.
      return `<a href="itgfolder://${localPath}" class="text-decoration-none itgfolder-link" data-folder-id="${localPath}" style="color: #0d6efd;" title="탐색기에서 열기 (프로토콜 필요)">${localPath}</a>`;
    }
    return `<a href="#" class="text-decoration-none folder-open-link" data-project-code="${projectCode}" style="color: #0d6efd;">${localPath}</a>`;
  }

  /**
   * itgfolder:// 링크 클릭 감지 (2026-07-08).
   * 클릭 후 1.5초 뒤에도 탭이 활성 상태이면 프로토콜 미설치 → 설치 안내 알림.
   * ProjectRowAccordion 초기화 시 한 번 등록.
   */
  bindItgfolderProtocolHandler() {
    // 전역 플래그로 중복 등록 방지 (여러 컴포넌트 인스턴스 대응)
    if (window._itgfolderHandlerBound) return;
    window._itgfolderHandlerBound = true;

    document.addEventListener('click', (e) => {
      const link = e.target.closest('.itgfolder-link');
      if (!link) return;
      const folderId = link.dataset.folderId;
      if (!folderId) return;

      // 이미 프로토콜 정상 작동 확인된 브라우저는 매번 안내하지 않음
      if (localStorage.getItem('itg_folder_protocol_ok') === '1') return;

      // 프로토콜이 실행되면 탐색기(외부 앱)로 focus가 이동 → window blur 발생.
      // 1.5초 안에 blur 없으면 프로토콜 미설치 유력.
      let focusLost = false;
      const onBlur = () => { focusLost = true; };
      window.addEventListener('blur', onBlur, { once: true });

      setTimeout(() => {
        window.removeEventListener('blur', onBlur);
        if (focusLost) {
          // 성공: 다시 안내 안 하도록 저장
          localStorage.setItem('itg_folder_protocol_ok', '1');
          return;
        }
        // 미설치 유력
        const proceed = confirm(
          '폴더가 열리지 않았나요?\n\n' +
          '탐색기 프로토콜 미설치일 수 있습니다.\n' +
          '설치 방법:\n' +
          '  1) 회사에서 배포한 install-itg-folder.bat 파일을 실행\n' +
          '  2) 브라우저를 완전히 종료 후 재실행\n' +
          '  3) 다시 폴더 링크 클릭\n\n' +
          '설치 가이드가 필요하면 관리자(kiko@itg-aircon.com)에게 문의하세요.\n\n' +
          '확인을 누르면 폴더 ID를 클립보드에 복사합니다 (Drive에서 직접 열기용).'
        );
        if (proceed && navigator.clipboard) {
          navigator.clipboard.writeText(folderId).catch(() => {});
        }
      }, 1500);
    });
  }

  generateDocumentSection(rowData) {
    const localPath = (rowData['견적서 및 계약서 폴더 경로'] || rowData['AK'] || '').trim();
    const projectCode = rowData['프로젝트 코드'];

    return `
      <div class="legacy-card mt-3 document-card ${localPath ? '' : 'document-card-empty'}">
        <div class="legacy-card-row">
          <div class="legacy-card-main">
            <span class="legacy-card-label">
              <i class="fab fa-google-drive me-2" style="color: #4285f4;"></i>
              문서 폴더
            </span>
            <div class="editable-value" data-field="견적서 및 계약서 폴더 경로" data-original-value="${localPath}">
              ${this.renderFolderLink(localPath, projectCode)}
            </div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * 수금 특이사항 섹션 생성 (통합 편집 모드 적용)
   */
  generateCollectionNotesSection(rowData) {
    const notes = (rowData['수금 관련 특이사항'] || '').trim();

    return `
      <div class="legacy-card mt-3 collection-card ${notes ? '' : 'collection-card-empty'}">
        <div class="legacy-card-row">
          <div class="legacy-card-main">
            <span class="legacy-card-label">
              <i class="fas fa-sticky-note me-2 text-success"></i>
              수금 특이사항
            </span>
            <div class="editable-value" data-field="수금 관련 특이사항" data-original-value="${notes}">
              ${notes || '특이사항이 없습니다.'}
            </div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * 실시간 계산 이벤트 바인딩
   */
  bindCalculationEvents(projectCode) {

    // 총액1 입력 시 실시간 계산 (자체 구현)
    this.accordionContainer.addEventListener('input', (e) => {
      const target = e.target;
      if (!target.matches('input[data-field="총액1"], input[data-field="총액 1"]')) return;

      // 모던 모듈 자체 계산 로직 사용
      this.calculateTotal2FromTotal1(projectCode, e.target.value);
    });

    // 계약금, 중도금, 잔금 입력 시 미수금 실시간 계산 (순수 모던 구현)
    this.accordionContainer.addEventListener('input', (e) => {
      const target = e.target;
      if (!target.matches('input[data-field="계약금"], input[data-field="중도금"], input[data-field="잔금"]')) return;
      const fieldName = e.target.getAttribute('data-field');

      // 트리거된 필드 추적 (잔금 입력 시 수금확인 자동 설정용)
      this.lastTriggeredField = fieldName;

      // 모던 모듈 자체 미수금 계산 로직 사용
      setTimeout(() => {
        this.calculateOutstandingAmount(projectCode);
      }, 100);
    });

    // 총액2는 calculated-field이므로 input 이벤트가 없음 (총액1 + 부가세로 자동 계산됨)
  }

  /**
   * 카드 이벤트 바인딩
   */
  bindCardEvents() {
    // === 통합 편집 버튼 이벤트 ===

    // 통합 편집 버튼 클릭
    this.accordionContainer.addEventListener('click', (e) => {
      if (!e.target.closest('.unified-edit-btn')) return;
      const target = e.target.closest('.unified-edit-btn');
      const projectCode = target.dataset.projectCode;
      this.enableUnifiedEditMode(projectCode);
    });

    // 통합 저장 버튼 클릭
    this.accordionContainer.addEventListener('click', (e) => {
      if (!e.target.closest('.unified-save-btn')) return;
      const target = e.target.closest('.unified-save-btn');
      const projectCode = target.dataset.projectCode;
      this.saveAllChanges(projectCode);
    });

    // 통합 취소 버튼 클릭
    this.accordionContainer.addEventListener('click', (e) => {
      if (!e.target.closest('.unified-cancel-btn')) return;
      const target = e.target.closest('.unified-cancel-btn');
      const projectCode = target.dataset.projectCode;
      this.cancelAllChanges(projectCode);
    });

    // === 개별 카드 편집 이벤트 제거됨 - 통합 편집 모드만 사용 ===
    // .edit-card-btn, .save-card-btn, .cancel-card-btn 이벤트 리스너 완전 제거
    // generateEditButtons()가 빈 문자열을 반환하므로 버튼 자체가 렌더링되지 않음

    // 공사 취소 버튼 이벤트
    this.accordionContainer.addEventListener('click', (e) => {
      const target = e.target.closest('.cancel-construction-btn');
      if (!target) return;
      const projectCode = target.dataset.projectCode;
      this.cancelConstruction(projectCode);
    });

    // 공사 재개 버튼 이벤트 (취소 상태에서)
    this.accordionContainer.addEventListener('click', (e) => {
      const target = e.target.closest('.resume-construction-btn');
      if (!target) return;
      const projectCode = target.dataset.projectCode;
      this.resumeConstruction(projectCode);
    });

    // 폴더 열기 링크 이벤트
    this.accordionContainer.addEventListener('click', (e) => {
      const target = e.target.closest('.folder-open-link');
      if (!target) return;

      e.preventDefault(); // 기본 링크 동작 방지
      logger.debug('[폴더 열기] 링크 클릭됨:', target);

      const projectCode = target.dataset.projectCode;
      logger.debug('[폴더 열기] 프로젝트 코드:', projectCode);

      if (!projectCode) {
        logger.error('[폴더 열기] 프로젝트 코드가 없습니다');
        this.showMessage('프로젝트 코드를 찾을 수 없습니다.', 'error');
        return;
      }

      this.openFolder(projectCode);
    });

    // 폴더 편집 버튼 이벤트
    this.accordionContainer.addEventListener('click', (e) => {
      if (!e.target.closest('.edit-folder-btn')) return;
      const target = e.target.closest('.edit-folder-btn');
      const documentCard = target.closest('.document-card');
      this.toggleFolderEditMode(documentCard, true);
    });

    // 폴더 저장 버튼 이벤트
    this.accordionContainer.addEventListener('click', (e) => {
      if (!e.target.closest('.save-folder-btn')) return;
      const target = e.target.closest('.save-folder-btn');
      const documentCard = target.closest('.document-card');
      this.saveFolderPath(documentCard);
    });

    // 폴더 취소 버튼 이벤트
    this.accordionContainer.addEventListener('click', (e) => {
      if (!e.target.closest('.cancel-folder-btn')) return;
      const target = e.target.closest('.cancel-folder-btn');
      const documentCard = target.closest('.document-card');
      this.toggleFolderEditMode(documentCard, false);
    });

    // 노트 편집 버튼 이벤트
    this.accordionContainer.addEventListener('click', (e) => {
      if (!e.target.closest('.edit-notes-btn')) return;
      const target = e.target.closest('.edit-notes-btn');
      const collectionCard = target.closest('.collection-card');
      this.toggleNotesEditMode(collectionCard, true);
    });

    // 노트 저장 버튼 이벤트
    this.accordionContainer.addEventListener('click', (e) => {
      if (!e.target.closest('.save-notes-btn')) return;
      const target = e.target.closest('.save-notes-btn');
      const collectionCard = target.closest('.collection-card');
      this.saveCollectionNotes(collectionCard);
    });

    // 노트 취소 버튼 이벤트
    this.accordionContainer.addEventListener('click', (e) => {
      if (!e.target.closest('.cancel-notes-btn')) return;
      const target = e.target.closest('.cancel-notes-btn');
      const collectionCard = target.closest('.collection-card');
      this.toggleNotesEditMode(collectionCard, false);
    });


    // 키보드 네비게이션 (레거시 동일)
    this.accordionContainer.addEventListener('keydown', (e) => {
      if (!e.target.closest('.inline-edit-input')) return;

      const infoCard = e.target.closest('.info-card');
      if (!infoCard || !infoCard.id) return; // null 체크 추가

      const projectCode = this.currentProject;
      const cardType = infoCard.id.split('-')[1];

      if (e.key === 'Enter') {
        e.preventDefault();
        this.saveCardChanges(projectCode, cardType);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.cancelCardEditing(projectCode, cardType);
      }
    });

    // 클릭 외부 감지 - 통합 편집 시스템에서는 외부 클릭 시 자동 저장하지 않음
    // 사용자가 명시적으로 "저장" 또는 "취소" 버튼을 클릭해야 함
    // 중복 등록 방지
    if (!this.documentClickHandler) {
      this.documentClickHandler = (e) => {
        if (!this.isOpen) return;

        // 테이블 관련 요소 클릭은 제외 (이벤트 버블링 방지)
        const isTableClick = e.target.closest('table, tbody, tr, td, .dataTables_wrapper');
        if (isTableClick) return;

        // 아코디언 내부 클릭인지 확인
        // 통합 편집 시스템에서는 외부 클릭 시 아무 동작 안 함 (명시적 저장/취소만 허용)
      };
      document.addEventListener('click', this.documentClickHandler);
    }
  }

  /**
   * 카드 편집 모드 활성화 (레거시 동일)
   */
  enableCardEditing(projectCode, cardType) {
    const card = document.getElementById(`card-${cardType}-${projectCode}`);

    // 권한이 없는 카드는 워터마크만 추가하고 편집 모드 진입하지 않음
    const userRole = window.userRole || 'viewer';
    if (!this.canEditCard(cardType, userRole)) {
      // 워터마크만 추가
      if (!card.querySelector('.readonly-watermark')) {
        const watermark = document.createElement('div');
        watermark.className = 'readonly-watermark';
        watermark.innerHTML = '<i class="fas fa-lock"></i>';
        card.insertBefore(watermark, card.firstChild);
        card.classList.add('readonly-card');
      }
      return; // 편집 모드 진입 중단
    }

    // 편집 모드 클래스 추가
    card.classList.add('editing');

    // 편집 가능한 필드들을 입력 필드로 변환
    card.querySelectorAll('.editable-value').forEach((element, index) => {
      const field = element;
      const fieldName = field.dataset.field;
      // 현재 값 가져오기 - 모든 필드 통일 처리
      let currentValue = field.textContent.trim();

      // textContent가 빈 값이면 기존 input에서 값 추출 시도 (모든 필드 타입)
      if (!currentValue || currentValue === '') {
        const existingInput = field.querySelector('input, textarea, select');
        if (existingInput) {
          if (existingInput.type === 'checkbox') {
            currentValue = existingInput.checked ? '포함' : '미포함';
          } else {
            currentValue = existingInput.value || '';
          }
        }

        // 그래도 값이 없으면 프로젝트 데이터에서 가져오기
        if (!currentValue || currentValue === '') {
          const originalData = this.currentProject;
          currentValue = originalData[fieldName] || '';
        }
      }

      // 원본 값 저장 (취소 시 복원용)
      // 부가세, 수금확인은 HTML 템플릿에서 data-original-value가 이미 설정되어 있음
      // 나머지 필드는 현재 화면에 표시된 HTML을 저장 (인풋 생성 전)
      const hasInput = field.querySelector('input, textarea, select');
      if (!hasInput) {
        // 텍스트 상태인 경우만 원본값 저장
        if (fieldName !== '부가세' && fieldName !== '수금 확인') {
          // 현재 화면에 표시된 텍스트 또는 HTML 저장
          field.dataset.originalValue = field.innerHTML;
        }
      }

      // 계산 필드인지 확인 (편집하지 않음)
      const isCalculatedField = this.isCalculatedField(fieldName);

      if (isCalculatedField) {
        // 계산 필드는 편집하지 않고 현재 값 유지하지만, 편집 시작 시점의 원본 HTML 저장
        const originalHTML = field.innerHTML;
        field.setAttribute('data-edit-original-value', originalHTML);        return;
      }

      // 필드 타입에 따른 입력 요소 생성
      const inputElement = this.createInputElement(fieldName, currentValue);
      field.innerHTML = inputElement;
    });

    // 실시간 콤마 포맷팅 이벤트 리스너 추가 (레거시 동일)
    card.querySelectorAll('.money-field').forEach((element, index) => {
      const input = element;

      // 입력 시 실시간 콤마 포맷팅
      input.addEventListener('input', (e) => {
        let cursorPosition = e.target.selectionStart;
        let value = e.target.value.replace(/[^0-9]/g, '');
        let beforeCommas = (e.target.value.substring(0, cursorPosition).match(/,/g) || []).length;

        if (value) {
          let formattedValue = parseInt(value).toLocaleString('ko-KR');
          e.target.value = formattedValue;

          // 커서 위치 재조정 (콤마 추가/제거에 따른)
          let afterCommas = (formattedValue.substring(0, cursorPosition).match(/,/g) || []).length;
          let newPosition = cursorPosition + (afterCommas - beforeCommas);
          e.target.setSelectionRange(newPosition, newPosition);
        } else {
          e.target.value = '';
        }
      });

      // 포커스 시 선택 (레거시 동일)
      input.addEventListener('focus', (e) => {
        setTimeout(() => e.target.select(), 50);
      });

      // 페이스트 이벤트 처리
      input.addEventListener('paste', (e) => {
        setTimeout(() => {
          let value = e.target.value.replace(/[^0-9]/g, '');
          if (value) {
            e.target.value = parseInt(value).toLocaleString('ko-KR');
          }
        }, 10);
      });
    });

    // textarea 이벤트 처리 (레거시 동일)
    card.querySelectorAll('textarea.inline-edit-input').forEach((element, index) => {
      const textarea = element;

      // 자동 높이 조절
      textarea.addEventListener('input', function() {
        this.style.height = 'auto';
        this.style.height = this.scrollHeight + 'px';
      });
    });

    // Financial 카드 편집 시작 시, Collection 카드의 미수금 필드에도 원본값 저장
    if (cardType === 'financial') {
      const collectionCard = document.getElementById(`card-collection-${projectCode}`);
      if (collectionCard) {
        const outstandingField = collectionCard.querySelector('.editable-value[data-field="미수금"]');
        if (outstandingField && !outstandingField.getAttribute('data-edit-original-value')) {
          const originalHTML = outstandingField.innerHTML;
          outstandingField.setAttribute('data-edit-original-value', originalHTML);        }
      }
    }

    // 자동 계산 트리거 이벤트 리스너 추가 (레거시 동일)
    this.addCalculationEventListeners(card, projectCode);

    // 버튼 상태 변경 및 애니메이션 (통합 편집 모드에서는 스킵)
    // === 개별 카드 버튼 참조 제거됨 - 통합 편집 모드만 사용 ===
    // .edit-card-btn, .save-card-btn, .cancel-card-btn 제거됨

    // 🆕 EditState 업데이트 (카드 레벨 통합 이벤트 핸들러)
    // 중복 등록 방지: 이미 리스너가 등록되어 있으면 스킵
    if (this.editState && this.editState.isActive && !card.dataset.editstateListenerAttached) {
      // 공통 핸들러 함수 (input/change 이벤트 모두 처리)
      const updateEditState = (e) => {
        // 편집 가능한 필드에서 발생한 이벤트만 처리
        const editableField = e.target.closest('.editable-value');
        if (editableField && editableField.dataset.field) {
          const fieldName = editableField.dataset.field;
          let value;

          // 필드 타입에 따른 값 추출
          if (e.target.type === 'checkbox') {
            value = e.target.checked ? 'true' : 'false';
          } else if (e.target.classList.contains('money-field')) {
            // 금액 필드: 콤마 제거 후 저장
            value = e.target.value.replace(/,/g, '');
          } else {
            value = e.target.value;
          }

          // EditState에 필드 업데이트 기록
          this.editState.updateField(fieldName, value);
          logger.debug(`[EditState] ${fieldName} 업데이트:`, value);
        }
      };

      // 이벤트 위임: input과 change 이벤트 모두 처리
      card.addEventListener('input', updateEditState, { capture: true });
      card.addEventListener('change', updateEditState, { capture: true }); // 날짜, 체크박스 등

      // 리스너 등록 완료 플래그
      card.dataset.editstateListenerAttached = 'true';
      logger.debug(`[EditState] 카드 input/change 리스너 등록 완료: ${cardType}`);
    }

    // 날짜 필드 생성이 완료된 후 공사 시작일/종료일 검증 설정
    setTimeout(() => {
      this.setupDateRangeValidationForCard(projectCode, cardType);
    }, 150);
  }

  /**
   * 카드의 공사 시작일/종료일 검증 설정
   */
  setupDateRangeValidationForCard(projectCode, cardType) {
    const card = document.getElementById(`card-${cardType}-${projectCode}`);
    if (!card) return;

    // 시작일과 종료일 입력 필드 찾기
    const startDateInput = card.querySelector('input[data-date-type="start"]');
    const endDateInput = card.querySelector('input[data-date-type="end"]');

    if (!startDateInput || !endDateInput) return;

    // 이미 이벤트가 설정되어 있는지 확인 (중복 방지)
    if (startDateInput._dateValidationSet || endDateInput._dateValidationSet) {
      return;
    }

    // 시작일 변경 시 종료일의 최소값 설정
    const startChangeHandler = () => {
      if (startDateInput.value) {
        endDateInput.min = startDateInput.value;

        // 현재 종료일이 시작일보다 이른 경우 초기화
        if (endDateInput.value && endDateInput.value < startDateInput.value) {
          endDateInput.value = '';
        }
      } else {
        // 시작일이 비어있으면 min 제거
        endDateInput.removeAttribute('min');
      }
    };
    startDateInput.addEventListener('change', startChangeHandler);
    startDateInput._dateValidationSet = true;

    // 종료일은 검증 플래그만 설정 (시작일 제한 없음)
    endDateInput._dateValidationSet = true;

    // 초기 로드 시 min 설정
    if (startDateInput.value) {
      endDateInput.min = startDateInput.value;
    }
  }

  /**
   * 드롭다운 필드 옵션 반환 (레거시 동일)
   */
  getDropdownOptions(fieldName) {
    const dropdownOptions = {
      // 기본 정보 카드
      '사업자': ['글로벌', '글로벌그룹'],
      '도급 구분': ['외주', '내부', '일당', '기타'],

      // 공사 정보 카드
      '공사 구분': ['설치', '이전', '세척', '철거', 'A/S', '판매', '매입', '기타'],
      '기계 분류': ['신품', '중고', '기존제품', '기타'],
      '브랜드': ['LG', '삼성', '캐리어', 'FCU', '덕트', '기타'],
      '계약 형태': ['일반', '리스', '렌탈', '기타'],

      // 시공자 옵션 - DB에서 로드 (window.__CONSTRUCTORS_CACHE__),
      // 캐시가 없을 때만 아래 하드코딩 폴백 사용 (시드 데이터와 일치)
      '시공자': (window.__CONSTRUCTORS_CACHE__) ? window.__CONSTRUCTORS_CACHE__ : {
        categories: [
          { name: '메인', options: ['고승빈', '구상모', '노성현', '남진열', '최태식', '한현규', '현재호'] },
          { name: '서브', options: ['김재광', '박성준', '우성덕트', '송파팀장'] },
          { name: '내부', options: ['아이티', '김종연', '김태현', '일당'] },
          { name: '세척', options: ['김석홍'] }
        ]
      },

      // 계산서 옵션 - 카테고리별 구조화 (각 카테고리에서 1개만 선택, 총 3개까지)
      '계산서': {
        categories: [
          { name: '일반', options: ['계약금', '중도금', '잔금'] },
          { name: 'N입금', options: ['계약금', '중도금', '잔금'] },
          { name: '카드', options: ['계약금', '중도금', '잔금'] }
        ],
        special: ['미발행']
      }
    };

    return dropdownOptions[fieldName] || null;
  }

  /**
   * 계산서 드롭다운 생성 (카테고리별 체크박스)
   */
  createBillStatusDropdown(fieldName, currentValue, dropdownOptions) {
    const dropdownId = `bill-status-${Math.random().toString(36).substr(2, 9)}`;

    // 현재 값 파싱
    // 형식: "현금결제-계약금, N입금-중도금, 카드결제-잔금"
    const selectedItems = {};  // { '현금결제-계약금': true, 'N입금-중도금': true }
    let isMibalhaeng = false;

    if (currentValue === '미발행' || !currentValue || currentValue === '-') {
      isMibalhaeng = currentValue === '미발행';
    } else {
      // 콤마로 분리해서 각 항목을 파싱
      const items = currentValue.split(',').map(s => s.trim());
      items.forEach(item => {
        selectedItems[item] = true;
      });
    }

    // 표시할 텍스트
    const displayText = currentValue && currentValue !== '-' ? currentValue : '선택';

    // 미발행 옵션 (맨 위에 표시)
    const specialHtml = `
      <li class="p-0">
        <label class="d-flex align-items-center w-100 m-0 cursor-pointer" style="padding: 0.5rem 0.75rem; gap: 0.5rem; border-bottom: 2px solid #dee2e6;">
          <span style="font-size: 0.875rem;">미발행</span>
          <input class="form-check-input m-0 bill-special-checkbox" type="checkbox" value="미발행"
                 ${isMibalhaeng ? 'checked' : ''}>
        </label>
      </li>
    `;

    // 카테고리별 체크박스 HTML 생성 (한 줄로)
    const categoryHtml = dropdownOptions.categories.map((category, catIndex) => `
      <li class="p-0 bill-category-group" data-category="${category.name}">
        <div style="display: flex; align-items: center; padding: 0.5rem 0.75rem; gap: 0.75rem;">
          <strong style="font-size: 0.8125rem; color: #495057; min-width: 60px;">[${category.name}]</strong>
          <div style="display: flex; gap: 0.75rem; flex: 1;">
            ${category.options.map(option => {
              const itemKey = `${category.name}-${option}`;
              return `
                <label class="d-flex align-items-center m-0 cursor-pointer" style="gap: 0.25rem;">
                  <span style="font-size: 0.8125rem; white-space: nowrap;">${option}</span>
                  <input class="form-check-input m-0 bill-stage-checkbox" type="checkbox"
                         value="${option}"
                         data-category="${category.name}"
                         data-item-key="${itemKey}"
                         ${selectedItems[itemKey] ? 'checked' : ''}>
                </label>
              `;
            }).join('')}
          </div>
        </div>
      </li>
    `).join('');

    return `
      <div class="dropdown multi-select-dropdown inline-edit-input" data-field="${fieldName}">
        <button class="form-select form-select-sm dropdown-toggle text-start multi-select-btn"
                type="button" id="${dropdownId}" data-bs-toggle="dropdown"
                data-bs-auto-close="outside"
                data-bs-boundary="viewport"
                aria-expanded="false"
                style="height: auto; white-space: normal; padding-top: 0.375rem; padding-bottom: 0.375rem;">
          <span class="selected-text">${displayText}</span>
        </button>
        <ul class="dropdown-menu bill-status-dropdown" aria-labelledby="${dropdownId}"
            style="max-width: 400px; min-width: 300px; z-index: 9999;">
          ${specialHtml}
          ${categoryHtml}
        </ul>
      </div>
    `;
  }

  /**
   * 시공자 드롭다운 생성 (카테고리별 체크박스, 한 줄에 5개씩)
   */
  createConstructorDropdown(fieldName, currentValue, dropdownOptions) {
    const dropdownId = `constructor-${Math.random().toString(36).substr(2, 9)}`;

    // 현재 값 파싱 (콤마로 구분)
    const selectedItems = currentValue ? currentValue.split(',').map(s => s.trim()).filter(v => v && v !== '-') : [];

    // 활성 / 비활성 옵션 분리
    const allActiveOptions = dropdownOptions.categories.flatMap(cat => cat.options || []);
    const allInactiveOptions = dropdownOptions.categories.flatMap(cat => cat.inactiveOptions || []);

    // 진짜 "기타": 활성도 비활성도 아닌 값
    const otherValues = selectedItems.filter(v => !allActiveOptions.includes(v) && !allInactiveOptions.includes(v));
    const otherValue = otherValues.length > 0 ? otherValues[0] : '';

    // 정상 체크 대상: 활성 옵션에 있는 값들
    const normalValues = selectedItems.filter(v => allActiveOptions.includes(v));

    // 비활성 시공자 중 현재 프로젝트에 저장된 값들 (회색 옵션으로 표시)
    const inactiveSelected = selectedItems.filter(v => allInactiveOptions.includes(v));

    // 표시할 텍스트
    const displayText = selectedItems.length > 0 ? selectedItems.join(', ') : '선택';

    // 카테고리별 체크박스 HTML 생성 (계산서와 100% 동일 구조, flex-wrap: wrap만 추가)
    // 비활성 시공자 중 현재 프로젝트에 저장된 것은 해당 카테고리에 회색으로 추가 표시
    // 기타는 별도 [기타] 카테고리 row로 분리 (otherHtml에서 처리)
    const categoryHtml = dropdownOptions.categories.map(category => {
      // 이 카테고리의 비활성 시공자 중 현재 프로젝트에 저장된 것만 표시
      const inactiveInThisCategory = (category.inactiveOptions || []).filter(name => inactiveSelected.includes(name));
      const inactiveOptionsHtml = inactiveInThisCategory.map(option => `
        <label class="d-flex align-items-center m-0 cursor-pointer" style="gap: 0.25rem; opacity: 0.55;" title="비활성 시공자 (거래 중단)">
          <span style="font-size: 0.8125rem; white-space: nowrap; color: #6c757d;">${option}</span>
          <input class="form-check-input m-0 constructor-checkbox" type="checkbox"
                 value="${option}"
                 data-category="${category.name}"
                 data-inactive="true"
                 checked>
        </label>
      `).join('');

      return `
        <li class="p-0 bill-category-group constructor-category-group" data-category="${category.name}">
          <div style="display: flex; align-items: flex-start; padding: 0.5rem 0.75rem; gap: 0.75rem;">
            <strong style="font-size: 0.8125rem; color: #495057; min-width: 60px; padding-top: 0.15rem; flex-shrink: 0;">[${category.name}]</strong>
            <div style="display: flex; gap: 0.75rem; flex: 1; flex-wrap: wrap; row-gap: 1rem;">
              ${category.options.map(option => `
                <label class="d-flex align-items-center m-0 cursor-pointer" style="gap: 0.25rem; min-width: 80px; flex-shrink: 0;">
                  <span style="font-size: 0.8125rem; white-space: nowrap;">${option}</span>
                  <input class="form-check-input m-0 constructor-checkbox" type="checkbox"
                         value="${option}"
                         data-category="${category.name}"
                         ${normalValues.includes(option) ? 'checked' : ''}>
                </label>
              `).join('')}
              ${inactiveOptionsHtml}
            </div>
          </div>
        </li>
      `;
    }).join('');

    // [기타] 카테고리 row (다른 카테고리와 동일한 형식 + 체크박스 옆에 입력란)
    const otherHtml = `
      <li class="p-0 bill-category-group constructor-category-group" data-category="기타">
        <div style="display: flex; align-items: center; padding: 0.5rem 0.75rem; gap: 0.75rem;">
          <strong style="font-size: 0.8125rem; color: #495057; min-width: 60px; flex-shrink: 0;">[기타]</strong>
          <div style="display: flex; gap: 0.75rem; flex: 1; align-items: center;">
            <label class="d-flex align-items-center m-0 cursor-pointer" for="${dropdownId}-other-check" style="gap: 0.25rem; min-width: 80px; flex-shrink: 0;">
              <span style="font-size: 0.8125rem; white-space: nowrap;">직접 입력</span>
              <input class="form-check-input m-0 constructor-other-check" type="checkbox"
                     id="${dropdownId}-other-check" ${otherValue ? 'checked' : ''}>
            </label>
            <input type="text" class="form-control form-control-sm constructor-other-input ${otherValue ? '' : 'd-none'}"
                   id="${dropdownId}-other-input" placeholder="이름 입력" value="${otherValue}"
                   style="font-size: 0.8125rem; height: 28px; padding: 0.25rem 0.5rem; flex: 1; min-width: 0; max-width: 250px;">
          </div>
        </div>
      </li>
    `;

    return `
      <div class="dropdown multi-select-dropdown inline-edit-input" data-field="${fieldName}">
        <button class="form-select form-select-sm dropdown-toggle text-start multi-select-btn"
                type="button" id="${dropdownId}" data-bs-toggle="dropdown"
                data-bs-auto-close="outside"
                aria-expanded="false"
                style="height: auto; white-space: normal; padding-top: 0.375rem; padding-bottom: 0.375rem;">
          <span class="selected-text">${displayText}</span>
        </button>
        <ul class="dropdown-menu bill-status-dropdown constructor-dropdown" aria-labelledby="${dropdownId}"
            style="z-index: 9999;">
          ${categoryHtml}
          ${otherHtml}
        </ul>
      </div>
    `;
  }

  /**
   * 입력 요소 생성 (멀티셀렉트 지원)
   */
  createInputElement(fieldName, currentValue) {
    // 멀티셀렉트가 필요한 필드들 (최대 2개 선택, 시공자는 최대 3개)
    const multiSelectFields = ['공사 구분', '기계 분류', '브랜드', '도급 구분', '시공자'];

    // 드롭다운 필드들
    const dropdownOptions = this.getDropdownOptions(fieldName);
    if (dropdownOptions) {
      // 계산서 필드 특별 처리 - 카테고리별 체크박스
      if (fieldName === '계산서' && dropdownOptions.categories) {
        return this.createBillStatusDropdown(fieldName, currentValue, dropdownOptions);
      }

      // 시공자 필드 특별 처리 - 카테고리별 체크박스
      if (fieldName === '시공자' && dropdownOptions.categories) {
        logger.debug('[시공자] createConstructorDropdown 호출', { fieldName, hasCategories: !!dropdownOptions.categories });
        return this.createConstructorDropdown(fieldName, currentValue, dropdownOptions);
      }

      // 멀티셀렉트 필드 처리 - 체크박스 드롭다운
      if (multiSelectFields.includes(fieldName)) {
        // 현재 값을 쉼표로 split (예: "신축, 증축" → ['신축', '증축'])
        const selectedValues = currentValue ? currentValue.split(',').map(v => v.trim()).filter(v => v && v !== '-') : [];
        const dropdownId = `multiselect-${fieldName.replace(/\s/g, '-')}-${Math.random().toString(36).substr(2, 9)}`;

        // 시공자 필드: 드롭다운 옵션에 없는 값 찾기 (기타로 입력한 값)
        let otherValue = '';
        let normalValues = selectedValues;

        if (fieldName === '시공자') {
          // 옵션에 없는 값들을 찾음
          const otherValues = selectedValues.filter(v => !dropdownOptions.includes(v));
          if (otherValues.length > 0) {
            otherValue = otherValues[0]; // 첫 번째 기타 값만 사용
            normalValues = selectedValues.filter(v => dropdownOptions.includes(v));
            logger.debug(`📝 [편집 모드] 시공자 "기타" 값 감지: "${otherValue}"`);
          }
        }

        // 선택된 값들을 표시할 텍스트
        const displayText = selectedValues.length > 0 ? selectedValues.join(', ') : '선택';

        // 체크박스 옵션 HTML (일반 드롭다운과 동일한 스타일)
        const checkboxesHtml = dropdownOptions.map(option => `
          <li class="p-0">
            <label class="d-flex align-items-center w-100 m-0 cursor-pointer" for="${dropdownId}-${option}"
                   style="padding: 0.5rem 0.75rem; gap: 0.5rem;">
              <span style="flex: 1; font-size: 0.875rem;">${option}</span>
              <input class="form-check-input m-0" type="checkbox" value="${option}"
                     id="${dropdownId}-${option}"
                     ${normalValues.includes(option) ? 'checked' : ''}>
            </label>
          </li>
        `).join('');

        // 시공자 필드에 "기타" 체크박스 + 입력창 추가
        const otherCheckboxHtml = fieldName === '시공자' ? `
          <li class="p-0">
            <label class="d-flex align-items-center w-100 m-0 cursor-pointer" for="${dropdownId}-other-check"
                   style="padding: 0.5rem 0.75rem; gap: 0.5rem;">
              <span style="flex: 1; font-size: 0.875rem;">기타</span>
              <input class="form-check-input m-0 accordion-constructor-other-check" type="checkbox"
                     id="${dropdownId}-other-check" ${otherValue ? 'checked' : ''}>
            </label>
          </li>
          <li class="p-0 ${otherValue ? '' : 'd-none'} accordion-constructor-other-input-container" id="${dropdownId}-other-input-container" style="margin-top: -0.25rem;">
            <div style="padding: 0.25rem 0.75rem 0.5rem 0.75rem;">
              <input type="text" class="form-control form-control-sm accordion-constructor-other-input"
                     id="${dropdownId}-other-input" placeholder="이름 입력"
                     value="${otherValue}"
                     style="font-size: 0.875rem;">
            </div>
          </li>
        ` : '';

        // Bootstrap 드롭다운 + 체크박스 (form-select 스타일 모방)
        // 시공자 필드는 가로 배치를 위해 인라인 스타일 제거
        const dropdownStyle = fieldName === '시공자' ? '' : 'style="max-height: 200px; overflow-y: auto;"';

        // 시공자 필드: 3명 이상 선택 시 자연스럽게 줄바꿈되도록 버튼 스타일 추가
        const buttonStyle = fieldName === '시공자' ? 'style="height: auto; white-space: normal; padding-top: 0.375rem; padding-bottom: 0.375rem;"' : '';

        return `
          <div class="dropdown multi-select-dropdown inline-edit-input" data-field="${fieldName}">
            <button class="form-select form-select-sm dropdown-toggle text-start multi-select-btn"
                    type="button" id="${dropdownId}" data-bs-toggle="dropdown"
                    data-bs-auto-close="outside"
                    aria-expanded="false"
                    ${buttonStyle}>
              <span class="selected-text">${displayText}</span>
            </button>
            <ul class="dropdown-menu" aria-labelledby="${dropdownId}"
                ${dropdownStyle}>
              ${checkboxesHtml}
              ${otherCheckboxHtml}
            </ul>
          </div>
        `;
      }

      // 일반 단일 선택 드롭다운
      const optionsHtml = dropdownOptions.map(option =>
        `<option value="${option}" ${currentValue === option ? 'selected' : ''}>${option}</option>`
      ).join('');
      return `<select class="form-select form-select-sm inline-edit-input">${optionsHtml}</select>`;
    }

    // 이메일 필드 (레거시 동일 - placeholder 없음)
    if (fieldName.includes('이메일')) {
      return `<input type="email" class="form-control form-control-sm inline-edit-input"
                     value="${currentValue}">`;
    }

    // 연락처 필드 (레거시 동일 - placeholder 없음)
    if (fieldName.includes('연락처')) {
      return `<input type="tel" class="form-control form-control-sm inline-edit-input input-manager"
                     value="${currentValue}"
                     pattern="[0-9]{3}-[0-9]{4}-[0-9]{4}">`;
    }

    // 날짜 필드들 - 날짜 검증 추가
    if (fieldName.includes('날짜') || fieldName.includes('시작') || fieldName.includes('종료')) {
      const dateValue = this.parseDate(currentValue);
      const inputId = 'date-input-' + Math.random().toString(36).substr(2, 9);

      // 공사 시작일/종료일 구분을 위한 data 속성
      const isStartDate = fieldName.includes('시작');
      const isEndDate = fieldName.includes('종료');
      const dateType = isStartDate ? 'start' : (isEndDate ? 'end' : '');

      // 나중에 attachDateValidator로 검증 적용
      setTimeout(() => {
        const dateInput = document.getElementById(inputId);
        if (dateInput) {
          attachDateValidator(dateInput, (normalized) => {          });
        }
      }, 100);

      return `<input type="date"
                     id="${inputId}"
                     class="form-control form-control-sm inline-edit-input date-validated"
                     value="${dateValue}"
                     data-date-type="${dateType}"
                     data-field-name="${fieldName}"
                     style="cursor: pointer;"
                     lang="ko-KR"
                     onclick="this.showPicker ? this.showPicker() : this.focus()">`;
    }

    // 마진율 필드 (레거시 동일 - 계산 필드, 편집 불가)
    if (fieldName === '마진율') {
      return `<span class="calculated-field text-muted">
                ${currentValue}
              </span>`;
    }

    // 견적서 및 계약서 폴더 경로 필드
    if (fieldName === '견적서 및 계약서 폴더 경로') {
      return `<div class="folder-input-wrapper" style="position: relative; display: inline-block; width: 100%;">
                <input type="text" class="form-control form-control-sm inline-edit-input folder-path-input"
                       value="${currentValue}"
                       placeholder="Google Drive 링크 또는 폴더 ID를 입력하세요"
                       style="width: 100%; padding-right: 35px;">
                <i class="fas fa-check-circle folder-success-icon"
                   style="position: absolute; right: 10px; top: 50%; transform: translateY(-50%);
                          color: #23923c; font-size: 18px; display: none; pointer-events: none;"></i>
              </div>`;
    }

    // 수금 특이사항 필드 - 텍스트 인풋으로 변경
    if (fieldName.includes('특이사항') && fieldName.includes('수금')) {
      return `<input type="text" class="form-control form-control-sm inline-edit-input collection-notes-input"
                     placeholder="수금 관련 특이사항을 입력하세요."
                     value="${currentValue}">`;
    }

    // 금액 필드들 - 실시간 콤마 포맷팅 (레거시 동일 - placeholder 있음)
    if (fieldName.includes('총액') || fieldName.includes('제품대') || fieldName.includes('도급비') ||
        fieldName.includes('자재비') || fieldName.includes('기타비') || fieldName.includes('잔금') ||
        fieldName.includes('계약금') || fieldName.includes('중도금') || fieldName.includes('미수금')) {
      const numericValue = this.parseNumeric(currentValue);
      const inputId = 'money-input-' + Math.random().toString(36).substr(2, 9);
      return `<input type="text" class="form-control form-control-sm inline-edit-input money-field"
                     id="${inputId}"
                     value="${numericValue ? parseInt(numericValue).toLocaleString('ko-KR') : ''}"
                     placeholder="금액 입력"
                     class="inline-edit-select"
                     data-field-type="money">`;
    }

    // 부가세 필드 - Bootstrap form-switch with badge (토글+뱃지)
    if (fieldName === '부가세') {
      const isChecked = currentValue.includes('포함') && !currentValue.includes('미포함');
      const checkboxId = 'vat-checkbox-' + Math.random().toString(36).substr(2, 9);
      return `
        <div class="form-check form-switch">
          <input type="checkbox" class="form-check-input inline-edit-input vat-checkbox"
                 id="${checkboxId}"
                 data-field="부가세"
                 ${isChecked ? 'checked' : ''}
                 onchange="
                   const badge = this.nextElementSibling;
                   const isChecked = this.checked;
                   badge.textContent = isChecked ? '포함' : '미포함';
                   badge.className = 'badge ms-2 vat-status-badge';
                   badge.style.backgroundColor = isChecked ? '#cfe2ff' : '#f8f9fa';
                   badge.style.color = isChecked ? '#084298' : '#6c757d';
                   badge.dataset.vatStatus = isChecked ? 'true' : 'false';
                 ">
          <span class="badge ms-2 vat-status-badge"
                style="background-color: ${isChecked ? '#cfe2ff' : '#f8f9fa'}; color: ${isChecked ? '#084298' : '#6c757d'};"
                data-vat-status="${isChecked ? 'true' : 'false'}">
            ${isChecked ? '포함' : '미포함'}
          </span>
        </div>
      `;
    }

    // 수금확인 필드 - Bootstrap form-switch with badge (토글+뱃지)
    if (fieldName === '수금 확인') {
      const isCompleted = currentValue === 'true' || currentValue === '완료' || currentValue === '✓' || currentValue === 'Y';
      const checkboxId = 'collection-checkbox-' + Math.random().toString(36).substr(2, 9);
      return `
        <div class="form-check form-switch">
          <input type="checkbox" class="form-check-input inline-edit-input collection-checkbox"
                 id="${checkboxId}"
                 data-field="수금 확인"
                 ${isCompleted ? 'checked' : ''}
                 onchange="
                   const badge = this.nextElementSibling;
                   const isChecked = this.checked;
                   badge.textContent = isChecked ? '완료' : '대기';
                   badge.className = 'badge ms-2 collection-status-badge';
                   badge.style.backgroundColor = isChecked ? '#d1edcc' : '#fff3cd';
                   badge.style.color = isChecked ? '#0f5132' : '#664d03';
                   badge.dataset.status = isChecked ? 'true' : 'false';
                 ">
          <span class="badge ms-2 collection-status-badge"
                style="background-color: ${isCompleted ? '#d1edcc' : '#fff3cd'}; color: ${isCompleted ? '#0f5132' : '#664d03'};"
                data-status="${isCompleted ? 'true' : 'false'}">
            ${isCompleted ? '완료' : '대기'}
          </span>
        </div>
      `;
    }

    // 기본 텍스트 필드 (레거시 동일 - placeholder 없음)
    return `<input type="text" class="form-control form-control-sm inline-edit-input text-input"
                   value="${currentValue}">`;
  }

  /**
   * 카드 변경사항 저장 (레거시 동일)
   */
  async saveCardChanges(projectCode, cardType) {
    // 🚫 개별 카드 저장 기능은 더 이상 지원되지 않습니다
    logger.warn(
      '[레거시] saveCardChanges() 호출이 차단되었습니다.\n' +
      '개별 카드 편집 기능은 제거되었으며, 통합 편집 모드를 사용해야 합니다.\n' +
      `호출 위치: projectCode=${projectCode}, cardType=${cardType}`
    );
    this.showMessage('개별 카드 저장은 더 이상 지원되지 않습니다. 상단의 "편집" 버튼을 사용해주세요.', 'warning');
    return;

    // === 아래 코드는 실행되지 않음 (레거시 코드 보존) ===
    const card = document.getElementById(`card-${cardType}-${projectCode}`);

    // 카드 요소 존재 확인
    if (!card) {
      logger.error(`❌ [ERROR] 카드를 찾을 수 없음: card-${cardType}-${projectCode}`);
      return;
    }

    const saveBtn = card.querySelector('.save-card-btn');
    const cancelBtn = card.querySelector('.cancel-card-btn');

    // 버튼 요소 존재 확인
    if (!saveBtn || !cancelBtn) {
      logger.error(`❌ [ERROR] 버튼을 찾을 수 없음: saveBtn=${!!saveBtn}, cancelBtn=${!!cancelBtn}`);
      return;
    }
    const changes = {};

    // 버튼 비활성화 및 로딩 상태
    saveBtn.disabled = true;
    cancelBtn.disabled = true;
    const originalSaveText = saveBtn.innerHTML;
    saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i>저장중...';

    // 변경된 필드들 수집 및 검증 (금액 필드 콤마 제거)
    let validationErrors = [];
    card.querySelectorAll('.editable-value').forEach((element, index) => {
      const field = element;
      const fieldName = field.dataset.field;
      const input = field.querySelector('input, select, textarea, .inline-edit-input');

      if (input) {
        let value;

        // VAT 체크박스 처리 - 레거시 동일 (문자열 'true'/'false')
        if (input.classList.contains('vat-checkbox')) {
          value = input.checked ? 'true' : 'false';
        }
        // 수금확인 체크박스 처리 - 부가세와 동일 (문자열 'true'/'false')
        else if (input.classList.contains('collection-checkbox')) {
          value = input.checked ? 'true' : 'false';
        }
        else {
          value = input.value; // input, select, textarea 모두 .value 지원
        }

        // 필드별 검증
        const validation = this.validateField(fieldName, value);
        if (!validation.isValid) {
          validationErrors.push(`${fieldName}: ${validation.message}`);
          input.classList.add('is-invalid');
          return;
        } else {
          input.classList.remove('is-invalid');
        }

        // 금액 필드인 경우 콤마 제거
        if (input.classList.contains('money-field')) {
          value = value.replace(/[^0-9]/g, '');
        }
        // 날짜 필드는 YYYY-MM-DD 형식 그대로 전송 (백엔드/구글시트에서 처리)
        // else if (fieldName.includes('날짜') || fieldName.includes('시작') || fieldName.includes('종료')) {
        //   value = this.formatDateForSave(value);  // 불필요한 이중 변환 제거
        // }
        changes[fieldName] = value;
      }
    });

    // 검증 오류가 있으면 중단
    if (validationErrors.length > 0) {
      saveBtn.innerHTML = originalSaveText;
      saveBtn.disabled = false;
      cancelBtn.disabled = false;
      this.showMessage(`입력 오류: ${validationErrors.join(', ')}`, 'error');
      return;
    }

    // 계산된 필드들은 구글 시트 수식으로 관리되므로 전송하지 않음
    // this.addCalculatedFieldsToChanges(card, changes);  // 백엔드에서 필터링하므로 주석 처리

    try {
      // API를 통해 변경사항 저장
      const _to = ('AbortSignal' in window && typeof AbortSignal.timeout === 'function')
        ? { signal: AbortSignal.timeout(120000) } : {};
      const response = await fetch(`/api/projects/${projectCode}`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json'
        },
        credentials: 'same-origin',
        body: JSON.stringify(changes),
        ..._to,
      });

      if (!response.ok) {
        // 서버의 구체적인 에러 메시지 읽기
        let errorMessage = `저장 중 오류가 발생했습니다 (HTTP ${response.status})`;
        try {
          const errorData = await response.json();
          if (errorData.message) {
            errorMessage = errorData.message;
          } else if (errorData.error) {
            errorMessage = errorData.error;
            // 필드 정보가 있으면 추가
            if (errorData.field) {
              errorMessage += ` (필드: ${errorData.field})`;
            }
          }
        } catch (e) {
          // JSON 파싱 실패 시 기본 메시지 사용
          console.error('Failed to parse error response:', e);
        }
        throw new Error(errorMessage);
      }

      const result = await response.json();

      if (result.success || Array.isArray(result)) {
        // 성공 애니메이션
        saveBtn.innerHTML = '<i class="fas fa-check me-1"></i>완료!';
        saveBtn.style.transform = 'scale(1.05)';
        setTimeout(() => {
          saveBtn.style.transform = 'scale(1)';
        }, 200);

        // 성공 시 편집 모드 해제 (업데이트된 데이터로)
        // 딜레이 제거로 깜빡임 방지
        if (result.project) {
          // 전체 프로젝트 데이터로 카드 업데이트
          this.updateCardWithProjectData(projectCode, cardType, result.project);
        } else {
          // 변경사항만으로 업데이트 (레거시 방식)
          this.disableCardEditing(projectCode, cardType, changes);
        }

        // 🔥 실시간 계산 실행 (카드 업데이트 이후 - DOM 값 정확히 반영)
        this.triggerRealTimeCalculations(projectCode, Object.keys(changes));

        // 부분 업데이트 이벤트 발송 (전체 새로고침 방지)
        window.dispatchEvent(new CustomEvent('projectUpdated', {
          detail: {
            project: result.project || result,
            projectCode: projectCode,
            action: 'update',
            partialUpdate: true  // 부분 업데이트 플래그
          }
        }));

        this.showMessage('변경사항이 저장되었습니다.', 'success');
      } else {
        throw new Error(result.message || '저장에 실패했습니다.');
      }

    } catch (error) {
      // 오류 상태 복원
      saveBtn.innerHTML = originalSaveText;
      saveBtn.disabled = false;
      cancelBtn.disabled = false;

      this.showMessage(error.message || '저장 중 오류가 발생했습니다.', 'error');
    }
  }

  /**
   * 카드 편집 취소 (값 복원)
   * ⚠️ 통합 취소(cancelAllChanges)에서만 호출됨 - 개별 버튼 제거됨
   */
  cancelCardEditing(projectCode, cardType) {
    const card = document.getElementById(`card-${cardType}-${projectCode}`);
    const cancelBtn = card.querySelector('.cancel-card-btn');

    // 취소 버튼 애니메이션 (카드별 버튼이 있는 경우에만)
    if (cancelBtn && !cancelBtn.disabled) {
      cancelBtn.disabled = true;
      cancelBtn.style.transform = 'scale(0.95)';
      setTimeout(() => {
        cancelBtn.style.transform = 'scale(1)';
        setTimeout(() => cancelBtn.disabled = false, 500);
      }, 100);
    }

    // 해당 카드의 원본 값으로 복원 (편집 가능 필드 + 계산 필드 모두)
    if (!card) {
      logger.error(`[ERROR] 카드를 찾을 수 없음: card-${cardType}-${projectCode}`);
      return;
    }

    // 편집 중인 모든 필드 찾기 (.editable-value와 그 안의 input 모두)
    const allFields = card.querySelectorAll('.editable-value');

    allFields.forEach((element, index) => {
      const fieldName = element.dataset.field;
      const isCalculatedField = this.isCalculatedField(fieldName);
      const editOriginalValue = element.getAttribute('data-edit-original-value');
      const originalValue = element.dataset.originalValue;

      if (isCalculatedField) {
        // 계산 필드는 편집 시작 시점의 원본 값으로 복원
        if (editOriginalValue !== null) {
          element.innerHTML = editOriginalValue;
        }
      } else if (fieldName === '부가세') {
        // 부가세는 뱃지로 복원
        
        const vatValue = originalValue === 'true' ? '포함' : '미포함';
        const badge = this.unifiedBadgeSystem.createBadge('vat', vatValue);
        
        element.innerHTML = badge;
      } else if (fieldName === '수금 확인') {
        // 수금 확인은 뱃지로 복원
        
        const collectionValue = originalValue === 'true' ? '완료' : '대기';
        const badge = this.formatCollectionStatus(collectionValue);
        
        element.innerHTML = badge;
      } else {
        // 일반 편집 필드는 원본 값으로 복원 (날짜 필드도 동일)
        if (originalValue) {
          element.innerHTML = originalValue;
        } else {
          // originalValue가 없으면 빈 값 대신 '-' 표시
          element.innerHTML = '-';
        }
      }
    });

    // 드롭다운 체크박스 리셋 (멀티셀렉트 필드)
    const multiSelectDropdowns = card.querySelectorAll('.multi-select-dropdown');
    multiSelectDropdowns.forEach(dropdown => {
      const fieldName = dropdown.dataset.field;
      const dropdownMenu = dropdown.querySelector('.dropdown-menu');
      const dropdownButton = dropdown.querySelector('.dropdown-toggle');

      if (!dropdownMenu) return;

      // Bootstrap 드롭다운이 열려있으면 닫기
      if (dropdownButton && dropdownMenu.classList.contains('show')) {
        const bsDropdown = bootstrap.Dropdown.getInstance(dropdownButton);
        if (bsDropdown) {
          bsDropdown.hide();
        }
      }

      // 원본 값 가져오기
      const displayElement = card.querySelector(`[data-field="${fieldName}"]`);
      const originalValue = displayElement?.dataset?.originalValue || '';

      // 원본 값을 쉼표로 split하여 배열로 변환
      const originalValues = originalValue ? originalValue.split(',').map(v => v.trim()).filter(v => v && v !== '-') : [];

      logger.debug(`📋 [편집 취소] "${fieldName}" 체크박스 리셋: ${originalValues.join(', ')}`);

      // 모든 체크박스를 원본 상태로 복원
      const checkboxes = dropdownMenu.querySelectorAll('.form-check-input:not(.accordion-constructor-other-check)');
      checkboxes.forEach(checkbox => {
        checkbox.checked = originalValues.includes(checkbox.value);
      });

      // "기타" 체크박스 및 입력창 리셋 (시공자 필드)
      if (fieldName === '시공자') {
        const otherCheck = dropdownMenu.querySelector('.accordion-constructor-other-check');
        const otherInputContainer = dropdownMenu.querySelector('.accordion-constructor-other-input-container');
        const otherInput = dropdownMenu.querySelector('.accordion-constructor-other-input');

        if (otherCheck) {
          otherCheck.checked = false;
        }
        if (otherInputContainer) {
          otherInputContainer.classList.add('d-none');
        }
        if (otherInput) {
          otherInput.value = '';
        }

        logger.debug('✅ [편집 취소] "기타" 체크박스 및 입력창 리셋 완료');
      }
    });

    // Financial 카드 편집 취소 시, Collection 카드의 미수금도 복원 (계산 연동)
    if (cardType === 'financial') {
      const collectionCard = document.getElementById(`card-collection-${projectCode}`);
      if (collectionCard) {
        const outstandingField = collectionCard.querySelector('.editable-value[data-field="미수금"]');
        if (outstandingField) {
          const editOriginalValue = outstandingField.getAttribute('data-edit-original-value');
          const originalValue = outstandingField.dataset.originalValue;

          if (editOriginalValue !== null) {
            outstandingField.innerHTML = editOriginalValue;
          } else if (originalValue) {
            outstandingField.innerHTML = originalValue;
          }
        }
      }
    }

    this.disableCardEditing(projectCode, cardType);
  }

  /**
   * 카드 편집 모드 해제
   */
  disableCardEditing(projectCode, cardType, newValues = null) {
    const card = document.getElementById(`card-${cardType}-${projectCode}`);

    // 새 값이 있으면 적용
    if (newValues) {
      card.querySelectorAll('.editable-value').forEach((element, index) => {
        const field = element;
        const fieldName = field.dataset.field;

        if (newValues[fieldName] !== undefined) {
          let displayValue = newValues[fieldName];
          let colorClass = '';

          // 필드 타입에 따른 표시 형식 변환 및 색상 적용 (레거시 동일)
          if (fieldName.includes('총액') || fieldName.includes('제품대') || fieldName.includes('도급비') || fieldName.includes('자재비') || fieldName.includes('기타비') || fieldName.includes('잔금') || fieldName.includes('계약금') || fieldName.includes('중도금')) {
            displayValue = this.formatCurrency(displayValue);
          } else if (fieldName === '미수금') {
            // 미수금 색상: 0원이면 녹색, 있으면 빨간색
            displayValue = this.formatCurrency(displayValue);
            const numValue = parseFloat(displayValue.replace(/[^0-9]/g, '')) || 0;
            colorClass = numValue === 0 ? 'text-success fw-bold' : 'text-danger fw-bold';
          } else if (fieldName === '순익' || fieldName === '마진율') {
            // 순익/마진율 색상: 양수면 녹색, 음수면 빨간색, 0이면 회색
            const numValue = parseFloat(displayValue) || 0;
            if (fieldName === '순익') {
              displayValue = this.formatCurrency(displayValue);
            }
            colorClass = numValue > 0 ? 'text-success fw-bold' :
                        numValue < 0 ? 'text-danger fw-bold' : 'text-muted';
          } else if (fieldName.includes('날짜') || fieldName.includes('시작') || fieldName.includes('종료')) {
            displayValue = this.formatDate(displayValue);
          } else if (fieldName === '부가세') {
            // 🔥 createBadge를 사용하여 일관된 뱃지 표시 (normalizeStatus 내장)
            displayValue = this.unifiedBadgeSystem.createBadge('vat', newValues[fieldName]);
          } else if (fieldName === '수금 확인') {
            // 🔥 normalizeStatus를 사용하여 다양한 형식("완료", "true", true 등) 처리
            displayValue = this.formatCollectionStatus(newValues[fieldName]);
          }

          // 색상 클래스 적용
          if (colorClass) {
            field.innerHTML = `<span class="${colorClass}">${displayValue}</span>`;
          } else {
            field.innerHTML = displayValue;
          }

          // data-original-value 업데이트 (부가세, 수금 확인 뱃지 필드는 실제 값을 저장)
          if (fieldName === '부가세' || fieldName === '수금 확인') {
            // 🔥 normalizeStatus를 사용하여 일관된 boolean 값으로 저장
            const badgeType = fieldName === '부가세' ? 'vatBadge' : 'collectionBadge';
            field.dataset.originalValue = this.unifiedBadgeSystem[badgeType].normalizeStatus(newValues[fieldName]);
          } else {
            field.dataset.originalValue = field.textContent.trim();
          }
        }
      });
    }

    // 편집 완료 후 백업 속성 정리
    card.querySelectorAll('.editable-value').forEach((element, index) => {
      const field = element;
      field.removeAttribute('data-edit-original-value');
    });

    // 🆕 드롭다운 인스턴스 정리
    const dropdownButtons = card.querySelectorAll('.dropdown-toggle');
    dropdownButtons.forEach(button => {
      const bsDropdown = bootstrap.Dropdown.getInstance(button);
      if (bsDropdown) {
        bsDropdown.dispose();
      }
    });

    // 드롭다운 메뉴를 원래 위치로 복귀 (body에 추가된 경우)
    const dropdownMenus = card.querySelectorAll('.multi-select-dropdown .dropdown-menu');
    dropdownMenus.forEach(menu => {
      const dropdown = menu.closest('.multi-select-dropdown');
      if (dropdown && menu.parentElement === document.body) {
        dropdown.appendChild(menu);
      }
    });

    // 편집 모드 클래스 제거
    card.classList.remove('editing');

    // 버튼 상태 변경 (카드별 버튼이 있는 경우에만)
    const editBtn = card.querySelector('.edit-card-btn');
    const saveBtn = card.querySelector('.save-card-btn');
    const cancelBtn = card.querySelector('.cancel-card-btn');

    if (editBtn && saveBtn && cancelBtn) {
      editBtn.classList.remove('d-none');
      saveBtn.classList.add('d-none');
      cancelBtn.classList.add('d-none');
    }
  }

  /**
   * 전체 프로젝트 데이터로 카드 업데이트 (서버에서 받은 최신 데이터 사용)
   * - 깜빡임 없이 부드럽게 편집 모드에서 일반 모드로 전환
   */
  updateCardWithProjectData(projectCode, cardType, projectData) {
    const card = document.getElementById(`card-${cardType}-${projectCode}`);

    if (!card) {
      logger.error(`[updateCardWithProjectData] 카드를 찾을 수 없음: card-${cardType}-${projectCode}`);
      return;
    }

    logger.debug(`[updateCardWithProjectData] ${cardType} 카드 업데이트 시작:`, projectData);

    // 🆕 드롭다운 인스턴스 정리 (편집 모드 종료 전에 먼저 정리)
    const dropdownButtons = card.querySelectorAll('.dropdown-toggle');
    dropdownButtons.forEach(button => {
      const bsDropdown = bootstrap.Dropdown.getInstance(button);
      if (bsDropdown) {
        bsDropdown.dispose();
      }
    });

    // 드롭다운 메뉴를 원래 위치로 복귀 (body에 추가된 경우)
    const dropdownMenus = card.querySelectorAll('.multi-select-dropdown .dropdown-menu');
    dropdownMenus.forEach(menu => {
      const dropdown = menu.closest('.multi-select-dropdown');
      if (dropdown && menu.parentElement === document.body) {
        dropdown.appendChild(menu);
      }
    });

    // 🔥 editing 클래스는 HTML 업데이트 후에 제거 (깜빡임 방지)
    // 순서: 1) innerHTML 업데이트로 input 제거 → 2) editing 클래스 제거

    // 모든 편집 가능한 필드를 순회하면서 업데이트
    card.querySelectorAll('.editable-value').forEach((element) => {
      const fieldName = element.dataset.field;
      const newValue = projectData[fieldName];

      // 🔥 메모 필드(계약금/중도금/잔금)는 여기서 값과 아이콘을 함께 처리 (한 번에!)
      const isMemoField = fieldName === '계약금' || fieldName === '중도금' || fieldName === '잔금';
      if (isMemoField) {
        const amountFieldName = fieldName + ' (금액)';
        // "(금액)" 필드가 비어있거나 빈 문자열이면 원본 필드(계약금/중도금/잔금)에서 fallback
        const amountValueRaw = projectData[amountFieldName] || projectData[fieldName] || '';
        const memoData = projectData[fieldName + '_메모'];
        const hasMemo = memoData && memoData.trim() !== '';
        const hasAmount = AmountCalculator.safeParseCurrency(amountValueRaw) !== 0;

        // 금액 + 메모 아이콘을 한 번에 설정
        element.removeAttribute('data-bs-toggle');
        element.removeAttribute('data-bs-title');
        element.style.cursor = '';

        if (!hasAmount) {
          element.innerHTML = this.formatCurrency(amountValueRaw);
        } else {
          const tooltipText = hasMemo ? this.escapeHTML(memoData) : '메모를 작성해주세요';
          const iconType = hasMemo ? 'fas' : 'far';
          const stateClass = hasMemo ? 'has-memo' : 'no-memo';
          const iconColorClass = 'text-success';  // 금액이 있으면 항상 초록색
          const ariaLabel = hasMemo ? '메모 보기' : '메모를 작성해주세요';

          element.innerHTML = `
            <span class="memo-value-wrapper" data-bs-toggle="tooltip" data-bs-title="${tooltipText}" aria-label="${ariaLabel}">
              <span class="memo-amount-text">${this.formatCurrency(amountValueRaw)}</span>
              <span class="memo-tooltip-trigger ${stateClass}">
                <i class="${iconType} fa-sticky-note ms-1 ${iconColorClass}"></i>
              </span>
            </span>
          `;
        }

        // data-original-value 업데이트
        element.dataset.originalValue = amountValueRaw;
        return; // 메모 필드 처리 완료
      }

      if (newValue !== undefined && newValue !== null) {
        let displayValue = newValue;
        let colorClass = '';

        logger.debug(`  [필드 업데이트] ${fieldName}: ${newValue}`);

        // 필드 타입에 따른 표시 형식 변환 및 색상 적용 (disableCardEditing과 동일 로직)
        if (fieldName.includes('총액') || fieldName.includes('제품대') || fieldName.includes('도급비') ||
            fieldName.includes('자재비') || fieldName.includes('기타비')) {
          displayValue = this.formatCurrency(displayValue);
        } else if (fieldName === '미수금') {
          displayValue = this.formatCurrency(displayValue);
          const numValue = parseFloat(displayValue.replace(/[^0-9]/g, '')) || 0;
          colorClass = numValue === 0 ? 'text-success fw-semibold' : 'text-danger fw-semibold';
        } else if (fieldName === '순익' || fieldName === '마진율') {
          const numValue = parseFloat(displayValue) || 0;
          if (fieldName === '순익') {
            displayValue = this.formatCurrency(displayValue);
          }
          colorClass = numValue > 0 ? 'text-success fw-semibold' :
                      numValue < 0 ? 'text-danger fw-semibold' : 'text-muted';
        } else if (fieldName.includes('날짜') || fieldName.includes('시작') || fieldName.includes('종료')) {
          displayValue = this.formatDate(displayValue);
        } else if (fieldName === '부가세') {
          // 🔥 createBadge를 사용하여 일관된 뱃지 표시 (normalizeStatus 내장)
          displayValue = this.unifiedBadgeSystem.createBadge('vat', newValue);
        } else if (fieldName === '수금 확인') {
          // 🔥 normalizeStatus를 사용하여 다양한 형식("완료", "true", true 등) 처리
          displayValue = this.formatCollectionStatus(newValue);
        }

        // 색상 클래스 적용하여 innerHTML 업데이트
        if (colorClass) {
          element.innerHTML = `<span class="${colorClass}">${displayValue}</span>`;
        } else {
          element.innerHTML = displayValue;
        }

        // data-original-value 업데이트 (다음 편집을 위해)
        if (fieldName === '부가세' || fieldName === '수금 확인') {
          // 🔥 normalizeStatus를 사용하여 일관된 boolean 값으로 저장
          const badgeType = fieldName === '부가세' ? 'vatBadge' : 'collectionBadge';
          element.dataset.originalValue = this.unifiedBadgeSystem[badgeType].normalizeStatus(newValue);
        } else {
          element.dataset.originalValue = element.textContent.trim();
        }


      }
    });

    // 편집 완료 후 백업 속성 정리 (data-edit-original-value만 제거)
    card.querySelectorAll('.editable-value').forEach((element) => {
      element.removeAttribute('data-edit-original-value');
    });

    // 버튼 상태 변경 (개별 카드 버튼이 있는 경우만)
    const editBtn = card.querySelector('.edit-card-btn');
    if (editBtn) {
      editBtn.classList.remove('d-none');
    }
    card.querySelectorAll('.save-card-btn, .cancel-card-btn').forEach(btn => btn.classList.add('d-none'));

    // HTML 업데이트 완료 후 editing 클래스 제거 (깜빡임 방지)
    card.classList.remove('editing');

  }

  /**
   * 메모 버튼 아이콘 업데이트 (저장 후 메모 상태 반영)
   */
  updateMemoButtonIcons(projectCode, projectData) {
    logger.debug('[updateMemoButtonIcons] 시작:', {
      projectCode,
      projectData,
      hasFieldMemoButton: !!this.fieldMemoButton
    });

    if (!this.fieldMemoButton) {
      logger.warn('[updateMemoButtonIcons] FieldMemoButton이 초기화되지 않음');
      return;
    }

    const collectionCard = document.getElementById(`card-collection-${projectCode}`);
    if (!collectionCard) {
      logger.warn(`[updateMemoButtonIcons] 수금 카드를 찾을 수 없음: card-collection-${projectCode}`);
      return;
    }

    // 메모 가능한 필드 (계약금, 중도금, 잔금)
    const memoableFields = ['계약금', '중도금', '잔금'];

    memoableFields.forEach(fieldName => {
      const memoKey = `${fieldName}_메모`;
      const memo = projectData[memoKey] || '';
      const amount = AmountCalculator.safeParseCurrency(projectData[fieldName] || 0);

      logger.debug(`  [메모 필드] ${fieldName}:`, {
        amount,
        memo: memo ? `"${memo}"` : '(없음)',
        memoLength: memo.length
      });

      // 메모 버튼 아이콘 업데이트 (편집 모드용)
      const button = collectionCard.querySelector(`.field-memo-btn[data-field="${fieldName}"]`);
      if (button) {
        this.fieldMemoButton.updateButtonIcon(button, memo);
        logger.debug(`[메모 버튼 아이콘 업데이트] ${fieldName}: ${memo ? '메모 있음' : '메모 없음'}`);
      }

      // .editable-value 내부의 아이콘 업데이트 (읽기 모드용)
      const valueElement = collectionCard.querySelector(`.editable-value[data-field="${fieldName}"]`);
      if (valueElement) {
        const formattedAmount = this.formatCurrency(amount);
        valueElement.removeAttribute('data-bs-toggle');
        valueElement.removeAttribute('data-bs-title');
        valueElement.style.cursor = '';

        if (amount > 0) {
          const tooltipText = memo ? this.escapeHTML(memo) : '메모를 작성해주세요';
          const iconType = memo ? 'fas' : 'far';
          const stateClass = memo ? 'has-memo' : 'no-memo';
          const iconColorClass = 'text-success';  // 금액이 있으면 항상 초록색

          valueElement.innerHTML = `
            <span class="memo-value-wrapper" data-bs-toggle="tooltip" data-bs-title="${tooltipText}" aria-label="${memo ? '메모 보기' : '메모를 작성해주세요'}">
              <span class="memo-amount-text">${formattedAmount}</span>
              <span class="memo-tooltip-trigger ${stateClass}">
                <i class="${iconType} fa-sticky-note ms-1 ${iconColorClass}"></i>
              </span>
            </span>
          `;
        } else {
          valueElement.textContent = formattedAmount;
        }

        valueElement.dataset.originalValue = formattedAmount;

        logger.debug(`[메모 필드 업데이트 완료] ${fieldName}: ${formattedAmount}, 메모: ${memo ? '있음' : '없음'}`);
      } else {
        logger.warn(`[updateMemoButtonIcons] editable-value 요소를 찾을 수 없음: ${fieldName}`);
      }
    });

    // Bootstrap tooltip 재초기화 (DOM 변경 후 필수)
    try {
      // 수금 카드 내의 모든 tooltip 재초기화
      const tooltipTriggerList = collectionCard.querySelectorAll('[data-bs-toggle="tooltip"]');
      tooltipTriggerList.forEach(tooltipTriggerEl => {
        // 기존 tooltip 인스턴스 제거 (있으면)
        const existingTooltip = window.bootstrap?.Tooltip.getInstance(tooltipTriggerEl);
        if (existingTooltip) {
          existingTooltip.dispose();
        }
        // 새 tooltip 인스턴스 생성
        new window.bootstrap.Tooltip(tooltipTriggerEl);
      });
      logger.debug('[Bootstrap Tooltip] 재초기화 완료');
    } catch (error) {
      logger.error('[Bootstrap Tooltip] 재초기화 실패:', error);
    }
  }

  /**
   * 아코디언 내부 필드 업데이트 (공사 취소/재개 시)
   */
  updateAccordionFields(projectCode, updatedProject) {
    // accordionContainer 자체를 사용 (project-accordion-content는 존재하지 않음)
    const accordion = this.accordionContainer;
    if (!accordion) {
      logger.warn(`[updateAccordionFields] 아코디언 컨테이너를 찾을 수 없음: ${projectCode}`);
      return;
    }

    // 1. 수금 관련 특이사항 업데이트
    const receivableNotesField = accordion.querySelector('[data-field="수금 관련 특이사항"] .editable-value');
    if (receivableNotesField) {
      receivableNotesField.textContent = updatedProject['수금 관련 특이사항'] || '-';
    }

    // 2. 수금 확인 체크박스 업데이트
    const collectionConfirmedCheckbox = accordion.querySelector('[data-field="수금 확인"] input[type="checkbox"]');
    if (collectionConfirmedCheckbox) {
      // 🔥 normalizeStatus를 사용하여 다양한 형식("완료", "true", true 등) 일관되게 처리
      const isConfirmed = this.unifiedBadgeSystem.collectionBadge.normalizeStatus(updatedProject['수금 확인']) === 'true';
      collectionConfirmedCheckbox.checked = isConfirmed;
    }

    // 3. 공사 확정일 업데이트
    const constructionConfirmedField = accordion.querySelector('[data-field="공사 확정"] .editable-value');
    if (constructionConfirmedField) {
      const newDate = updatedProject['공사 확정'] || updatedProject['공사 확정일'] || '';
      constructionConfirmedField.textContent = newDate || '-';
    }

    logger.debug(`[updateAccordionFields] 필드 업데이트 완료: ${projectCode}`);
  }

  /**
   * 공사 취소 처리
   */
  async cancelConstruction(projectCode) {
    if (!confirm('공사를 취소하시겠습니까?\n취소 시 편집이 불가합니다.')) {
      return;
    }

    // 버튼 찾기 및 로딩 상태 설정
    const cancelBtn = document.querySelector(`.cancel-construction-btn[data-project-code="${projectCode}"]`);
    if (!cancelBtn) {
      logger.error('[cancelConstruction] 취소 버튼을 찾을 수 없습니다.');
      return;
    }

    const originalBtnHTML = cancelBtn.innerHTML;
    cancelBtn.disabled = true;
    cancelBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i>취소 중...';

    // 원래 상태 저장 (롤백용)
    const originalData = window.projectListApp?.stateManager?.findProject(projectCode);

    try {
      const _to = ('AbortSignal' in window && typeof AbortSignal.timeout === 'function')
        ? { signal: AbortSignal.timeout(120000) } : {};
      const response = await fetch('/api/project/cancel', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        credentials: 'same-origin',
        body: JSON.stringify({ projectCode }),
        ..._to,
      });

      const result = await response.json();

      if (result.success && result.updated_project) {
        // 1. 아코디언 내부 UI 업데이트 (수금확인, 공사확정일, 특이사항)
        this.updateAccordionFields(projectCode, result.updated_project);

        // 2. 버튼 업데이트
        const buttonContainer = document.querySelector(`.project-title-section[data-project-code="${projectCode}"] .project-actions-container`);
        if (buttonContainer) {
          buttonContainer.innerHTML = this.generateCancelResumeButton(projectCode, result.updated_project);
        }

        // 3. 취소 스타일 적용 (그레이스케일, 워터마크)
        this.applyCancelledProjectStyles(projectCode);

        // 4. StateManager 동기화 — 필터 재적용 함께 실행 (필터 우선)
        // 취소로 필터 조건에 안 맞아진 프로젝트가 필터 리스트에 남아있으면 UX 혼란.
        // 필터 정확성 우선. 리스트 리렌더로 아코디언 리셋될 수 있는데 이는 사용자가
        // 직접 조작한 결과라 자연스러움.
        if (window.projectListApp?.stateManager) {
          window.projectListApp.stateManager.updateSingleProject(projectCode, result.updated_project);
        }

        // 5. 메인 테이블 행 업데이트
        this.updateMainTableRow(projectCode, '공사취소');

        // 6. 전역 이벤트 발생
        window.dispatchEvent(new CustomEvent('projectUpdated', {
          detail: {
            projectCode: projectCode,
            action: 'cancel_construction',
            updatedProject: result.updated_project,
            partialUpdate: true
          }
        }));

        this.showMessage('공사가 취소되었습니다.', 'success');
      } else {
        throw new Error(result.error || '공사 취소에 실패했습니다.');
      }

    } catch (error) {
      logger.error('[cancelConstruction] 오류:', error);

      // 롤백 시도
      if (originalData && window.projectListApp?.stateManager) {
        logger.warn('[cancelConstruction] 롤백 시도 중...');
        window.projectListApp.stateManager.updateSingleProject(projectCode, originalData);
      }

      // 버튼 상태 복원
      if (cancelBtn) {
        cancelBtn.disabled = false;
        cancelBtn.innerHTML = originalBtnHTML;
      }

      this.showMessage(error.message || '공사 취소 중 오류가 발생했습니다.', 'error');
    }
  }

  /**
   * 공사 재개 처리 (취소 상태에서)
   */
  async resumeConstruction(projectCode) {
    if (!confirm('공사를 재개하시겠습니까?')) {
      return;
    }

    // 버튼 찾기 및 로딩 상태 설정
    const resumeBtn = document.querySelector(`.resume-construction-btn[data-project-code="${projectCode}"]`);
    if (!resumeBtn) {
      logger.error('[resumeConstruction] 재개 버튼을 찾을 수 없습니다.');
      return;
    }

    const originalBtnHTML = resumeBtn.innerHTML;
    resumeBtn.disabled = true;
    resumeBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i>재개 중...';

    // 원래 상태 저장 (롤백용)
    const originalData = window.projectListApp?.stateManager?.findProject(projectCode);

    try {
      const _to = ('AbortSignal' in window && typeof AbortSignal.timeout === 'function')
        ? { signal: AbortSignal.timeout(120000) } : {};
      const response = await fetch('/api/project/resume', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        credentials: 'same-origin',
        body: JSON.stringify({ projectCode }),
        ..._to,
      });

      const result = await response.json();

      if (result.success && result.updated_project) {
        // 1. 아코디언 내부 UI 업데이트 (필드 복원)
        this.updateAccordionFields(projectCode, result.updated_project);

        // 2. 버튼 업데이트
        const buttonContainer = document.querySelector(`.project-title-section[data-project-code="${projectCode}"] .project-actions-container`);
        if (buttonContainer) {
          buttonContainer.innerHTML = this.generateCancelResumeButton(projectCode, result.updated_project);
        }

        // 3. 취소 스타일 제거
        this.removeCancelledProjectStyles(projectCode);

        // 4. StateManager 동기화 — 필터 재적용 함께 실행 (필터 우선, 취소와 대칭)
        if (window.projectListApp?.stateManager) {
          window.projectListApp.stateManager.updateSingleProject(projectCode, result.updated_project);
        }

        // 5. 메인 테이블 행 업데이트 (상태 재계산)
        this.updateMainTableRow(projectCode, null);

        // 6. 전역 이벤트 발생
        window.dispatchEvent(new CustomEvent('projectUpdated', {
          detail: {
            projectCode: projectCode,
            action: 'resume_construction',
            updatedProject: result.updated_project,
            partialUpdate: true
          }
        }));

        this.showMessage('공사가 재개되었습니다.', 'success');
      } else if (result.success && result.already_active) {
        // 이미 정상 상태 (백엔드가 200 + already_active=true 반환) — 실패 아님
        // 사용자에게 정보만 알리고 재개 버튼 로딩 해제
        this.showMessage('이미 정상 상태입니다.', 'info');
        resumeBtn.disabled = false;
        resumeBtn.innerHTML = originalBtnHTML;
      } else {
        throw new Error(result.error || '공사 재개에 실패했습니다.');
      }

    } catch (error) {
      logger.error('[resumeConstruction] 오류:', error);

      // 롤백 시도
      if (originalData && window.projectListApp?.stateManager) {
        logger.warn('[resumeConstruction] 롤백 시도 중...');
        window.projectListApp.stateManager.updateSingleProject(projectCode, originalData);
      }

      // 버튼 상태 복원
      if (resumeBtn) {
        resumeBtn.disabled = false;
        resumeBtn.innerHTML = originalBtnHTML;
      }

      this.showMessage(error.message || '공사 재개 중 오류가 발생했습니다.', 'error');
    }
  }

  isProjectCancelled(projectData) {
    if (!projectData) return false;
    const notes = projectData['수금 관련 특이사항'] || projectData['AG'] || '';
    return typeof notes === 'string' && /공사\s*취소/.test(notes);
  }

  toggleCancelledStyles(projectCode, isCancelled) {
    if (isCancelled) {
      this.applyCancelledProjectStyles(projectCode);
    } else {
      this.removeCancelledProjectStyles(projectCode);
    }
  }

  getAccordionShell() {
    return this.accordionContainer?.querySelector('.accordion-shell') || null;
  }

  /**
   * 취소된 프로젝트 스타일 적용
   */
  applyCancelledProjectStyles(projectCode) {
    const shell = this.getAccordionShell();
    if (!shell) {
      logger.warn(`[applyCancelledProjectStyles] 아코디언 컨테이너를 찾을 수 없음: ${projectCode}`);
      return;
    }

    shell.classList.add('project-cancelled');


    // 편집 버튼만 비활성화 (재개 버튼은 제외)
    const editButtons = shell.querySelectorAll('.unified-edit-btn');
    editButtons.forEach(btn => {
      btn.disabled = true;
      btn.classList.add('disabled');
    });

    // 편집 가능한 필드 비활성화
    const editableFields = shell.querySelectorAll('.editable-value');
    editableFields.forEach(field => {
      field.setAttribute('contenteditable', 'false');
      field.style.pointerEvents = 'none';
    });

    // 워터마크 추가
    const cardBody = shell.querySelector('.card-body');
    if (cardBody && !cardBody.querySelector('.cancelled-watermark')) {
      const watermark = document.createElement('div');
      watermark.className = 'cancelled-watermark';
      watermark.textContent = '공사 취소';
      cardBody.style.position = 'relative';
      cardBody.appendChild(watermark);
    }

    // 재개 버튼이 있으면 명시적으로 활성화 (CSS에서 비활성화되지 않도록)
    const resumeBtn = shell.querySelector('.resume-construction-btn');
    if (resumeBtn) {
      resumeBtn.disabled = false;
      resumeBtn.style.pointerEvents = 'auto';
      resumeBtn.classList.remove('disabled');
    }
  }

  /**
   * 취소된 프로젝트 스타일 제거
   */
  removeCancelledProjectStyles(projectCode) {
    const shell = this.getAccordionShell();
    if (!shell) {
      logger.warn(`[removeCancelledProjectStyles] 아코디언 컨테이너를 찾을 수 없음: ${projectCode}`);
      return;
    }

    shell.classList.remove('project-cancelled');


    const editButtons = shell.querySelectorAll('.unified-edit-btn');
    editButtons.forEach(btn => {
      btn.disabled = false;
      btn.classList.remove('disabled');
    });

    const editableFields = shell.querySelectorAll('.editable-value');
    editableFields.forEach(field => {
      field.removeAttribute('contenteditable');
      field.style.pointerEvents = '';
    });

    const watermark = shell.querySelector('.cancelled-watermark');
    if (watermark) {
      watermark.remove();
    }
  }

  /**
   * 메인 테이블 행 업데이트 (공사 취소/재개 시)
   * @param {string} projectCode - 프로젝트 코드
   * @param {string|null} newStatus - 새 상태 ('공사취소' 또는 null for 재계산)
   */
  updateMainTableRow(projectCode, newStatus) {
    try {
      // ProjectTable 컴포넌트의 table 인스턴스 가져오기
      const table = window.projectListApp?.components?.table?.table;
      if (!table) {
        logger.warn('[updateMainTableRow] DataTable 인스턴스를 찾을 수 없습니다.');
        return;
      }

      const accordionWasOpen = this.isOpen && this.currentProject?.['프로젝트 코드'] === projectCode;

      // 해당 프로젝트 코드의 행 찾기 (필터/검색 무시하고 모든 데이터 검색)
      const rowData = table.rows({search: 'removed'}).data().toArray().find(row =>
        row['프로젝트 코드'] === projectCode
      );

      if (!rowData) {
        logger.warn(`[updateMainTableRow] 프로젝트 ${projectCode}을(를) 찾을 수 없습니다.`);
        return;
      }

      // 데이터 업데이트
      if (newStatus === '공사취소') {
        rowData['수금 관련 특이사항'] = '공사 취소';
        rowData['수금 확인'] = false;  // 수금 확인 false로 설정
        rowData['공사 확정'] = '';      // 공사 확정일 초기화
      } else if (newStatus === null) {
        // 재개 시 특이사항 초기화
        rowData['수금 관련 특이사항'] = '';

        if (window.dayjs) {
          rowData['공사 확정'] = window.dayjs().format('YYYY-MM-DD');
        } else {
          const now = new Date();
          const yyyy = now.getFullYear();
          const mm = String(now.getMonth() + 1).padStart(2, '0');
          const dd = String(now.getDate()).padStart(2, '0');
          rowData['공사 확정'] = `${yyyy}-${mm}-${dd}`;
        }
      }

      // DataTable 행 찾기 및 업데이트
      const rowIndex = table.rows().data().toArray().findIndex(row =>
        row['프로젝트 코드'] === projectCode
      );

      if (rowIndex !== -1) {
        const row = table.row(rowIndex);
        row.data(rowData);

        if (accordionWasOpen) {
          // 아코디언이 열린 상태에서는 draw를 지연하고 상태를 재설정
          this.pendingTableUpdate = true;
        } else {
          row.invalidate().draw(false);
        }

        // 행 DOM 요소 가져오기
        const rowNode = row.node();
        if (rowNode) {
          const isCancelledNow = this.isProjectCancelled(rowData);
          if (isCancelledNow) {
            rowNode.classList.add('project-cancelled-row');
          } else {
            rowNode.classList.remove('project-cancelled-row');
          }

          if (!isCancelledNow) {
            rowNode.classList.add('row-flash');
            setTimeout(() => {
              rowNode.classList.remove('row-flash');
            }, 1000);
          }
        }

        if (accordionWasOpen) {
          // 안전한 방식: draw로 모든 렌더러 적용 후 아코디언 재오픈
          const savedProjectCode = projectCode;
          const savedProjectData = { ...rowData };

          // 테이블 업데이트 (배지, 날짜, 금액 등 모든 렌더러 정상 작동)
          row.invalidate().draw(false);

          // draw 완료 후 취소 스타일 클래스 재적용 + 아코디언 재오픈
          // (2026-07-07): draw가 tr DOM을 새로 생성해서 line 3130에서 붙인
          // .project-cancelled-row 클래스가 유실됨. rowCallback이 실행되긴 하지만
          // 취소 상태 매칭이 확실치 않으므로 여기서 명시적으로 재적용.
          setTimeout(() => {
            const freshRow = table.row(rowIndex);
            const freshRowNode = freshRow?.node();
            if (freshRowNode) {
              const isCancelled = this.isProjectCancelled(rowData);
              if (isCancelled) {
                freshRowNode.classList.add('project-cancelled-row');
              } else {
                freshRowNode.classList.remove('project-cancelled-row');
              }
            }
            this.reopenAccordion(savedProjectCode, savedProjectData);
          }, 0);
        }

        logger.debug(`[updateMainTableRow] 프로젝트 ${projectCode} 행 업데이트 완료`);
      }

    } catch (error) {
      logger.error('[updateMainTableRow] 오류:', error);
    }
  }

  /**
   * 아코디언 상태를 강제로 초기화 (테이블이 재렌더링 되었을 때 사용)
   */
  hardResetAccordionState() {
    this.isOpen = false;
    this.currentProject = null;
    this.originalProjectCode = null;
    this.currentRowNumber = null;
    this.pendingTableUpdate = false;

    if (this.accordionContainer) {
      this.accordionContainer.classList.remove('show', 'accordion-slide-down', 'accordion-slide-up');
    }

    document.querySelectorAll('.accordion-row').forEach(row => row.remove());
    document.querySelectorAll('tbody tr.table-active').forEach(row => row.classList.remove('table-active'));
  }

  /**
   * 폴더 열기 - Flask API를 통해 Windows 탐색기 또는 브라우저에서 폴더 열기
   */
  async openFolder(projectCode) {
    if (!projectCode) {
      logger.error('[폴더 열기] 프로젝트 코드가 없습니다');
      return;
    }

    logger.debug('[폴더 열기] API 호출 시작:', projectCode);

    try {
      const apiUrl = `/api/folder/open/${projectCode}`;
      logger.debug('[폴더 열기] API URL:', apiUrl);

      const response = await fetch(apiUrl, {
        credentials: 'same-origin'
      });
      logger.debug('[폴더 열기] Response status:', response.status);

      const data = await response.json();
      logger.debug('[폴더 열기] Response data:', data);

      if (data.success) {
        const folderType = data.folder_type === 'google_drive' ? 'Google Drive' : '탐색기';
        logger.debug(`[폴더 열기] 성공: ${folderType}에서 폴더를 열었습니다.`);
      } else {
        logger.error('[폴더 열기] API 실패:', data.message || data.error);
      }
    } catch (error) {
      logger.error('[폴더 열기 오류]', error);
    }
  }

  /**
   * 유틸리티 메서드들
   */
  formatCollectionStatus(status) {
    return this.unifiedBadgeSystem.createBadge('collection', status);
  }

  formatDate(dateString) {
    if (!dateString || dateString === '-') return '-';

    try {
      if (typeof dateString === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
        return dateString;
      }

      const date = new Date(dateString);
      if (isNaN(date.getTime())) return dateString;

      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');

      return `${year}-${month}-${day}`;
    } catch (error) {
      return dateString;
    }
  }

  parseDate(dateString) {
    // 엄격한 빈 값 처리 - 다양한 빈 값 형태 처리
    if (!dateString ||
        dateString === '-' ||
        dateString === '' ||
        dateString === null ||
        dateString === undefined ||
        dateString === 'undefined' ||
        dateString === 'null' ||
        String(dateString).trim() === '' ||
        String(dateString).trim() === '-' ||
        String(dateString).trim().toLowerCase() === 'null') {
      return '';
    }

    // 문자열로 변환하고 trim
    const dateStr = String(dateString).trim();

    // 빈 문자열이면 반환
    if (dateStr === '') {
      return '';
    }

    // YYYY/M/D 형식을 YYYY-MM-DD로 변환
    if (dateStr.includes('/')) {
      const parts = dateStr.split('/');
      if (parts.length === 3 && parts.every(part => !isNaN(parseInt(part)))) {
        const year = parts[0];
        const month = parts[1].padStart(2, '0');
        const day = parts[2].padStart(2, '0');
        return `${year}-${month}-${day}`;
      }
    }

    // YYYY-MM-DD 형식은 그대로 반환 (더 엄격한 검증)
    if (dateStr.includes('-') && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      // 실제 유효한 날짜인지 검증
      const [year, month, day] = dateStr.split('-').map(Number);
      if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        return dateStr;
      }
    }

    // 다른 형식 시도 - 더 엄격한 검증
    try {
      const date = new Date(dateStr);
      if (isNaN(date.getTime()) || date.getFullYear() < 1900 || date.getFullYear() > 2100) {
        return '';
      }

      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    } catch (error) {
      return '';
    }
  }

  /**
   * 날짜를 구글 시트 저장용 형식(YYYY/MM/DD)으로 변환 (레거시 동일)
   */
  formatDateForSave(dateString) {
    if (!dateString || dateString === '-') return '';

    // YYYY-MM-DD → YYYY/MM/DD 변환
    if (dateString.includes('-')) {
      const parts = dateString.split('-');
      if (parts.length === 3) {
        return `${parts[0]}/${parts[1]}/${parts[2]}`;
      }
    }

    return dateString;
  }

  /**
   * 수금 날짜를 오늘로 설정 (레거시 동일)
   */
  setCollectionDateToToday(projectCode) {
    const today = new Date().toISOString().split('T')[0]; // YYYY-MM-DD 형식
    const collectionCard = document.getElementById(`card-collection-${projectCode}`);

    if (collectionCard) {
      const dateField = collectionCard.querySelector('[data-field="수금 날짜"]');
      if (dateField) {
        // 편집 모드인지 확인
        const dateInput = dateField.querySelector('input[type="date"]');
        if (dateInput) {
          dateInput.value = today;        }
      }
    }
  }

  /**
   * 수금 날짜를 초기화 (레거시 동일)
   */
  resetCollectionDate(projectCode) {
    const collectionCard = document.getElementById(`card-collection-${projectCode}`);

    if (collectionCard) {
      const dateField = collectionCard.querySelector('[data-field="수금 날짜"]');
      if (dateField) {
        // 편집 모드인지 확인
        const dateInput = dateField.querySelector('input[type="date"]');
        if (dateInput) {
          dateInput.value = '';        }
      }
    }
  }

  formatCurrency(amount) {
    if (!amount || amount === '-') return '-';

    const numValue = typeof amount === 'string' ?
      parseFloat(amount.replace(/[^0-9.-]/g, '')) :
      parseFloat(amount);

    if (isNaN(numValue) || numValue === 0) return '-';

    return numValue.toLocaleString() + '원';
  }

  parseNumeric(value) {
    if (!value || value === '-') return '';

    const numValue = typeof value === 'string' ?
      parseFloat(value.replace(/[^0-9.-]/g, '')) :
      parseFloat(value);

    return isNaN(numValue) ? '' : numValue.toString();
  }

  /**
   * HTML 특수문자 이스케이프 (XSS 방지)
   */
  escapeHTML(str) {
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
   * 아코디언 내부의 모든 툴팁 초기화
   * Bootstrap 5 표준 방식: data-bs-toggle="tooltip" 속성 사용
   */
  initializeAccordionTooltips() {
    if (!this.accordionContainer) return;

    // 기존 툴팁 인스턴스 제거 (메모리 누수 방지)
    const existingTooltips = this.accordionContainer.querySelectorAll('[data-bs-toggle="tooltip"]');
    existingTooltips.forEach(el => {
      const instance = window.bootstrap?.Tooltip.getInstance(el);
      if (instance) {
        instance.dispose();
      }
    });

    // 새로 렌더링된 모든 [data-bs-toggle="tooltip"] 요소에 대해 툴팁 초기화
    const tooltipElements = this.accordionContainer.querySelectorAll('[data-bs-toggle="tooltip"]');
    tooltipElements.forEach(el => {
      if (window.bootstrap && window.bootstrap.Tooltip) {
        new window.bootstrap.Tooltip(el, {
          container: 'body',  // 툴팁이 테이블 바깥에 표시되도록 함
          trigger: 'hover',
          placement: 'top',
          html: false
        });
      }
    });

    if (tooltipElements.length > 0) {
      logger.debug(`[ProjectRowAccordion] 툴팁 ${tooltipElements.length}개 초기화 완료`);
    }
  }

  formatBillStatus(rowData) {
    const billValue = rowData['계산서'] || rowData['X'] || '';

    if (billValue === true || billValue === 'TRUE' || billValue === '✓' || billValue === 'true') {
      return '발행완료';
    }
    if (billValue === false || billValue === 'FALSE' || billValue === '' || billValue === '-') {
      return '미발행';
    }
    if (typeof billValue === 'string' && billValue.trim() && !['true', 'false', 'TRUE', 'FALSE'].includes(billValue)) {
      // 계산서 값에서 카테고리 추출하여 아이콘 추가
      const text = billValue;
      const icons = this.extractBillIcons(billValue);
      return `${text}${icons}`;
    }

    return '미발행';
  }

  /**
   * 계산서 값에서 카테고리별 아이콘 추출
   * 예: "일반-계약금, N입금-중도금" → 일반 아이콘 + N입금 아이콘
   */
  extractBillIcons(billValue) {
    if (!billValue || billValue === '-' || billValue === '미발행') {
      return '';
    }

    // 사용된 카테고리를 Set으로 수집 (중복 제거)
    const categories = new Set();
    const items = billValue.split(',').map(s => s.trim());

    items.forEach(item => {
      const parts = item.split('-');
      if (parts.length === 2) {
        const category = parts[0].trim();
        categories.add(category);
      }
    });

    // 각 카테고리에 대한 아이콘 생성
    let icons = '';
    categories.forEach(category => {
      icons += this.getBillStatusIcon(category);
    });

    return icons;
  }

  calculateMarginRate(rowData) {
    // 마진율 계산 로직 - Google Sheets와 동일하게 총액 1 기준
    const totalAmount = parseFloat(rowData['총액 1'] || rowData['총액1'] || rowData['총액'] || 0);
    const costs = parseFloat(rowData['제품대'] || 0) +
                 parseFloat(rowData['도급비'] || 0) +
                 parseFloat(rowData['자재비'] || 0) +
                 parseFloat(rowData['기타비'] || 0);

    if (totalAmount === 0) return '<span class="text-muted">-</span>';

    const marginRate = ((totalAmount - costs) / totalAmount) * 100;
    const colorClass = marginRate > 0 ? 'text-success fw-semibold' :
                       marginRate < 0 ? 'text-danger fw-semibold' : 'text-muted';

    return `<span class="${colorClass}">${marginRate.toFixed(1)}%</span>`;
  }

  /**
   * 계산된 필드들을 변경사항에 추가 (저장 시 계산 결과 포함)
   */
  addCalculatedFieldsToChanges(card, changes) {
    // 계산 필드들의 현재 값을 수집
    const calculatedFields = ['총액 2', '미수금', '마진율', '순익'];

    calculatedFields.forEach(fieldName => {
      const field = card.querySelector(`[data-field="${fieldName}"].editable-value.calculated-field`);
      if (field) {
        // textContent에서 값 추출 (span 태그 내부 텍스트)
        let value = field.textContent.trim();

        // 화폐 단위 제거 (1,000원 → 1000)
        if (fieldName === '총액 2' || fieldName === '미수금' || fieldName === '순익') {
          value = value.replace(/[^0-9-]/g, '');
        }
        // 퍼센트 제거 (15.5% → 15.5)
        else if (fieldName === '마진율') {
          value = value.replace(/[^0-9.-]/g, '');
        }

        // 값이 있으면 변경사항에 추가 (미수금은 0도 저장)
        if (value && value !== '-') {
          changes[fieldName] = value;
        }
      }
    });

    // 수금확인 자동 설정된 값도 포함 (잔금 입력 시 자동 설정됨)
    if (this.lastTriggeredField === '잔금') {
      const confirmBadge = card.querySelector('[data-field="수금 확인"] .collection-status-badge');
      const dateField = card.querySelector('[data-field="수금 날짜"]');

      if (confirmBadge) {
        const confirmValue = confirmBadge.dataset.status === 'true' ? 'true' : 'false';
        changes['수금 확인'] = confirmValue;
      }

      if (dateField) {
        const dateInput = dateField.querySelector('input');
        if (dateInput) {
          const dateValue = dateInput.value;
          if (dateValue && dateValue !== '-') {
            changes['수금 날짜'] = dateValue;
          }
        }
      }
    }
  }

  /**
   * 계산 필드인지 확인 (자동 계산되는 읽기 전용 필드)
   */
  isCalculatedField(fieldName) {
    const calculatedFields = [
      '총액 2',      // 총액1 + 부가세로 계산
      '총액2',       // 공백 없는 버전
      '미수금',      // 총액2 - (계약금 + 중도금 + 잔금)으로 계산 (백엔드에서 통일됨)
      '마진율',     // 마진율 계산 결과 (마진율 %)
      '순익'        // 순익 계산 결과 (총액1 - 모든 비용)
    ];

    return calculatedFields.includes(fieldName);
  }

  /**
   * 편집 중 실시간 계산 이벤트 리스너 추가 (레거시 동일)
   */
  addCalculationEventListeners(card, projectCode) {
    // 기존 계산 이벤트 리스너 제거 (중복 방지)
    this.removeCalculationEventListeners(card);

    // 편집 모드에서 총액1 변경 시 → 총액2 자동 계산 (개별 input 바인딩)
    const total1Input = card.querySelector('[data-field="총액 1"] input');
    if (total1Input) {
      const total1Handler = () => {
        this.lastTriggeredField = '총액 1';
        
        this.calculateTotal2FromTotal1(projectCode);
      };
      total1Input.addEventListener('input', total1Handler);
      total1Input._calculationHandler = total1Handler; // 제거용 참조 저장
    }

    // 부가세 체크박스 변경 시 → 총액2 자동 계산
    const vatCheckbox = card.querySelector('[data-field="부가세"] .vat-checkbox');
    if (vatCheckbox) {
      const vatHandler = () => {
        this.lastTriggeredField = '부가세';
        this.calculateTotal2FromTotal1(projectCode);
      };
      vatCheckbox.addEventListener('change', vatHandler);
      vatCheckbox._calculationHandler = vatHandler;
    }

    // 편집 모드에서 결제 관련 필드 변경 시 → 미수금 자동 계산 (개별 input 바인딩)
    const paymentFields = ['계약금', '중도금', '잔금']; // 총액2 제거 (계산 필드이므로 input 이벤트 없음)
    paymentFields.forEach(fieldName => {
      const paymentInput = card.querySelector(`[data-field="${fieldName}"] input`);
      if (paymentInput) {
        const paymentHandler = () => {
          this.lastTriggeredField = fieldName;
          
          this.calculateOutstandingAmount(projectCode);
        };
        paymentInput.addEventListener('input', paymentHandler);
        paymentInput._calculationHandler = paymentHandler;
      }
    });

    // 비용 관련 필드 변경 시 → 순익/마진율 자동 계산
    const costFields = ['제품대', '도급비', '자재비', '기타비'];
    costFields.forEach(fieldName => {
      const costInput = card.querySelector(`[data-field="${fieldName}"] input`);
      if (costInput) {
        const costHandler = () => {
          this.lastTriggeredField = fieldName;
          this.calculateProfitFields(projectCode);
        };
        costInput.addEventListener('input', costHandler);
        costInput._calculationHandler = costHandler;
      }
    });
  }

  /**
   * 기존 계산 이벤트 리스너 제거 (중복 방지)
   */
  removeCalculationEventListeners(card) {
    const allInputs = card.querySelectorAll('input, .vat-checkbox');
    allInputs.forEach(input => {
      if (input._calculationHandler) {
        const eventType = input.classList.contains('vat-checkbox') ? 'change' : 'input';
        input.removeEventListener(eventType, input._calculationHandler);
        delete input._calculationHandler;
      }
    });
  }

  /**
   * 자동 계산식 1: 총액1 → 총액2 (부가세 포함/미포함) - 레거시 동일
   */
  calculateTotal2FromTotal1(projectCode, inputValue = null) {
    const financialCard = document.getElementById(`card-financial-${projectCode}`);
    if (!financialCard) {
      logger.warn(`[ERROR] [계산] 금액정보 카드를 찾을 수 없음: ${projectCode}`);
      return;
    }

    // 편집 모드인지 확인
    const isEditMode = financialCard.classList.contains('editing');

    // 편집 모드가 아니면 실행하지 않음 (API 값 사용)
    if (!isEditMode) {
      return;
    }

    // 총액1 값 가져오기 (매개변수가 있으면 사용, 없으면 필드에서 가져오기)
    let total1Value;
    if (inputValue !== null) {
      total1Value = parseFloat(inputValue.replace(/[^0-9.-]/g, '')) || 0;
    } else {
      const total1Input = financialCard.querySelector('[data-field="총액 1"] input');
      total1Value = parseFloat(total1Input.value?.replace(/[^0-9]/g, '') || '0') || 0;
    }

    

    // 부가세 체크박스 상태 가져오기
    const vatCheckbox = financialCard.querySelector('[data-field="부가세"] .vat-checkbox');
    const vatIncluded = vatCheckbox.checked;

    // 총액2 계산: 한국 회계 실무 방식
    let total2Value;
    if (vatIncluded) {
      // 부가세 = ROUND(공급가액 × 0.1, 0)
      const vat = Math.round(total1Value * 0.1);
      // 총액2 = 공급가액 + 부가세
      total2Value = total1Value + vat;
    } else {
      // 부가세 미포함 시 공급가액 그대로
      total2Value = total1Value;
    }

    // NaN/Infinity 검증 및 안전 처리
    if (!Number.isFinite(total2Value)) {
      logger.warn(`[계산 필드 검증] 총액2 계산 오류: total1=${total1Value}, vatIncluded=${vatIncluded}, result=${total2Value}`);
      total2Value = total1Value || 0;  // 안전값으로 총액1 사용
    }

    // 음수 값 검증
    if (total2Value < 0) {
      logger.warn(`[계산 필드 검증] 총액2가 음수: ${total2Value} → 0으로 처리`);
      total2Value = 0;
    }

    

    // 총액2 필드 업데이트 (calculated-field이므로 display만 존재)
    const total2Display = financialCard.querySelector('[data-field="총액 2"].editable-value.calculated-field') ||
                          financialCard.querySelector('[data-field="총액2"].editable-value.calculated-field');

    if (total2Display) {
      const displayValue = total2Value > 0 ? `${parseInt(total2Value).toLocaleString('ko-KR')}원` : '-';
      total2Display.textContent = displayValue;
      // data-original-value는 저장 시에만 사용되므로 편집 중에는 업데이트 안 함
    }

    // 연쇄 계산 트리거 (즉시 실행)
    this.calculateOutstandingAmount(projectCode); // 미수금 재계산
    this.calculateProfitFields(projectCode);      // 마진율, 순익 재계산
  }

  /**
   * 자동 계산식 2: 미수금 및 수금확인/수금날짜 자동 설정 - 레거시 동일
   */
  calculateOutstandingAmount(projectCode) {
    const financialCard = document.getElementById(`card-financial-${projectCode}`);
    const collectionCard = document.getElementById(`card-collection-${projectCode}`);
    if (!financialCard || !collectionCard) return;

    // 총액2 가져오기 - 총액2는 읽기 전용 필드이므로 항상 display에서 가져옴
    let totalAmount = 0;

    // 다양한 selector로 총액2 요소 찾기 시도
    let total2Display = financialCard.querySelector('[data-field="총액 2"].editable-value');
    if (!total2Display) {
      total2Display = financialCard.querySelector('[data-field="총액2"].editable-value');
    }
    if (!total2Display) {
      total2Display = financialCard.querySelector('.editable-value[data-field="총액 2"]');
    }
    if (!total2Display) {
      total2Display = financialCard.querySelector('.editable-value[data-field="총액2"]');
    }
    if (total2Display) {
      // 총액2는 계산 필드이므로 display 요소에서만 값 추출
      const displayText = total2Display.textContent || total2Display.innerText || '';
      totalAmount = parseFloat(displayText.replace(/[^0-9]/g, '') || '0') || 0;
      
    } else {
      logger.error(`❌ [미수금] 총액2 display 요소를 찾을 수 없음 (financial 카드: ${financialCard?.id})`);

      // 디버깅: financial 카드의 모든 요소 출력
      if (financialCard) {
        const allEditableValues = financialCard.querySelectorAll('.editable-value');      }
    }

    // 결제 금액들 가져오기 - 편집 모드와 표시 모드 모두 고려
    let contractAmount = 0, midAmount = 0, finalAmount = 0;

    const contractInput = collectionCard.querySelector('[data-field="계약금"] input');
    const contractDisplay = collectionCard.querySelector('[data-field="계약금"] .editable-value');
    if (contractInput && contractInput.offsetParent !== null) {
      contractAmount = parseFloat(contractInput.value.replace(/[^0-9]/g, '') || '0') || 0;
    } else if (contractDisplay) {
      contractAmount = parseFloat((contractDisplay.textContent || '').replace(/[^0-9]/g, '') || '0') || 0;
    }

    const midInput = collectionCard.querySelector('[data-field="중도금"] input');
    const midDisplay = collectionCard.querySelector('[data-field="중도금"] .editable-value');
    if (midInput && midInput.offsetParent !== null) {
      midAmount = parseFloat(midInput.value.replace(/[^0-9]/g, '') || '0') || 0;
    } else if (midDisplay) {
      midAmount = parseFloat((midDisplay.textContent || '').replace(/[^0-9]/g, '') || '0') || 0;
    }

    const finalInput = collectionCard.querySelector('[data-field="잔금"] input');
    const finalDisplay = collectionCard.querySelector('[data-field="잔금"] .editable-value');
    if (finalInput && finalInput.offsetParent !== null) {
      finalAmount = parseFloat(finalInput.value.replace(/[^0-9]/g, '') || '0') || 0;
    } else if (finalDisplay) {
      finalAmount = parseFloat((finalDisplay.textContent || '').replace(/[^0-9]/g, '') || '0') || 0;
    }

    // 미수금 계산
    let outstandingAmount = totalAmount - contractAmount - midAmount - finalAmount;

    // NaN/Infinity 검증 및 안전 처리
    if (!Number.isFinite(outstandingAmount)) {
      logger.warn(`[계산 필드 검증] 미수금 계산 오류: totalAmount=${totalAmount}, payments=${contractAmount + midAmount + finalAmount}, result=${outstandingAmount}`);
      outstandingAmount = 0;
    }

    // 음수 허용 (과수금 상태일 수 있음) - 하지만 극단적인 값은 제한
    if (Math.abs(outstandingAmount) > 1e15) {  // 1000조 이상의 비정상 값
      logger.warn(`[계산 필드 검증] 미수금이 비정상적으로 큼: ${outstandingAmount} → 0으로 처리`);
      outstandingAmount = 0;
    }

    

    // 미수금 필드 업데이트 (calculated-field이므로 display만 존재)
    const outstandingDisplay = collectionCard.querySelector('[data-field="미수금"].editable-value.calculated-field');

    if (outstandingDisplay) {
      const colorClass = outstandingAmount === 0 ? 'text-success fw-semibold' : 'text-danger fw-semibold';
      const displayValue = this.formatCurrency(outstandingAmount);
      outstandingDisplay.innerHTML = `<span class="${colorClass}">${displayValue}</span>`;
      
    }

    // 미수금 변경 시 수금확인 및 수금날짜 자동 설정 (개선: 모든 필드 변경에 반응)
    // 미수금에 영향을 주는 모든 필드: 총액1, 부가세, 계약금, 중도금, 잔금
    const fieldsAffectingOutstanding = ['총액 1', '총액1', '부가세', '계약금', '중도금', '잔금'];

    if (fieldsAffectingOutstanding.includes(this.lastTriggeredField)) {
      const confirmBadge = collectionCard.querySelector('[data-field="수금 확인"] .collection-status-badge');
      const dateInput = collectionCard.querySelector('[data-field="수금 날짜"] input');

      if (outstandingAmount === 0) {
        // 미수금이 0이면 수금확인 = 완료, 수금날짜 = 오늘
        if (confirmBadge) {
          confirmBadge.dataset.status = 'true';
          confirmBadge.textContent = '완료';
          confirmBadge.className = 'badge ms-2 collection-status-badge';
          confirmBadge.style.backgroundColor = '#d1edcc';
          confirmBadge.style.color = '#0f5132';

          // 토글 버튼도 함께 업데이트
          const collectionCheckbox = collectionCard.querySelector('[data-field="수금 확인"] .collection-checkbox');
          if (collectionCheckbox) {
            collectionCheckbox.checked = true;
          }

          // 🆕 EditState에 수금 확인 기록
          if (this.editState && this.editState.isActive) {
            this.editState.updateField('수금 확인', 'true');
            logger.debug('[EditState] 수금 확인 자동 설정: true');
          }
        }
        // 수금 날짜 자동 설정 (레거시 setCollectionDateToToday 동일)
        this.setCollectionDateToToday(projectCode);

        // 🆕 EditState에 수금 날짜 기록
        if (this.editState && this.editState.isActive) {
          const today = new Date().toISOString().split('T')[0];
          this.editState.updateField('수금 날짜', today);
          logger.debug('[EditState] 수금 날짜 자동 설정:', today);
        }
      } else {
        // 미수금이 0보다 크면 수금확인 = 대기, 수금날짜 클리어
        if (confirmBadge) {
          confirmBadge.dataset.status = 'false';
          confirmBadge.textContent = '대기';
          confirmBadge.className = 'badge ms-2 collection-status-badge';
          confirmBadge.style.backgroundColor = '#fff3cd';
          confirmBadge.style.color = '#664d03';

          // 토글 버튼도 함께 업데이트
          const collectionCheckbox = collectionCard.querySelector('[data-field="수금 확인"] .collection-checkbox');
          if (collectionCheckbox) {
            collectionCheckbox.checked = false;
          }

          // 🆕 EditState에 수금 확인 기록
          if (this.editState && this.editState.isActive) {
            this.editState.updateField('수금 확인', 'false');
            logger.debug('[EditState] 수금 확인 자동 설정: false');
          }
        }
        // 수금 날짜 초기화 (레거시 resetCollectionDate 동일)
        this.resetCollectionDate(projectCode);

        // 🆕 EditState에 수금 날짜 기록 (빈 값)
        if (this.editState && this.editState.isActive) {
          this.editState.updateField('수금 날짜', '');
          logger.debug('[EditState] 수금 날짜 초기화: (빈 값)');
        }
      }
    }
  }

  /**
   * 자동 계산식 2: 총액2 계산 (총액1 + 부가세) - 레거시 동일
   */
  calculateAmountFields(projectCode) {
    const financialCard = document.getElementById(`card-financial-${projectCode}`);
    if (!financialCard) return;

    // 총액1 값 가져오기
    const total1Field = financialCard.querySelector('[data-field="총액 1"], [data-field="총액1"]');
    let total1 = 0;
    if (total1Field) {
      const total1Text = total1Field.textContent || '';
      total1 = parseFloat(total1Text.replace(/[^\d.-]/g, '')) || 0;
    }

    // 부가세 포함 여부 확인
    const vatField = financialCard.querySelector('[data-field="부가세"]');
    let vatIncluded = false;
    if (vatField) {
      const vatText = vatField.textContent || '';
      vatIncluded = (vatText.includes('포함') && !vatText.includes('미포함')) ||
                   vatText.includes('VAT 포함') ||
                   vatField.classList.contains('vat-included');
    }

    // 총액2 계산: 부가세 포함이면 +10%, 아니면 그대로
    const total2 = vatIncluded ? total1 + Math.round(total1 * 0.1) : total1;

    // 총액2 필드 업데이트
    const total2Field = financialCard.querySelector('[data-field="총액 2"], [data-field="총액2"]');
    if (total2Field) {
      const formattedTotal2 = this.formatCurrency(total2);
      total2Field.textContent = formattedTotal2;

      // 계산 결과임을 표시
      total2Field.classList.add('calculated-field', 'text-primary');
      // total2Field.setAttribute('title', '자동 계산된 값 (총액1 + 부가세)'); // 툴팁 제거
    }
  }

  /**
   * 자동 계산식 3: 순익 및 마진율 계산 - 레거시 동일
   */
  calculateProfitFields(projectCode) {
    const financialCard = document.getElementById(`card-financial-${projectCode}`);
    const profitCard = document.getElementById(`card-profit-${projectCode}`);
    if (!financialCard || !profitCard) return;

    // 편집 모드인지 확인
    const isEditMode = profitCard.classList.contains('editing');

    // 편집 모드가 아니면 실행하지 않음 (API 값 사용)
    if (!isEditMode) {
      // 편집 모드가 아니므로 재계산 스킵
      return;
    }

    // 총액1 가져오기 (부가세 제외 수입)
    const total1Input = financialCard.querySelector('[data-field="총액 1"] input');
    const totalAmount = parseFloat(total1Input?.value?.replace(/[^0-9]/g, '') || '0') || 0;

    // 비용 항목들 가져오기
    const productInput = profitCard.querySelector('[data-field="제품대"] input');
    const contractInput = profitCard.querySelector('[data-field="도급비"] input');
    const materialInput = profitCard.querySelector('[data-field="자재비"] input');
    const otherInput = profitCard.querySelector('[data-field="기타비"] input');

    const productCost = parseFloat(productInput?.value?.replace(/[^0-9]/g, '') || '0') || 0;
    const contractCost = parseFloat(contractInput?.value?.replace(/[^0-9]/g, '') || '0') || 0;
    const materialCost = parseFloat(materialInput?.value?.replace(/[^0-9]/g, '') || '0') || 0;
    const otherCost = parseFloat(otherInput?.value?.replace(/[^0-9]/g, '') || '0') || 0;

    // 순익 계산
    const totalCosts = productCost + contractCost + materialCost + otherCost;
    let netProfit = totalAmount - totalCosts;

    // NaN/Infinity 검증 및 안전 처리
    if (!Number.isFinite(netProfit)) {
      logger.warn(`[계산 필드 검증] 순익 계산 오류: totalAmount=${totalAmount}, totalCosts=${totalCosts}, result=${netProfit}`);
      netProfit = 0;
    }

    // 마진율 계산
    let marginRate;
    if (totalAmount === 0) {
      // 총액이 0인 경우 마진율 0%
      marginRate = 0;
    } else if (totalAmount > 0) {
      marginRate = (netProfit / totalAmount) * 100;

      // NaN/Infinity 검증
      if (!Number.isFinite(marginRate)) {
        logger.warn(`[계산 필드 검증] 마진율 계산 오류: netProfit=${netProfit}, totalAmount=${totalAmount}, result=${marginRate}`);
        marginRate = 0;
      }

      // 극단적인 값 제한 (-1000% ~ 1000%)
      if (marginRate > 1000) {
        logger.warn(`[계산 필드 검증] 마진율이 1000%를 초과: ${marginRate.toFixed(1)}% → 1000%로 제한`);
        marginRate = 1000;
      } else if (marginRate < -1000) {
        logger.warn(`[계산 필드 검증] 마진율이 -1000% 미만: ${marginRate.toFixed(1)}% → -1000%로 제한`);
        marginRate = -1000;
      }
    } else {
      // 총액이 음수인 경우 (비정상)
      logger.warn(`[계산 필드 검증] 총액이 음수: ${totalAmount}, 마진율 0%로 처리`);
      marginRate = 0;
    }

    // UI 업데이트 (calculated-field이므로 display만 존재)
    const profitDisplay = profitCard.querySelector('[data-field="마진율"].editable-value.calculated-field');
    if (profitDisplay) {
      const colorClass = marginRate > 0 ? 'text-success fw-semibold' :
                        marginRate < 0 ? 'text-danger fw-semibold' : 'text-muted';
      profitDisplay.innerHTML = `<span class="${colorClass}">${marginRate.toFixed(1)}%</span>`;
    }

    // 순익 표시 (calculated-field이므로 display만 존재)
    const netProfitDisplay = profitCard.querySelector('[data-field="순익"].editable-value.calculated-field');
    if (netProfitDisplay) {
      const colorClass = netProfit > 0 ? 'text-success fw-semibold' :
                        netProfit < 0 ? 'text-danger fw-semibold' : 'text-muted';
      const displayValue = netProfit >= 0 ? this.formatCurrency(netProfit) : `-${this.formatCurrency(Math.abs(netProfit))}`;
      netProfitDisplay.innerHTML = `<span class="${colorClass}">${displayValue}</span>`;
    }
  }

  /**
   * 필드별 검증 (레거시 동일)
   */
  validateField(fieldName, value) {
    // 필수 필드 검증 (레거시 동일)
    const requiredFields = ['사업자', '발주처 담당자', '공사 구분'];
    if (requiredFields.includes(fieldName) && (!value || value.trim() === '')) {
      return { isValid: false, message: '필수 입력 항목입니다' };
    }

    // 이메일 검증 (레거시 동일)
    if (fieldName.includes('이메일')) {
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (value && value.trim() && !emailRegex.test(value)) {
        return { isValid: false, message: '올바른 이메일 형식을 입력하세요 (예: example@company.com)' };
      }
    }

    // 연락처 검증 (레거시 동일)
    if (fieldName.includes('연락처')) {
      if (value && value.trim()) {
        // 숫자와 하이픈만 허용
        const phoneRegex = /^[0-9-]+$/;
        if (!phoneRegex.test(value)) {
          return { isValid: false, message: '숫자와 하이픈만 입력 가능합니다' };
        }
        // 휴대폰 형식 확인
        const mobileRegex = /^01[0-9]-\d{4}-\d{4}$/;
        const landlineRegex = /^0[2-9][0-9]-\d{3,4}-\d{4}$/;
        if (!mobileRegex.test(value) && !landlineRegex.test(value)) {
          return { isValid: false, message: '올바른 전화번호 형식을 입력하세요 (예: 010-1234-5678)' };
        }
      }
    }

    // 금액 검증 (레거시 동일)
    if (fieldName.includes('총액') || fieldName.includes('제품대') || fieldName.includes('도급비') ||
        fieldName.includes('자재비') || fieldName.includes('기타비') || fieldName.includes('계약금') ||
        fieldName.includes('중도금') || fieldName.includes('잔금') || fieldName.includes('미수금')) {
      if (value && value.trim()) {
        const numericValue = value.replace(/[^0-9]/g, '');
        if (isNaN(parseInt(numericValue)) || numericValue === '') {
          return { isValid: false, message: '숫자만 입력 가능합니다' };
        }
        const amount = parseInt(numericValue);
        if (amount < 0) {
          return { isValid: false, message: '음수는 입력할 수 없습니다' };
        }
        if (amount > 999999999999) { // 1조 원 제한
          return { isValid: false, message: '입력 가능한 최대 금액을 초과했습니다' };
        }
      }
    }

    // 날짜 검증 (레거시 동일)
    if (fieldName.includes('날짜') || fieldName.includes('시작') || fieldName.includes('종료')) {
      if (value && value.trim()) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
          return { isValid: false, message: '올바른 날짜 형식(YYYY-MM-DD)을 입력하세요' };
        }
        // 날짜 유효성 검사
        const date = new Date(value);
        if (isNaN(date.getTime())) {
          return { isValid: false, message: '존재하지 않는 날짜입니다' };
        }
        // 미래 날짜 제한 (공사 시작/종료일 제외)
        if (!fieldName.includes('시작') && !fieldName.includes('종료')) {
          const today = new Date();
          today.setHours(23, 59, 59, 999);
          if (date > today) {
            return { isValid: false, message: '미래 날짜는 입력할 수 없습니다' };
          }
        }
      }
    }

    // 텍스트 길이 검증 (레거시 동일)
    if (value && typeof value === 'string') {
      if (fieldName.includes('현장 주소') && value.length > 200) {
        return { isValid: false, message: '주소는 200자 이내로 입력하세요' };
      }
      if (fieldName.includes('공사 내용') && value.length > 500) {
        return { isValid: false, message: '공사 내용은 500자 이내로 입력하세요' };
      }
      if (fieldName.includes('특이사항') && value.length > 1000) {
        return { isValid: false, message: '특이사항은 1000자 이내로 입력하세요' };
      }
    }

    return { isValid: true };
  }

  /**
   * 메시지 표시
   */
  showMessage(message, type) {
    // Toast 컴포넌트 사용 (이미 구현된 것 활용)
    if (window.projectToast) {
      window.projectToast.show(message, type);
    } else {
      // 폴백: 간단한 alert
      alert(message);
    }
  }

  /**
   * 이벤트 바인딩 (레거시 동일)
   */
  bindEvents() {
    // ESC 키로 아코디언 닫기 (편집 중이 아닌 경우만)
    // 중복 등록 방지
    if (!this.documentKeydownHandler) {
      this.documentKeydownHandler = (e) => {
        if (e.key === 'Escape' && this.isOpen) {
          // 편집 중인 필드가 있는지 확인
          const editingInputs = this.accordionContainer.querySelectorAll('.inline-edit-input');
          if (editingInputs.length === 0) {
            this.closeAccordion();
          }
        }
      };
      document.addEventListener('keydown', this.documentKeydownHandler);
    }
  }

  /**
   * 자동 계산 필드 데이터 검증 시스템
   * 계산된 값과 구글 시트 원본 값을 비교하여 불일치 시 경고 표시
   */
  validateCalculatedFields(projectCode, rowData) {
    if (!projectCode || !rowData) return;

    const calculatedFields = [
      {
        name: '총액 2',
        type: 'currency',
        calculator: () => this.calculateTotal2(rowData),
        original: rowData['총액 2'] || rowData['총액2'] || rowData['S'] || 0
      },
      {
        name: '미수금',
        type: 'currency',
        calculator: () => this.calculateOutstandingAmountValue(rowData),
        original: rowData['미수금'] || 0
      },
      {
        name: '마진율',
        type: 'percentage',
        calculator: () => this.calculateMarginRateValue(rowData),
        original: rowData['마진율'] || 0
      },
      {
        name: '순익',
        type: 'currency',
        calculator: () => this.calculateNetProfitValue(rowData),
        original: rowData['순익'] || 0
      }
    ];

    calculatedFields.forEach(field => {
      this.validateSingleField(projectCode, field);
    });
  }

  /**
   * 개별 필드 검증
   */
  validateSingleField(projectCode, field) {
    const calculatedValue = field.calculator();
    const originalValue = this.parseNumericValue(field.original, field.type);
    const isMatch = this.compareValues(calculatedValue, originalValue, field.type);

    // 해당 필드 요소 찾기
    const $fieldElement = document.querySelector(`.accordion-container [data-field="${field.name}"]`);

    if ($fieldElement) {
      // 기존 검증 아이콘 제거
      $fieldElement.querySelectorAll('.data-validation-icon').forEach(el => el.remove());

      // 검증 UI 비활성화 - 일반 텍스트처럼 깔끔하게 표시
      // if (!isMatch) {
      //   this.addDataValidationIcon($fieldElement, field, calculatedValue, originalValue);
      // }
    }
  }

  /**
   * 데이터 불일치 시 경고 아이콘 추가
   */
  addDataValidationIcon($fieldElement, field, calculatedValue, originalValue) {
    const iconHtml = `
      <span class="data-validation-icon ms-1"
            data-bs-toggle="tooltip"
            data-bs-placement="top"
            data-bs-html="true"
            title="<strong>[WARN] 데이터 불일치</strong><br/>
                   계산값: <span class='text-warning'>${this.formatValue(calculatedValue, field.type)}</span><br/>
                   원본값: <span class='text-info'>${this.formatValue(originalValue, field.type)}</span><br/>
                   <small class='text-muted'>데이터 동기화가 필요할 수 있습니다</small>"
            class="help-icon">
        <i class="fas fa-exclamation-circle text-warning help-icon"></i>
      </span>
    `;

    $fieldElement.insertAdjacentHTML('beforeend', iconHtml);

    // Bootstrap tooltip 초기화
    const tooltip = new bootstrap.Tooltip($fieldElement.querySelector('.data-validation-icon'));
  }

  /**
   * 계산 함수들 (검증용)
   */
  calculateTotal2(rowData) {
    const total1 = parseFloat(rowData['총액'] || rowData['총액 1'] || 0);
    const vatValue = rowData['부가세'];
    const vatIncluded = vatValue && typeof vatValue === 'string' && vatValue.includes('포함');

    // AmountCalculator 사용 (FLOOR + 끝자리 1/9원 보정 적용)
    return AmountCalculator.calculateWithVAT(total1, vatIncluded);
  }

  calculateOutstandingAmountValue(rowData) {
    const total = parseFloat(rowData['총액 2'] || rowData['총액2'] || rowData['S'] || rowData['총액'] || 0);
    const contract = parseFloat(rowData['계약금'] || 0);
    const interim = parseFloat(rowData['중도금'] || 0);
    const final = parseFloat(rowData['잔금'] || 0);

    return Math.max(0, total - contract - interim - final);
  }

  calculateMarginRateValue(rowData) {
    // Google Sheets와 동일하게 총액 1 기준
    const totalAmount = parseFloat(rowData['총액 1'] || rowData['총액1'] || rowData['총액'] || 0);
    const costs = parseFloat(rowData['제품대'] || 0) +
                 parseFloat(rowData['도급비'] || 0) +
                 parseFloat(rowData['자재비'] || 0) +
                 parseFloat(rowData['기타비'] || 0);

    if (totalAmount === 0) return 0;
    return ((totalAmount - costs) / totalAmount) * 100;
  }

  calculateNetProfitValue(rowData) {
    // Google Sheets와 동일하게 총액 1 기준
    const totalAmount = parseFloat(rowData['총액 1'] || rowData['총액1'] || rowData['총액'] || 0);
    const costs = parseFloat(rowData['제품대'] || 0) +
                 parseFloat(rowData['도급비'] || 0) +
                 parseFloat(rowData['자재비'] || 0) +
                 parseFloat(rowData['기타비'] || 0);

    return totalAmount - costs;
  }

  /**
   * 숫자 값 파싱
   */
  parseNumericValue(value, type) {
    if (!value) return 0;

    if (typeof value === 'string') {
      if (type === 'percentage') {
        return parseFloat(value.replace('%', '')) || 0;
      } else {
        return parseFloat(value.replace(/[^0-9.-]/g, '')) || 0;
      }
    }

    return parseFloat(value) || 0;
  }

  /**
   * 값 비교 (오차범위 고려)
   */
  compareValues(calculated, original, type) {
    const tolerance = type === 'percentage' ? 0.1 : 100; // 퍼센트는 0.1%, 금액은 100원 오차 허용
    return Math.abs(calculated - original) <= tolerance;
  }

  /**
   * 값 포맷팅
   */
  formatValue(value, type) {
    if (type === 'percentage') {
      return value.toFixed(1) + '%';
    } else {
      return this.formatCurrency(value);
    }
  }

  /**
   * 실시간 계산 트리거 - 변경된 필드에 따라 적절한 계산 실행
   */
  triggerRealTimeCalculations(projectCode, changedFields) {
    if (!projectCode || !changedFields || changedFields.length === 0) return;

    // 금액 관련 필드가 변경된 경우 총액2 계산
    const amountFields = ['총액 1', '총액1', '부가세'];
    const shouldCalculateAmount = changedFields.some(field =>
      amountFields.some(amountField => field.includes(amountField))
    );

    // 손익 관련 필드가 변경된 경우 순익/마진율 계산
    const profitFields = ['총액 1', '총액1', '제품대', '도급비', '자재비', '기타비'];
    const shouldCalculateProfit = changedFields.some(field =>
      profitFields.some(profitField => field.includes(profitField))
    );

    // 미수금 관련 필드가 변경된 경우 미수금 계산 (향후 구현)
    const receivableFields = ['총액 2', '총액2', '계약금', '중도금', '잔금'];
    const shouldCalculateReceivable = changedFields.some(field =>
      receivableFields.some(receivableField => field.includes(receivableField))
    );

    

    // 계산 실행 (순서 중요: 총액2 → 순익 → 미수금)
    if (shouldCalculateAmount) {      this.calculateAmountFields(projectCode);
    }

    if (shouldCalculateProfit) {      this.calculateProfitFields(projectCode);
    }

    if (shouldCalculateReceivable) {      this.calculateOutstandingAmount(projectCode);
    }

    // 계산 완료 후 시각적 피드백 (계산된 필드 강조)
    setTimeout(() => {
      const calculatedFields = ['.calculated-field', '[data-field="총액 2"]', '[data-field="순익"]', '[data-field="마진율"]'];
      calculatedFields.forEach(selector => {
        const element = document.querySelector(`#accordion-${projectCode} ${selector}`);
        if (element) {
          element.classList.add('calculation-highlight');
          setTimeout(() => element.classList.remove('calculation-highlight'), 1500);
        }
      });
    }, 100);
  }

  /**
   * 폴더 편집 모드 토글
   */
  toggleFolderEditMode(documentCard, enable) {
    if (!documentCard) return;

    const cardRow = documentCard.querySelector('.legacy-card-row');
    const cardEdit = documentCard.querySelector('.legacy-card-edit');
    const input = cardEdit?.querySelector('.folder-path-input');

    if (enable) {
      // 편집 모드로 전환
      cardRow?.classList.add('d-none');
      cardEdit?.classList.remove('d-none');
      input?.focus();
      
      // Google Drive 링크 자동 파싱 이벤트 추가
      if (input && !input.dataset.folderParsingAttached) {
        // 폴더 경로 필드는 folder-path-input 클래스 기반이라 다른 필드처럼
        // .editable-value 부모의 dataset.field로 잡히는 EditState 위임 리스너를
        // 거치지 않는다. 여기서 명시적으로 EditState.updateField 호출해야 저장 시
        // 변경사항으로 인식됨 (2026-07-08 매니저 배포 첫날 실측 사고 대응).
        const FOLDER_FIELD = '견적서 및 계약서 폴더 경로';
        const syncEditState = () => {
          this.parseDriveFolderUrl(input);
          if (this.editState && this.editState.isActive) {
            this.editState.updateField(FOLDER_FIELD, input.value.trim());
          }
        };
        input.addEventListener('input', syncEditState);
        input.addEventListener('paste', () => {
          // paste는 브라우저가 값 세팅한 뒤에 잡아야 하니 미세 지연
          setTimeout(syncEditState, 10);
        });
        input.dataset.folderParsingAttached = 'true';

        // 편집 모드 진입 시 기존 값 검증
        if (input.value.trim()) {
          setTimeout(() => this.parseDriveFolderUrl(input), 100);
        }
      }
    } else {
      // 보기 모드로 전환
      cardRow?.classList.remove('d-none');
      cardEdit?.classList.add('d-none');
    }
  }

  /**
   * 폴더 경로 저장
   */
  saveFolderPath(documentCard) {
    if (!documentCard) return;

    const input = documentCard.querySelector('.folder-path-input');
    const valueDisplay = documentCard.querySelector('[data-role="folder-display"]');
    const projectCode = this.currentProject?.['프로젝트 코드'];

    if (!input || !valueDisplay || !projectCode) return;

    const newPath = input.value.trim();
    const fieldName = '견적서 및 계약서 폴더 경로';

    // UI 업데이트
    valueDisplay.textContent = newPath || '폴더 경로가 설정되지 않았습니다.';
    valueDisplay.title = newPath;

    // 카드 스타일 업데이트
    if (newPath) {
      documentCard.classList.remove('document-card-empty');
    } else {
      documentCard.classList.add('document-card-empty');
    }

    // 편집 모드 종료
    this.toggleFolderEditMode(documentCard, false);

    // 데이터는 이미 통합 편집 모드의 EditState를 통해 관리됨
  }

  /**
   * Google Drive URL에서 폴더 ID 자동 추출
   */
  parseDriveFolderUrl(input) {
    console.log('[parseDriveFolderUrl] 함수 호출됨, input:', input);
    if (!input) {
      console.error('[parseDriveFolderUrl] input이 없습니다!');
      return;
    }

    const value = input.value.trim();
    console.log('[parseDriveFolderUrl] 입력 값:', value);

    // 체크 아이콘 찾기
    const wrapper = input.closest('.folder-input-wrapper');
    const successIcon = wrapper?.querySelector('.folder-success-icon');
    console.log('[parseDriveFolderUrl] wrapper 찾기:', wrapper);
    console.log('[parseDriveFolderUrl] successIcon 찾기:', successIcon);

    // 값이 비어있으면 주황색 테두리, 아이콘 숨김
    if (!value) {
      console.log('[parseDriveFolderUrl] 값이 비어있습니다 - 주황색 테두리 + 아이콘 숨김');
      input.style.borderColor = '#fd7e14'; // 주황색
      input.classList.remove('valid-folder-id'); // 유효 상태 제거
      if (successIcon) {
        successIcon.style.display = 'none';
        console.log('[parseDriveFolderUrl] 아이콘 숨김 완료');
      } else {
        console.warn('[parseDriveFolderUrl] successIcon을 찾을 수 없습니다!');
      }
      return;
    }

    // 1. 로컬 경로 감지 (C:\ through I:\) - 경고 및 주황색 테두리
    const localPathPattern = /^[C-I]:[\\/]/i;
    if (localPathPattern.test(value)) {
      logger.warn('[폴더 경로] 로컬 경로는 권장하지 않습니다:', value);
      input.style.borderColor = '#fd7e14'; // 주황색
      input.classList.remove('valid-folder-id'); // 유효 상태 제거
      if (successIcon) {
        successIcon.style.display = 'none';
      }
      return;
    }

    // 2. Google Drive URL에서 폴더 ID 추출
    let folderId = null;

    // 형식 1: https://drive.google.com/drive/folders/{ID}
    const urlPattern1 = /drive\.google\.com\/drive\/(?:u\/\d+\/)?folders\/([a-zA-Z0-9_-]+)/;
    const urlMatch1 = value.match(urlPattern1);

    // 형식 2: https://drive.google.com/open?id={ID} (윈도우 탐색기 "링크를 클립보드로 복사")
    const urlPattern2 = /drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/;
    const urlMatch2 = value.match(urlPattern2);

    if (urlMatch1) {
      folderId = urlMatch1[1];
    } else if (urlMatch2) {
      folderId = urlMatch2[1];
    }

    if (folderId) {
      // 폴더 ID로 자동 변경
      console.log('[parseDriveFolderUrl] 폴더 ID 추출 성공:', folderId);
      input.value = folderId;
      logger.info('[폴더 경로] Google Drive 링크에서 폴더 ID 추출:', folderId);

      // 초록색 테두리 + 체크 아이콘 표시 (저장할 때까지 유지)
      input.style.borderColor = '#23923c'; // 초록색
      input.style.transition = 'border-color 0.3s';
      input.classList.add('valid-folder-id'); // 유효 상태 추가 (hover 방지)

      if (successIcon) {
        successIcon.style.display = 'block';
        successIcon.style.opacity = '1';
      }

      return;
    }

    // 3. 이미 폴더 ID 형식인 경우 (Google Drive ID는 정확히 33자)
    const folderIdPattern = /^[a-zA-Z0-9_-]{33}$/;
    if (folderIdPattern.test(value)) {
      logger.debug('[폴더 경로] 올바른 폴더 ID 형식:', value);
      // 올바른 ID 형식 - 초록색 테두리 + 체크 아이콘
      input.style.borderColor = '#23923c';
      input.style.transition = 'border-color 0.3s';
      input.classList.add('valid-folder-id'); // 유효 상태 추가 (hover 방지)
      if (successIcon) {
        successIcon.style.display = 'block';
        successIcon.style.opacity = '1';
      }
      return;
    }

    // 잘못된 형식 - 주황색 테두리, 아이콘 숨김
    console.log('[parseDriveFolderUrl] 잘못된 형식 - 주황색 테두리 + 아이콘 숨김');
    input.style.borderColor = '#fd7e14';
    input.classList.remove('valid-folder-id'); // 유효 상태 제거
    if (successIcon) {
      successIcon.style.display = 'none';
      console.log('[parseDriveFolderUrl] 아이콘 숨김 완료');
    } else {
      console.warn('[parseDriveFolderUrl] successIcon을 찾을 수 없습니다!');
    }
  }


  /**
   * 노트 편집 모드 토글
   */
  toggleNotesEditMode(collectionCard, enable) {
    if (!collectionCard) return;

    const cardRow = collectionCard.querySelector('.legacy-card-row');
    const cardEdit = collectionCard.querySelector('.legacy-card-edit');
    const textarea = cardEdit?.querySelector('.collection-notes-input');

    if (enable) {
      // 편집 모드로 전환
      cardRow?.classList.add('d-none');
      cardEdit?.classList.remove('d-none');
      textarea?.focus();
    } else {
      // 보기 모드로 전환
      cardRow?.classList.remove('d-none');
      cardEdit?.classList.add('d-none');
    }
  }

  /**
   * 수금 특이사항 저장
   */
  saveCollectionNotes(collectionCard) {
    if (!collectionCard) return;

    const textarea = collectionCard.querySelector('.collection-notes-input');
    const valueDisplay = collectionCard.querySelector('[data-role="notes-display"]');
    const projectCode = this.currentProject?.['프로젝트 코드'];

    if (!textarea || !valueDisplay || !projectCode) return;

    const newNotes = textarea.value.trim();
    const fieldName = '수금 관련 특이사항';

    // UI 업데이트
    valueDisplay.textContent = newNotes || '특이사항이 없습니다.';
    valueDisplay.title = newNotes;

    // 카드 스타일 업데이트
    if (newNotes) {
      collectionCard.classList.remove('collection-card-empty');
    } else {
      collectionCard.classList.add('collection-card-empty');
    }

    // 편집 모드 종료
    this.toggleNotesEditMode(collectionCard, false);

    // 데이터는 이미 통합 편집 모드의 EditState를 통해 관리됨
  }

  /**
   * 모든 검증 아이콘 제거
   */
  removeAllValidationIcons() {
    if (this.accordionContainer) {
      const validationIcons = this.accordionContainer.querySelectorAll('.data-validation-icon');
      validationIcons.forEach(icon => {
        // Bootstrap tooltip 제거
        const tooltip = bootstrap.Tooltip.getInstance(icon);
        if (tooltip) {
          tooltip.dispose();
        }
        icon.remove();
      });
    }
  }

  // ==============================
  // 통합 편집 모드 메서드들
  // ==============================

  /**
   * 통합 편집 모드 진입 (권한별 카드 활성화)
   */
  async enableUnifiedEditMode(projectCode) {


    // 통합 편집 진입 시 현재 프로젝트 코드로 originalProjectCode 재설정
    // (이전 저장에서 코드가 변경되었을 수 있으므로 최신 상태로 동기화)
    const currentCode = this.currentProject?.['프로젝트 코드'] || projectCode;
    if (this.originalProjectCode !== currentCode) {

      this.originalProjectCode = currentCode;
    }

    // 원본 메모 값 저장 (임시 저장 비교용)
    this.originalMemos = {
      '계약금_메모': this.currentProject?.['계약금_메모'] || '',
      '중도금_메모': this.currentProject?.['중도금_메모'] || '',
      '잔금_메모': this.currentProject?.['잔금_메모'] || ''
    };
    logger.debug('[편집 모드] 원본 메모 저장:', this.originalMemos);

    // 🆕 메모 변경사항 추적 초기화
    this.pendingMemoChanges = {};
    logger.debug('[편집 모드] 메모 변경사항 추적 초기화');

    // 세션 체크는 authInterceptor.js가 자동으로 처리하므로 별도 체크 불필요
    // 모든 fetch 요청에서 401 응답 시 자동으로 로그인 페이지로 리다이렉트됨
    // (이전에는 여기서 /api/projects/list?limit=1 호출로 세션 체크를 했으나,
    //  불필요한 API 요청으로 편집 모드 진입이 느려지는 문제가 있었음)

    const userRole = window.userRole || 'viewer';

    if (userRole === 'viewer') {
      this.showMessage('편집 권한이 없습니다.', 'error');
      return;
    }

    // 2026-07-08 draft 복구 체크 — 이전 편집 세션에서 저장된 미완료 값이 있으면 안내
    try {
      const draftKey = `itg_draft_${projectCode}`;
      const draftRaw = sessionStorage.getItem(draftKey);
      if (draftRaw) {
        const draft = JSON.parse(draftRaw);
        const draftAt = new Date(draft._savedAt || 0);
        const ageMin = (Date.now() - draftAt.getTime()) / 60000;
        // 30분 이내 draft만 복구 제안 (오래된 것은 stale 위험)
        if (ageMin < 30 && confirm(
          `이전 편집 세션의 미저장 데이터가 있습니다.\n` +
          `저장 시각: ${draftAt.toLocaleString('ko-KR')} (${Math.round(ageMin)}분 전)\n\n` +
          `복구하시겠습니까?\n(취소하면 draft 삭제)`
        )) {
          this._pendingDraft = draft;
        } else {
          sessionStorage.removeItem(draftKey);
        }
      }
    } catch (e) {
      logger.debug('[편집 모드] draft 복구 시도 오류 (무시):', e);
    }

    // draft 자동 저장 시작 (5초 주기, 편집 종료 시 clear)
    if (this._draftAutoSaveTimer) clearInterval(this._draftAutoSaveTimer);
    this._draftAutoSaveTimer = setInterval(() => {
      try {
        const shell = this.getAccordionShell?.();
        if (!shell) return;
        const inputs = shell.querySelectorAll('input[data-field], textarea[data-field], select[data-field]');
        if (inputs.length === 0) return;
        const snapshot = { _savedAt: new Date().toISOString(), _projectCode: projectCode, fields: {} };
        inputs.forEach(el => {
          const field = el.dataset.field;
          if (field) snapshot.fields[field] = el.value;
        });
        sessionStorage.setItem(`itg_draft_${projectCode}`, JSON.stringify(snapshot));
      } catch (e) {
        logger.debug('[편집 모드] draft 자동 저장 실패 (무시):', e);
      }
    }, 5000);

    // 프로젝트 잠금 획득 시도
    try {
      logger.debug(`🔒 [잠금] 프로젝트 잠금 획득 시도: ${projectCode} (탭 ID: ${this.tabId})`);
      const lockResponse = await fetch('/api/project-lock/acquire', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({
          project_code: projectCode,
          tab_id: this.tabId
        })
      });

      const lockResult = await lockResponse.json();

      if (!lockResult.success) {
        // 다른 사용자가 편집 중
        logger.warn(`⚠️ [잠금] 잠금 획득 실패: ${lockResult.message}`);
        this.showMessage(lockResult.message, 'warning');
        return;
      }

      logger.debug(`✅ [잠금] 프로젝트 잠금 획득 성공: ${projectCode}`);

      // Heartbeat 시작 (5분마다 자동 연장)
      this.startLockHeartbeat(projectCode);
    } catch (error) {
      logger.error('❌ [잠금] 잠금 획득 중 오류:', error);
      this.showMessage('편집 모드 진입 중 오류가 발생했습니다.', 'error');
      return;
    }

    // 🆕 EditState 초기화 (잠금 획득 성공 후)
    this.editState = new EditState();
    this.editState.initialize(this.currentProject);
    logger.debug('[EditState] 편집 상태 초기화 완료');

    // A안: EditState 초기화 검증 (실패 시 편집 모드 진입 차단)
    if (!this.editState || !this.editState.isActive) {
      logger.error('[EditState] ❌ 초기화 실패 - 편집 모드를 시작할 수 없습니다.');
      this.showMessage('편집 시스템 오류가 발생했습니다. 페이지를 새로고침한 후 다시 시도해주세요.', 'error');

      // 획득한 락 즉시 해제
      await this.releaseProjectLock(projectCode);
      return;
    }

    // UI 수정을 try-catch로 감싸서 오류 시 잠금 해제 (레이스 컨디션 방지)
    try {
      // 모든 카드 목록 (권한 없는 카드도 포함하여 워터마크 표시용)
      const allCardTypes = ['basic', 'construction', 'financial', 'collection', 'profit'];



      // 1. 헤더 필드를 편집 모드로 전환
      const headerContainer = document.getElementById(`header-${projectCode}`);
      if (headerContainer) {        headerContainer.innerHTML = this.generateHeaderFields(projectCode, this.currentProject, true);
        headerContainer.classList.add('editing');

        // 헤더 필드 이벤트 바인딩 (프로젝트 코드 실시간 업데이트)
        this.bindHeaderFieldEvents(projectCode);
      }

      // 2. 각 카드를 편집 모드로 전환 (권한 없는 카드는 워터마크 표시)
      allCardTypes.forEach(cardType => {
        const card = document.getElementById(`card-${cardType}-${projectCode}`);
        if (card) {
          this.enableCardEditing(projectCode, cardType, true); // 버튼 업데이트 스킵
        }
      });

      // 3. 문서 폴더와 수금 특이사항도 권한에 따라 활성화
      if (this.canEditCard('document', userRole)) {
        const documentCard = this.accordionContainer.querySelector('.document-card');
        if (documentCard) {
          documentCard.classList.add('editing');
          const editableField = documentCard.querySelector('.editable-value');
          if (editableField) {
            const fieldName = editableField.dataset.field;
            const currentValue = editableField.dataset.originalValue || '';
            const inputElement = this.createInputElement(fieldName, currentValue);
            editableField.innerHTML = inputElement;
            
            // Google Drive 링크 자동 파싱 + EditState 동기화 이벤트 추가
            // (2026-07-08 fix: 폴더 경로 붙여넣기 → 저장 시 '변경사항 없음' 오인식 사고 대응.
            //  통합 편집 모드는 enableCardEditing을 안 거쳐서 카드 레벨 위임 리스너가 없다.
            //  folder-path-input에서 직접 EditState.updateField 호출해야 dirtyFields에 반영됨.)
            setTimeout(() => {
              const folderInput = editableField.querySelector('.folder-path-input');
              if (folderInput) {
                const syncFolderEditState = () => {
                  this.parseDriveFolderUrl(folderInput);
                  if (this.editState && this.editState.isActive) {
                    this.editState.updateField(fieldName, folderInput.value.trim());
                  }
                };
                folderInput.addEventListener('input', syncFolderEditState);
                folderInput.addEventListener('paste', () => {
                  // paste는 브라우저가 값 세팅 후에 잡아야 함
                  setTimeout(syncFolderEditState, 10);
                });

                // 편집 모드 진입 시 기존 값 검증
                if (folderInput.value.trim()) {
                  this.parseDriveFolderUrl(folderInput);
                }
              } else {
                logger.warn('[문서 폴더] folder-path-input을 찾을 수 없음');
              }
            }, 100);
          }
        }
      }

      if (this.canEditCard('notes', userRole)) {
        const collectionCard = this.accordionContainer.querySelector('.collection-card');
        if (collectionCard) {
          collectionCard.classList.add('editing');
          const editableField = collectionCard.querySelector('.editable-value');
          if (editableField) {
            const fieldName = editableField.dataset.field;
            const currentValue = editableField.dataset.originalValue || '';
            const inputElement = this.createInputElement(fieldName, currentValue);
            editableField.innerHTML = inputElement;
          }
        }
      }

      // 통합 버튼 상태 변경
      this.updateUnifiedButtons(projectCode, 'editing');

      // 편집 모드 시 공사 취소/재개 버튼 숨김
      const shell = this.accordionContainer.querySelector(`.accordion-shell[data-project-code="${projectCode}"]`);
      if (shell) {
        const cancelResumeBtn = shell.querySelector('.construction-action-btn');
        if (cancelResumeBtn) {
          cancelResumeBtn.style.display = 'none';
        }
      }

      // 🆕 ModeManager를 통한 편집 모드 활성화
      const modeChangeSuccess = this.modeManager.setAccordionMode(ACCORDION_MODE.EDIT);
      if (!modeChangeSuccess) {
        throw new Error('편집 모드 전환에 실패했습니다.');
      }

      // 🔥 Tom Select 멀티셀렉트 초기화 전에 기존 드롭다운 인스턴스 정리
      setTimeout(() => {
        // 기존 Bootstrap Dropdown 인스턴스 제거
        allCardTypes.forEach(cardType => {
          const card = document.getElementById(`card-${cardType}-${projectCode}`);
          if (card) {
            const dropdownButtons = card.querySelectorAll('.dropdown-toggle');
            dropdownButtons.forEach(button => {
              const bsDropdown = bootstrap.Dropdown.getInstance(button);
              if (bsDropdown) {
                bsDropdown.dispose();
              }
            });
          }
        });

        // 새로운 드롭다운 초기화
        this.initializeTomSelectMultiSelects();

        // 2026-07-08 draft 실제 적용 — _pendingDraft(line 4715)를 실제 폼 필드에 주입.
        // 이전엔 confirm에서 '확인'을 눌러도 값이 채워지지 않아 매니저가 다시 처음부터
        // 입력해야 했음 (반쪽 구현). 편집 UI가 완전히 렌더된 이 시점에 주입해야 안전.
        if (this._pendingDraft && this._pendingDraft.fields) {
          try {
            const shell = this.accordionContainer.querySelector(
              `.accordion-shell[data-project-code="${projectCode}"]`
            );
            let applied = 0;
            if (shell) {
              Object.entries(this._pendingDraft.fields).forEach(([field, value]) => {
                if (value === undefined || value === null) return;
                const el = shell.querySelector(
                  `input[data-field="${field}"], textarea[data-field="${field}"], select[data-field="${field}"]`
                );
                if (el) {
                  el.value = value;
                  // 값 변경 이벤트 발생 — Tom Select·계산식·validation 등이 반응해야 함
                  el.dispatchEvent(new Event('input', { bubbles: true }));
                  el.dispatchEvent(new Event('change', { bubbles: true }));
                  applied++;
                }
              });
            }
            logger.info(`[편집 모드] draft 값 ${applied}개 폼 필드에 복구 완료`);
            if (applied > 0) {
              this.showMessage(`이전 미저장 편집 내용 ${applied}개 필드 복구됨`, 'info');
            }
          } catch (draftErr) {
            logger.warn('[편집 모드] draft 폼 적용 실패:', draftErr);
          } finally {
            this._pendingDraft = null;
          }
        }
      }, 100);

    } catch (uiError) {
      // UI 수정 중 오류 발생 시 잠금 해제 및 롤백
      logger.error('❌ [UI 수정] 편집 모드 UI 변경 중 오류 발생:', uiError);

      // 잠금 해제 시도
      try {
        this.stopLockHeartbeat();
        await fetch('/api/project-lock/release', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          credentials: 'same-origin',
          body: JSON.stringify({
            project_code: projectCode,
            tab_id: this.tabId
          })
        });
        logger.debug(`✅ [롤백] 잠금 해제 완료 (UI 오류로 인한 롤백)`);
      } catch (releaseError) {
        logger.error('❌ [롤백] 잠금 해제 실패:', releaseError);
      }

      this.showMessage('편집 모드 UI 로드 중 오류가 발생했습니다. 페이지를 새로고침 해주세요.', 'error');
      return;
    }

    // 🆕 beforeunload 경고 활성화 (편집 중 페이지 이탈 방지)
    window.addEventListener('beforeunload', this.beforeUnloadHandler);
    logger.debug('✅ [beforeunload] 페이지 이탈 경고 활성화');

  }

  /**
   * 통합 버튼 상태 업데이트
   */
  updateUnifiedButtons(projectCode, mode) {
    // 최신 코드로 먼저 찾고, 없으면 원본 코드로 찾기 (코드 변경 시 대응)
    let buttonsContainer = this.accordionContainer.querySelector(`.unified-edit-buttons[data-project-code="${projectCode}"]`);
    if (!buttonsContainer && this.originalProjectCode) {
      buttonsContainer = this.accordionContainer.querySelector(`.unified-edit-buttons[data-project-code="${this.originalProjectCode}"]`);
    }
    if (!buttonsContainer) {
      logger.warn(`⚠️ [통합 버튼] 버튼 컨테이너를 찾을 수 없음: ${projectCode}`);
      return;
    }

    const editBtn = buttonsContainer.querySelector('.unified-edit-btn');
    const saveBtn = buttonsContainer.querySelector('.unified-save-btn');
    const cancelBtn = buttonsContainer.querySelector('.unified-cancel-btn');

    if (mode === 'editing') {
      editBtn.classList.add('d-none');
      saveBtn.classList.remove('d-none');
      cancelBtn.classList.remove('d-none');
    } else {
      editBtn.classList.remove('d-none');
      saveBtn.classList.add('d-none');
      cancelBtn.classList.add('d-none');
    }
  }

  /**
   * 전체 카드 변경사항 수집
   */
  collectAllChanges(projectCode) {

    // 🆕 EditState가 활성화되어 있으면 EditState.collectChanges() 사용 (단일 진실 소스)
    if (this.editState && this.editState.isActive) {
      logger.debug('[EditState] ✅ EditState.collectChanges() 사용 - 단일 진실 소스에서 변경사항 수집');
      const changes = this.editState.collectChanges();

      // 🆕 메모 변경사항 병합 (FieldMemoButton을 통해 저장된 메모)
      if (Object.keys(this.pendingMemoChanges).length > 0) {
        Object.assign(changes, this.pendingMemoChanges);
        logger.debug(`[메모 병합] ${Object.keys(this.pendingMemoChanges).length}개 메모 변경사항 추가`);
      }

      logger.debug('[EditState] 수집된 변경사항 (메모 포함):', changes);
      return changes;
    }

    // 레거시 로직 (EditState 비활성 시 fallback)
    // ⚠️ EditState가 비활성화된 경우 편집 시스템 오류로 간주
    logger.error('[EditState] ❌ 비정상 상태: EditState가 비활성화되었습니다. 편집이 불가능합니다.');
    logger.error('[EditState] isEditMode:', this.modeManager.isEditMode());
    logger.error('[EditState] editState:', this.editState);
    logger.error('[EditState] editState.isActive:', this.editState?.isActive);

    // 🆕 메시지는 호출하는 쪽(try-catch)에서 표시하도록 함 (중복 방지)
    // 편집 차단 (저장 불가)
    throw new Error('EditState가 비활성화되어 저장할 수 없습니다.');
  }

  /**
   * 통합 저장 (한 번에 전송, 아코디언 유지)
   */
  async saveAllChanges(projectCode) {
    // 동시 저장 요청 큐잉 처리
    if (this.isSavingInProgress) {
      // 큐에 이미 같은 프로젝트가 있으면 제거 (최신 요청만 유지)
      this.saveQueue = this.saveQueue.filter(item => item.projectCode !== projectCode);

      // 큐에 추가
      return new Promise((resolve, reject) => {
        this.saveQueue.push({
          projectCode,
          resolve,
          reject,
          timestamp: Date.now()
        });

        const saveBtn = this.accordionContainer.querySelector('.unified-save-btn');
        if (saveBtn) {
          saveBtn.innerHTML = `<i class="fas fa-clock me-1"></i>대기 중 (${this.saveQueue.length})`;
        }
      });
    }

    // 저장 진행 중 플래그 설정
    this.isSavingInProgress = true;

    try {
      await this._executeSave(projectCode);
    } catch (error) {
      throw error;
    } finally {
      // 성공/실패 여부와 관계없이 플래그 해제
      this.isSavingInProgress = false;

      // 큐에 대기 중인 요청 처리 (플래그 해제 후)
      this._processNextInQueue();
    }
  }

  /**
   * 실제 저장 로직 실행
   */
  async _executeSave(projectCode) {

    // 변경사항 수집 (EditState 에러 시 자동 편집 종료)
    let allChanges;
    try {
      allChanges = this.collectAllChanges(projectCode);
    } catch (error) {
      logger.error('[collectAllChanges] ❌ 변경사항 수집 중 오류:', error);
      this.showMessage('편집 시스템 오류가 발생했습니다. 편집 모드를 종료합니다.', 'error');

      // 🆕 자동으로 편집 모드 종료 (사용자가 직접 취소하지 않아도 됨)
      this.disableUnifiedEditMode(projectCode);
      throw error; // 저장 프로세스 중단
    }    

    // 날짜 범위 검증 (공사 시작 <= 공사 종료)
    const startDate = allChanges['공사 시작'] || this.currentProject['공사 시작'];
    const endDate = allChanges['공사 종료'] || this.currentProject['공사 종료'];

    if (startDate && endDate) {
      const rangeValidation = validateDateRange(startDate, endDate);
      if (!rangeValidation.valid) {
        this.showMessage(`날짜 검증 실패: ${rangeValidation.error}`, 'error');
        throw new Error(`날짜 검증 실패: ${rangeValidation.error}`);
      }
    }

    const saveBtn = this.accordionContainer.querySelector('.unified-save-btn');
    const cancelBtn = this.accordionContainer.querySelector('.unified-cancel-btn');

    if (Object.keys(allChanges).length === 0) {
      // 변경사항 없음 - 버튼에 피드백 표시 후 편집 모드 해제
      saveBtn.innerHTML = '<i class="fas fa-info-circle me-1"></i>변경사항 없음';
      saveBtn.classList.remove('btn-outline-success');
      saveBtn.classList.add('btn-secondary');

      setTimeout(() => {
        // 버튼 원래 상태로 복구
        saveBtn.innerHTML = '<i class="fas fa-check me-1"></i>저장';
        saveBtn.classList.remove('btn-secondary');
        saveBtn.classList.add('btn-outline-success');

        // 편집 모드 해제
        this.cancelAllChanges(projectCode);
      }, 800);

      return;  // 플래그 해제는 finally에서
    }

    if (!saveBtn || !cancelBtn) {
      logger.error('❌ [통합 저장] 버튼을 찾을 수 없습니다.');
      throw new Error('저장 버튼을 찾을 수 없습니다.');
    }

    // 버튼 비활성화 및 로딩 상태
    saveBtn.disabled = true;
    cancelBtn.disabled = true;
    const originalSaveText = saveBtn.innerHTML;
    saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i>저장중...';

    // 메모 필드 분리 (별도 API로 처리) - catch 블록에서도 접근 가능하도록 try 밖에서 선언
    const memoChanges = {};
    const projectChanges = {};

    try {

      Object.keys(allChanges).forEach(key => {
        if (key.endsWith('_메모')) {
          memoChanges[key] = allChanges[key];
        } else {
          projectChanges[key] = allChanges[key];
        }
      });

      logger.debug('[통합 저장] 프로젝트 변경:', Object.keys(projectChanges));
      logger.debug('[통합 저장] 메모 변경:', Object.keys(memoChanges));

      // 1. 프로젝트 필드 저장 (메모 제외)
      let result = null;
      if (Object.keys(projectChanges).length > 0) {
        const apiProjectCode = this.originalProjectCode || projectCode;

        // 🔒 Optimistic Lock: 현재 버전 추가
        const currentVersion = this.currentProject?._version || '0';
        projectChanges._version = currentVersion;
        logger.debug(`[Optimistic Lock] 버전 포함 저장 요청: ${currentVersion}`);

        const startTime = performance.now();
        const response = await fetch(`/api/projects/${apiProjectCode}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          credentials: 'same-origin',
          body: JSON.stringify(projectChanges)
        });

        const endTime = performance.now();
        logger.debug(`⏱️ [통합 저장] 프로젝트 필드 저장 API 소요 시간: ${(endTime - startTime).toFixed(0)}ms`);

        // 🔒 Optimistic Lock: 409 Conflict 처리 (병합 UI)
        if (response.status === 409) {
          const conflictResult = await response.json();
          logger.warn(`[Optimistic Lock] 버전 충돌 감지: ${conflictResult.message}`);

          // 최신 데이터로 UI 업데이트 (서버에서 current_data 반환)
          if (conflictResult.current_data) {
            this.updateAllCardsWithProjectData(projectCode, conflictResult.current_data);
            this.currentProject = conflictResult.current_data;
            logger.info('[409 병합] 최신 데이터로 UI 업데이트 완료');
          }

          // 병합 UI: 사용자에게 선택권 제공
          const userChoice = confirm(
            '⚠️ 다른 사용자가 이 프로젝트를 먼저 수정했습니다.\n\n' +
            '최신 버전으로 업데이트되었습니다.\n' +
            '계속 편집하시겠습니까?\n\n' +
            '[확인] 계속 편집 (최신 데이터 기준)\n' +
            '[취소] 편집 모드 종료'
          );

          if (userChoice) {
            // 계속 편집: 편집 모드 유지, 최신 데이터로 작업
            this.showMessage(
              '최신 버전으로 업데이트되었습니다. 계속 편집할 수 있습니다.',
              'info',
              3000
            );
            logger.info('[409 병합] 사용자가 계속 편집 선택');
            // 편집 모드 유지 - 아무것도 하지 않음
          } else {
            // 편집 종료: 편집 모드 해제
            this.disableUnifiedEditMode(projectCode);
            this.showMessage(
              '편집 모드가 종료되었습니다.',
              'info',
              3000
            );
            logger.info('[409 병합] 사용자가 편집 종료 선택');
          }

          // 충돌 에러를 throw하여 저장 프로세스 중단
          throw new Error('version_conflict');
        }

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        result = await response.json();

        if (!result.success) {
          throw new Error(result.message || '저장에 실패했습니다.');
        }
      } else {
        // 프로젝트 변경사항이 없으면 가짜 result 생성
        // 메모만 변경된 경우를 대비하여 현재 프로젝트 데이터 사용
        logger.debug('[통합 저장] 프로젝트 필드 변경 없음, 메모만 업데이트');
        result = { success: true, project: null };  // null로 설정하여 나중에 currentProject 사용
      }

      // 프로젝트 코드가 변경되었는지 확인 (담당자/사업자 변경 시)
      const finalProjectCode = result?.project_code || projectCode;
      if (finalProjectCode !== projectCode) {
        logger.debug(`[통합 저장] 프로젝트 코드 변경 감지: ${projectCode} → ${finalProjectCode}`);
      }

      // 2. 메모 저장 (Batch API - 순차 처리로 SSL 에러 방지)
      if (Object.keys(memoChanges).length > 0) {
        saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i>메모 저장중...';

        // 메모 배열 구성
        const memos = Object.entries(memoChanges).map(([memoKey, memoValue]) => ({
          field_name: memoKey.replace('_메모', ''),
          memo: memoValue || null
        }));

        logger.debug(`[통합 저장] Batch API로 ${memos.length}개 메모 저장 요청 (프로젝트 코드: ${finalProjectCode})`);

        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 30000); // 타임아웃 30초 (3개 메모 * 10초)

        try {
          const memoStartTime = performance.now();
          const batchResponse = await fetch('/api/projects/field-memos/batch', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            credentials: 'same-origin',
            signal: controller.signal,
            body: JSON.stringify({
              project_code: finalProjectCode,  // 변경된 프로젝트 코드 사용
              memos: memos
            })
          });

          clearTimeout(timeoutId);
          const memoEndTime = performance.now();
          logger.debug(`⏱️ [통합 저장] 메모 일괄 저장 API 소요 시간: ${(memoEndTime - memoStartTime).toFixed(0)}ms`);

          if (!batchResponse.ok) {
            const errorText = await batchResponse.text();
            throw new Error(`메모 일괄 저장 실패: HTTP ${batchResponse.status} - ${errorText}`);
          }

          const batchResult = await batchResponse.json();

          // 결과 로깅
          logger.debug(`[통합 저장] Batch 결과: ${batchResult.message}`);
          logger.debug(`[통합 저장] 성공: ${batchResult.success_count}개, 실패: ${batchResult.failed_count}개`);

          // 실패한 메모가 있으면 에러 발생
          if (batchResult.failed_count > 0) {
            const failedFields = batchResult.results
              .filter(r => !r.success)
              .map(r => `${r.field_name} (${r.message})`)
              .join(', ');
            throw new Error(`일부 메모 저장 실패: ${failedFields}`);
          }

          // 성공한 메모들의 로컬 데이터 업데이트
          Object.entries(memoChanges).forEach(([memoKey, memoValue]) => {
            this.currentProject[memoKey] = memoValue;
          });

          logger.debug(`[통합 저장] 모든 메모 저장 완료: ${batchResult.success_count}개`);

        } catch (error) {
          clearTimeout(timeoutId);

          if (error.name === 'AbortError') {
            throw new Error(`메모 저장 타임아웃: 30초 초과`);
          }

          // 에러를 상위로 전파
          throw error;
        }

        saveBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i>저장중...';
      }

      if (result.success) {
        // 성공 애니메이션 - 테두리에서 채움 초록색으로 변경
        // 🆕 Google Sheets 반영 상태 UI 피드백
        saveBtn.innerHTML = '<i class="fas fa-database me-1"></i>DB 저장 완료 · <i class="fas fa-table me-1"></i>Sheets 반영 중...';
        saveBtn.classList.remove('btn-outline-success');
        saveBtn.classList.add('btn-success');
        saveBtn.style.transform = 'scale(1.05)';
        setTimeout(() => {
          saveBtn.style.transform = 'scale(1)';
        }, 200);

        // 0.3초 후 완전 완료 메시지로 변경
        setTimeout(() => {
          saveBtn.innerHTML = '<i class="fas fa-check-double me-1"></i>저장 완료!';
        }, 300);

        

        // 서버 응답을 반영하여 currentProject 업데이트
        let updatedProjectData;
        if (result.project) {
          const memoBackup = {};
          const memoFields = ['계약금_메모', '중도금_메모', '잔금_메모'];
          memoFields.forEach(field => {
            if (this.currentProject[field] !== undefined) {
              memoBackup[field] = this.currentProject[field];
              logger.debug(`[메모 백업] ${field} = "${this.currentProject[field]}"`);
            }
          });

          updatedProjectData = { ...result.project };
          Object.assign(updatedProjectData, memoBackup);
          logger.debug('[저장 완료] 서버 데이터 반영 (메모 복원 포함)');
        } else {
          updatedProjectData = { ...this.currentProject };
          Object.assign(updatedProjectData, projectChanges);
          logger.debug('[저장 완료] 서버가 project 데이터를 반환하지 않아 변경분만 반영');
        }

        this.currentProject = updatedProjectData;

        if (result.project_code) {
          this.currentProject['프로젝트 코드'] = result.project_code;
        }

        logger.debug('[저장 완료] currentProject 갱신 완료:', this.currentProject);

        // 프로젝트 코드 변경 확인
        const finalProjectCode = this.currentProject['프로젝트 코드'] || projectCode;
        const projectCodeChanged = result.old_project_code && result.old_project_code !== finalProjectCode;

        if (projectCodeChanged) {
          logger.debug(`[코드 변경] ${result.old_project_code} → ${finalProjectCode}`);

          // 1. 구 코드로 획득한 락 즉시 해제 (새 코드로 락이 남지 않도록)
          this.releaseProjectLock(result.old_project_code).catch(error => {
            logger.warn('⚠️ [잠금] 구 코드 락 해제 중 오류 (무시됨):', error);
          });

          // 2. DataTable 행 업데이트
          const rowUpdateSuccess = this.updateDataTableRow(result.old_project_code, finalProjectCode, this.currentProject);

          // 3. StateManager 전역 상태 업데이트 (프로젝트 코드 변경 시에도 필수)
          if (window.projectListApp?.stateManager) {
            const updateSuccess = window.projectListApp.stateManager.updateSingleProject(
              finalProjectCode,
              this.currentProject,
              result.old_project_code  // 이전 프로젝트 코드로 검색
            );
            if (updateSuccess) {
              logger.debug('[데이터 동기화] StateManager 업데이트 완료 (프로젝트 코드 변경됨)');
            } else {
              logger.warn('[데이터 동기화] StateManager 업데이트 실패 - 프로젝트를 찾을 수 없습니다');
            }
          }

          // 4. DOM 내 data-project-code 및 ID 동기화
          this.syncDomProjectCode(result.old_project_code, finalProjectCode);

          // 5. 아코디언 행을 새 위치로 이동 (정렬 변경 대응)
          const repositioned = this.moveAccordionRow(finalProjectCode);

          const projectUpdatedDetail = {
            partialUpdate: true,
            project: this.currentProject,
            projectCode: finalProjectCode,
            oldProjectCode: result.old_project_code
          };

          if (!rowUpdateSuccess || !repositioned) {
            logger.warn('[코드 변경] 아코디언 위치 재배치에 실패하여 안전 모드로 재오픈 수행');
            this.reopenAccordion(finalProjectCode, this.currentProject);

            window.dispatchEvent(new CustomEvent('projectUpdated', { detail: projectUpdatedDetail }));

            saveBtn.innerHTML = '<i class="fas fa-check me-1"></i>저장';
            saveBtn.classList.remove('btn-success');
            saveBtn.classList.add('btn-outline-success');
            saveBtn.disabled = false;
            cancelBtn.disabled = false;

            this.disableUnifiedEditMode(finalProjectCode, true);
            return;
          } else {
            // 새 행에 활성화 스타일 적용
            const targetRow = this.dataTable?.row(`#row-${finalProjectCode}`).node();
            if (targetRow) {
              document.querySelectorAll('tbody tr').forEach(row => row.classList.remove('table-active'));
              targetRow.classList.add('table-active');
            }
          }

          // 6. 편집 모드 정리 (이후 DOM 복원 작업 없음)
          this.disableUnifiedEditMode(finalProjectCode, true);

          // 7. 카드 및 헤더를 최신 데이터로 갱신
          this.updateAllCardsWithProjectData(finalProjectCode, this.currentProject);

          const headerContainerForCodeChange = document.getElementById(`header-${finalProjectCode}`);
          if (headerContainerForCodeChange) {
            headerContainerForCodeChange.innerHTML = this.generateHeaderFields(finalProjectCode, this.currentProject, false);
          }

          // 8. 이벤트 발송 (다른 컴포넌트 동기화용)
          window.dispatchEvent(new CustomEvent('projectUpdated', { detail: projectUpdatedDetail }));

          // 저장 버튼 상태 복원
          saveBtn.innerHTML = '<i class="fas fa-check me-1"></i>저장';
          saveBtn.classList.remove('btn-success');
          saveBtn.classList.add('btn-outline-success');
          saveBtn.disabled = false;
          cancelBtn.disabled = false;

          return;
        }

        // 저장된 값으로 카드 업데이트 및 편집 모드 해제
        // 버튼 원래 상태로 복구
        saveBtn.innerHTML = '<i class="fas fa-check me-1"></i>저장';
        saveBtn.classList.remove('btn-success');
        saveBtn.classList.add('btn-outline-success');
        saveBtn.disabled = false;
        cancelBtn.disabled = false;

        // 서버에서 반환된 전체 프로젝트 데이터로 업데이트 (계산 필드 포함)
        logger.debug('[저장 완료] 서버 응답:', {
          hasProject: !!result.project,
          project: result.project,
          memoChanges: Object.keys(memoChanges)
        });

        // 최신 프로젝트 코드 사용 (실시간 변경 대응)
        const currentCode = this.currentProject['프로젝트 코드'] || projectCode;

        if (currentCode !== projectCode) {
          this.syncDomProjectCode(projectCode, currentCode);
        }

        // 1. 편집 모드 정리 (heartbeat, beforeunload, 이벤트 정리 등)
        this.disableUnifiedEditMode(currentCode, true);

        // 2. 서버 데이터로 모든 카드 재렌더링 (메모 아이콘 포함)
        this.updateAllCardsWithProjectData(currentCode, this.currentProject);

        // 3. 헤더도 서버 데이터로 재렌더링
        const headerContainer = document.getElementById(`header-${currentCode}`);
        if (headerContainer) {
          headerContainer.innerHTML = this.generateHeaderFields(currentCode, this.currentProject, false);
        }

        // DataTable 행 업데이트 (모든 필드 변경 반영)
        // 프로젝트 코드가 변경되지 않은 경우에도 다른 필드 변경사항을 테이블에 반영
        this.updateDataTableRow(currentCode, currentCode, this.currentProject);

        // localStorage 캐시 무효화 (F5 새로고침 시 최신 데이터 표시되도록)
        try {
          const namespace = 'itg_dashboard';
          const keysToRemove = [];
          for (let key in localStorage) {
            if (key.startsWith(namespace)) {
              keysToRemove.push(key);
            }
          }
          keysToRemove.forEach(key => localStorage.removeItem(key));
          logger.debug('[캐시 무효화] localStorage 캐시 삭제 완료:', keysToRemove.length);
        } catch (error) {
          logger.warn('[캐시 무효화] localStorage 클리어 실패:', error);
        }

        // StateManager 전역 상태 업데이트 (토글 시 최신 데이터 표시 보장)
        if (window.projectListApp?.stateManager) {
          const updateSuccess = window.projectListApp.stateManager.updateSingleProject(
            currentCode,
            this.currentProject
          );
          if (updateSuccess) {
            logger.debug('[데이터 동기화] StateManager 업데이트 완료 - 토글 시 최신 데이터 표시됨');
          } else {
            logger.warn('[데이터 동기화] StateManager 업데이트 실패 - 프로젝트를 찾을 수 없습니다');
          }
        }

        // 부분 업데이트 이벤트 발송 (아코디언 유지)
        window.dispatchEvent(new CustomEvent('projectUpdated', {
          detail: { partialUpdate: true, project: this.currentProject, projectCode: currentCode }
        }));

        // Toast 제거 - 저장 버튼의 시각적 피드백만으로 충분
        // this.showMessage('변경사항이 저장되었습니다.', 'success');

      } else {
        throw new Error(result.message || '저장에 실패했습니다.');
      }

    } catch (error) {
      // 상세 에러 로그 (디버깅용)
      logger.error('❌ [통합 저장] 저장 실패:', {
        errorMessage: error.message,
        errorStack: error.stack,
        projectCode: projectCode,
        hasProjectChanges: Object.keys(projectChanges).length > 0,
        projectChangesKeys: Object.keys(projectChanges),
        hasMemoChanges: Object.keys(memoChanges).length > 0,
        memoChangesKeys: Object.keys(memoChanges),
        url: window.location.href,
        timestamp: new Date().toISOString()
      });

      // 오류 상태 복원
      saveBtn.innerHTML = originalSaveText;
      saveBtn.disabled = false;
      cancelBtn.disabled = false;

      // 사용자 친화적 에러 메시지 표시
      const friendlyError = getProjectSaveError(error);
      this.showMessage(
        `${friendlyError.message}\n\n💡 ${friendlyError.action}`,
        'error'
      );

      // 상세 에러는 콘솔에만 기록
      logger.error('[상세 에러]', friendlyError.originalError);
    }
    // 플래그 해제는 saveAllChanges()의 finally에서 처리
  }

  /**
   * 큐에 대기 중인 다음 저장 요청 처리
   */
  _processNextInQueue() {
    if (this.saveQueue.length === 0) {
      
      return;
    }

    const nextRequest = this.saveQueue.shift();
    // 다음 저장 실행
    this.saveAllChanges(nextRequest.projectCode)
      .then(() => nextRequest.resolve())
      .catch((error) => nextRequest.reject(error));
  }

  /**
   * 모든 카드를 서버 데이터로 업데이트
   */
  updateAllCardsWithProjectData(projectCode, projectData) {


    let updatedCount = 0;

    ProjectRowAccordion.CARD_TYPES.forEach(cardType => {
      const card = document.getElementById(`card-${cardType}-${projectCode}`);
      if (card) {
        this.updateCardWithProjectData(projectCode, cardType, projectData);
        updatedCount++;
      }
    });

    
  }

  /**
   * 통합 취소 (모든 카드 원본 복원, 아코디언 유지)
   */
  cancelAllChanges(projectCode) {



    // 먼저 각 카드의 모든 드롭다운 버튼 찾아서 닫기
    let closedCount = 0;
    ProjectRowAccordion.CARD_TYPES.forEach(cardType => {
      const card = document.getElementById(`card-${cardType}-${projectCode}`);
      if (card) {
        const dropdownButtons = card.querySelectorAll('.dropdown-toggle');
        dropdownButtons.forEach(button => {
          const bsDropdown = bootstrap.Dropdown.getInstance(button);
          if (bsDropdown) {
            bsDropdown.hide();
            closedCount++;
          }
        });
      }
    });
    logger.debug(`🔽 [통합 취소] ${closedCount}개 드롭다운 닫기 완료`);

    // 모든 카드를 원본 값으로 복원
    ProjectRowAccordion.CARD_TYPES.forEach(cardType => {
      const card = document.getElementById(`card-${cardType}-${projectCode}`);
      if (card && card.classList.contains('editing')) {
        this.cancelCardEditing(projectCode, cardType);
      }
    });

    // 문서 폴더 복원
    const documentCard = this.accordionContainer.querySelector('.document-card');
    if (documentCard && documentCard.classList.contains('editing')) {
      const editableField = documentCard.querySelector('.editable-value');
      if (editableField) {
        const originalValue = editableField.dataset.originalValue || '';
        const projectCode = this.currentProject?.['프로젝트 코드'] || '';
        editableField.innerHTML = this.renderFolderLink(originalValue, projectCode);
      }
      documentCard.classList.remove('editing');
    }

    // 수금 특이사항 복원
    const collectionCard = this.accordionContainer.querySelector('.collection-card');
    if (collectionCard && collectionCard.classList.contains('editing')) {
      const editableField = collectionCard.querySelector('.editable-value');
      if (editableField) {
        const originalValue = editableField.dataset.originalValue || '';
        editableField.innerHTML = originalValue || '특이사항이 없습니다.';
      }
      collectionCard.classList.remove('editing');
    }

    // 편집 모드 해제 (헤더 필드도 data-original-value로 복원됨)
    this.disableUnifiedEditMode(projectCode);

    
  }

  /**
   * 통합 편집 모드 해제
   */
  disableUnifiedEditMode(projectCode, skipDataRestore = false) {

    // 2026-07-08 draft 자동 저장 timer 종료 + 저장된 draft 삭제 (편집 완료·취소 시)
    if (this._draftAutoSaveTimer) {
      clearInterval(this._draftAutoSaveTimer);
      this._draftAutoSaveTimer = null;
    }
    try {
      sessionStorage.removeItem(`itg_draft_${projectCode}`);
    } catch (e) { /* ignore */ }

    // 현재 프로젝트 코드 사용 (실시간 업데이트로 변경되었을 수 있음)
    const currentCode = this.currentProject?.['프로젝트 코드'] || projectCode;

    // 🆕 모든 카드의 드롭다운 정리 (저장 후 편집 모드 해제 시)
    let disposedCount = 0;
    ProjectRowAccordion.CARD_TYPES.forEach(cardType => {
      const card = document.getElementById(`card-${cardType}-${currentCode}`);
      if (card) {
        const dropdownButtons = card.querySelectorAll('.dropdown-toggle');
        dropdownButtons.forEach(button => {
          const bsDropdown = bootstrap.Dropdown.getInstance(button);
          if (bsDropdown) {
            // 드롭다운 닫고 인스턴스 제거
            bsDropdown.hide();
            bsDropdown.dispose();
            disposedCount++;
          }
        });

        // 드롭다운 메뉴를 원래 위치로 복귀 (body에 추가된 경우)
        const dropdownMenus = card.querySelectorAll('.multi-select-dropdown .dropdown-menu');
        dropdownMenus.forEach(menu => {
          const dropdown = menu.closest('.multi-select-dropdown');
          if (dropdown && menu.parentElement === document.body) {
            dropdown.appendChild(menu);
          }
        });
      }
    });
    if (disposedCount > 0) {
      logger.debug(`🔽 [편집 모드 해제] ${disposedCount}개 드롭다운 정리 완료`);
    }

    // 🆕 beforeunload 경고 비활성화
    window.removeEventListener('beforeunload', this.beforeUnloadHandler);
    logger.debug('✅ [beforeunload] 페이지 이탈 경고 비활성화');

    // Heartbeat 정지
    this.stopLockHeartbeat();

    // 🆕 EditState 정리 (편집 모드 해제 시)
    if (this.editState && this.editState.isActive) {
      this.editState.destroy();
      this.editState = null;
      logger.debug('[EditState] 편집 상태 정리 완료');
    }

    // 🆕 메모 변경사항 추적 초기화
    this.pendingMemoChanges = {};
    logger.debug('[편집 모드 해제] 메모 변경사항 추적 초기화');

    // 🆕 메모 원본으로 복원 (취소 시)
    if (!skipDataRestore && this.originalMemos && this.currentProject) {
      const currentProjectCode = this.currentProject['프로젝트 코드'];
      logger.debug('[메모 복원] 원본 메모로 복원 시작:', this.originalMemos);

      // 모든 데이터 소스의 메모를 원본으로 복원
      ['계약금_메모', '중도금_메모', '잔금_메모'].forEach(memoKey => {
        const originalMemo = this.originalMemos[memoKey] || '';

        // 1. DataTables 캐시 복원
        if (window.projectListApp?.components?.table?.table) {
          const table = window.projectListApp.components.table.table;
          const rows = table.rows().data().toArray();
          rows.forEach((row, idx) => {
            if (row['프로젝트 코드'] === currentProjectCode) {
              row[memoKey] = originalMemo;
              table.row(idx).data(row);
            }
          });
        }

        // 2. filters.currentData 복원
        if (window.projectListApp?.components?.filters?.currentData) {
          const targetProject = window.projectListApp.components.filters.currentData.find(
            p => p['프로젝트 코드'] === currentProjectCode
          );
          if (targetProject) targetProject[memoKey] = originalMemo;
        }

        // 3. StateManager.currentData 복원
        if (window.projectListApp?.stateManager?.currentData) {
          const targetProject = window.projectListApp.stateManager.currentData.find(
            p => p['프로젝트 코드'] === currentProjectCode
          );
          if (targetProject) targetProject[memoKey] = originalMemo;
        }

        // 4. window.projectsData 복원
        if (window.projectsData) {
          const targetProject = window.projectsData.find(
            p => p['프로젝트 코드'] === currentProjectCode
          );
          if (targetProject) targetProject[memoKey] = originalMemo;
        }

        // 5. currentProject 복원
        this.currentProject[memoKey] = originalMemo;
      });

      logger.debug('[메모 복원] 원본 메모 복원 완료');
    }

    // 🆕 모든 카드의 EditState 리스너 플래그 제거 (중복 등록 방지 플래그 초기화)
    ProjectRowAccordion.CARD_TYPES.forEach(cardType => {
      const card = document.getElementById(`card-${cardType}-${currentCode}`);
      if (card && card.dataset.editstateListenerAttached) {
        delete card.dataset.editstateListenerAttached;
      }
    });

    // 🆕 헤더 컨테이너의 EditState 리스너 플래그 제거
    const headerContainerForCleanup = document.getElementById(`header-${currentCode}`);
    if (headerContainerForCleanup && headerContainerForCleanup.dataset.editstateListenerAttached) {
      delete headerContainerForCleanup.dataset.editstateListenerAttached;
    }

    logger.debug('[EditState] 카드/헤더 리스너 플래그 정리 완료');

    // Tom Select 인스턴스 정리 (네이티브 멀티셀렉트는 정리 불필요)
    this.tomSelectInstances = [];

    // 1. 헤더 필드 처리
    const headerContainer = document.getElementById(`header-${currentCode}`);
    if (headerContainer && headerContainer.classList.contains('editing')) {
      if (!skipDataRestore) {
        // 취소 시: data-original-value로 복원
        const managerField = headerContainer.querySelector('[data-field="담당자"]');
        const addressField = headerContainer.querySelector('[data-field="현장 주소"]');
        const contentField = headerContainer.querySelector('[data-field="공사 내용"]');

        const restoredData = {
          '담당자': managerField?.dataset.originalValue || this.currentProject['담당자'] || '',
          '현장 주소': addressField?.dataset.originalValue || this.currentProject['현장 주소'] || '',
          '공사 내용': contentField?.dataset.originalValue || this.currentProject['공사 내용'] || ''
        };

        // 원본 프로젝트 코드로 헤더 생성 (취소 시 원래 코드로 복원)
        headerContainer.innerHTML = this.generateHeaderFields(this.originalProjectCode, restoredData, false);
        headerContainer.id = `header-${this.originalProjectCode}`;
      }

      headerContainer.classList.remove('editing');
    }

    // 2. 모든 카드의 편집 모드 해제 (cardTypes 재사용)
    ProjectRowAccordion.CARD_TYPES.forEach(cardType => {
      // 현재 코드로 카드 찾기 (실시간 업데이트로 변경되었을 수 있음)
      let card = document.getElementById(`card-${cardType}-${currentCode}`);

      // 현재 코드로 못 찾으면 원본 코드로 시도
      if (!card) {
        card = document.getElementById(`card-${cardType}-${this.originalProjectCode}`);
      }

      // readonly 카드는 워터마크만 제거
      if (card && card.classList.contains('readonly-card')) {
        const watermark = card.querySelector('.readonly-watermark');
        if (watermark) {
          watermark.remove();
        }
        card.classList.remove('readonly-card');
        return;
      }

      if (card && card.classList.contains('editing')) {
        

        if (!skipDataRestore) {
          // 취소 시: .editable-value를 읽기 전용 뷰로 복원
          card.querySelectorAll('.editable-value').forEach(field => {
            const fieldName = field.dataset.field;
            const editBackup = field.getAttribute('data-edit-original-value');
            const originalValue = field.dataset.originalValue;

            // 1. 백업된 HTML이 있으면 복원
            if (editBackup !== null) {
              field.innerHTML = editBackup;
              return;
            }

            // 2. 부가세 뱃지 복원
            if (fieldName === '부가세') {
              const vatValue = originalValue === 'true';
              const badge = this.unifiedBadgeSystem
                ? this.unifiedBadgeSystem.createBadge('vat', vatValue ? '포함' : '미포함')
                : (vatValue ? '포함' : '미포함');
              field.innerHTML = badge;
              return;
            }

            // 3. 수금 확인 뱃지 복원
            if (fieldName === '수금 확인') {
              const collectionValue = originalValue === 'true';
              const status = collectionValue ? '완료' : '대기';
              field.innerHTML = this.formatCollectionStatus(status);
              return;
            }

            // 4. data-original-value로 복원
            if (originalValue !== undefined) {
              let displayValue = originalValue;
              let colorClass = '';

              // 필드 타입별 포맷팅
              if (fieldName.includes('총액') || fieldName.includes('제품대') || fieldName.includes('도급비') ||
                  fieldName.includes('자재비') || fieldName.includes('기타비') || fieldName.includes('잔금') ||
                  fieldName.includes('계약금') || fieldName.includes('중도금')) {
                displayValue = this.formatCurrency(displayValue);
              } else if (fieldName === '미수금') {
                displayValue = this.formatCurrency(displayValue);
                const numValue = parseFloat(displayValue.replace(/[^0-9]/g, '')) || 0;
                colorClass = numValue === 0 ? 'text-success fw-bold' : 'text-danger fw-bold';
              } else if (fieldName === '순익' || fieldName === '마진율') {
                const numValue = parseFloat(displayValue) || 0;
                if (fieldName === '순익') {
                  displayValue = this.formatCurrency(displayValue);
                }
                colorClass = numValue > 0 ? 'text-success fw-bold' :
                            numValue < 0 ? 'text-danger fw-bold' : 'text-muted';
              } else if (fieldName.includes('날짜') || fieldName.includes('시작') || fieldName.includes('종료')) {
                displayValue = this.formatDate(displayValue);
              }

              if (colorClass) {
                field.innerHTML = `<span class="${colorClass}">${displayValue}</span>`;
              } else {
                field.innerHTML = displayValue || '-';
              }
              return;
            }

            // 5. 마지막 수단: 현재 input/select 값 사용
            const activeSelect = field.querySelector('select');
            if (activeSelect && activeSelect.selectedOptions.length > 0) {
              field.innerHTML = activeSelect.selectedOptions[0].textContent;
            } else {
              const activeInput = field.querySelector('input, textarea');
              field.innerHTML = activeInput ? activeInput.value : '';
            }
          });
        }

        // 백업 속성 정리 (취소 시에만)
        if (!skipDataRestore) {
          card.querySelectorAll('.editable-value').forEach(field => {
            field.removeAttribute('data-edit-original-value');
          });

          // 메모 필드 아이콘 제거 (취소 시에만)
          if (cardType === 'collection') {
            card.querySelectorAll('.editable-value[data-field="계약금"], .editable-value[data-field="중도금"], .editable-value[data-field="잔금"]').forEach(field => {
              const icon = field.querySelector('i.fa-sticky-note');
              if (icon) {
                icon.remove();
              }
              field.removeAttribute('data-bs-toggle');
              field.removeAttribute('data-bs-title');
              field.style.cursor = '';
            });
          }
        }

        // editing 클래스 제거 (취소 시에만 - 저장 시에는 이미 updateCardWithProjectData에서 제거됨)
        if (!skipDataRestore) {
          card.classList.remove('editing');
        } else {
          card.classList.remove('editing');
        }

      }

      // 카드 ID 복원 (취소 시만)
      if (!skipDataRestore && card && currentCode !== this.originalProjectCode) {
        card.id = `card-${cardType}-${this.originalProjectCode}`;
      }
    });

    // currentProject 프로젝트 코드 복원 (취소 시만)
    if (!skipDataRestore && this.currentProject) {
      this.currentProject['프로젝트 코드'] = this.originalProjectCode;
    }

    // 문서 폴더 editing 클래스 제거
    const documentCard = this.accordionContainer.querySelector('.document-card');
    if (documentCard && documentCard.classList.contains('editing')) {

      // 취소 시에만 원래 값으로 복원 (저장 시에는 updateAllCardsWithProjectData가 처리)
      if (!skipDataRestore) {
        const editableField = documentCard.querySelector('.editable-value');
        if (editableField) {
          const input = editableField.querySelector('input, textarea');
          if (input) {
            const originalValue = editableField.dataset.originalValue || '';
            const projectCode = this.currentProject?.['프로젝트 코드'] || '';
            editableField.innerHTML = this.renderFolderLink(originalValue, projectCode);
          }
        }
      }

      documentCard.classList.remove('editing');

    }

    // 수금 특이사항 editing 클래스 제거
    const collectionCard = this.accordionContainer.querySelector('.collection-card');
    if (collectionCard && collectionCard.classList.contains('editing')) {

      // 취소 시에만 원래 값으로 복원 (저장 시에는 updateAllCardsWithProjectData가 처리)
      if (!skipDataRestore) {
        const editableField = collectionCard.querySelector('.editable-value');
        if (editableField) {
          const input = editableField.querySelector('input, textarea');
          if (input) {
            const originalValue = editableField.dataset.originalValue || '';
            editableField.innerHTML = originalValue || '특이사항이 없습니다.';
          }
        }
      }

      collectionCard.classList.remove('editing');

    }

    // 통합 버튼 상태 변경
    this.updateUnifiedButtons(projectCode, 'view');

    // 편집 모드 해제 시 공사 취소/재개 버튼 다시 표시
    const shell = this.accordionContainer.querySelector(`.accordion-shell[data-project-code="${projectCode}"]`);
    if (shell) {
      const cancelResumeBtn = shell.querySelector('.construction-action-btn');
      if (cancelResumeBtn) {
        cancelResumeBtn.style.display = '';
      }
    }

    // 툴팁 인스턴스가 dispose된 뒤 뷰 모드 마크업으로 복원되므로 재초기화
    this.initializeAccordionTooltips();

    // 🆕 ModeManager를 통한 편집 모드 해제
    this.modeManager.setAccordionMode(ACCORDION_MODE.VIEW);

    // 원본 메모 참조 정리
    this.originalMemos = null;

    // 프로젝트 잠금 해제 (항상 originalProjectCode로 해제)
    // 프로젝트 코드가 저장 중 변경되었더라도 락은 원본 코드로 획득했으므로
    // 원본 코드로 해제해야 락이 풀립니다.
    const lockCodeToRelease = this.originalProjectCode || projectCode;
    this.releaseProjectLock(lockCodeToRelease).catch(error => {
      logger.warn('⚠️ [잠금] 잠금 해제 중 오류 (무시됨):', error);
    });

  }

  /**
   * 프로젝트 잠금 해제 헬퍼 함수
   */
  async releaseProjectLock(projectCode) {
    try {
      const response = await fetch('/api/project-lock/release', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({
          project_code: projectCode,
          tab_id: this.tabId
        })
      });

      const result = await response.json();

      if (result.success) {
        logger.debug(`🔓 [잠금] 프로젝트 잠금 해제 성공: ${projectCode}`);
      } else {
        logger.warn(`⚠️ [잠금] 잠금 해제 실패: ${result.message}`);
      }

      return result;
    } catch (error) {
      logger.error('❌ [잠금] 잠금 해제 중 오류:', error);
      throw error;
    }
  }

  /**
   * 편집 버튼 잠금 상태 업데이트
   */
  async updateEditButtonLockStatus(projectCode) {
    try {
      // 잠금 상태 조회
      const response = await fetch(`/api/project-lock/status/${projectCode}`, {
        credentials: 'same-origin'
      });
      const result = await response.json();

      if (!result.success) {
        return;
      }

      const lockInfo = result.lock_info;
      const isLocked = result.is_locked;

      // 편집 버튼 찾기
      const editBtn = this.accordionContainer.querySelector('.unified-edit-btn');
      if (!editBtn) {
        return;
      }

      if (isLocked) {
        // 다른 사용자가 편집 중
        const lockedByName = lockInfo.user_name;
        editBtn.disabled = true;
        editBtn.innerHTML = `<i class="fas fa-lock me-1"></i>${lockedByName}님이 편집 중...`;
        editBtn.title = `${lockedByName}님이 현재 이 프로젝트를 편집하고 있습니다`;
      } else {
        // 잠금 없음 - 편집 가능
        editBtn.disabled = false;
        editBtn.innerHTML = '<i class="fas fa-edit"></i><span>편집</span>';
        editBtn.title = '편집';
      }
    } catch (error) {
      logger.warn('⚠️ [잠금] 잠금 상태 확인 중 오류 (무시됨):', error);
    }
  }

  /**
   * 컴포넌트 정리
   */
  destroy() {
    // ModeManager 이벤트 리스너 제거 (메모리 누수 방지)
    if (this.boundHandleAccordionModeChange) {
      this.modeManager.off('accordionModeChanged', this.boundHandleAccordionModeChange);
      this.boundHandleAccordionModeChange = null;
    }

    // Heartbeat 타이머 정리 (메모리 누수 방지)
    this.stopLockHeartbeat();

    // 모든 input 필드의 계산 핸들러 제거
    if (this.accordionContainer) {
      this.accordionContainer.querySelectorAll('.inline-edit-input').forEach(input => {
        if (input._calculationHandler) {
          ['input', 'change', 'blur'].forEach(eventType => {
            input.removeEventListener(eventType, input._calculationHandler);
          });
          delete input._calculationHandler;
        }

        // 날짜 검증 핸들러 제거
        if (input._dateValidatorAttached) {
          input.removeEventListener('blur', input._dateBlurHandler);
          input.removeEventListener('input', input._dateInputHandler);
          delete input._dateValidatorAttached;
          delete input._dateBlurHandler;
          delete input._dateInputHandler;
        }
      });

      // DOM 제거 (이벤트 리스너 자동 정리)
      this.accordionContainer.remove();
      this.accordionContainer = null;
    }

    // 전역 document 이벤트 리스너 제거
    if (this.documentClickHandler) {
      document.removeEventListener('click', this.documentClickHandler);
      this.documentClickHandler = null;
    }
    if (this.documentKeydownHandler) {
      document.removeEventListener('keydown', this.documentKeydownHandler);
      this.documentKeydownHandler = null;
    }

    // Tom Select 인스턴스 정리
    this.tomSelectInstances = [];

    // 저장 큐 정리
    this.saveQueue = [];
    this.isSavingInProgress = false;

    // 참조 정리
    this.isOpen = false;
    this.currentProject = null;
    this.projectCards = {};
    this.dataTable = null;
    this.unifiedBadgeSystem = null;  }

  /**
   * ========================================
   * 헤더 편집 기능 (담당자, 현장주소, 공사내용)
   * ========================================
   */

  /**
   * 활성 담당자 목록 가져오기 (퇴사자 제외, 실제 사용 중인 담당자만)
   */
  getActiveManagersFromData() {
    // 1. ManagerBadge에서 활성 담당자 기본 목록
    const baseManagers = this.unifiedBadgeSystem.managerBadge.getActiveManagers();

    // 2. DataTable에서 실제 사용 중인 담당자 추출
    const dataTableManagers = new Set();

    if (this.dataTable) {
      this.dataTable.rows().every(function() {
        const rowData = this.data();
        const manager = rowData['담당자'];
        if (manager) {
          dataTableManagers.add(manager.trim());
        }
      });
    }

    // 3. 기본 목록 + 실제 사용 중인 담당자 합치기
    const resignedTeam = this.unifiedBadgeSystem.managerBadge.resignedTeam;
    const allManagers = new Set([...baseManagers, ...dataTableManagers]);

    // 4. 퇴사자 및 경영지원팀 제외
    const excludedManagers = ['김단이', '심장원', '아이티', '이근혁', '황샛별', '황해승'];
    const activeManagers = Array.from(allManagers).filter(m => !excludedManagers.includes(m));

    // 5. 가나다순 정렬
    return activeManagers.sort((a, b) => a.localeCompare(b, 'ko-KR'));
  }

  /**
   * 프로젝트 행번호 가져오기 (Google Sheets 기준)
   */
  getProjectRowNumber(projectCode) {
    if (!this.dataTable) return 0;

    const row = this.dataTable.row(`#row-${projectCode}`);
    if (row && row.index() !== undefined) {
      // DataTable index는 0-based, Google Sheets는 2부터 시작 (1행은 헤더)
      return row.index() + 2;
    }

    return 0;
  }

  /**
   * 프로젝트 코드 생성 (클라이언트 사이드)
   * - 백엔드 project_code_generator.py와 동일한 로직
   */
  generateProjectCode(rowNumber, company, manager) {
    // 1. 사업자 → Prefix
    const companyPrefix = {
      '플렌트': 'P',
      '글로벌': 'G',
      '글로벌그룹': 'R'
    };
    const prefix = companyPrefix[company] || '';

    // 2. 행번호 4자리
    const numberPart = String(rowNumber).padStart(4, '0');

    // 3. 담당자 → Suffix
    const managerSuffix = {
      '박용구': '-YG', '박정우': '-JW', '강성환': '-SH', '박민우': '-MW',
      '이근혁': '-GH', '김호중': '-HJ', '아이티': '-IT', '김단이': '-DN',
      '권태훈': '-TH', '주영민': '-YM', '심장원': '-SJW', '빈승정': '-SJ',
      '박민재': '-MJ', '조성헌': '-JSH', '황해승': '-HS', '강민석': '-MS'
    };
    const suffix = managerSuffix[manager] || '';

    // 4. 조합
    return `${prefix}${numberPart}${suffix}`;
  }

  /**
   * 헤더 필드 생성 (읽기/편집 모드)
   */
  generateHeaderFields(projectCode, projectData, isEditMode = false) {
    const manager = projectData['담당자'] || '';
    const address = projectData['현장 주소'] || '';
    const content = projectData['공사 내용'] || '';

    if (!isEditMode) {
      // 읽기 모드 - 빈 값일 때 안내 메시지 표시
      const displayAddress = address || '주소 정보 없음';
      const displayContent = content || '공사내용 없음';

      return `
        <div class="project-code-badge">
          <i class="fas fa-tag me-2"></i>${projectCode}
        </div>
        <div class="project-manager-info">
          <i class="fas fa-user me-2"></i>${manager || '담당자 미지정'}
        </div>
        <div class="project-address-info" title="${displayAddress}">
          <i class="fas fa-map-marker-alt me-2"></i>${displayAddress}
        </div>
        <div class="project-content-info" title="${displayContent}">
          <i class="fas fa-tools me-2"></i>${displayContent}
        </div>
      `;
    } else {
      // 편집 모드 - 실제 값만 사용 (빈 값이면 placeholder 표시)
      const activeManagers = this.getActiveManagersFromData();

      return `
        <div class="project-code-badge readonly" id="code-badge-${projectCode}">
          <i class="fas fa-tag me-2"></i>${projectCode}
        </div>
        <div class="project-manager-info editable">
          <i class="fas fa-user me-2"></i>
          <select class="header-edit-select" data-field="담당자" data-original-value="${manager}">
            <option value="">담당자 선택</option>
            ${activeManagers.map(m => `
              <option value="${m}" ${m === manager ? 'selected' : ''}>${m}</option>
            `).join('')}
          </select>
        </div>
        <div class="project-address-info editable wide-field">
          <i class="fas fa-map-marker-alt me-2"></i>
          <input type="text" class="header-edit-input"
                 data-field="현장 주소"
                 data-original-value="${address}"
                 value="${address}"
                 placeholder="현장 주소를 입력하세요">
        </div>
        <div class="project-content-info editable wide-field">
          <i class="fas fa-tools me-2"></i>
          <input type="text" class="header-edit-input"
                 data-field="공사 내용"
                 data-original-value="${content}"
                 value="${content}"
                 placeholder="공사 내용을 입력하세요">
        </div>
      `;
    }
  }

  /**
   * 헤더 프로젝트 코드 실시간 업데이트
   */
  updateHeaderProjectCode(oldProjectCode) {
    // 저장된 행 번호 사용 (아코디언 열 때 저장됨)
    const rowNumber = this.currentRowNumber;
    if (!rowNumber) {
      logger.warn('[헤더 업데이트] 행번호를 찾을 수 없음');
      return oldProjectCode;
    }

    // 현재 선택된 사업자/담당자 가져오기
    const headerContainer = document.getElementById(`header-${oldProjectCode}`);
    const basicCard = document.getElementById(`card-basic-${oldProjectCode}`);

    let company = this.currentProject['사업자'];
    let manager = this.currentProject['담당자'];
    // 편집 중인 값 가져오기 (사업자는 기본정보 카드에 있음)
    if (basicCard) {
      const companyField = basicCard.querySelector('[data-field="사업자"]');
      // 필드 컨테이너(div) 안에 select 요소가 있을 수 있으므로 직접 찾아야 함
      const companySelect = companyField?.querySelector('select') ||
                           (companyField?.tagName === 'SELECT' ? companyField : null);
      if (companySelect) {
        company = companySelect.value;      }
    }

    // 담당자는 헤더에 있음
    if (headerContainer) {
      const managerSelect = headerContainer.querySelector('[data-field="담당자"]');
      if (managerSelect && managerSelect.tagName === 'SELECT') {
        manager = managerSelect.value;
      }
    }
    // 새 프로젝트 코드 생성
    const newProjectCode = this.generateProjectCode(rowNumber, company, manager);

    // 헤더 배지 텍스트만 업데이트 (미리보기 용도)
    const codeBadge = document.getElementById(`code-badge-${oldProjectCode}`);
    if (codeBadge) {
      codeBadge.innerHTML = `<i class="fas fa-tag me-2"></i>${newProjectCode}`;

      // 변경되었음을 시각적으로 표시
      if (newProjectCode !== oldProjectCode) {
        codeBadge.classList.add('code-changed');
        setTimeout(() => codeBadge.classList.remove('code-changed'), 1000);
      }

      // NOTE: 배지 ID는 변경하지 않음 (미리보기만)
    }

    // NOTE: 카드 ID와 currentProject['프로젝트 코드']는 여기서 변경하지 않음
    // - 저장 전 미리보기 용도이므로 DOM ID는 원본 유지
    // - 저장 성공 시 _executeSave()에서 서버 응답 기반으로 안전하게 업데이트
    // - 이렇게 해야 취소/재편집 시 DOM 요소를 정상적으로 찾을 수 있음

    return newProjectCode;
  }

  /**
   * 헤더 필드 이벤트 바인딩
   */
  bindHeaderFieldEvents(projectCode) {
    const headerContainer = document.getElementById(`header-${projectCode}`);
    if (!headerContainer) return;

    // 🆕 EditState: 헤더 컨테이너 전체에 이벤트 위임 (모든 헤더 필드 추적)
    // 중복 등록 방지: 이미 리스너가 등록되어 있으면 스킵
    if (this.editState && this.editState.isActive && !headerContainer.dataset.editstateListenerAttached) {
      headerContainer.addEventListener('input', (e) => {
        const field = e.target.closest('[data-field]');
        if (field) {
          const fieldName = field.dataset.field;
          let value;

          if (e.target.tagName === 'SELECT') {
            value = e.target.value;
          } else if (e.target.type === 'checkbox') {
            value = e.target.checked;
          } else {
            value = e.target.value;
          }

          this.editState.updateField(fieldName, value);
          logger.debug(`[EditState] 헤더 필드 업데이트: ${fieldName} =`, value);
        }
      });

      // change 이벤트도 처리 (select, checkbox 등)
      headerContainer.addEventListener('change', (e) => {
        const field = e.target.closest('[data-field]');
        if (field) {
          const fieldName = field.dataset.field;
          let value;

          if (e.target.tagName === 'SELECT') {
            value = e.target.value;
          } else if (e.target.type === 'checkbox') {
            value = e.target.checked;
          } else {
            value = e.target.value;
          }

          this.editState.updateField(fieldName, value);
          logger.debug(`[EditState] 헤더 필드 변경: ${fieldName} =`, value);
        }
      });

      // 리스너 등록 완료 플래그
      headerContainer.dataset.editstateListenerAttached = 'true';
      logger.debug('[EditState] 헤더 input 리스너 등록 완료');
    }

    // 담당자 변경 시 (프로젝트 코드 재계산)
    const managerSelect = headerContainer.querySelector('[data-field="담당자"]');
    if (managerSelect) {
      managerSelect.addEventListener('change', () => {
        // 항상 최신 프로젝트 코드 사용
        const currentCode = this.currentProject['프로젝트 코드'];
        const newCode = this.updateHeaderProjectCode(currentCode);
        // NOTE: 헤더 컨테이너 ID는 여기서 변경하지 않음
        // 저장 후 서버 응답에 따라 _executeSave에서 안전하게 업데이트
        // (코드 변경이 거부될 수 있으므로 미리 ID를 바꾸면 불일치 발생)
      });
    }

    // 기본정보 카드의 사업자 변경 시 (프로젝트 코드 재계산)
    const basicCard = document.getElementById(`card-basic-${projectCode}`);
    if (basicCard) {
      const companySelect = basicCard.querySelector('[data-field="사업자"]');
      if (companySelect) {
        companySelect.addEventListener('change', () => {
          // 항상 최신 프로젝트 코드 사용
          const currentCode = this.currentProject['프로젝트 코드'];
          const newCode = this.updateHeaderProjectCode(currentCode);
          // NOTE: 헤더 컨테이너 ID는 여기서 변경하지 않음
          // 저장 후 서버 응답에 따라 _executeSave에서 안전하게 업데이트
          // (코드 변경이 거부될 수 있으므로 미리 ID를 바꾸면 불일치 발생)
        });
      }
    }
  }

  /**
   * DataTable 행 업데이트 (프로젝트 코드 변경 시)
   */
  updateDataTableRow(oldCode, newCode, projectData) {
    if (!this.dataTable) {
      logger.warn('⚠️ [DataTable] DataTable 인스턴스 없음');
      return false;
    }

    logger.debug(`🔍 [DataTable] 행 찾기 시작: oldCode=${oldCode}, newCode=${newCode}`);

    // 방법 1: ID 선택자로 찾기
    let row = this.dataTable.row(`#row-${oldCode}`);
    logger.debug(`🔍 [DataTable] 방법1 (ID #row-${oldCode}):`, row.index());

    // 방법 2: oldCode로 못 찾으면 newCode로 시도
    if (!row || row.index() === undefined) {
      logger.debug(`⚠️ [DataTable] oldCode로 찾기 실패, newCode로 재시도: ${newCode}`);
      row = this.dataTable.row(`#row-${newCode}`);
      logger.debug(`🔍 [DataTable] 방법2 (ID #row-${newCode}):`, row.index());
    }

    // 방법 3: data-project-code 속성으로 DOM 검색
    if (!row || row.index() === undefined) {
      logger.debug(`⚠️ [DataTable] ID로 찾기 실패, DOM 검색 시도`);
      const rowElement = document.querySelector(`tr[data-project-code="${oldCode}"]`);
      if (rowElement) {
        row = this.dataTable.row(rowElement);
        logger.debug(`🔍 [DataTable] 방법3 (data-project-code=${oldCode}):`, row.index());
      }
    }

    // 방법 4: 데이터 검색 (프로젝트 코드 필드로 찾기)
    if (!row || row.index() === undefined) {
      logger.debug(`⚠️ [DataTable] DOM 검색 실패, 데이터 검색 시도`);
      row = this.dataTable.row((idx, data) => {
        return data['프로젝트 코드'] === oldCode || data['프로젝트 코드'] === newCode;
      });
      logger.debug(`🔍 [DataTable] 방법4 (데이터 검색):`, row.index());
    }

    // 여전히 못 찾으면 실패 반환
    if (!row || row.index() === undefined) {
      logger.error(`❌ [DataTable] 모든 방법으로 행을 찾을 수 없음: ${oldCode}/${newCode}`);
      return false;
    }

    logger.debug(`✅ [DataTable] 행 찾기 성공, 인덱스:`, row.index());

    // 1. 행 데이터 업데이트
    row.data(projectData); // 데이터만 업데이트

    // 2. 프로젝트 코드가 변경된 경우 무조건 draw (reopenAccordion이 행을 찾을 수 있도록)
    const codeChanged = oldCode !== newCode;

    if (codeChanged) {
      // 프로젝트 코드 변경 시 즉시 draw (reopenAccordion에서 새 행을 찾을 수 있도록)
      this.dataTable.draw(false);
      logger.debug('✅ [DataTable] 테이블 재렌더링 완료 (프로젝트 코드 변경)');
    } else if (!this.isOpen || this.currentProject?.['프로젝트 코드'] !== newCode) {
      // 아코디언이 닫혀있거나 다른 프로젝트인 경우에만 draw
      this.dataTable.draw(false); // false = 현재 페이지 유지
      logger.debug('✅ [DataTable] 테이블 재렌더링 완료 (아코디언 닫힘)');
    } else {
      // 아코디언이 열려있으면 draw 연기 (플래그 설정)
      this.pendingTableUpdate = true;
      logger.debug('⏭️ [DataTable] draw() 건너뜀 (아코디언 열려있음, 닫힌 후 자동 업데이트됨)');
    }

    // 3. 행 ID 업데이트 (DOM 조작)
    const rowNode = row.node();
    if (rowNode) {
      rowNode.id = `row-${newCode}`;
      rowNode.setAttribute('data-project-code', newCode);
    }

    logger.debug(`✅ [DataTable] 행 업데이트 성공: ${oldCode} → ${newCode}`);
    return true;
  }

  /**
   * 프로젝트 코드 변경 시 DOM 식별자 동기화
   */
  syncDomProjectCode(oldCode, newCode) {
    if (!oldCode || !newCode || oldCode === newCode) {
      return;
    }

    const shell = this.getAccordionShell();
    if (shell) {
      shell.dataset.projectCode = newCode;
      shell.querySelectorAll('[data-project-code]').forEach(element => {
        if (element.dataset.projectCode === oldCode) {
          element.dataset.projectCode = newCode;
        }
      });
    }

    if (this.accordionContainer) {
      this.accordionContainer.dataset.currentProjectCode = newCode;
    } else {
      logger.warn('[syncDomProjectCode] accordionContainer가 정의되지 않음');
    }

    const buttonsContainer = this.accordionContainer?.querySelector(`.unified-edit-buttons[data-project-code="${oldCode}"]`);
    if (buttonsContainer) {
      buttonsContainer.dataset.projectCode = newCode;
    }

    ['.unified-edit-btn', '.unified-save-btn', '.unified-cancel-btn'].forEach(selector => {
      const button = this.accordionContainer?.querySelector(`${selector}[data-project-code="${oldCode}"]`);
      if (button) {
        button.dataset.projectCode = newCode;
      }
    });

    const titleSection = this.accordionContainer?.querySelector(`.project-title-section[data-project-code="${oldCode}"]`);
    if (titleSection) {
      titleSection.dataset.projectCode = newCode;
    }

    const headerContainer = document.getElementById(`header-${oldCode}`);
    if (headerContainer) {
      headerContainer.id = `header-${newCode}`;
    }

    const cards = [
      'basic', 'schedule', 'payment', 'construction', 'management',
      'files', 'financial', 'collection', 'profit'
    ];
    cards.forEach(cardType => {
      const card = document.getElementById(`card-${cardType}-${oldCode}`);
      if (card) {
        card.id = `card-${cardType}-${newCode}`;
      }
    });

    const codeBadge = document.getElementById(`code-badge-${oldCode}`);
    if (codeBadge) {
      codeBadge.id = `code-badge-${newCode}`;
    }

    this.originalProjectCode = newCode;
  }

  /**
   * 데이터테이블 정렬 변경 후 아코디언 행을 새 위치로 이동
   */
  moveAccordionRow(newProjectCode) {
    const accordionRow = this.accordionContainer?.closest('.accordion-row');
    if (!accordionRow) {
      logger.warn('[moveAccordionRow] 아코디언 행을 찾을 수 없습니다.');
      return false;
    }
    accordionRow.dataset.projectCode = newProjectCode;

    if (!this.dataTable || typeof this.dataTable.row !== 'function') {
      logger.warn('[moveAccordionRow] DataTable 인스턴스를 찾을 수 없습니다.');
      return false;
    }

    const targetRow = this.dataTable.row(`#row-${newProjectCode}`).node();
    if (!targetRow || !targetRow.parentNode) {
      logger.warn(`[moveAccordionRow] 대상 테이블 행을 찾을 수 없습니다: ${newProjectCode}`);
      return false;
    }

    // 아코디언 행을 새 위치로 이동
    targetRow.parentNode.insertBefore(accordionRow, targetRow.nextSibling);
    this.accordionContainer.classList.add('show');
    this.isOpen = true;
    return true;
  }

  /**
   * 아코디언 재오픈 (프로젝트 코드 변경 후)
   */
  reopenAccordion(newProjectCode, projectData) {


    // 1. 기존 아코디언 닫기 (즉시)
    document.querySelectorAll('.accordion-row').forEach(row => row.remove());
    this.accordionContainer.classList.remove('show');
    document.querySelectorAll('tbody tr').forEach(row => row.classList.remove('table-active'));

    // 2. 새 프로젝트 코드로 아코디언 열기
    setTimeout(() => {
      let newRow = this.dataTable.row(`#row-${newProjectCode}`).node();

      // ID로 못 찾으면 데이터 검색
      if (!newRow) {
        logger.debug(`🔍 [Accordion] ID로 찾기 실패, 데이터 검색: ${newProjectCode}`);
        const row = this.dataTable.row((idx, data) => {
          return data['프로젝트 코드'] === newProjectCode;
        });
        newRow = row.node();
      }

      if (newRow) {
        this.openAccordion(newRow, projectData);
      } else {
        // DataTable에서 행을 찾지 못한 경우: 페이지 새로고침
        logger.warn(`⚠️ [Accordion] DataTable에서 행을 찾을 수 없음: ${newProjectCode}, 페이지 새로고침`);
        window.location.reload();
      }
    }, 150);
  }

  /**
   * 체크박스 드롭다운 멀티셀렉트 초기화
   * - Bootstrap 드롭다운 + 체크박스
   * - 최대 2개 선택 제한
   */
  initializeTomSelectMultiSelects() {
    const multiSelectDropdowns = this.accordionContainer.querySelectorAll('.multi-select-dropdown');
    logger.debug(`🔧 [CheckboxDropdown] 초기화 대상: ${multiSelectDropdowns.length}개 필드`);

    multiSelectDropdowns.forEach(dropdown => {
      const fieldName = dropdown.dataset.field || '';
      const button = dropdown.querySelector('.dropdown-toggle');
      const selectedText = dropdown.querySelector('.selected-text');
      const checkboxes = dropdown.querySelectorAll('.form-check-input');
      const dropdownMenu = dropdown.querySelector('.dropdown-menu');

      // Bootstrap Dropdown 초기화 (popperConfig로 body에 추가)
      const bsDropdown = new bootstrap.Dropdown(button, {
        popperConfig: {
          strategy: 'fixed',
          modifiers: [
            {
              name: 'preventOverflow',
              options: {
                boundary: 'viewport'
              }
            }
          ]
        }
      });

      // 드롭다운 열릴 때 body에 추가
      button.addEventListener('shown.bs.dropdown', () => {
        document.body.appendChild(dropdownMenu);

        // 위치 계산
        const rect = button.getBoundingClientRect();
        dropdownMenu.style.position = 'fixed';
        dropdownMenu.style.top = `${rect.bottom + 2}px`;
        dropdownMenu.style.left = `${rect.left}px`;

        // 시공자 필드: 한 줄 6명 한도, 7명 이상은 자동 wrap
        // (max-width로 박스 폭 제한 → 옵션 폭 80 × 6 + 라벨/padding ≈ 720px)
        if (fieldName === '시공자') {
          dropdownMenu.style.width = 'max-content';
          dropdownMenu.style.minWidth = '420px';
          dropdownMenu.style.maxWidth = '680px';
          dropdownMenu.style.maxHeight = 'none';
          dropdownMenu.style.overflowY = 'visible';
          dropdownMenu.style.padding = '0.75rem 0 0.75rem 0.5rem';  // 위/좌/아래 여백 추가, 우측은 그대로
        } else if (fieldName === '계산서') {
          // 계산서 필드는 카테고리별 레이아웃 (이미 HTML에서 설정됨)
          dropdownMenu.style.width = 'auto';
          dropdownMenu.style.minWidth = '300px';
          dropdownMenu.style.maxWidth = '400px';
          dropdownMenu.style.maxHeight = 'none';
          dropdownMenu.style.overflowY = 'visible';
        } else {
          dropdownMenu.style.width = `${rect.width}px`;
          dropdownMenu.style.minWidth = `${rect.width}px`;
          dropdownMenu.style.maxWidth = `${rect.width}px`;
        }

        dropdownMenu.style.zIndex = '9999';
        dropdownMenu.style.display = 'block';
      });

      // 드롭다운 닫힐 때 원래 위치로 복귀
      button.addEventListener('hidden.bs.dropdown', () => {
        dropdown.appendChild(dropdownMenu);
        // 스타일 초기화
        dropdownMenu.style.position = '';
        dropdownMenu.style.top = '';
        dropdownMenu.style.left = '';
        dropdownMenu.style.width = '';
        dropdownMenu.style.minWidth = '';
        dropdownMenu.style.maxWidth = '';
        dropdownMenu.style.maxHeight = '';
        dropdownMenu.style.overflowY = '';
        dropdownMenu.style.padding = '';
        dropdownMenu.style.flexDirection = '';
        dropdownMenu.style.flexWrap = '';
        dropdownMenu.style.gap = '';
        dropdownMenu.style.zIndex = '';
        dropdownMenu.style.display = '';

        // 시공자 필드의 li 스타일도 초기화
        if (fieldName === '시공자') {
          const listItems = dropdownMenu.querySelectorAll('li');
          listItems.forEach(li => {
            li.style.flex = '';
            li.style.width = '';
            li.style.listStyle = '';
            li.style.padding = '';
            li.style.display = '';

            const label = li.querySelector('label');
            if (label) {
              label.style.justifyContent = '';
              label.style.width = '';
              label.style.padding = '';
              label.style.margin = '';
              label.style.display = '';
            }
          });
        }
      });

      // 계산서 필드: 카테고리별 체크박스 이벤트 처리
      if (fieldName === '계산서') {
        const billStageCheckboxes = dropdownMenu.querySelectorAll('.bill-stage-checkbox');
        const billSpecialCheckbox = dropdownMenu.querySelector('.bill-special-checkbox');

        // 단계별 체크박스 이벤트 (계약금, 중도금, 잔금)
        billStageCheckboxes.forEach(checkbox => {
          checkbox.addEventListener('change', () => {
            const category = checkbox.dataset.category;
            const stage = checkbox.value;

            if (checkbox.checked) {
              // 미발행 체크박스 해제
              if (billSpecialCheckbox) {
                billSpecialCheckbox.checked = false;
              }

              // 같은 카테고리의 다른 체크박스 모두 해제 (각 카테고리에서 1개만 선택)
              billStageCheckboxes.forEach(cb => {
                if (cb.dataset.category === category && cb !== checkbox) {
                  cb.checked = false;
                }
              });

              // 같은 단계의 다른 카테고리 체크박스 해제 (각 단계는 1개 카테고리만)
              billStageCheckboxes.forEach(cb => {
                if (cb.value === stage && cb !== checkbox) {
                  cb.checked = false;
                }
              });

              // 전체 선택 개수 확인 (최대 3개)
              const checkedCount = Array.from(billStageCheckboxes).filter(cb => cb.checked).length;
              if (checkedCount > 3) {
                checkbox.checked = false;
                logger.debug('✋ [계산서] 최대 3개까지만 선택 가능합니다');
                return;
              }
            }

            this.updateBillStatusSelection(billStageCheckboxes, billSpecialCheckbox, selectedText, fieldName);
          });
        });

        // 미발행 체크박스 이벤트
        if (billSpecialCheckbox) {
          billSpecialCheckbox.addEventListener('change', () => {
            if (billSpecialCheckbox.checked) {
              // 모든 단계별 체크박스 해제
              billStageCheckboxes.forEach(cb => {
                cb.checked = false;
              });
            }

            this.updateBillStatusSelection(billStageCheckboxes, billSpecialCheckbox, selectedText, fieldName);
          });
        }
      }

      // 시공자 필드: 카테고리별 체크박스 이벤트 처리
      if (fieldName === '시공자') {
        const constructorCheckboxes = dropdownMenu.querySelectorAll('.constructor-checkbox');
        const otherCheck = dropdownMenu.querySelector('.constructor-other-check');
        const otherInput = dropdownMenu.querySelector('.constructor-other-input');

        // 일반 체크박스 이벤트
        constructorCheckboxes.forEach(checkbox => {
          checkbox.addEventListener('change', () => {
            // 전체 선택 개수 확인 (최대 3개)
            const checkedCount = Array.from(constructorCheckboxes).filter(cb => cb.checked).length;
            const otherChecked = otherCheck && otherCheck.checked ? 1 : 0;

            if (checkbox.checked && (checkedCount + otherChecked) > 3) {
              checkbox.checked = false;
              logger.debug('✋ [시공자] 최대 3개까지만 선택 가능합니다');
              return;
            }

            this.updateConstructorSelection(constructorCheckboxes, otherCheck, otherInput, selectedText, fieldName);
          });
        });

        // 기타 체크박스 이벤트 (입력란이 같은 row의 우측에 표시됨)
        if (otherCheck && otherInput) {
          otherCheck.addEventListener('change', () => {
            const normalCheckedCount = Array.from(constructorCheckboxes).filter(cb => cb.checked).length;

            if (otherCheck.checked && normalCheckedCount >= 3) {
              otherCheck.checked = false;
              logger.debug('✋ [시공자] 최대 3개까지만 선택 가능합니다');
              return;
            }

            if (otherCheck.checked) {
              otherInput.classList.remove('d-none');
              otherInput.focus();
            } else {
              otherInput.classList.add('d-none');
              otherInput.value = '';
            }

            this.updateConstructorSelection(constructorCheckboxes, otherCheck, otherInput, selectedText, fieldName);
          });

          // 입력창 값 변경 시 선택된 값 업데이트
          otherInput.addEventListener('input', () => {
            this.updateConstructorSelection(constructorCheckboxes, otherCheck, otherInput, selectedText, fieldName);
          });
        }
      }

      // 체크박스 변경 이벤트 (계산서 필드는 별도 이벤트 리스너 있으므로 제외)
      if (fieldName !== '계산서') {
        checkboxes.forEach(checkbox => {
          checkbox.addEventListener('change', () => {
            // 시공자 필드: 기타 체크박스를 제외한 일반 체크박스만 카운트
            let effectiveCount;
            if (fieldName === '시공자') {
              const normalCheckedBoxes = Array.from(checkboxes).filter(cb =>
                cb.checked && !cb.classList.contains('accordion-constructor-other-check')
              );
              effectiveCount = normalCheckedBoxes.length;

              // 기타 체크박스가 체크되어 있으면 +1
              const otherCheck = dropdownMenu.querySelector('.accordion-constructor-other-check');
              if (otherCheck && otherCheck.checked) {
                effectiveCount++;
              }
            } else {
              const checkedBoxes = Array.from(checkboxes).filter(cb => cb.checked);
              effectiveCount = checkedBoxes.length;
            }

            // 최대 선택 제한 (시공자: 3개, 나머지: 2개)
            const maxLimit = fieldName === '시공자' ? 3 : 2;
            if (effectiveCount > maxLimit) {
              checkbox.checked = false;
              logger.debug(`✋ [CheckboxDropdown] "${fieldName}" 최대 ${maxLimit}개까지만 선택 가능`);
              return;
            }

            // 시공자 필드는 별도 메서드로 처리 (기타 입력값 포함)
            if (fieldName === '시공자') {
              const otherCheck = dropdownMenu.querySelector('.accordion-constructor-other-check');
              const otherInput = dropdownMenu.querySelector('.accordion-constructor-other-input');
              this.updateAccordionConstructorSelection(checkboxes, selectedText, otherCheck, otherInput, fieldName);
            } else {
              // 선택된 값 업데이트
              const checkedBoxes = Array.from(checkboxes).filter(cb => cb.checked);
              const selectedValues = checkedBoxes.map(cb => cb.value);
              selectedText.textContent = selectedValues.length > 0 ? selectedValues.join(', ') : '선택';
              logger.debug(`✅ [CheckboxDropdown] "${fieldName}" 선택된 항목: ${selectedValues.length}개`);

              // 🆕 EditState: 드롭다운 선택값 업데이트
              if (this.editState && this.editState.isActive && fieldName) {
                const finalValue = selectedValues.length > 0 ? selectedValues.join(', ') : '';
                this.editState.updateField(fieldName, finalValue);
                logger.debug(`[EditState] 드롭다운 업데이트: ${fieldName} =`, finalValue);
              }
            }
          });
        });
      }

      logger.debug(`✅ [CheckboxDropdown] "${fieldName}" 초기화 완료`);
    });

    logger.debug(`🎉 [CheckboxDropdown] 총 ${multiSelectDropdowns.length}개 필드 초기화 완료`);
  }

  /**
   * 아코디언 시공자 선택 업데이트 (일반 체크박스 + 기타 입력값 포함)
   */
  updateAccordionConstructorSelection(checkboxes, selectedText, otherCheck, otherInput, fieldName) {
    // 기타 체크박스를 제외한 일반 체크박스만 수집
    const checkedBoxes = Array.from(checkboxes).filter(cb =>
      cb.checked && !cb.classList.contains('accordion-constructor-other-check')
    );
    const selectedValues = checkedBoxes.map(cb => cb.value);

    // 기타 체크박스가 체크되어 있고 입력값이 있으면 추가
    if (otherCheck && otherCheck.checked && otherInput && otherInput.value.trim()) {
      selectedValues.push(otherInput.value.trim());
    }

    // 선택된 값 표시 업데이트
    selectedText.textContent = selectedValues.length > 0 ? selectedValues.join(', ') : '선택';

    // 🆕 EditState: 드롭다운 선택값 업데이트
    if (this.editState && this.editState.isActive && fieldName) {
      const finalValue = selectedValues.length > 0 ? selectedValues.join(', ') : '';
      this.editState.updateField(fieldName, finalValue);
      logger.debug(`[EditState] 드롭다운 업데이트: ${fieldName} =`, finalValue);
    }

    logger.debug(`✅ [CheckboxDropdown] "${fieldName}" 선택된 항목: ${selectedValues.length}개`, selectedValues);
  }

  /**
   * 계산서 선택 업데이트 (카테고리-단계 형식)
   * 형식: "현금결제-계약금, N입금-중도금, 카드결제-잔금"
   */
  /**
   * 시공자 선택 값 업데이트
   */
  updateConstructorSelection(constructorCheckboxes, otherCheck, otherInput, selectedText, fieldName) {
    const selectedItems = [];

    // 일반 체크박스에서 선택된 값
    constructorCheckboxes.forEach(cb => {
      if (cb.checked) {
        selectedItems.push(cb.value);
      }
    });

    // 기타 입력값 추가
    if (otherCheck && otherCheck.checked && otherInput && otherInput.value.trim()) {
      selectedItems.push(otherInput.value.trim());
    }

    // 선택된 값 표시
    const displayText = selectedItems.length > 0 ? selectedItems.join(', ') : '선택';
    selectedText.textContent = displayText;

    // EditState 업데이트
    if (this.editState && this.editState.isActive && fieldName) {
      const finalValue = selectedItems.length > 0 ? selectedItems.join(', ') : '';
      this.editState.updateField(fieldName, finalValue);
      logger.debug(`[EditState] ${fieldName} 업데이트: ${finalValue}`);
    }
  }

  updateBillStatusSelection(billStageCheckboxes, billSpecialCheckbox, selectedText, fieldName) {
    // 미발행 체크 여부 확인
    if (billSpecialCheckbox && billSpecialCheckbox.checked) {
      selectedText.textContent = '미발행';
      logger.debug(`✅ [BillStatus] 선택: 미발행`);

      // 🆕 EditState: 계산서 필드 업데이트 (fieldName 파라미터 사용)
      if (this.editState && this.editState.isActive && fieldName) {
        this.editState.updateField(fieldName, '미발행');
        logger.debug(`[EditState] ${fieldName} 업데이트: 미발행`);
      }
      return;
    }

    // 선택된 항목 수집 (각 체크박스의 "카테고리-단계" 형식)
    const selectedItems = [];
    const categories = new Set();
    billStageCheckboxes.forEach(cb => {
      if (cb.checked) {
        const category = cb.dataset.category;
        const stage = cb.value;
        selectedItems.push(`${category}-${stage}`);
        categories.add(category);
      }
    });

    // 선택된 값 포맷팅: "카테고리-단계, 카테고리-단계"
    // 편집 모드에서는 아이콘 없이 텍스트만 표시
    if (selectedItems.length > 0) {
      const displayText = selectedItems.join(', ');

      // 편집 모드에서는 텍스트만 표시 (아이콘 제거)
      selectedText.textContent = displayText;
      logger.debug(`✅ [BillStatus] 선택: ${displayText}`);

      // 🆕 EditState: 계산서 필드 업데이트 (fieldName 파라미터 사용)
      if (this.editState && this.editState.isActive && fieldName) {
        this.editState.updateField(fieldName, displayText);
        logger.debug(`[EditState] ${fieldName} 업데이트:`, displayText);
      }
    } else {
      selectedText.textContent = '선택';
      logger.debug(`✅ [BillStatus] 선택 해제`);

      // 🆕 EditState: 계산서 필드 업데이트 (빈 값, fieldName 파라미터 사용)
      if (this.editState && this.editState.isActive && fieldName) {
        this.editState.updateField(fieldName, '');
        logger.debug(`[EditState] ${fieldName} 업데이트: (빈 값)`);
      }
    }
  }

  /**
   * 프로젝트 잠금 heartbeat 시작 (30초마다 자동 연장)
   * @param {string} projectCode - 프로젝트 코드
   */
  startLockHeartbeat(projectCode) {
    // 기존 heartbeat 정리
    this.stopLockHeartbeat();

    logger.debug(`🫀 [Heartbeat] 시작: ${projectCode} (30초마다 잠금 연장)`);

    // 30초(30,000ms)마다 잠금 연장
    this.heartbeatInterval = setInterval(async () => {
      try {
        logger.debug(`🫀 [Heartbeat] 잠금 연장 요청: ${projectCode} (탭 ID: ${this.tabId})`);
        const response = await fetch('/api/project-lock/heartbeat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          credentials: 'same-origin',
          body: JSON.stringify({
            project_code: projectCode,
            tab_id: this.tabId
          })
        });

        const result = await response.json();

        if (result.success) {
          logger.debug(`✅ [Heartbeat] 잠금 연장 성공: ${projectCode}`);
          // 성공 시 실패 카운터 리셋
          this.heartbeatFailureCount = 0;
        } else {
          logger.warn(`⚠️ [Heartbeat] 잠금 연장 실패: ${result.message}`);
          // 다른 사용자가 잠금을 빼앗은 경우 편집 모드 종료
          this.stopLockHeartbeat();
          this.showMessage('다른 사용자가 편집 중입니다. 편집 모드가 종료됩니다.', 'warning');
          // 편집 모드 강제 종료
          this.disableUnifiedEditMode(projectCode);
        }
      } catch (error) {
        logger.error('❌ [Heartbeat] 네트워크 오류:', error);

        // 연속 실패 횟수 증가
        this.heartbeatFailureCount++;
        logger.warn(`⚠️ [Heartbeat] 연속 실패 횟수: ${this.heartbeatFailureCount}`);

        // 3회 연속 실패 시 경고
        if (this.heartbeatFailureCount === 3) {
          // 네트워크 상태 = 시스템 알림 → 사이트 최상단 헤더
          if (window.showSystemAlert) window.showSystemAlert('⚠️ 네트워크가 불안정합니다. 편집 내용을 저장하세요.', 'warning');
          else this.showToast('⚠️ 네트워크가 불안정합니다. 편집 내용을 저장하세요.', 'warning');
          logger.warn('⚠️ [Heartbeat] 3회 연속 실패 - 사용자에게 경고');
        }

        // 5회 연속 실패 시 편집 모드 자동 종료
        if (this.heartbeatFailureCount >= 5) {
          logger.error('❌ [Heartbeat] 5회 연속 실패 - 편집 모드 자동 종료');
          this.stopLockHeartbeat();
          this.showMessage('네트워크 연결이 불안정하여 편집 모드가 종료됩니다. 작업 내용을 확인해주세요.', 'error');
          this.disableUnifiedEditMode(projectCode);
        }
      }
    }, 30 * 1000); // 30초
  }

  /**
   * 프로젝트 잠금 heartbeat 정지
   */
  stopLockHeartbeat() {
    if (this.heartbeatInterval) {
      clearInterval(this.heartbeatInterval);
      this.heartbeatInterval = null;
      this.heartbeatFailureCount = 0;  // 실패 카운터 리셋
      logger.debug(`🫀 [Heartbeat] 정지`);
    }
  }

  /**
   * Toast 알림 표시
   * @param {string} message - 표시할 메시지
   * @param {string} type - 알림 타입 ('success', 'error', 'warning', 'info')
   */
  showToast(message, type = 'info') {
    // Toast 컴포넌트 동적 로딩
    import('./Toast.js').then(({ default: Toast }) => {
      new Toast().show(message, type);
    }).catch(error => {
      // Toast 로딩 실패 시 콘솔에 메시지 출력
      logger.error('[Toast] 로딩 실패:', error);
      logger.debug(`[알림] ${message}`);
    });
  }
}
