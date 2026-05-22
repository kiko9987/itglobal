import logger from '../utils/logger.js';

/**
 * Audit Log Modal 컴포넌트 (완전한 레거시 포팅)
 * 감사 로그 모달 - 페이지네이션, 필터링, 복잡한 툴팁 시스템
 */
export default class AuditLogModal {
  constructor() {
    this.modal = null;
    this.modalElement = null;
    this.currentPage = 1;
    this.totalPages = 1;
    this.perPage = 15;
    this.init();
  }

  /**
   * 컴포넌트 초기화
   */
  init() {
    this.modalElement = document.getElementById('auditLogsModal');
    if (!this.modalElement) {
      logger.debug('[AUDIT_LOG] auditLogsModal 요소가 없어서 생성합니다.');
      this.createModalHTML();
      this.modalElement = document.getElementById('auditLogsModal');
    }

    // table-striped 클래스 제거 (Bootstrap 자동 줄무늬 방지)
    const table = this.modalElement.querySelector('.audit-logs-table');
    if (table) {
      table.classList.remove('table-striped');
    }

    // Bootstrap 모달 인스턴스 생성
    this.modal = new bootstrap.Modal(this.modalElement, {
      backdrop: true,
      keyboard: true
    });

    // 이벤트 설정
    this.setupEventListeners();
  }

  /**
   * 이벤트 리스너 설정
   */
  setupEventListeners() {
    // 날짜 선택 변경 이벤트
    const daysSelect = document.getElementById('auditLogsDays');
    if (daysSelect) {
      daysSelect.addEventListener('change', () => this.loadAuditLogs(1));
    }

    // 새로고침 버튼 (공통 헬퍼 사용)
    const refreshBtn = document.getElementById('refreshAuditLogsBtn');
    if (refreshBtn && typeof window.runRefreshAnimation === 'function') {
      refreshBtn.addEventListener('click', () => {
        window.runRefreshAnimation(refreshBtn, () => this.loadAuditLogs(1));
      });
    }

    // 모달 숨김 이벤트
    this.modalElement.addEventListener('hidden.bs.modal', () => {
      this.handleModalHidden();
    });

    // 페이지네이션 클릭 이벤트 위임
    const paginationContainer = document.getElementById('auditLogsPagination');
    if (paginationContainer) {
      paginationContainer.addEventListener('click', (e) => {
        e.preventDefault();
        const pageLink = e.target.closest('a[data-page]');
        if (pageLink && !pageLink.closest('.disabled')) {
          const page = parseInt(pageLink.dataset.page);
          this.changePage(page);
        }
      });
    }
  }

  /**
   * 감사 로그 모달 열기
   */
  async open() {
    try {
      // 페이지네이션 초기화
      this.currentPage = 1;
      this.totalPages = 1;

      // 모달 이벤트 리스너 정리 및 재설정
      this.modalElement.removeEventListener('hidden.bs.modal', this.handleModalHidden);
      this.modalElement.addEventListener('hidden.bs.modal', () => this.handleModalHidden());

      this.modal.show();
      await this.loadAuditLogs(1); // 첫 페이지부터 시작
    } catch (error) {
      logger.error('감사 로그 모달 열기 실패:', error);
      this.showAlert('모달을 여는 중 오류가 발생했습니다.', 'error');
    }
  }

  /**
   * 감사 로그 로드
   */
  async loadAuditLogs(page = 1) {
    this.currentPage = page;
    const days = document.getElementById('auditLogsDays')?.value || 7;
    const tableBody = document.getElementById('auditLogsTableBody');
    const countElement = document.getElementById('auditLogsCount');
    const pageInfoElement = document.getElementById('auditLogsPageInfo');

    if (!tableBody) return;

    // 로딩 상태 표시
    tableBody.innerHTML = `
      <tr>
        <td colspan="8" class="text-center">
          <i class="fas fa-spinner fa-spin"></i> 로딩 중...
        </td>
      </tr>
    `;

    try {
      const response = await fetch(`/api/audit-logs?days=${days}&page=${this.currentPage}&per_page=${this.perPage}`, {
        credentials: 'include'
      });
      const data = await response.json();

      if (data.success) {
        this.displayAuditLogs(data.logs);

        // 페이지네이션 정보 업데이트
        const pagination = data.pagination;
        this.totalPages = pagination.total_pages;

        if (countElement) {
          countElement.textContent = `로그 개수: ${pagination.total_count}개`;
        }
        if (pageInfoElement) {
          pageInfoElement.textContent = `페이지: ${this.currentPage}/${this.totalPages}`;
        }

        // 페이지네이션 UI 업데이트
        this.updatePaginationUI(pagination);
      } else {
        this.showAlert(data.message || '로그를 불러오는데 실패했습니다.', 'error');
        tableBody.innerHTML = `
          <tr>
            <td colspan="8" class="text-center text-danger">
              <i class="fas fa-exclamation-triangle"></i> 로그를 불러오지 못했습니다.
            </td>
          </tr>
        `;
        this.resetPaginationUI();
      }
    } catch (error) {
      logger.error('감사 로그 로딩 오류:', error);
      this.showAlert('네트워크 오류가 발생했습니다.', 'error');
      tableBody.innerHTML = `
        <tr>
          <td colspan="8" class="text-center text-danger">
            <i class="fas fa-wifi"></i> 네트워크 오류
          </td>
        </tr>
      `;
      this.resetPaginationUI();
    }
  }

  /**
   * 감사 로그 표시
   */
  displayAuditLogs(logs) {
    const tableBody = document.getElementById('auditLogsTableBody');

    if (logs.length === 0) {
      tableBody.innerHTML = `
        <tr>
          <td colspan="8" class="text-center text-muted">
            <i class="fas fa-info-circle"></i> 표시할 로그가 없습니다.
          </td>
        </tr>
      `;
      return;
    }

    tableBody.innerHTML = logs.map((log, index) => {
      const timestamp = new Date(log.timestamp);
      const formattedTime = timestamp.toLocaleString('ko-KR', {
        timeZone: 'Asia/Seoul',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      });

      const actionBadge = this.getActionBadge(log.action);
      const roleBadge = this.getRoleBadge(log.user_role);

      // 긴 값들을 스마트하게 처리 (객체 반환)
      const oldValueFormatted = this.formatLogValue(log.old_value, log.field_name);
      const newValueFormatted = this.formatLogValue(log.new_value, log.field_name);

      // HTML 이스케이프 함수
      const escapeHtml = (text) => {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = String(text);
        return div.innerHTML;
      };

      // 안전하게 속성 추출 (객체가 아닌 경우 대비)
      const oldDisplay = oldValueFormatted?.display || String(oldValueFormatted || '-');
      const oldOriginal = oldValueFormatted?.original || '';
      const oldTruncated = oldValueFormatted?.truncated || false;

      const newDisplay = newValueFormatted?.display || String(newValueFormatted || '-');
      const newOriginal = newValueFormatted?.original || '';
      const newTruncated = newValueFormatted?.truncated || false;

      // 행 번호가 0-based index 기준: 첫 행(index 0) 흰색, 두 번째 행(index 1) 회색
      // index 0 -> 첫 번째 데이터 행 -> 흰색 ✓
      // index 1 -> 두 번째 데이터 행 -> 회색 ✓
      const rowBgColor = index % 2 === 0 ? '#ffffff' : 'rgba(0, 0, 0, 0.015)';
      const rowHoverColor = 'rgba(0, 0, 0, 0.025)'; // 모든 행 호버 시 동일한 색상

      // 툴팁 스타일 - 축약된 경우 help 커서 표시
      const oldValueStyle = oldTruncated ? 'cursor: help; text-decoration: underline dotted;' : '';
      const newValueStyle = newTruncated ? 'cursor: help; text-decoration: underline dotted;' : '';

      // TD에 적용할 배경색 스타일
      const tdBgStyle = `background-color: ${rowBgColor} !important;`;

      return `
        <tr data-row-bg="${rowBgColor}" data-row-hover="${rowHoverColor}">
          <td class="text-center" style="${tdBgStyle}">${formattedTime}</td>
          <td class="text-center" style="${tdBgStyle}" title="${log.user_email}">${log.user_name}</td>
          <td class="text-center" style="${tdBgStyle}">${roleBadge}</td>
          <td class="text-center" style="${tdBgStyle}">${actionBadge}</td>
          <td class="text-center" style="${tdBgStyle}">${log.project_code || '-'}</td>
          <td class="text-center" style="${tdBgStyle}">${this.getDisplayFieldName(log.field_name) || '-'}</td>
          <td class="text-center" style="${tdBgStyle} ${oldValueStyle}" ${oldTruncated ? `title="${escapeHtml(oldOriginal)}"` : ''}>${oldDisplay}</td>
          <td class="text-center" style="${tdBgStyle} ${newValueStyle}" ${newTruncated ? `title="${escapeHtml(newOriginal)}"` : ''}>${newDisplay}</td>
        </tr>
      `;
    }).join('');

    // Bootstrap 기본 툴팁 초기화
    setTimeout(() => {
      const tooltipTriggerList = document.querySelectorAll('#auditLogsTableBody [title]');
      [...tooltipTriggerList].forEach(tooltipTriggerEl => {
        new bootstrap.Tooltip(tooltipTriggerEl, {
          placement: 'top',
          trigger: 'hover',
          boundary: 'viewport'
        });
      });

      // 행 색상 및 호버 효과 추가 (CSS !important 덮어쓰기 위해 setProperty 사용)
      const rows = document.querySelectorAll('#auditLogsTableBody tr');
      rows.forEach(row => {
        const normalBg = row.dataset.rowBg;
        const hoverBg = row.dataset.rowHover;
        const tds = row.querySelectorAll('td');

        // 초기 색상 설정 (CSS !important 덮어쓰기)
        tds.forEach(td => {
          td.style.setProperty('background-color', normalBg, 'important');
        });

        row.addEventListener('mouseenter', () => {
          tds.forEach(td => {
            td.style.setProperty('background-color', hoverBg, 'important');
          });
        });

        row.addEventListener('mouseleave', () => {
          tds.forEach(td => {
            td.style.setProperty('background-color', normalBg, 'important');
          });
        });
      });
    }, 100);
  }

  /**
   * 필드명 표시명 매핑 함수 (로그에서만 사용)
   */
  getDisplayFieldName(fieldName) {
    if (!fieldName) return '-';

    // 로그 표시용 필드명 매핑
    const displayFieldNames = {
      // 영문 필드명 매핑
      'status': '상태',
      'folder_path': '폴더 경로',

      // 긴 필드명 축약
      '견적서 및 계약서 폴더 경로': '문서 폴더'
    };

    return displayFieldNames[fieldName] || fieldName;
  }

  /**
   * 숫자에 콤마 추가 (금액 포맷팅)
   */
  formatCurrency(value) {
    // 이미 포맷된 값이거나 빈 값인 경우
    if (!value || value === '-' || value === 'null' || value === 'undefined') {
      return value;
    }

    const strValue = String(value);

    // 숫자만 추출 (₩, 콤마, 공백 등 제거)
    const cleanValue = strValue.replace(/[₩,\s]/g, '');

    // 숫자가 아닌 경우 원본 반환
    if (!/^-?\d+(\.\d+)?$/.test(cleanValue)) {
      return strValue;
    }

    // 숫자를 콤마로 포맷팅
    const number = parseFloat(cleanValue);
    const formatted = number.toLocaleString('ko-KR');

    // 원화 기호가 없으면 추가
    return strValue.includes('₩') ? `₩${formatted}` : formatted;
  }

  /**
   * 로그 값 포맷팅
   * @returns {Object} { display: 표시할 텍스트, original: 원본 텍스트, truncated: 축약 여부 }
   */
  formatLogValue(value, fieldName) {
    if (!value || value === 'null' || value === 'undefined') {
      return { display: '-', original: '', truncated: false };
    }

    const strValue = String(value);

    // 금액 필드인 경우 콤마 포맷팅
    const amountFields = [
      '계약금', '중도금', '잔금',
      '총액 1', '총액 2', '총액2',
      '미수금',
      '제품대', '도급비', '자재비', '기타비',
      '순익'
    ];

    if (fieldName && amountFields.includes(fieldName)) {
      const formattedAmount = this.formatCurrency(strValue);
      return { display: formattedAmount, original: formattedAmount, truncated: false };
    }

    // 폴더 경로인 경우 특별 처리 (컬럼이 23%로 넓어졌으므로 더 많이 표시)
    if (fieldName && (fieldName.includes('폴더') || fieldName.includes('경로'))) {
      if (strValue.length > 60) {
        // 경로를 앞부분과 끝부분만 보여주기
        const parts = strValue.split(/[\\\/]/);
        if (parts.length > 2) {
          const firstPart = parts[0];
          const lastPart = parts[parts.length - 1];
          // 첫 부분과 마지막 부분을 합쳐서 55자 이내로 제한
          if ((firstPart + lastPart).length > 50) {
            return {
              display: `${firstPart.substring(0, 20)}/.../${lastPart.substring(0, 25)}...`,
              original: strValue,
              truncated: true
            };
          }
          return {
            display: `${firstPart}/.../${lastPart}...`,
            original: strValue,
            truncated: true
          };
        } else {
          return {
            display: strValue.substring(0, 55) + '...',
            original: strValue,
            truncated: true
          };
        }
      }
      return { display: strValue, original: strValue, truncated: false };
    }

    // 일반 값은 30자로 제한
    const truncated = strValue.length > 30;
    return {
      display: truncated ? strValue.substring(0, 30) + '...' : strValue,
      original: strValue,
      truncated: truncated
    };
  }

  /**
   * 액션 뱃지 생성
   */
  getActionBadge(action) {
    const actionMap = {
      // 프로젝트 관련
      'UPDATE_FIELD': { text: '필드 수정', style: 'background-color: #e3f2fd; color: #1976d2; border: 1px solid #bbdefb;' },
      'UPDATE_FIELD_MEMO': { text: '메모 수정', style: 'background-color: #fff3cd; color: #856404; border: 1px solid #ffeaa7;' },
      'CREATE_PROJECT': { text: '프로젝트 생성', style: 'background-color: #e8f5e8; color: #388e3c; border: 1px solid #c8e6c8;' },
      'DELETE_PROJECT': { text: '프로젝트 삭제', style: 'background-color: #ffebee; color: #c62828; border: 1px solid #ffcdd2;' },
      'CANCEL_PROJECT': { text: '공사 취소', style: 'background-color: #fff3e0; color: #f57c00; border: 1px solid #ffcc80;' },
      'RESUME_PROJECT': { text: '공사 재개', style: 'background-color: #e0f7fa; color: #00838f; border: 1px solid #b2ebf2;' },

      // 인증 관련
      'LOGIN': { text: '로그인', style: 'background-color: #e0f2f1; color: #00796b; border: 1px solid #b2dfdb;' },
      'LOGOUT': { text: '로그아웃', style: 'background-color: #f3e5f5; color: #7b1fa2; border: 1px solid #e1bee7;' },

      // 사용자 관리 관련
      'USER_PERMISSION_UPDATE': { text: '권한 변경', style: 'background-color: #fff3e0; color: #e65100; border: 1px solid #ffe0b2;' },
      'USER_STATUS_CHANGE': { text: '상태 변경', style: 'background-color: #f1f8e9; color: #558b2f; border: 1px solid #dcedc8;' },
      'USER_CREATE': { text: '사용자 생성', style: 'background-color: #e8eaf6; color: #3f51b5; border: 1px solid #c5cae9;' },
      'USER_DELETE': { text: '사용자 삭제', style: 'background-color: #fce4ec; color: #c2185b; border: 1px solid #f8bbd0;' },

      // 폴더 관련
      'FOLDER_OPEN': { text: '폴더 열기', style: 'background-color: #e1f5fe; color: #0277bd; border: 1px solid #b3e5fc;' }
    };

    const actionInfo = actionMap[action] || { text: action, style: 'background-color: #f5f5f5; color: #666; border: 1px solid #ddd;' };
    return `<span class="badge" style="font-size: 0.7rem; ${actionInfo.style}">${actionInfo.text}</span>`;
  }

  /**
   * 권한 뱃지 생성
   */
  getRoleBadge(role) {
    const roleMap = {
      'admin': { text: 'Admin', style: 'background-color: #ffeaa7; color: #d63031; border: 1px solid #fdcb6e;' },
      'super_admin': { text: 'Admin', style: 'background-color: #ffeaa7; color: #d63031; border: 1px solid #fdcb6e;' },
      'editor': { text: 'Editor', style: 'background-color: #fff3cd; color: #856404; border: 1px solid #ffeaa7;' },
      'viewer': { text: 'Viewer', style: 'background-color: #f8f9fa; color: #6c757d; border: 1px solid #dee2e6;' }
    };

    const roleInfo = roleMap[role] || { text: role, style: 'background-color: #f8f9fa; color: #6c757d; border: 1px solid #dee2e6;' };
    return `<span class="badge" style="font-size: 0.7rem; ${roleInfo.style}">${roleInfo.text}</span>`;
  }

  /**
   * 툴팁 속성 생성 함수
   */
  createTooltipAttribute(oldValue, newValue, fieldName, isOldValue = false) {
    // 값이 없는 경우 처리
    if (!oldValue && !newValue) return '';
    if (!oldValue) return '';
    if (!newValue) return '';

    const oldStr = String(oldValue);
    const newStr = String(newValue);
    const displayValue = isOldValue ? oldStr : newStr;

    // 폴더 경로 필드가 아닌 경우 툴팁 없음
    if (!fieldName || (!fieldName.includes('폴더') && !fieldName.includes('경로'))) {
      return '';
    }

    // 이전값 컬럼이거나 값이 동일한 경우 - 변경 강조 없이 표시
    if (isOldValue || oldStr === newStr) {
      return `class="custom-tooltip" data-tooltip-content="${displayValue.replace(/"/g, '&quot;')}" data-old-value="" data-common-start="-1"`;
    }

    // 새값 컬럼에서 변경사항이 있는 경우 - 차이점 찾기
    let commonStart = 0;
    const maxLength = Math.min(oldStr.length, newStr.length);
    for (let i = 0; i < maxLength; i++) {
      if (oldStr[i] !== newStr[i]) {
        break;
      }
      commonStart = i + 1;
    }

    return `class="custom-tooltip" data-tooltip-content="${displayValue.replace(/"/g, '&quot;')}" data-old-value="${oldStr.replace(/"/g, '&quot;')}" data-common-start="${commonStart}"`;
  }

  /**
   * 커스텀 툴팁 생성
   */
  createCustomTooltips() {
    const tooltipElements = document.querySelectorAll('.custom-tooltip');

    tooltipElements.forEach(element => {
      // 기존 툴팁 제거
      const existingTooltip = element.querySelector('.custom-tooltiptext');
      if (existingTooltip) {
        existingTooltip.remove();
      }

      const content = element.getAttribute('data-tooltip-content');
      const oldValue = element.getAttribute('data-old-value');
      const commonStart = parseInt(element.getAttribute('data-common-start'));

      if (content) {
        const tooltipDiv = document.createElement('div');
        tooltipDiv.className = 'custom-tooltiptext';

        // 이전값이거나 변경사항 없는 경우 (commonStart === -1)
        if (!oldValue || commonStart === -1) {
          tooltipDiv.textContent = content;
        } else if (commonStart >= 0) {
          // 새값에서 변경사항이 있는 경우
          const unchangedPart = content.substring(0, commonStart);
          const changedPart = content.substring(commonStart);
          tooltipDiv.innerHTML = `${unchangedPart}<span class="tooltip-changed">${changedPart}</span>`;
        } else {
          tooltipDiv.textContent = content;
        }

        element.appendChild(tooltipDiv);
      }
    });

    // 툴팁 스타일을 동적으로 추가 (한 번만)
    if (!document.getElementById('custom-tooltip-styles')) {
      const style = document.createElement('style');
      style.id = 'custom-tooltip-styles';
      style.textContent = `
        .custom-tooltip {
          position: relative;
          cursor: help;
        }
        .custom-tooltip .custom-tooltiptext {
          visibility: hidden;
          width: 400px;
          background-color: #333;
          color: #fff;
          text-align: left;
          border-radius: 6px;
          padding: 8px;
          position: absolute;
          z-index: 1;
          bottom: 125%;
          left: 50%;
          margin-left: -200px;
          opacity: 0;
          transition: opacity 0.3s;
          font-size: 0.85rem;
          line-height: 1.4;
          white-space: pre-wrap;
          word-break: break-all;
        }
        .custom-tooltip .custom-tooltiptext::after {
          content: "";
          position: absolute;
          top: 100%;
          left: 50%;
          margin-left: -5px;
          border-width: 5px;
          border-style: solid;
          border-color: #333 transparent transparent transparent;
        }
        .custom-tooltip:hover .custom-tooltiptext {
          visibility: visible;
          opacity: 1;
        }
        .tooltip-changed {
          background-color: #ffeb3b;
          color: #333;
          padding: 0 2px;
          border-radius: 3px;
        }
      `;
      document.head.appendChild(style);
    }
  }

  /**
   * 페이지네이션 UI 업데이트
   */
  updatePaginationUI(pagination) {
    const paginationContainer = document.getElementById('auditLogsPagination');
    const prevButton = document.getElementById('auditLogsPrevPage');
    const nextButton = document.getElementById('auditLogsNextPage');

    if (!paginationContainer || !prevButton || !nextButton) return;

    // 이전/다음 버튼 활성화/비활성화
    if (pagination.has_prev) {
      prevButton.classList.remove('disabled');
      prevButton.querySelector('a').setAttribute('data-page', this.currentPage - 1);
    } else {
      prevButton.classList.add('disabled');
      prevButton.querySelector('a').removeAttribute('data-page');
    }

    if (pagination.has_next) {
      nextButton.classList.remove('disabled');
      nextButton.querySelector('a').setAttribute('data-page', this.currentPage + 1);
    } else {
      nextButton.classList.add('disabled');
      nextButton.querySelector('a').removeAttribute('data-page');
    }

    // 기존 페이지 번호들 제거
    const existingPages = paginationContainer.querySelectorAll('.page-number');
    existingPages.forEach(page => page.remove());

    // 새 페이지 번호들 생성
    const maxPagesToShow = 5;
    let startPage = Math.max(1, this.currentPage - Math.floor(maxPagesToShow / 2));
    let endPage = Math.min(pagination.total_pages, startPage + maxPagesToShow - 1);

    if (endPage - startPage + 1 < maxPagesToShow) {
      startPage = Math.max(1, endPage - maxPagesToShow + 1);
    }

    // 페이지 번호 버튼들을 다음 버튼 앞에 삽입
    for (let i = startPage; i <= endPage; i++) {
      const pageItem = document.createElement('li');
      pageItem.className = `page-item page-number ${i === this.currentPage ? 'active' : ''}`;
      pageItem.innerHTML = `<a class="page-link" href="#" data-page="${i}">${i}</a>`;
      paginationContainer.insertBefore(pageItem, nextButton);
    }
  }

  /**
   * 페이지네이션 UI 초기화
   */
  resetPaginationUI() {
    const paginationContainer = document.getElementById('auditLogsPagination');
    const prevButton = document.getElementById('auditLogsPrevPage');
    const nextButton = document.getElementById('auditLogsNextPage');
    const pageInfoElement = document.getElementById('auditLogsPageInfo');

    if (!paginationContainer) return;

    // 모든 페이지 번호 제거
    const existingPages = paginationContainer.querySelectorAll('.page-number');
    existingPages.forEach(page => page.remove());

    // 버튼 비활성화
    if (prevButton) {
      prevButton.classList.add('disabled');
      prevButton.querySelector('a').removeAttribute('data-page');
    }
    if (nextButton) {
      nextButton.classList.add('disabled');
      nextButton.querySelector('a').removeAttribute('data-page');
    }

    // 페이지 정보 초기화
    if (pageInfoElement) {
      pageInfoElement.textContent = '페이지 정보';
    }
  }

  /**
   * 페이지 변경
   */
  changePage(page) {
    if (page < 1 || page > this.totalPages || page === this.currentPage) {
      return;
    }
    this.loadAuditLogs(page);
  }

  /**
   * 모달 숨김 이벤트 처리
   */
  handleModalHidden() {
    // 모달이 닫힐 때 실행할 정리 작업
    this.clearAlert();
  }

  /**
   * 알림 메시지 표시 (감사 로그 모달 전용)
   */
  showAlert(message, type = 'success') {
    const alertContainer = document.getElementById('auditLogsAlert');
    if (!alertContainer) return;

    // 기존 알림 제거
    alertContainer.innerHTML = '';

    // 컴팩트 인라인 알림 (헤더용, 시공자/사용자 모달 패턴과 통일)
    const styleMap = {
      success:   { bg: '#198754', color: '#fff', icon: 'fa-check-circle' },
      secondary: { bg: '#6c757d', color: '#fff', icon: 'fa-minus-circle' },
      danger:    { bg: '#dc3545', color: '#fff', icon: 'fa-exclamation-circle' },
      error:     { bg: '#dc3545', color: '#fff', icon: 'fa-exclamation-circle' },
      warning:   { bg: '#ffc107', color: '#212529', icon: 'fa-exclamation-triangle' },
      info:      { bg: '#0d6efd', color: '#fff', icon: 'fa-info-circle' }
    };
    const s = styleMap[type] || styleMap.info;
    const alertId = 'auditLogsAlertItem_' + Date.now();

    alertContainer.innerHTML = `
      <div id="${alertId}"
           style="display: inline-flex;
                  align-items: center;
                  gap: 0.4rem;
                  background-color: ${s.bg};
                  color: ${s.color};
                  border-radius: 0.375rem;
                  padding: 0.35rem 0.7rem;
                  font-size: 0.8125rem;
                  font-weight: 500;
                  white-space: nowrap;
                  box-shadow: 0 1px 3px rgba(0,0,0,0.08);
                  opacity: 0;
                  transition: opacity 0.2s ease-in-out;">
        <i class="fas ${s.icon}" style="font-size: 0.85rem;"></i>
        <span>${message}</span>
      </div>
    `;

    // fade-in
    requestAnimationFrame(() => {
      const el = document.getElementById(alertId);
      if (el) el.style.opacity = '1';
    });

    // 성공/일반 정보는 3초 후 자동 제거
    if (type === 'success' || type === 'info' || type === 'secondary') {
      setTimeout(() => {
        const el = document.getElementById(alertId);
        if (el) {
          el.style.opacity = '0';
          setTimeout(() => this.clearAlert(), 250);
        }
      }, 3000);
    }
  }

  /**
   * 알림 메시지 제거
   */
  clearAlert() {
    const alertContainer = document.getElementById('auditLogsAlert');
    if (alertContainer) {
      alertContainer.innerHTML = '';
    }
  }

  /**
   * 모달 HTML 생성
   */
  createModalHTML() {
    const modalHTML = `
      <!-- Audit Logs Modal -->
      <div class="modal fade" id="auditLogsModal" tabindex="-1" aria-labelledby="auditLogsModalLabel" data-bs-backdrop="true" data-bs-keyboard="true">
        <div class="modal-dialog audit-logs-modal">
          <div class="modal-content">
            <div class="modal-header" style="min-height: 68px; display: grid; grid-template-columns: auto 1fr auto auto auto; align-items: center; gap: 0.5rem; padding: 1rem;">
              <h5 class="modal-title mb-0 d-flex align-items-center" id="auditLogsModalLabel" style="height: 36px;">
                <i class="fas fa-history me-2"></i>작업 로그
              </h5>
              <!-- 알림 영역: 고정 높이로 layout shift 방지 -->
              <div id="auditLogsAlert" class="audit-logs-alert d-flex align-items-center" style="height: 36px; margin-left: 0.5rem;"></div>
              <select id="auditLogsDays" class="form-select form-select-sm audit-logs-days-select" style="height: 36px;">
                <option value="1">오늘</option>
                <option value="3">3일</option>
                <option value="7" selected>7일</option>
                <option value="14">14일</option>
                <option value="30">30일</option>
              </select>
              <button type="button" class="btn btn-sm btn-outline-secondary d-flex align-items-center" id="refreshAuditLogsBtn"
                      style="height: 36px;">
                <i class="fas fa-sync-alt me-1"></i>새로고침
              </button>
              <button type="button" class="btn-close d-flex align-items-center justify-content-center" data-bs-dismiss="modal" aria-label="닫기"
                      style="background: none; opacity: 1; padding: 0; margin: 0 !important; height: 36px; width: 36px; position: static;">
                <i class="fas fa-times" style="font-size: 1.5rem; color: #6c757d; line-height: 1;"></i>
              </button>
            </div>
            <div class="modal-body audit-logs-body">
              <div class="table-responsive">
                <table class="table table-sm table-hover table-striped audit-logs-table">
                  <thead class="table-light audit-logs-thead">
                    <tr>
                      <th>시간</th>
                      <th>사용자</th>
                      <th>권한</th>
                      <th>작업</th>
                      <th>프로젝트</th>
                      <th>필드명</th>
                      <th>이전값</th>
                      <th>새값</th>
                    </tr>
                  </thead>
                  <tbody id="auditLogsTableBody">
                    <tr>
                      <td colspan="8" class="text-center">
                        <i class="fas fa-spinner fa-spin"></i> 로딩 중...
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
            <div class="modal-footer">
              <!-- 한 줄에 좌측(로그 정보) - 중앙(페이지네이션) - 우측(닫기 버튼) -->
              <div class="d-flex align-items-center w-100 audit-logs-footer">
                <!-- 좌측: 로그 정보 -->
                <div class="d-flex align-items-center audit-logs-info">
                  <small class="text-muted me-3" id="auditLogsCount">로그 개수: 0개</small>
                  <small class="text-muted" id="auditLogsPageInfo">페이지 정보</small>
                </div>

                <!-- 중앙: 페이지네이션 (절대 중앙 배치) -->
                <div class="audit-logs-pagination-center">
                  <nav aria-label="감사 로그 페이지네이션">
                    <ul class="pagination pagination-sm mb-0" id="auditLogsPagination">
                      <li class="page-item disabled" id="auditLogsPrevPage">
                        <a class="page-link" href="#" tabindex="-1" aria-disabled="true">
                          <i class="fas fa-chevron-left"></i>
                        </a>
                      </li>
                      <!-- 페이지 번호들이 여기에 동적으로 추가됩니다 -->
                      <li class="page-item disabled" id="auditLogsNextPage">
                        <a class="page-link" href="#" tabindex="-1" aria-disabled="true">
                          <i class="fas fa-chevron-right"></i>
                        </a>
                      </li>
                    </ul>
                  </nav>
                </div>

                <!-- 우측: 빈 공간 유지 (헤더 X 버튼으로 닫기) -->
                <div class="audit-logs-close" aria-hidden="true"></div>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;

    // DOM에 모달 HTML 추가
    document.body.insertAdjacentHTML('beforeend', modalHTML);
  }

  /**
   * 컴포넌트 정리
   */
  destroy() {
    if (this.modal) {
      this.modal.dispose();
    }
  }
}