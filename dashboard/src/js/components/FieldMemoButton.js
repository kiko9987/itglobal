/**
 * Field Memo Button Component
 * 필드별 메모 추가/편집/삭제 기능을 제공하는 버튼 컴포넌트
 * 구글 시트의 메모 기능과 동일한 UX
 */
import logger from '../utils/logger.js';

export default class FieldMemoButton {
  constructor() {
    this.activePopover = null;
    this.currentFieldName = null;  // fieldKey → fieldName으로 변경
    this.currentProjectCode = null;
    this.currentMemo = null;
    this.popoverElement = null;
    this.previousActiveElement = null;  // 포커스 복원용

    // 메모 가능한 필드 목록 (계약금, 중도금, 잔금)
    this.memoableFields = ['계약금', '중도금', '잔금'];

    // 툴팁 요소 캐시
    this.activeTooltips = new Map();

    this.init();
  }

  /**
   * 초기화
   */
  init() {
    this.createPopoverTemplate();
    this.bindEvents();
  }

  /**
   * 팝오버 템플릿 생성 (재사용)
   */
  createPopoverTemplate() {
    const template = document.createElement('div');
    template.className = 'field-memo-popover';
    template.innerHTML = `
      <div class="memo-popover-header">
        <span class="memo-field-name"></span>
        <button type="button" class="memo-close-btn" aria-label="닫기">
          <i class="fas fa-times"></i>
        </button>
      </div>
      <div class="memo-popover-body">
        <textarea class="memo-textarea form-control"
                  placeholder="입금 정보를 메모하세요."
                  rows="4"></textarea>
      </div>
      <div class="memo-popover-footer">
        <button type="button" class="btn btn-sm btn-outline-danger memo-delete-btn" style="display:none;">
          <i class="fas fa-trash me-1"></i>삭제
        </button>
        <div class="ms-auto">
          <button type="button" class="btn btn-sm btn-secondary memo-cancel-btn"
                  style="background-color: #e9ecef; border-color: #dee2e6; color: #495057; padding: 0.25rem 0.75rem; font-size: 0.875rem;">취소</button>
          <button type="button" class="btn btn-sm btn-primary memo-save-btn"
                  style="background-color: #d0e7ff; border-color: #b6d4fe; color: #0066cc; padding: 0.25rem 0.75rem; font-size: 0.875rem;">저장</button>
        </div>
      </div>
    `;

    // body에 추가 (숨김 상태)
    template.style.display = 'none';
    document.body.appendChild(template);

    this.popoverElement = template;
  }

  /**
   * 이벤트 바인딩
   */
  bindEvents() {
    // 팝오버 내부 버튼 이벤트
    this.popoverElement.querySelector('.memo-close-btn').addEventListener('click', () => this.hidePopover());
    this.popoverElement.querySelector('.memo-cancel-btn').addEventListener('click', () => this.hidePopover());
    this.popoverElement.querySelector('.memo-save-btn').addEventListener('click', () => this.saveMemo());
    this.popoverElement.querySelector('.memo-delete-btn').addEventListener('click', () => this.deleteMemo());

    // 외부 클릭으로 팝오버 닫기
    document.addEventListener('click', (e) => {
      if (this.activePopover &&
          !this.popoverElement.contains(e.target) &&
          !e.target.closest('.field-memo-btn')) {
        this.hidePopover();
      }
    });

    // ESC 키로 팝오버 닫기
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && this.activePopover) {
        this.hidePopover();
      }
    });
  }

  /**
   * 메모 버튼 HTML 생성
   * @param {string} fieldName - 필드명 (계약금, 중도금, 잔금)
   * @param {string} projectCode - 프로젝트 코드
   * @param {string|null} memo - 기존 메모 내용
   * @param {number} amount - 해당 필드의 금액 (0원이면 버튼 숨김)
   * @param {boolean} isEditMode - 편집 모드 여부 (편집 모드에서는 항상 버튼 표시)
   * @returns {string} 버튼 HTML
   */
  createButton(fieldName, projectCode, memo = null, amount = 0, isEditMode = false) {
    if (!this.memoableFields.includes(fieldName)) {
      return '';
    }

    const hasMemo = memo && memo.trim();
    const hasAmount = amount && parseFloat(amount) > 0;

    // UX 개선: 편집 모드가 아닐 때만 금액 체크
    // - 편집 모드: 금액 입력 중일 수 있으므로 항상 버튼 표시
    // - 일반 모드: 금액이 없고 메모도 없으면 버튼 숨김 (메모가 있으면 삭제용으로 표시)
    if (!isEditMode && !hasAmount && !hasMemo) {
      return '';
    }

    const icon = hasMemo ? 'fa-sticky-note' : 'fa-plus';
    const title = hasMemo ? '메모 보기/수정' : '메모 추가';
    const btnClass = hasMemo ? 'has-memo' : 'no-memo';

    return `
      <button type="button"
              class="field-memo-btn ${btnClass}"
              data-field="${fieldName}"
              data-project-code="${projectCode}"
              title="${title}"
              aria-label="${title}">
        <i class="fas ${icon}"></i>
      </button>
    `;
  }

  /**
   * 메모 버튼에 이벤트 리스너 연결
   * @param {HTMLElement} container - 버튼이 포함된 컨테이너
   * @param {Object|Function} memoDataOrGetter - 메모 데이터 객체 또는 최신 데이터를 반환하는 getter 함수
   */
  attachButtonListeners(container, memoDataOrGetter = {}) {
    const buttons = container.querySelectorAll('.field-memo-btn');

    buttons.forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();

        const fieldName = btn.dataset.field;
        const projectCode = btn.dataset.projectCode;

        // 클릭 시점에 최신 메모 데이터 가져오기
        const memoData = typeof memoDataOrGetter === 'function'
          ? memoDataOrGetter()
          : memoDataOrGetter;

        // memoData에서 해당 필드의 메모 찾기 (예: "계약금_메모")
        const memoKey = `${fieldName}_메모`;
        const memo = memoData[memoKey] || null;

        this.showPopover(btn, fieldName, projectCode, memo);
      });
    });
  }

  /**
   * 팝오버 표시
   */
  showPopover(buttonElement, fieldName, projectCode, currentMemo) {
    // 팝오버 열기 전 현재 포커스된 요소 저장 (저장 후 복원용)
    this.previousActiveElement = document.activeElement;

    this.currentFieldName = fieldName;  // fieldKey 대신 fieldName 저장
    this.currentProjectCode = projectCode;
    this.currentMemo = currentMemo;

    // 편집 모드 확인
    const collectionCard = document.getElementById(`card-collection-${projectCode}`);
    const isEditMode = collectionCard?.classList.contains('editing');

    // 팝오버 내용 설정
    const header = this.popoverElement.querySelector('.memo-field-name');
    const textarea = this.popoverElement.querySelector('.memo-textarea');
    const deleteBtn = this.popoverElement.querySelector('.memo-delete-btn');
    const saveBtn = this.popoverElement.querySelector('.memo-save-btn');
    const cancelBtn = this.popoverElement.querySelector('.memo-cancel-btn');

    header.textContent = `${fieldName} 입금 메모`;
    textarea.value = currentMemo || '';

    // 읽기 모드: 저장/삭제 버튼 숨김, textarea 읽기 전용
    if (!isEditMode) {
      textarea.setAttribute('readonly', 'true');
      textarea.classList.add('readonly');
      deleteBtn.style.display = 'none';
      saveBtn.style.display = 'none';
      cancelBtn.textContent = '닫기';
      header.textContent = `${fieldName} 입금 메모 (읽기 전용)`;
      logger.debug('[MEMO] 읽기 모드 - 팝오버 읽기 전용으로 표시');
    } else {
      // 편집 모드: 저장/삭제 버튼 표시, textarea 편집 가능
      textarea.removeAttribute('readonly');
      textarea.classList.remove('readonly');
      deleteBtn.style.display = currentMemo ? 'inline-block' : 'none';
      saveBtn.style.display = 'inline-block';
      cancelBtn.textContent = '취소';
      logger.debug('[MEMO] 편집 모드 - 팝오버 편집 가능으로 표시');
    }

    // 팝오버 위치 계산 및 표시
    this.positionPopover(buttonElement);

    // 팝오버 표시
    this.popoverElement.style.display = 'block';
    this.activePopover = buttonElement;

    // 편집 모드일 때만 textarea에 포커스
    if (isEditMode) {
      setTimeout(() => textarea.focus(), 100);
    }
  }

  /**
   * 팝오버 위치 계산
   */
  positionPopover(anchorElement) {
    const rect = anchorElement.getBoundingClientRect();
    const popover = this.popoverElement;

    // 기본 위치: 버튼 우측 하단
    let top = rect.bottom + 5;
    let left = rect.left;

    // 화면 밖으로 나가는지 체크
    const popoverRect = popover.getBoundingClientRect();
    const viewportWidth = window.innerWidth;
    const viewportHeight = window.innerHeight;

    // 오른쪽 overflow 체크
    if (left + 300 > viewportWidth) {
      left = rect.right - 300;
    }

    // 하단 overflow 체크
    if (top + popoverRect.height > viewportHeight) {
      top = rect.top - popoverRect.height - 5;
    }

    popover.style.top = `${top}px`;
    popover.style.left = `${left}px`;
  }

  /**
   * 팝오버 숨기기
   */
  hidePopover() {
    if (this.popoverElement) {
      this.popoverElement.style.display = 'none';
    }
    this.activePopover = null;
    this.currentFieldName = null;  // fieldKey → fieldName
    this.currentProjectCode = null;
    this.currentMemo = null;
    this.previousActiveElement = null;  // 포커스 복원용 초기화
  }

  /**
   * 메모 저장 (구글 시트 셀 메모로 저장)
   */
  async saveMemo() {
    const textarea = this.popoverElement.querySelector('.memo-textarea');
    const newMemo = textarea.value.trim();

    if (!this.currentProjectCode || !this.currentFieldName) {
      logger.error('[ERROR] 메모 저장 실패: 프로젝트 코드 또는 필드명 없음');
      return;
    }

    // 편집 모드 확인: 아코디언이 editing 클래스를 가지고 있는지 체크
    const collectionCard = document.getElementById(`card-collection-${this.currentProjectCode}`);
    const isEditMode = collectionCard?.classList.contains('editing');

    // 편집 모드일 때: 임시 저장만 (API 호출 안 함)
    if (isEditMode) {
      logger.debug('[MEMO] 편집 모드 - 임시 저장 (API 호출 건너뜀)');

      // 로컬 데이터만 업데이트
      this.updateLocalMemoData(newMemo);

      // 버튼 아이콘 업데이트
      this.updateButtonIcon(this.activePopover, newMemo);

      // 팝오버 닫기
      this.hidePopover();

      logger.debug('[INFO] 메모 임시 저장 완료 (통합 저장 시 반영됨):', this.currentFieldName, '→', newMemo ? '작성됨' : '삭제됨');
      return;
    }

    // 읽기 모드일 때: 즉시 API 호출 (기존 동작)
    try {
      // 구글 시트 셀 메모 API 호출 (field_name만 전송)
      const response = await fetch('/api/projects/field-memo', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        credentials: 'same-origin',
        body: JSON.stringify({
          project_code: this.currentProjectCode,
          field_name: this.currentFieldName,  // field_key → field_name으로 변경
          memo: newMemo || null // 빈 메모는 null로 저장 (셀 메모 삭제)
        })
      });

      if (!response.ok) {
        throw new Error(`메모 저장 실패: ${response.status}`);
      }

      const result = await response.json();

      if (result.success) {
        // 버튼 아이콘 업데이트
        this.updateButtonIcon(this.activePopover, newMemo);

        // 커스텀 이벤트 발생 (다른 컴포넌트에서 메모 업데이트 감지)
        document.dispatchEvent(new CustomEvent('fieldMemoUpdated', {
          detail: {
            projectCode: this.currentProjectCode,
            fieldName: this.currentFieldName,  // fieldKey → fieldName
            memo: newMemo
          }
        }));

        // 전역 데이터 동기화 (모든 데이터 소스 통합 관리)
        const memoKey = `${this.currentFieldName}_메모`;

        // 1. DataTables 캐시 동기화 (데이터만 갱신, draw 제거로 편집 모드 유지)
        if (window.projectListApp?.components?.table?.table) {
          const table = window.projectListApp.components.table.table;
          const rows = table.rows().data().toArray();

          rows.forEach((row, idx) => {
            if (row['프로젝트 코드'] === this.currentProjectCode) {
              row[memoKey] = newMemo;
              table.row(idx).data(row); // 데이터만 갱신

              // 수금 모드일 때: 해당 셀 HTML 직접 갱신 + 툴팁 재초기화
              const projectTable = window.projectListApp.components.table;
              if (projectTable?.modeManager.isReceivablesMode()) {
                const columnName = this.currentFieldName === '계약금' ? 'downPayment' :
                                   this.currentFieldName === '중도금' ? 'midPayment' :
                                   this.currentFieldName === '잔금' ? 'finalPayment' : null;

                if (columnName) {
                  const cellApi = table.cell(idx, `${columnName}:name`);
                  const cellNode = cellApi.node();

                  if (cellNode) {
                    // 셀 HTML 재렌더링
                    const newHtml = cellApi.render('display');
                    cellNode.innerHTML = newHtml;

                    // 전역 툴팁 초기화 함수 호출 (해당 셀만 대상)
                    if (projectTable.initializeTooltips) {
                      projectTable.initializeTooltips(cellNode);
                    }

                    logger.debug(`[MEMO] 수금 모드 셀 HTML 갱신 + 툴팁 재초기화: ${columnName}`);
                  }
                }
              }
            }
          });

          // draw(false) 제거: 메모 컬럼은 수금 모드에서만 툴팁으로 표시되므로
          // draw 없이 데이터만 갱신하면 아코디언/편집 모드가 유지됩니다.
          logger.debug(`[MEMO] DataTables 캐시 갱신 완료 (draw 생략): ${this.currentProjectCode} / ${this.currentFieldName}`);
        }

        // 2. filters.currentData 동기화 (필터/통계 정확도 유지)
        if (window.projectListApp?.components?.filters?.currentData) {
          const filtersData = window.projectListApp.components.filters.currentData;
          const targetProject = filtersData.find(p => p['프로젝트 코드'] === this.currentProjectCode);

          if (targetProject) {
            targetProject[memoKey] = newMemo;
            logger.debug(`[MEMO] filters.currentData 갱신 완료: ${this.currentProjectCode} / ${this.currentFieldName}`);
          }
        }

        // 3. StateManager.currentData 동기화 (전역 상태 관리)
        if (window.projectListApp?.stateManager?.currentData) {
          const stateData = window.projectListApp.stateManager.currentData;
          const targetProject = stateData.find(p => p['프로젝트 코드'] === this.currentProjectCode);

          if (targetProject) {
            targetProject[memoKey] = newMemo;
            logger.debug(`[MEMO] StateManager.currentData 갱신 완료: ${this.currentProjectCode} / ${this.currentFieldName}`);
          }
        }

        // 4. window.projectsData 동기화 (레거시 전역 데이터)
        if (window.projectsData && Array.isArray(window.projectsData)) {
          const targetProject = window.projectsData.find(p => p['프로젝트 코드'] === this.currentProjectCode);

          if (targetProject) {
            targetProject[memoKey] = newMemo;
            logger.debug(`[MEMO] window.projectsData 갱신 완료: ${this.currentProjectCode} / ${this.currentFieldName}`);
          }
        }

        // projectUpdated 이벤트는 발생시키지 않음 (메모는 셀 메모일 뿐 프로젝트 데이터 변경이 아님)
        // 이미 모든 데이터 소스(DataTables, filters, StateManager, window.projectsData)를
        // 직접 동기화했으므로 추가 이벤트 불필요
        // projectUpdated 이벤트는 updateDataTableRow()를 트리거해서 draw()를 호출하므로
        // 아코디언이 닫히는 부작용이 있음

        // localStorage 캐시 클리어 (F5 새로고침 시 최신 메모 표시 보장)
        if (window.projectListApp?.dataManager) {
          window.projectListApp.dataManager.clearCache();
          logger.debug('[MEMO] localStorage 캐시 클리어 완료 - F5 새로고침 시 최신 데이터 로드됨');
        }

        // 포커스 복원을 위해 원래 요소 참조 저장
        const elementToRestore = this.previousActiveElement || this.activePopover;

        // 팝오버 닫기
        this.hidePopover();

        // 원래 입력 필드로 포커스 복원 (메모 버튼이 아닌 원래 위치)
        // 단, 아코디언 내부 요소인 경우에만 복원 (아코디언 밖으로 포커스 이동 방지)
        if (elementToRestore && typeof elementToRestore.focus === 'function') {
          setTimeout(() => {
            try {
              // 아코디언 컨테이너 확인
              const accordionContainer = elementToRestore.closest('.project-accordion-container');

              // 아코디언 내부 요소이고, DOM에 여전히 존재하는 경우에만 포커스 복원
              if (accordionContainer && document.body.contains(elementToRestore)) {
                elementToRestore.focus();
                logger.debug('[MEMO] 포커스 복원 완료:', elementToRestore.tagName, elementToRestore.className);
              } else {
                logger.debug('[MEMO] 포커스 복원 스킵 (아코디언 외부 요소 또는 제거된 요소)');
              }
            } catch (error) {
              logger.warn('[MEMO] 포커스 복원 실패:', error);
            }
          }, 100);
        }

        logger.debug('[INFO] 셀 메모 저장 완료:', this.currentFieldName, '→', newMemo ? '작성됨' : '삭제됨');

      } else {
        throw new Error(result.message || '메모 저장 실패');
      }

    } catch (error) {
      logger.error('[ERROR] 메모 저장 오류:', error);
      alert('메모 저장 중 오류가 발생했습니다.');
    }
  }

  /**
   * 메모 삭제
   */
  async deleteMemo() {
    if (!confirm('메모를 삭제하시겠습니까?')) {
      return;
    }

    // 빈 메모로 저장 (삭제 효과)
    const textarea = this.popoverElement.querySelector('.memo-textarea');
    textarea.value = '';
    await this.saveMemo();
  }

  /**
   * 로컬 메모 데이터 업데이트 (API 호출 없이)
   * 편집 모드에서 임시 저장용
   */
  updateLocalMemoData(newMemo) {
    const memoKey = `${this.currentFieldName}_메모`;

    // 1. DataTables 캐시 동기화
    if (window.projectListApp?.components?.table?.table) {
      const table = window.projectListApp.components.table.table;
      const rows = table.rows().data().toArray();

      rows.forEach((row, idx) => {
        if (row['프로젝트 코드'] === this.currentProjectCode) {
          row[memoKey] = newMemo;
          table.row(idx).data(row); // 데이터만 갱신, draw 없음
        }
      });

      logger.debug(`[MEMO] DataTables 캐시 갱신 (로컬): ${this.currentProjectCode} / ${this.currentFieldName}`);
    }

    // 2. filters.currentData 동기화
    if (window.projectListApp?.components?.filters?.currentData) {
      const filtersData = window.projectListApp.components.filters.currentData;
      const targetProject = filtersData.find(p => p['프로젝트 코드'] === this.currentProjectCode);

      if (targetProject) {
        targetProject[memoKey] = newMemo;
        logger.debug(`[MEMO] filters.currentData 갱신 (로컬): ${this.currentProjectCode} / ${this.currentFieldName}`);
      }
    }

    // 3. StateManager.currentData 동기화
    if (window.projectListApp?.stateManager?.currentData) {
      const stateData = window.projectListApp.stateManager.currentData;
      const targetProject = stateData.find(p => p['프로젝트 코드'] === this.currentProjectCode);

      if (targetProject) {
        targetProject[memoKey] = newMemo;
        logger.debug(`[MEMO] StateManager.currentData 갱신 (로컬): ${this.currentProjectCode} / ${this.currentFieldName}`);
      }
    }

    // 4. window.projectsData 동기화
    if (window.projectsData && Array.isArray(window.projectsData)) {
      const targetProject = window.projectsData.find(p => p['프로젝트 코드'] === this.currentProjectCode);

      if (targetProject) {
        targetProject[memoKey] = newMemo;
        logger.debug(`[MEMO] window.projectsData 갱신 (로컬): ${this.currentProjectCode} / ${this.currentFieldName}`);
      }
    }

    // 5. ProjectRowAccordion의 currentProject 동기화 (가장 중요!)
    // 6. EditState 동기화 (편집 모드일 때만)
    // window.projectRowAccordion이 실제로 사용되는 accordion 인스턴스
    const accordion = window.projectRowAccordion || window.projectListApp?.components?.accordion;

    if (accordion?.currentProject) {
      const accordionProject = accordion.currentProject;
      if (accordionProject['프로젝트 코드'] === this.currentProjectCode) {
        accordionProject[memoKey] = newMemo;
        logger.debug(`[MEMO] Accordion.currentProject 갱신 (로컬): ${this.currentProjectCode} / ${this.currentFieldName}`);
      }
    }

    // EditState 동기화
    if (accordion?.editState?.isActive) {
      accordion.editState.updateField(memoKey, newMemo);
      logger.debug(`[EditState] ${memoKey} 메모 업데이트:`, newMemo);
    }

    // 🆕 pendingMemoChanges 동기화 (편집 모드에서 메모 변경 추적)
    if (accordion?.modeManager?.isEditMode() && accordion.editState?.isActive) {
      accordion.pendingMemoChanges[memoKey] = newMemo;
      logger.debug(`[메모 추적] ${memoKey} 변경 감지:`, newMemo?.substring(0, 50) || '(빈 값)');
    }
  }

  /**
   * 버튼 아이콘 업데이트
   */
  updateButtonIcon(buttonElement, memo) {
    if (!buttonElement) return;

    const icon = buttonElement.querySelector('i');
    if (!icon) {
      logger.warn('[MEMO] 버튼 아이콘 요소를 찾을 수 없습니다:', buttonElement);
      return;
    }

    const hasMemo = memo && memo.trim();

    if (hasMemo) {
      icon.className = 'fas fa-sticky-note';
      buttonElement.classList.add('has-memo');
      buttonElement.classList.remove('no-memo');
      buttonElement.title = '메모 보기/수정';
    } else {
      icon.className = 'fas fa-plus';
      buttonElement.classList.remove('has-memo');
      buttonElement.classList.add('no-memo');
      buttonElement.title = '메모 추가';
    }
  }

  /**
   * 툴팁 생성 및 표시 (읽기 전용)
   * @param {HTMLElement} element - 툴팁을 표시할 요소
   * @param {string} memo - 메모 내용
   */
  showTooltip(element, memo) {
    if (!memo || !memo.trim()) return;

    // 기존 툴팁 제거
    this.hideTooltip(element);

    // 툴팁 요소 생성
    const tooltip = document.createElement('div');
    tooltip.className = 'field-memo-tooltip';
    tooltip.textContent = memo;

    // 위치 계산
    const rect = element.getBoundingClientRect();
    tooltip.style.position = 'fixed';
    tooltip.style.top = `${rect.bottom + 5}px`;
    tooltip.style.left = `${rect.left}px`;

    document.body.appendChild(tooltip);
    this.activeTooltips.set(element, tooltip);

    // 일정 시간 후 자동 제거
    setTimeout(() => this.hideTooltip(element), 5000);
  }

  /**
   * 툴팁 숨기기
   */
  hideTooltip(element) {
    const tooltip = this.activeTooltips.get(element);
    if (tooltip && tooltip.parentNode) {
      tooltip.parentNode.removeChild(tooltip);
      this.activeTooltips.delete(element);
    }
  }

  /**
   * 모든 툴팁 제거
   */
  hideAllTooltips() {
    this.activeTooltips.forEach((tooltip, element) => {
      if (tooltip.parentNode) {
        tooltip.parentNode.removeChild(tooltip);
      }
    });
    this.activeTooltips.clear();
  }

  /**
   * 정리
   */
  destroy() {
    this.hidePopover();
    this.hideAllTooltips();

    if (this.popoverElement && this.popoverElement.parentNode) {
      this.popoverElement.parentNode.removeChild(this.popoverElement);
    }

    this.popoverElement = null;
    this.activePopover = null;
  }
}
