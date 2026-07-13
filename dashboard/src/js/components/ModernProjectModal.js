/**
 * ModernProjectModal - 레거시의 모든 기능을 완전히 포함한 현대화 모달
 */
import AmountCalculator from '../utils/AmountCalculator.js';
import FormValidator from '../utils/FormValidator.js';
import ProjectCodeGenerator from '../utils/ProjectCodeGenerator.js';
import { initLeadProjectLoader } from '../utils/leadProjectLoader.js';

import logger from '../utils/logger.js';
export default class ModernProjectModal {
  constructor() {
    this.modal = null;
    this.modalElement = null;
    this.form = null;
    this.toastComponent = null;

    // 유틸리티 인스턴스
    this.codeGenerator = new ProjectCodeGenerator();

    // 데이터 로드 재시도 설정
    this.loadRetryCount = 0;
    this.MAX_LOAD_RETRIES = 10; // 최대 2초 대기 (10 × 200ms)

    this.initialized = false;
  }

  /**
   * 모달 초기화
   */
  async init() {
    if (this.initialized) return;

    this.modalElement = document.getElementById('modernProjectModal');
    if (!this.modalElement) {
      logger.warn('현대화된 프로젝트 모달 엘리먼트를 찾을 수 없습니다.');
      return;
    }

    this.modal = new bootstrap.Modal(this.modalElement);
    this.form = document.getElementById('modernProjectForm');

    // 제출·재시도 진행 중 모달 닫기 방지 (X 버튼·ESC·backdrop 클릭 모두 차단)
    // Bootstrap 5: hide.bs.modal 이벤트에서 preventDefault() 로 취소 가능.
    this.modalElement.addEventListener('hide.bs.modal', (ev) => {
      if (this._submitInProgress) {
        ev.preventDefault();
        // 잠깐 안내 (연속 시도 방지 위해 debounce)
        if (!this._closePreventNotified) {
          this._closePreventNotified = true;
          this.showAlert('등록이 진행 중입니다. 완료 후 자동으로 닫힙니다.', 'info');
          setTimeout(() => { this._closePreventNotified = false; }, 3000);
        }
      }
    });

    this.bindEvents();
    this.initialized = true;
  }

  /**
   * 이벤트 바인딩
   */
  bindEvents() {
    // 프로젝트 코드 미리보기 - populateDropdowns에서 재바인딩됨

    // 금액 입력 - 콤마 포맷팅
    const amountDisplayInput = document.getElementById('modern-total-amount-display');
    const vatCheckbox = document.getElementById('modern-vat-included');

    if (amountDisplayInput) {
      amountDisplayInput.addEventListener('input', (e) => this.handleAmountInput(e));
      amountDisplayInput.addEventListener('blur', () => this.updateAmount());
    }
    if (vatCheckbox) {
      vatCheckbox.addEventListener('change', () => this.updateAmount());
    }

    // 빠른 금액 버튼
    const amountBtns = this.modalElement.querySelectorAll('.modern-amount-btn');
    amountBtns.forEach(btn => {
      btn.addEventListener('click', (e) => {
        const amount = parseInt(e.target.dataset.amount);
        this.addQuickAmount(amount);
      });
    });

    const resetBtn = document.getElementById('modern-amount-reset-btn');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => this.resetAmount());
    }

    // 공사 구분 체크박스 (최대 2개)
    const workCategoryCheckboxes = this.modalElement.querySelectorAll('.modern-work-category-checkbox');
    workCategoryCheckboxes.forEach(checkbox => {
      checkbox.addEventListener('change', (e) => this.handleCheckboxLimit(e, '.modern-work-category-checkbox', 'modern-work-category-error', 2, 'modern-work-category'));
    });

    // 기계 분류 체크박스 (최대 2개)
    const machineCategoryCheckboxes = this.modalElement.querySelectorAll('.modern-machine-category-checkbox');
    machineCategoryCheckboxes.forEach(checkbox => {
      checkbox.addEventListener('change', (e) => this.handleCheckboxLimit(e, '.modern-machine-category-checkbox', 'modern-machine-category-error', 2));
    });

    // 브랜드 체크박스 (최대 2개)
    const brandCategoryCheckboxes = this.modalElement.querySelectorAll('.modern-brand-category-checkbox');
    brandCategoryCheckboxes.forEach(checkbox => {
      checkbox.addEventListener('change', (e) => this.handleCheckboxLimit(e, '.modern-brand-category-checkbox', 'modern-brand-category-error', 2));
    });

    // 도급 구분 체크박스 (최대 2개)
    const contractCategoryCheckboxes = this.modalElement.querySelectorAll('.modern-contract-category-checkbox');
    contractCategoryCheckboxes.forEach(checkbox => {
      checkbox.addEventListener('change', (e) => this.handleCheckboxLimit(e, '.modern-contract-category-checkbox', 'modern-contract-category-error', 2));
    });

    // 시공자 체크박스 - 이벤트 위임 (체크박스가 동적으로 생성되므로 정적 바인딩 안 됨)
    // window.renderConstructorCheckboxes 헬퍼가 매번 새로 그리기 때문에
    // 모달 요소에 위임하여 동적 체크박스도 정상 작동
    this.modalElement.addEventListener('change', (e) => {
      // 일반 시공자 체크박스 (최대 3개)
      if (e.target.matches('.modern-constructor-checkbox')) {
        this.handleCheckboxLimit(e, '.modern-constructor-checkbox', 'modern-constructor-error', 3);
        this.updateConstructorHiddenInput();
        return;
      }
      // 기타 체크박스 - 입력창 토글
      if (e.target.id === 'modern-constructor-other-check') {
        const otherInput = document.getElementById('modern-constructor-other-input');
        if (otherInput) {
          if (e.target.checked) {
            otherInput.classList.remove('d-none');
            otherInput.focus();
          } else {
            otherInput.classList.add('d-none');
            otherInput.value = '';
          }
          this.updateConstructorHiddenInput();
        }
      }
    });

    // 기타 입력란 텍스트 변경 - 이벤트 위임
    this.modalElement.addEventListener('input', (e) => {
      if (e.target.id === 'modern-constructor-other-input') {
        this.updateConstructorHiddenInput();
      }
    });

    // 날짜 동기화 ('시작일과 동일' 체크박스)
    const sameDateCheck = document.getElementById('modern-same-date-check');
    const startDateInput = document.getElementById('modern-start-date');
    const endDateInput = document.getElementById('modern-end-date');

    if (sameDateCheck && startDateInput && endDateInput) {
      sameDateCheck.addEventListener('change', (e) => {
        if (e.target.checked) {
          endDateInput.value = startDateInput.value;
          endDateInput.disabled = true;
        } else {
          endDateInput.disabled = false;
        }
      });

      startDateInput.addEventListener('change', () => {
        // 종료일의 최소값을 시작일로 설정
        if (startDateInput.value) {
          endDateInput.min = startDateInput.value;
        }

        // 시작일과 동일 체크박스가 체크되어 있으면 종료일도 동일하게
        if (sameDateCheck.checked) {
          endDateInput.value = startDateInput.value;
        }
      });
    }

    // 이메일 검증
    const emailInput = document.getElementById('modern-manager-email');
    const emailUnknownCheck = document.getElementById('modern-email-unknown-check');

    if (emailInput && emailUnknownCheck) {
      emailInput.addEventListener('input', () => this.validateEmail());
      emailInput.addEventListener('blur', () => this.validateEmail());

      emailUnknownCheck.addEventListener('change', (e) => {
        if (e.target.checked) {
          emailInput.value = '-';
          emailInput.readOnly = true;  // disabled 대신 readOnly 사용 (FormData에 포함됨)
          emailInput.style.backgroundColor = '#e9ecef';  // 비활성화된 것처럼 보이게
          emailInput.style.cursor = 'not-allowed';
          this.hideEmailError();
        } else {
          emailInput.value = '';
          emailInput.readOnly = false;
          emailInput.style.backgroundColor = '';
          emailInput.style.cursor = '';
        }
      });
    }

    // 주소 검증 (도로명/지번)
    const addressInput = document.getElementById('modern-address');
    if (addressInput) {
      addressInput.addEventListener('blur', () => this.validateAddress());
    }

    // 유입 구분 ↔ 사업자명 잠금 연동
    // 거래처가 아니면 사업자명은 '-' 고정 + readonly (사업자등록증 미수령 전 상태)
    // 거래처면 편집 가능 + datalist 자동완성 복구
    const clientSelectForLock = document.getElementById('modern-client');
    if (clientSelectForLock) {
      clientSelectForLock.addEventListener('change', () => this.syncBusinessNameLock());
      // 초기 로드 시에도 상태 정리 (기본값이 비어있으면 lock 해제)
      this.syncBusinessNameLock();
    }

    // 날짜 입력 필드 클릭 이벤트 (아이콘 클릭 시 날짜 선택창 열기)
    const calendarIcons = this.modalElement.querySelectorAll('.calendar-icon');
    calendarIcons.forEach(icon => {
      icon.addEventListener('click', () => {
        const inputGroup = icon.closest('.input-group');
        const dateInput = inputGroup.querySelector('input[type="date"]');
        if (dateInput && !dateInput.disabled) {
          dateInput.showPicker();
        }
      });
    });

    // 날짜 입력 필드 직접 클릭 시에도 날짜 선택창 열기
    const dateInputs = this.modalElement.querySelectorAll('input[type="date"]');
    dateInputs.forEach(input => {
      input.addEventListener('click', () => {
        if (!input.disabled) {
          input.showPicker();
        }
      });
    });

    // 폴더 경로 입력 검증
    const documentPathInput = document.getElementById('modern-document-path');
    if (documentPathInput) {
      documentPathInput.addEventListener('input', () => this.validateFolderPath());
      documentPathInput.addEventListener('paste', () => {
        // paste 이벤트 후 input 값이 업데이트될 때까지 대기
        setTimeout(() => this.validateFolderPath(), 10);
      });
    }

    // 폼 제출
    if (this.form) {
      this.form.addEventListener('submit', (e) => {
        e.preventDefault();
        this.submitProject();
      });
    }

    // 모달 닫히기 전 이벤트 (포커스 관리 - 접근성)
    this.modalElement.addEventListener('hide.bs.modal', () => {
      // 모달 내부에 포커스가 있으면 안전한 곳으로 이동
      // aria-hidden이 설정되기 전에 포커스를 이동시켜 접근성 경고 방지
      const activeElement = document.activeElement;
      if (this.modalElement.contains(activeElement)) {
        // 모달을 연 버튼으로 포커스 이동 (없으면 body)
        const triggerButton = document.querySelector('[data-bs-target="#modernProjectModal"]');
        if (triggerButton && triggerButton.offsetParent !== null) {
          triggerButton.focus();
        } else {
          document.body.focus();
        }
      }
    });

    // 모달 닫힘 이벤트
    this.modalElement.addEventListener('hidden.bs.modal', () => {
      this.resetForm();
    });

    // 리드→프로젝트 로더 (기존 리드 불러오기, 방문 예약/견적 제출)
    initLeadProjectLoader({
      inputId: 'modern-lead-search-input',
      resultsId: 'modern-lead-search-results',
      leadNoInputId: 'modern-lead-no',
      linkedInfoId: 'modern-lead-linked-info',
      linkedInfoTextId: 'modern-lead-linked-info-text',
      unlinkBtnId: 'modern-lead-unlink-btn',
      badgeId: 'lead-linked-badge',
      badgeTextId: 'lead-linked-badge-text',
      modalId: 'modernProjectModal',
      emailUnknownCheckId: 'modern-email-unknown-check',
      targets: {
        client: 'modern-client',
        siteManager: 'modern-site-manager',
        phone: 'modern-manager-phone',
        email: 'modern-manager-email',
        address: 'modern-address',
        documentPath: 'modern-document-path',
      },
    });
    logger.info('[ModernProjectModal] Lead-project loader initialized');
  }

  /**
   * 모달 열기
   */
  async open() {
    logger.debug('[ModernProjectModal] open() 호출됨');
    try {
      if (!this.initialized) {
        logger.debug('[ModernProjectModal] 초기화 시작');
        await this.init();
      }

      if (!this.modal) {
        throw new Error('모달 초기화에 실패했습니다.');
      }

      // 시공자 체크박스 영역 동적 렌더링 (DB 캐시 기반 - 시공자 관리 변경 자동 반영)
      this.renderConstructorField();

      // 2026-07-10 이전 성공 상태 리셋 — 재오픈 시 버튼·플래그·idem-key 초기화.
      // 성공 시 finally 에서 이 값들을 유지하기 때문에 다음 오픈에서 반드시 초기화 필요.
      this._submitInProgress = false;
      this._currentIdempotencyKey = null;
      const submitBtn = document.getElementById('modern-submit-new-project');
      if (submitBtn) {
        submitBtn.disabled = false;
        if (submitBtn.innerHTML.includes('등록 완료')) {
          // 템플릿 원본 형태로 복원 (스피너 span + '프로젝트 등록' 텍스트)
          submitBtn.innerHTML =
            '<span class="spinner-border spinner-border-sm me-2 d-none" id="modern-submit-loading"></span>프로젝트 등록';
        }
      }

      // 모달 먼저 즉시 표시 (UX 개선)
      this.modal.show();

      // 데이터는 백그라운드에서 로딩 (await 제거)
      this.loadOptionsInBackground();

    } catch (error) {
      logger.error('[ModernProjectModal] 모달 열기 오류:', error);
      this.showAlert('모달을 여는 중 오류가 발생했습니다.', 'danger');
    }
  }

  /**
   * 시공자 체크박스 영역을 DB 캐시 기반으로 동적 렌더링
   * 시공자 관리 모달에서 변경하면 자동으로 반영됨
   * @param {Array<string>} selectedValues - 이미 선택된 시공자 이름 (편집 모드 등)
   */
  renderConstructorField(selectedValues = []) {
    if (typeof window.renderConstructorCheckboxes !== 'function') {
      logger.warn('[ModernProjectModal] renderConstructorCheckboxes 헬퍼 함수 없음');
      return;
    }
    window.renderConstructorCheckboxes('modern-constructor-categories', {
      selectedValues: selectedValues,
      checkboxClass: 'modern-constructor-checkbox',
      idPrefix: 'modern-constructor',
      showOther: true,
      otherCheckId: 'modern-constructor-other-check',
      otherInputId: 'modern-constructor-other-input',
    });
  }

  /**
   * 캐시 유효성 검사
   */
  isCacheValid() {
    return this.optionsCache &&
           (Date.now() - this.optionsCacheTime) < this.CACHE_TTL;
  }

  /**
   * 옵션 데이터 로딩 (프로젝트 데이터에서 즉시 추출)
   *
   * 변경 사항:
   * - API 호출 제거 → 페이지에 이미 로드된 프로젝트 데이터 재사용
   * - 로딩 상태 제거 → 동기적으로 즉시 완료
   * - 필터와 100% 동일한 데이터 사용
   * - 재시도 횟수 제한으로 무한 루프 방지
   */
  loadOptionsInBackground() {
    try {
      // ProjectListAPI.getCurrentData()가 fallback 체인 처리
      // (stateManager → filters → app.currentData)
      const projectData = window.ITGlobalApp?.api?.getCurrentData();

      if (!projectData || !Array.isArray(projectData) || projectData.length === 0) {
        // 최대 재시도 횟수 체크
        if (this.loadRetryCount < this.MAX_LOAD_RETRIES) {
          this.loadRetryCount++;
          logger.debug(`[ModernProjectModal] 데이터 대기 중 - 재시도 ${this.loadRetryCount}/${this.MAX_LOAD_RETRIES}`);
          setTimeout(() => this.loadOptionsInBackground(), 200);
          return;
        } else {
          logger.error('[ModernProjectModal] 최대 재시도 횟수 초과 - 데이터 로드 실패');
          return;
        }
      }

      // 성공 시 재시도 카운터 리셋
      this.loadRetryCount = 0;

      // 프로젝트 데이터에서 옵션 추출 (필터와 동일한 로직)
      const options = this.extractOptionsFromProjectData(projectData);

      logger.debug('[ModernProjectModal] 추출된 옵션:', {
        owners: options.owners?.length,
        companies: options.companies?.length,
        clients: options.clients?.length
      });

      // UI 업데이트 (캐시 불필요 - 프로젝트 데이터가 이미 캐시됨)
      this.populateDropdowns({ options });

      // 유입 구분은 시트 데이터 유효성 검사에서 다시 로드 (사용자가 시트에서 수정한 값 반영)
      this.refreshInflowOptions();

      logger.debug('[ModernProjectModal] 옵션 로드 완료 (0ms)');

    } catch (error) {
      logger.error('[ModernProjectModal] 옵션 로드 오류:', error);
    }
  }

  /**
   * 유입 구분 값에 따라 사업자명 필드 상태 동기화.
   * - 거래처: 편집 가능 + datalist 자동완성 유지 + '-' 이면 값 비움
   * - 그 외: '-' 자동 입력 + readonly + datalist 자동완성 비활성화
   */
  syncBusinessNameLock() {
    const clientSelect = document.getElementById('modern-client');
    const bizInput = document.getElementById('modern-business-name');
    if (!clientSelect || !bizInput) return;

    const isPartner = clientSelect.value === '거래처';

    if (isPartner) {
      // 거래처 → 편집 가능, 자동완성 복구
      bizInput.readOnly = false;
      bizInput.style.backgroundColor = '';
      bizInput.style.cursor = '';
      bizInput.setAttribute('list', 'modernBusinessNameList');
      bizInput.placeholder = '사업자등록증명';
      if (bizInput.value === '-') bizInput.value = '';
    } else {
      // 거래처 외 → '-' 고정, 편집 불가
      bizInput.value = '-';
      bizInput.readOnly = true;
      bizInput.style.backgroundColor = '#e9ecef';
      bizInput.style.cursor = 'not-allowed';
      bizInput.removeAttribute('list');
      bizInput.placeholder = '거래처 유입 시에만 입력';
    }
  }

  /**
   * 시트 D열의 데이터 유효성 검사 값으로 유입 구분 드롭다운 갱신.
   * 실패 시 기존 프로젝트 D열 유니크 값(=extractOptionsFromProjectData 산출값)이 유지됨.
   */
  async refreshInflowOptions() {
    try {
      console.log('[INFLOW] fetch /api/inflow-options 시작');
      const res = await fetch('/api/inflow-options', { credentials: 'same-origin' });
      console.log(`[INFLOW] response status: ${res.status}`);
      if (!res.ok) {
        const errText = await res.text().catch(() => '');
        console.warn(`[INFLOW] fetch 실패 (${res.status})`, errText);
        return;
      }
      const json = await res.json();
      console.log('[INFLOW] response body:', json);
      const options = Array.isArray(json.options) ? json.options.filter(Boolean) : [];
      if (!options.length) {
        console.warn('[INFLOW] options 배열이 비어있음 - 시트 D열 dataValidation 확인 필요');
        return;
      }

      const clientSelect = document.getElementById('modern-client');
      if (!clientSelect || clientSelect.tagName !== 'SELECT') return;

      const currentValue = clientSelect.value;
      while (clientSelect.children.length > 1) {
        clientSelect.removeChild(clientSelect.lastChild);
      }
      options.forEach((v) => {
        const opt = document.createElement('option');
        opt.value = v;
        opt.textContent = v;
        clientSelect.appendChild(opt);
      });
      // 선택값 복원 (여전히 목록에 있으면)
      if (currentValue && options.includes(currentValue)) {
        clientSelect.value = currentValue;
      }
      logger.debug(`[ModernProjectModal] 유입 구분 옵션 갱신: ${options.length}개`);
    } catch (err) {
      logger.warn('[ModernProjectModal] 유입 구분 옵션 갱신 예외:', err);
    }
  }

  /**
   * 프로젝트 데이터에서 옵션 추출 (필터와 동일한 로직)
   */
  extractOptionsFromProjectData(projectData) {
    // 담당자 목록 추출 (필터와 동일)
    const owners = [...new Set(
      projectData
        .map(item => (item['담당자'] || '').trim())
        .filter(owner => owner)
    )].sort();

    // 사업자 목록 추출 (필터와 동일)
    const companies = [...new Set(
      projectData
        .map(item => (item['사업자'] || '').trim())
        .filter(company => company)
    )].sort();

    // 유입 구분 초기값 = 기존 프로젝트 D열 유니크 값 (서버 fetch 전 fallback)
    // 실제 마스터 목록은 /projects/api/inflow-options 로 시트 validation을 조회해 덮어씀
    const clients = [...new Set(
      projectData
        .map(item => (item['유입 구분'] || '').trim())
        .filter(client => client)
    )].sort((a, b) => a.localeCompare(b, 'ko'));

    // 사업자명 목록 추출 (E열, 신규)
    const businessNames = [...new Set(
      projectData
        .map(item => (item['사업자명'] || '').trim())
        .filter(name => name)
    )].sort((a, b) => a.localeCompare(b, 'ko'));

    return { owners, companies, clients, businessNames };
  }

  /**
   * 드롭다운 로딩 상태 설정
   */
  setDropdownsLoading(isLoading) {
    const ownerSelect = document.getElementById('modern-owner');
    const companySelect = document.getElementById('modern-company');
    const clientInput = document.getElementById('modern-client');  // 이제 select
    const businessNameInput = document.getElementById('modern-business-name');

    if (isLoading) {
      // 로딩 중 - 비활성화 및 로딩 메시지
      if (ownerSelect) {
        ownerSelect.disabled = true;
        ownerSelect.innerHTML = '<option value="">로딩 중...</option>';
      }
      if (companySelect) {
        companySelect.disabled = true;
      }
      if (clientInput) {
        clientInput.disabled = true;
        // select는 placeholder 없음, 첫 옵션 텍스트 변경으로 안내
        if (clientInput.tagName === 'SELECT' && clientInput.children[0]) {
          clientInput.children[0].textContent = '로딩 중...';
        } else {
          clientInput.placeholder = '로딩 중...';
        }
      }
      if (businessNameInput) {
        businessNameInput.disabled = true;
        businessNameInput.placeholder = '로딩 중...';
      }
    } else {
      // 로딩 완료 - 활성화
      if (ownerSelect) {
        ownerSelect.disabled = false;
      }
      if (companySelect) {
        companySelect.disabled = false;
      }
      if (clientInput) {
        clientInput.disabled = false;
        if (clientInput.tagName === 'SELECT' && clientInput.children[0]) {
          clientInput.children[0].textContent = '선택';
        } else {
          clientInput.placeholder = '거래처 입력';
        }
      }
      if (businessNameInput) {
        businessNameInput.disabled = false;
        businessNameInput.placeholder = '사업자등록증명';
      }
    }
  }

  /**
   * 드롭다운 채우기 (담당자, 사업자, 거래처)
   */
  populateDropdowns(data) {
    console.log('[DEBUG] populateDropdowns 시작');

    // 사업자 목록
    const companies = data.options?.companies || data.companies || [];
    console.log('[DEBUG] 사업자 수:', companies.length);

    const companySelect = document.getElementById('modern-company');
    if (companySelect) {
      companySelect.disabled = false;
      companySelect.innerHTML = '<option value="">사업자 선택</option>';
      companies.forEach(company => {
        const option = document.createElement('option');
        option.value = company;
        option.textContent = company;
        companySelect.appendChild(option);
      });

      // innerHTML 변경 후 이벤트 리스너 재바인딩
      companySelect.removeEventListener('change', this._companyChangeHandler);
      this._companyChangeHandler = () => {
        console.log('[DEBUG] 사업자 변경됨:', companySelect.value);
        this.updateCodePreview();
      };
      companySelect.addEventListener('change', this._companyChangeHandler);
      console.log('[DEBUG] 사업자 이벤트 리스너 연결됨');
    } else {
      console.error('[DEBUG] modern-company 엘리먼트를 찾을 수 없음');
    }

    // 담당자 목록 (제외할 인원 필터링)
    const excludedOwners = ['김단이', '심장원', '아이티', '이근혁', '황샛별', '황해승'];
    const owners = data.options?.owners || data.owners || [];
    const filteredOwners = owners.filter(owner => !excludedOwners.includes(owner));
    console.log('[DEBUG] 담당자 수:', filteredOwners.length);

    const ownerSelect = document.getElementById('modern-owner');
    if (ownerSelect) {
      ownerSelect.disabled = false;
      ownerSelect.innerHTML = '<option value="">담당자 선택</option>';
      filteredOwners.forEach(owner => {
        const option = document.createElement('option');
        option.value = owner;
        option.textContent = owner;
        ownerSelect.appendChild(option);
      });

      // 2026-07-08 편의: 로그인한 사용자 본인으로 담당자 기본 선택.
      // 매니저는 대부분 자기 프로젝트를 등록하니 수동 클릭 한 번을 줄임.
      // 수동 수정은 그대로 가능. 시트 담당자 목록에 없거나 excludedOwners에 있으면 미선택.
      const currentUserName = (window.userDisplayName || '').trim();
      if (currentUserName && filteredOwners.includes(currentUserName)) {
        ownerSelect.value = currentUserName;
      }

      // innerHTML 변경 후 이벤트 리스너 재바인딩
      ownerSelect.removeEventListener('change', this._ownerChangeHandler);
      this._ownerChangeHandler = () => {
        console.log('[DEBUG] 담당자 변경됨:', ownerSelect.value);
        this.updateCodePreview();
      };
      ownerSelect.addEventListener('change', this._ownerChangeHandler);
      console.log('[DEBUG] 담당자 이벤트 리스너 연결됨');
    } else {
      console.error('[DEBUG] modern-owner 엘리먼트를 찾을 수 없음');
    }

    // 거래처 (D열, 유입채널) - select 드롭다운
    const clients = data.options?.clients || data.clients || [];
    const clientSelect = document.getElementById('modern-client');
    if (clientSelect && clientSelect.tagName === 'SELECT') {
      // 첫 옵션 "선택"만 유지하고 나머지 제거
      while (clientSelect.children.length > 1) {
        clientSelect.removeChild(clientSelect.lastChild);
      }
      clients.forEach(client => {
        const option = document.createElement('option');
        option.value = client;
        option.textContent = client;
        clientSelect.appendChild(option);
      });
    }

    // 사업자명 (E열, 신규) - datalist 자동완성
    const businessNames = data.options?.businessNames || data.businessNames || [];
    const businessNameList = document.getElementById('modernBusinessNameList');
    if (businessNameList) {
      businessNameList.innerHTML = '';
      businessNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        businessNameList.appendChild(option);
      });
    }
  }

  /**
   * 페이지 로드 시 프리페치 (더 이상 필요 없음 - 프로젝트 데이터 재사용)
   *
   * 변경 사항: API 호출 제거, 호환성 유지를 위해 빈 함수로 유지
   */
  async prefetchOptions() {
    // 프로젝트 데이터에서 직접 추출하므로 프리페치 불필요
    logger.debug('[ModernProjectModal] prefetchOptions 호출됨 (프로젝트 데이터 재사용으로 인해 불필요)');
  }

  /**
   * 프로젝트 코드 미리보기 업데이트
   */
  updateCodePreview() {
    console.log('[DEBUG] updateCodePreview 호출됨');
    const company = document.getElementById('modern-company')?.value;
    const owner = document.getElementById('modern-owner')?.value;
    const codeElement = document.getElementById('modern-generated-code');

    console.log('[DEBUG] 현재 값:', { company, owner, codeElement: !!codeElement });

    if (!codeElement) {
      console.error('[DEBUG] modern-generated-code 엘리먼트를 찾을 수 없음');
      return;
    }

    console.log('[DEBUG] previewCode 호출 시작');
    this.codeGenerator.previewCode(company, owner, (code, error) => {
      console.log('[DEBUG] previewCode 콜백:', { code, error });
      if (code === 'loading') {
        codeElement.innerHTML = '<span class="spinner-border spinner-border-sm text-warning me-2" role="status" aria-hidden="true"></span>코드 생성 중...';
        codeElement.className = 'generated-code loading';
        codeElement.style.fontFamily = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif";
        codeElement.style.letterSpacing = 'normal';
      } else if (error) {
        codeElement.textContent = error;
        codeElement.className = 'generated-code error';
        codeElement.style.fontFamily = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif";
        codeElement.style.letterSpacing = 'normal';
      } else {
        codeElement.textContent = code;
        codeElement.className = 'generated-code success';
        // 시스템 폰트 사용 (가독성 개선)
        codeElement.style.fontFamily = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif";
        codeElement.style.letterSpacing = 'normal';
        codeElement.style.fontWeight = '700';
      }
    });
  }

  /**
   * 금액 입력 처리 (콤마 포맷팅)
   */
  handleAmountInput(event) {
    let value = event.target.value.replace(/[^0-9]/g, ''); // 숫자만 추출

    if (value) {
      // 콤마 추가
      const formatted = parseInt(value).toLocaleString('ko-KR');
      event.target.value = formatted;

      // hidden input에 숫자값 저장
      const hiddenInput = document.getElementById('modern-total-amount');
      if (hiddenInput) {
        hiddenInput.value = value;
      }
    } else {
      const hiddenInput = document.getElementById('modern-total-amount');
      if (hiddenInput) {
        hiddenInput.value = '0';
      }
    }

    // 총액2 실시간 업데이트
    this.updateAmount();
  }

  /**
   * 금액 업데이트
   */
  updateAmount() {
    const hiddenInput = document.getElementById('modern-total-amount');
    const vatCheckbox = document.getElementById('modern-vat-included');
    const displayElement = document.getElementById('modern-display-amount');
    const koreanElement = document.getElementById('modern-amount-korean');

    if (!hiddenInput || !displayElement) return;

    const amount = parseFloat(hiddenInput.value) || 0;
    const includeVAT = vatCheckbox?.checked || false;

    const result = AmountCalculator.formatAmountDisplay(amount, includeVAT);

    displayElement.textContent = result.formatted;
    if (koreanElement) {
      koreanElement.textContent = result.korean;
    }
  }

  /**
   * 빠른 금액 추가
   */
  addQuickAmount(amount) {
    const hiddenInput = document.getElementById('modern-total-amount');
    const displayInput = document.getElementById('modern-total-amount-display');

    if (!hiddenInput || !displayInput) return;

    const current = parseFloat(hiddenInput.value) || 0;
    const newAmount = current + amount;

    hiddenInput.value = newAmount;
    displayInput.value = newAmount.toLocaleString('ko-KR');

    this.updateAmount();
  }

  /**
   * 금액 초기화
   */
  resetAmount() {
    const hiddenInput = document.getElementById('modern-total-amount');
    const displayInput = document.getElementById('modern-total-amount-display');

    if (hiddenInput) hiddenInput.value = '0';
    if (displayInput) displayInput.value = '';

    this.updateAmount();
  }

  /**
   * 체크박스 제한 처리 (범용)
   * @param {Event} event - 체크박스 change 이벤트
   * @param {string} selector - 체크박스 셀렉터
   * @param {string} errorId - 에러 메시지 요소 ID
   * @param {number} maxLimit - 최대 선택 개수 (기본값: 2)
   * @param {string} hiddenInputId - hidden input ID (선택 사항)
   */
  handleCheckboxLimit(event, selector, errorId, maxLimit = 2, hiddenInputId = null) {
    const checkboxes = this.modalElement.querySelectorAll(selector);
    const checkedBoxes = Array.from(checkboxes).filter(cb => cb.checked);
    const errorElement = document.getElementById(errorId);

    // 최대 개수 제한
    if (checkedBoxes.length > maxLimit) {
      event.target.checked = false;
      if (errorElement) {
        errorElement.classList.remove('d-none');
        setTimeout(() => {
          errorElement.classList.add('d-none');
        }, 3000);
      }
    } else {
      if (errorElement) {
        errorElement.classList.add('d-none');
      }
    }

    // hidden input이 있으면 선택된 값을 콤마로 구분하여 저장
    if (hiddenInputId) {
      const hiddenInput = document.getElementById(hiddenInputId);
      if (hiddenInput) {
        const values = checkedBoxes.map(cb => cb.value);
        hiddenInput.value = values.join(', ');
      }
    }
  }

  /**
   * 시공자 hidden input 업데이트 (일반 체크박스 + 기타 입력값 포함)
   */
  updateConstructorHiddenInput() {
    const checkboxes = this.modalElement.querySelectorAll('.modern-constructor-checkbox');
    const checkedBoxes = Array.from(checkboxes).filter(cb => cb.checked);
    const hiddenInput = document.getElementById('modern-constructor');

    // 일반 체크박스 값 수집
    const values = checkedBoxes.map(cb => cb.value);

    // 기타 체크박스가 체크되어 있고 입력값이 있으면 추가
    const otherCheck = document.getElementById('modern-constructor-other-check');
    const otherInput = document.getElementById('modern-constructor-other-input');

    if (otherCheck && otherCheck.checked && otherInput && otherInput.value.trim()) {
      values.push(otherInput.value.trim());
    }

    // hidden input 업데이트
    if (hiddenInput) {
      hiddenInput.value = values.join(', ');
    }
  }

  /**
   * 이메일 검증
   */
  validateEmail() {
    const emailInput = document.getElementById('modern-manager-email');
    const errorElement = document.getElementById('modern-email-error');

    if (!emailInput || !errorElement) return;

    const email = emailInput.value.trim();

    // 빈 값이거나 "확인 불가"인 경우는 검증하지 않음
    if (!email || email === '확인 불가') {
      this.hideEmailError();
      return;
    }

    // 이메일 정규식
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    if (!emailRegex.test(email)) {
      errorElement.classList.remove('d-none');
      return false;
    } else {
      this.hideEmailError();
      return true;
    }
  }

  hideEmailError() {
    const errorElement = document.getElementById('modern-email-error');
    if (errorElement) {
      errorElement.classList.add('d-none');
    }
  }

  /**
   * 주소 검증 (도로명/지번)
   */
  validateAddress() {
    const addressInput = document.getElementById('modern-address');
    const infoElement = document.getElementById('modern-address-info');
    const warningElement = document.getElementById('modern-address-warning');

    if (!addressInput) return;

    const address = addressInput.value.trim();

    if (!address) {
      if (infoElement) infoElement.classList.add('d-none');
      if (warningElement) warningElement.classList.add('d-none');
      return;
    }

    // 지번 주소 패턴 (숫자-숫자 형태)
    const jibunPattern = /\d+-\d+/;

    // 도로명 주소 키워드
    const doroKeywords = ['로', '길', '대로'];

    const hasJibun = jibunPattern.test(address);
    const hasDoro = doroKeywords.some(keyword => address.includes(keyword));

    if (hasJibun && !hasDoro) {
      // 지번 주소로 판단
      if (warningElement) {
        warningElement.classList.remove('d-none');
      }
      if (infoElement) {
        infoElement.classList.add('d-none');
      }
    } else if (hasDoro) {
      // 도로명 주소로 판단
      if (infoElement) {
        infoElement.classList.remove('d-none');
      }
      if (warningElement) {
        warningElement.classList.add('d-none');
      }

      // 자동 정리 (불필요한 공백 제거)
      addressInput.value = address.replace(/\s+/g, ' ');
    } else {
      // 판단 불가
      if (infoElement) infoElement.classList.add('d-none');
      if (warningElement) warningElement.classList.add('d-none');
    }
  }

  /**
   * 폴더 경로 입력 검증
   */
  validateFolderPath() {
    const input = document.getElementById('modern-document-path');
    if (!input) return;

    const value = input.value.trim();

    // 기존 에러 메시지 제거
    let errorDiv = document.getElementById('modern-folder-path-error');
    let successDiv = document.getElementById('modern-folder-path-success');

    if (errorDiv) errorDiv.remove();
    if (successDiv) successDiv.remove();

    // 입력값이 없으면 검증하지 않음
    if (!value) {
      input.style.borderColor = 'var(--gray-300)';
      return;
    }

    // 1. 로컬 경로 감지 (C:\ through I:\)
    const localPathPattern = /^[C-I]:[\\\/]/i;
    if (localPathPattern.test(value)) {
      this.showFolderPathError(
        '로컬 경로는 사용할 수 없습니다. 탐색기에서 폴더 우클릭 → "링크를 클립보드로 복사"를 사용해주세요.',
        input
      );
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
      input.value = folderId;
      this.showFolderPathSuccess('폴더 ID가 자동으로 추출되었습니다.', input);
      return;
    }

    // 3. 이미 폴더 ID 형식인 경우 (20자 이상의 영숫자와 -_)
    const folderIdPattern = /^[a-zA-Z0-9_-]{20,}$/;
    if (folderIdPattern.test(value)) {
      input.style.borderColor = '#28a745';
      this.showFolderPathSuccess('올바른 폴더 ID 형식입니다.', input);
      return;
    }

    // 4. 기타 경우 (잘못된 형식)
    input.style.borderColor = 'var(--gray-300)';
  }

  /**
   * 폴더 경로 에러 메시지 표시
   */
  showFolderPathError(message, inputElement) {
    inputElement.style.borderColor = '#dc3545';

    const errorDiv = document.createElement('div');
    errorDiv.id = 'modern-folder-path-error';
    errorDiv.className = 'mt-2 p-2 rounded';
    errorDiv.style.cssText = 'background-color: #f8d7da; border-left: 4px solid #dc3545; font-size: 0.875rem;';
    errorDiv.innerHTML = `
      <div style="display: flex; align-items: flex-start; gap: 0.5rem;">
        <i class="fas fa-exclamation-triangle" style="color: #721c24; margin-top: 2px;"></i>
        <div style="color: #721c24;">${message}</div>
      </div>
    `;

    inputElement.parentElement.appendChild(errorDiv);
  }

  /**
   * 폴더 경로 성공 메시지 표시
   */
  showFolderPathSuccess(message, inputElement) {
    inputElement.style.borderColor = '#28a745';

    const successDiv = document.createElement('div');
    successDiv.id = 'modern-folder-path-success';
    successDiv.className = 'mt-2 p-2 rounded';
    successDiv.style.cssText = 'background-color: #d4edda; border-left: 4px solid #28a745; font-size: 0.875rem;';
    successDiv.innerHTML = `
      <div style="display: flex; align-items: flex-start; gap: 0.5rem;">
        <i class="fas fa-check-circle" style="color: #155724; margin-top: 2px;"></i>
        <div style="color: #155724;">${message}</div>
      </div>
    `;

    inputElement.parentElement.appendChild(successDiv);
  }

  /**
   * 서버 오류를 사용자 친화적 메시지로 변환
   */
  getUserFriendlyErrorMessage(error) {
    const errorMessage = error.message || '';
    const statusCode = error.statusCode;

    // 구글 시트 연동 오류
    if (errorMessage.includes('Google Sheets') || errorMessage.includes('GOOGLE_SHEET_ID')) {
      return '구글 시트 연동에 일시적인 문제가 있습니다. 잠시 후 다시 시도하거나 관리자에게 문의하세요.';
    }

    // 네트워크 오류
    if (errorMessage.includes('Failed to fetch') || errorMessage.includes('NetworkError') || errorMessage.includes('네트워크')) {
      return '네트워크 연결에 문제가 있습니다. 인터넷 연결을 확인해주세요.';
    }

    // HTTP 상태 코드별 처리
    if (statusCode === 400) {
      // 2026-07-08 필드명 표시 시도: 서버 응답에서 어느 필드가 문제인지 추출
      // 지원 형식:
      //   - "필드 'X'는 필수입니다"
      //   - "X 필드가 필요합니다"
      //   - error.fieldErrors = { field: '메시지' }
      if (error.fieldErrors && typeof error.fieldErrors === 'object') {
        const parts = Object.entries(error.fieldErrors).map(([f, m]) => `• ${f}: ${m}`);
        return `다음 항목을 확인해주세요:\n${parts.join('\n')}`;
      }
      const fieldMatch = errorMessage.match(/(?:필드\s*['"]?([^'"\s]+)['"]?|['"]?([^'"\s]+)['"]?\s*필드)/);
      if (fieldMatch) {
        const field = fieldMatch[1] || fieldMatch[2];
        return `'${field}' 필드를 확인해주세요.\n(서버 메시지: ${errorMessage.slice(0, 200)})`;
      }
      // 서버 메시지가 이해 가능한 한글이면 그대로 노출
      if (errorMessage && errorMessage.length < 200 && /[가-힣]/.test(errorMessage)) {
        return `입력 오류: ${errorMessage}`;
      }
      return '입력하신 정보에 문제가 있습니다. 모든 필드를 다시 확인해주세요.';
    }
    if (statusCode === 401 || statusCode === 403) {
      return '로그인이 필요하거나 권한이 없습니다.';
    }
    if (statusCode === 404) {
      return '요청하신 정보를 찾을 수 없습니다. 페이지를 새로고침해주세요.';
    }
    if (statusCode === 409) {
      return '이미 존재하는 프로젝트입니다. 프로젝트 코드를 확인해주세요.';
    }
    if (statusCode === 500) {
      return '서버에서 일시적인 문제가 발생했습니다. 잠시 후 다시 시도해주세요.';
    }
    if (statusCode === 502 || statusCode === 503 || statusCode === 504) {
      return '서버가 일시적으로 응답하지 않습니다. 잠시 후 다시 시도해주세요.';
    }

    // 기타 오류
    if (errorMessage.includes('timeout') || errorMessage.includes('시간 초과')) {
      return '요청 시간이 초과되었습니다. 네트워크 상태를 확인하고 다시 시도해주세요.';
    }

    // 기본 메시지 (원본 메시지가 사용자에게 이해 가능한 경우)
    if (errorMessage && errorMessage.length < 100 && !errorMessage.includes('Error') && !errorMessage.includes('Exception')) {
      return errorMessage;
    }

    // 최종 폴백
    return '프로젝트 등록 중 문제가 발생했습니다. 잠시 후 다시 시도하거나 관리자에게 문의하세요.';
  }

  /**
   * 제출 전 최종 유효성 검증
   */
  validateBeforeSubmit() {
    const errors = [];

    // 1. 프로젝트 코드 확인
    const codeElement = document.getElementById('modern-generated-code');
    const code = codeElement?.textContent.trim();
    if (!code || code === '사업자와 담당자를 선택하세요' || code.includes('코드 생성 중') || code.includes('선택하세요')) {
      errors.push('프로젝트 코드가 생성되지 않았습니다. 사업자와 담당자를 선택해주세요.');
    }

    // 2. 필수 필드 확인
    const company = document.getElementById('modern-company')?.value;
    const owner = document.getElementById('modern-owner')?.value;
    const client = document.getElementById('modern-client')?.value;
    const address = document.getElementById('modern-address')?.value;

    if (!company) errors.push('사업자를 선택해주세요.');
    if (!owner) errors.push('담당자를 선택해주세요.');
    if (!client || !client.trim()) errors.push('거래처를 입력해주세요.');
    if (!address || !address.trim()) errors.push('현장 주소를 입력해주세요.');

    // 3. 금액 검증
    const totalAmount = document.getElementById('modern-total-amount')?.value;
    if (!totalAmount || parseFloat(totalAmount) <= 0) {
      errors.push('총액을 0원보다 크게 입력해주세요.');
    }

    // 4. 공사 구분 최소 선택 확인
    const workCategories = this.modalElement.querySelectorAll('.modern-work-category-checkbox:checked');
    if (workCategories.length === 0) {
      errors.push('공사 구분을 최소 1개 이상 선택해주세요.');
    }

    // 5. 시공자 "기타" 체크 시 입력값 확인
    const constructorOtherCheck = document.getElementById('modern-constructor-other-check');
    const constructorOtherInput = document.getElementById('modern-constructor-other-input');
    if (constructorOtherCheck?.checked && (!constructorOtherInput?.value || !constructorOtherInput.value.trim())) {
      errors.push('시공자 "기타"를 선택한 경우 이름을 입력해주세요.');
    }

    // 6. 날짜 검증 (이미 있지만 재확인)
    const startDate = document.getElementById('modern-start-date')?.value;
    const endDate = document.getElementById('modern-end-date')?.value;
    if (startDate && endDate) {
      const dateValidation = FormValidator.validateDateRange(startDate, endDate);
      if (!dateValidation.isValid) {
        errors.push(dateValidation.message);
      }
    }

    // 7. 이메일 검증 (이미 있지만 재확인)
    const emailInput = document.getElementById('modern-manager-email');
    if (emailInput?.value && emailInput.value !== '-' && !this.validateEmail()) {
      errors.push('올바른 이메일 형식을 입력해주세요.');
    }

    return {
      isValid: errors.length === 0,
      errors: errors
    };
  }

  /**
   * 폼 데이터 fingerprint 계산 (Idempotency Key 용).
   * 서버 dedup hash 와 동일한 필드 (사업자|담당자|주소|시작일|총액) + 오늘 날짜.
   * 같은 데이터로 같은 날 재클릭 → 무조건 같은 key → 서버 idem 캐시 100% 히트.
   * crypto.subtle 없거나 실패 시 UUID 로 fallback (기존 동작 유지).
   */
  async _computeFormFingerprint() {
    try {
      const fd = new FormData(this.form);
      const biz = (fd.get('사업자') || '').toString().trim();
      const owner = (fd.get('담당자') || '').toString().trim();
      const addr = (fd.get('현장 주소') || '').toString().trim();
      const start = (fd.get('공사 시작') || '').toString().trim().slice(0, 10);
      const amount = (fd.get('총액 1') || '').toString().replace(/[,₩\s]/g, '');
      const today = new Date().toISOString().slice(0, 10);
      const raw = `${biz}|${owner}|${addr}|${start}|${amount}|${today}`;
      if (crypto?.subtle) {
        const buf = new TextEncoder().encode(raw);
        const hashBuf = await crypto.subtle.digest('SHA-256', buf);
        const hex = Array.from(new Uint8Array(hashBuf))
          .map(b => b.toString(16).padStart(2, '0'))
          .join('');
        return `fp-${hex.slice(0, 32)}`;
      }
    } catch (err) {
      logger.debug('[ModernProjectModal] fingerprint 계산 실패, UUID fallback:', err);
    }
    return (typeof crypto !== 'undefined' && crypto.randomUUID)
      ? crypto.randomUUID()
      : `idem-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  }

  /**
   * 프로젝트 제출 (재시도 로직 포함)
   */
  async submitProject(retryCount = 0) {
    const maxRetries = 2; // 최대 2회 재시도
    const submitBtn = document.getElementById('modern-submit-new-project');
    const submitLoading = document.getElementById('modern-submit-loading');
    let willRetry = false; // 재시도 예정 플래그
    let didSucceed = false; // 성공 플래그 — finally 에서 버튼 원복 스킵 판단

    // 2026-07-08 재진입 방지: 최초 제출(retryCount=0)에서 이미 submit 진행 중이면 skip.
    // 폼 더블 클릭·엔터 연타로 인한 신규 프로젝트 중복 등록 방지.
    // 내부 재시도(retryCount > 0)는 예외 — 진행 중인 submit 자체가 재호출한 것.
    if (retryCount === 0) {
      if (this._submitInProgress) {
        logger.debug('[ModernProjectModal] submitProject 재진입 감지 - 스킵');
        return;
      }
      this._submitInProgress = true;
      // 2026-07-10 Idempotency Key — 폼 데이터 fingerprint 기반.
      //   기존: crypto.randomUUID() → 재클릭마다 새 key → 서버 idem 캐시 miss → 중복 등록 발생.
      //     실제 관측 (G3852/G3853, R3854/R3855): 서버 응답 후 프론트 상태 리셋되면서
      //     매니저 재클릭 시 새 UUID 발급 → 두 요청 페이로드는 100% 같은데 서버 idem miss.
      //   개선: 서버 dedup hash 와 동일 필드로 프론트 fingerprint 생성.
      //     같은 데이터면 무조건 같은 key → 서버 idem 캐시가 100% 히트 (Redis 정상시).
      //   날짜 접미로 다른 날 같은 데이터 재등록은 정상 허용.
      this._currentIdempotencyKey = await this._computeFormFingerprint();
    }

    // 0. 최초 제출 시에만 최종 유효성 검증 수행
    if (retryCount === 0) {
      const validation = this.validateBeforeSubmit();
      if (!validation.isValid) {
        // 검증 실패: 모든 오류를 하나의 메시지로 표시
        const errorMessage = validation.errors.join('\n• ');
        this.showAlert(`다음 항목을 확인해주세요:\n\n• ${errorMessage}`, 'danger');
        this._submitInProgress = false;  // 검증 실패도 클리어
        return;
      }
    }

    // 1. 버튼을 "추가 중..." 상태로 변경
    const originalButtonText = submitBtn ? submitBtn.innerHTML : '';
    if (submitBtn) {
      submitBtn.disabled = true;
      if (retryCount > 0) {
        submitBtn.innerHTML = `<i class="fas fa-circle-notch fa-spin me-1"></i>재시도 중... (${retryCount}/${maxRetries})`;
      } else {
        submitBtn.innerHTML = '<i class="fas fa-circle-notch fa-spin me-1"></i>추가 중...';
      }
    }
    if (submitLoading) submitLoading.classList.remove('d-none');

    try {
      // 2. 폼 데이터 수집
      const formData = new FormData(this.form);
      const projectData = {};

      for (let [key, value] of formData.entries()) {
        projectData[key] = value;
      }

      // 프로젝트 코드 추가
      const codeElement = document.getElementById('modern-generated-code');
      const code = codeElement?.textContent.trim();
      if (code && code !== '사업자와 담당자를 선택하세요') {
        projectData['프로젝트 코드'] = code;
      }

      // VAT 포함 여부 (백엔드 호환: boolean → "TRUE"/"FALSE" 문자열)
      // 필드명: '부가세' (google_sheets.py column_mapping의 R열과 일치)
      const vatCheckbox = document.getElementById('modern-vat-included');
      projectData['부가세'] = vatCheckbox?.checked ? 'TRUE' : 'FALSE';

      // 공사 구분 (hidden input에서 가져오기)
      const workCategoryHidden = document.getElementById('modern-work-category');
      if (workCategoryHidden && workCategoryHidden.value) {
        projectData['공사 구분'] = workCategoryHidden.value;
      }

      // 기계 분류 처리
      const selectedMachineCategories = [];
      this.modalElement.querySelectorAll('.modern-machine-category-checkbox:checked').forEach(cb => {
        selectedMachineCategories.push(cb.value);
      });
      if (selectedMachineCategories.length > 0) {
        projectData['기계 분류'] = selectedMachineCategories.join(', ');
      }

      // 브랜드 처리
      const selectedBrands = [];
      this.modalElement.querySelectorAll('.modern-brand-category-checkbox:checked').forEach(cb => {
        selectedBrands.push(cb.value);
      });
      if (selectedBrands.length > 0) {
        projectData['브랜드'] = selectedBrands.join(', ');
      }

      // 도급 구분 처리
      const selectedContracts = [];
      this.modalElement.querySelectorAll('.modern-contract-category-checkbox:checked').forEach(cb => {
        selectedContracts.push(cb.value);
      });
      if (selectedContracts.length > 0) {
        projectData['도급 구분'] = selectedContracts.join(', ');
      }

      // 3. 날짜 검증
      const startDate = projectData['공사 시작'];
      const endDate = projectData['공사 종료'];
      if (startDate && endDate) {
        const dateValidation = FormValidator.validateDateRange(startDate, endDate);
        if (!dateValidation.isValid) {
          // 검증 실패: 버튼 복구
          if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = originalButtonText;
          }
          if (submitLoading) submitLoading.classList.add('d-none');

          this.showAlert(dateValidation.message, 'danger');
          return;
        }
      }

      // 4. 이메일 검증
      const emailInput = document.getElementById('modern-manager-email');
      if (emailInput && emailInput.value && emailInput.value !== '-') {
        if (!this.validateEmail()) {
          // 검증 실패: 버튼 복구
          if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = originalButtonText;
          }
          if (submitLoading) submitLoading.classList.add('d-none');

          this.showAlert('올바른 이메일 형식을 입력해주세요.', 'danger');
          emailInput.focus();
          return;
        }
      }

      // 5. 총액 2는 구글 시트 수식으로 처리되므로 프론트엔드에서는 전송하지 않음
      // 백엔드에서 다음 수식이 자동 삽입됨: =IF(R:R=TRUE, Q:Q+ROUND(Q:Q*0.1,0), Q:Q)

      // 6. 진행 상황 표시 (모달 내부)
      if (retryCount === 0) {
        this.showProgressBanner('서버에 저장 중입니다...');
      } else {
        this.showProgressBanner(`재시도 중입니다... (${retryCount}/${maxRetries})`);
      }

      // 7. 서버 전송 — Idempotency-Key 헤더 + 120초 타임아웃 (Google Sheets 재시도 버짓 대비)
      const abortCtrl = ('AbortSignal' in window && typeof AbortSignal.timeout === 'function')
        ? { signal: AbortSignal.timeout(120000) }
        : {};
      const response = await fetch('/api/projects/auto', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Idempotency-Key': this._currentIdempotencyKey || '',
        },
        credentials: 'same-origin',
        body: JSON.stringify(projectData),
        ...abortCtrl,
      });

      // 응답 body 안전 파싱 (재시작 중 body 잘림 방어)
      const _readJson = async (resp) => {
        try {
          const text = await resp.text();
          return text ? JSON.parse(text) : {};
        } catch (_) {
          return {};
        }
      };

      if (!response.ok) {
        const errorData = await _readJson(response);
        const statusCode = response.status;

        if (statusCode === 401) {
          // 인증 오류는 특별 처리 (로그인 페이지로 리다이렉트)
          throw new Error('AUTH_REQUIRED');
        }

        // 재시도 가능한 에러인지 표시하기 위해 상태 코드 포함
        const error = new Error(errorData.error || errorData.message || `HTTP ${statusCode}`);
        error.statusCode = statusCode; // 재시도 로직에서 사용
        if (errorData.error_id) error.error_id = errorData.error_id;
        throw error;
      }

      const result = await _readJson(response);

      // 백엔드 응답 포맷 호환: success 또는 ok 필드 체크
      const isSuccess = result.success === true || result.ok === true;

      if (isSuccess) {
        didSucceed = true;
        // 성공 시 버튼을 "✓ 등록 완료" 상태로 즉시 표시하고 disabled 유지.
        // finally 블록에서 원복 스킵 → 재클릭 자체 물리적으로 불가능.
        if (submitBtn) {
          submitBtn.disabled = true;
          submitBtn.innerHTML = '<i class="fas fa-check me-1"></i>등록 완료';
        }
        // 8. 진행 배너 메시지 변경 (데이터 동기화 중)
        // 2026-07-10 UX 개선 — 모달 즉시 닫기 + 데이터 동기화는 백그라운드로.
        //   기존: 동기화 완료 대기 후 닫음 → refreshData 8초+ 걸리면 매니저 대기 중
        //   재클릭 → 새 idempotency key → 서버 dedup 놓치면 중복 등록 (G3852/G3853 관측).
        //   개선: 서버 응답 받자마자 모달 닫음. 동기화는 fire-and-forget.
        const projectCode = result.project_code;
        const projectData = result.project_data;

        this.hideProgressBanner();
        // hide.bs.modal 이벤트 preventDefault 방지를 위해 먼저 flag 내림.
        // 재클릭 방지는 submitBtn.disabled + '등록 완료' 표시로 이미 커버됨.
        // (2026-07-13 사용자 관측: 등록 성공 후 모달이 자동으로 안 닫히고 X 버튼 필요)
        this._submitInProgress = false;
        this.modal.hide();

        // 백그라운드 동기화 — await 하지 않음
        (async () => {
          let syncSuccess = false;
          try {
            if (projectData && window.DataSyncManager && window.DataSyncManager.addProjectRealTime) {
              window.DataSyncManager.addProjectRealTime(projectData);
              syncSuccess = true;
            } else if (window.projectListApp && window.projectListApp.refreshData) {
              await window.projectListApp.refreshData(true, false, false);
              syncSuccess = true;
            }
          } catch (syncError) {
            logger.error('[ModernProjectModal] 백그라운드 동기화 오류:', syncError);
          }
          if (!syncSuccess) {
            // 동기화 실패 시 사용자에게 새로고침 유도
            if (window.showPageAlert) {
              window.showPageAlert(
                '프로젝트가 저장됐지만 화면 갱신에 실패했어요. 잠시 후 새로고침됩니다.',
                'warning'
              );
            }
            setTimeout(() => window.location.reload(), 3000);
          }
        })();

        // 13. 성공 메시지 — 프로젝트 코드 포함 (사용자 확인용)
        const successMsg = result.deduped
          ? `동일 프로젝트가 이미 등록되어 있어 기존 코드로 연결했습니다: ${projectCode}`
          : `새 프로젝트가 등록되었습니다: ${projectCode}`;
        if (window.showPageAlert) {
          window.showPageAlert(successMsg, 'success');
        } else {
          this.showSuccessMessage(successMsg);
        }

        // 13-b. 수금관리 모드에서 완납(미수금=0) 프로젝트는 필터에 안 잡혀 안 보임 → 사용자 안내
        try {
          const receivable = parseFloat(projectData?.['미수금'] || 0);
          const isReceivablesMode =
            document.getElementById('outstandingFilter')?.value === 'outstanding';
          if (isReceivablesMode && receivable <= 0) {
            // 새 프로젝트 결과 안내 → 페이지 헤더
            const msg = '완납 상태로 등록되어 수금관리 모드에서는 숨겨집니다. 상단 토글을 끄시면 확인할 수 있어요.';
            if (window.showPageAlert) {
              window.showPageAlert(msg, 'info');
            } else {
              const toast = await this.getToastComponent();
              toast.show(msg, 'info', 6000);
            }
          }
        } catch (hiddenNoticeError) {
          logger.debug('[ModernProjectModal] 완납 안내 표시 스킵:', hiddenNoticeError);
        }

        // localStorage 이벤트 (cross-tab sync)
        const eventData = {
          type: 'project_created',
          project_code: result.project_code,
          timestamp: new Date().toISOString()
        };
        localStorage.setItem('newProjectCreated', JSON.stringify(eventData));
        localStorage.removeItem('newProjectCreated');

      } else {
        throw new Error(result.message || result.error || '등록에 실패했습니다.');
      }

    } catch (error) {
      logger.error('프로젝트 등록 오류:', error);

      // 인증 오류 특별 처리 (리다이렉트)
      if (error.message === 'AUTH_REQUIRED') {
        this.hideProgressBanner();
        this.showAlert('로그인이 필요합니다. 로그인 페이지로 이동합니다.', 'warning');
        setTimeout(() => window.location.href = '/login', 2000);
        return;
      }

      // 14. 재시도 로직 - 네트워크 오류 및 일시적 서버 오류에 대해서만 재시도
      // 단, Google Sheets 연동 비활성화 같은 설정 문제는 재시도 제외
      const isConfigurationError =
        error.message.includes('Google Sheets 연동') ||
        error.message.includes('GOOGLE_SHEET_ID');

      const isRetryableError =
        !isConfigurationError && (
          // 네트워크 오류 (fetch 실패)
          error.message.includes('Failed to fetch') ||
          error.message.includes('NetworkError') ||
          error.message.includes('네트워크') ||
          // 일시적 서버 오류 (502, 504)
          // 503은 메시지 내용으로 판단 (일시적 오류 vs 설정 오류)
          error.statusCode === 502 ||
          error.statusCode === 504 ||
          (error.statusCode === 503 && !isConfigurationError)
        );

      if (isRetryableError && retryCount < maxRetries) {
        logger.debug(`[ModernProjectModal] 재시도 ${retryCount + 1}/${maxRetries}`);

        // 재시도 플래그 설정 (finally에서 버튼 복구 건너뛰기 위함)
        willRetry = true;

        // 재시도 카운트다운 메시지 표시 (모달 내부 배너)
        const retryDelay = retryCount + 1;
        this.showProgressBanner(`네트워크 오류. ${retryDelay}초 후 자동으로 재시도합니다... (${retryCount + 1}/${maxRetries})`);

        // 지수 백오프: 1초, 2초
        setTimeout(() => {
          this.submitProject(retryCount + 1);
        }, 1000 * retryDelay);

        return; // finally 블록으로 이동 (재시도 중이므로 버튼은 복구하지 않음)
      }

      // 15. 최종 실패 - 진행 배너 숨김 및 사용자 친화적 에러 메시지 표시
      // 모달은 열린 채로 유지하여 사용자가 수정/재시도 가능하도록 함
      this.hideProgressBanner();
      const userFriendlyMessage = this.getUserFriendlyErrorMessage(error);
      // 서버가 반환한 error_id 있으면 사용자 안내에 포함 (관리자 문의 시 로그 추적용)
      const errorIdSuffix = error.error_id
        ? `\n\n오류 ID: ${error.error_id} (관리자에게 전달)`
        : '';
      this.showAlert(userFriendlyMessage + errorIdSuffix, 'danger');
    } finally {
      // 14. 버튼과 스피너 복구
      // - 성공 (didSucceed=true): 버튼 "✓ 등록 완료" 상태 유지, 모달 닫히면 재열림 시 초기화됨.
      //     이걸 유지해야 재클릭이 물리적으로 불가능 → 중복 등록 원천 차단 (2026-07-10).
      // - 재시도 예정 (willRetry=true): 버튼 "재시도 중..." 상태 유지, 재귀 호출이 처리.
      // - 검증 실패 / 서버 실패: 버튼 원복, 사용자가 수정·재시도 가능.
      if (!willRetry && !didSucceed) {
        if (submitBtn) {
          submitBtn.disabled = false;
          if (submitBtn.innerHTML.includes('추가 중') || submitBtn.innerHTML.includes('재시도 중')) {
            submitBtn.innerHTML = originalButtonText;
          }
        }
        if (submitLoading) {
          submitLoading.classList.add('d-none');
        }
        // 재진입 방지 플래그 해제 — 실패 시에만 (성공/재시도는 계속 유지)
        this._submitInProgress = false;
        // Idempotency key 정리 — 다음 최초 제출 시 새 key 발급되도록
        this._currentIdempotencyKey = null;
      } else if (didSucceed) {
        // 성공 후에도 스피너는 숨김 (버튼은 "✓ 등록 완료" 유지)
        if (submitLoading) {
          submitLoading.classList.add('d-none');
        }
        // _submitInProgress / _currentIdempotencyKey 는 유지 → 모달 재오픈까지 재클릭 100% 차단
      }
    }
  }

  /**
   * 폼 초기화
   */
  resetForm() {
    if (this.form) {
      this.form.reset();
    }

    // 유입 구분 ↔ 사업자명 잠금 재동기화 (form.reset 은 change 이벤트를 fire 하지 않음)
    this.syncBusinessNameLock();

    // 리드 연결 UI 초기화
    const leadNo = document.getElementById('modern-lead-no');
    if (leadNo) leadNo.value = '';
    const leadLinked = document.getElementById('modern-lead-linked-info');
    if (leadLinked) leadLinked.classList.add('d-none');
    const leadBadge = document.getElementById('lead-linked-badge');
    if (leadBadge) leadBadge.classList.add('d-none');
    const leadSearchInput = document.getElementById('modern-lead-search-input');
    if (leadSearchInput) leadSearchInput.value = '';
    const leadResults = document.getElementById('modern-lead-search-results');
    if (leadResults) {
      leadResults.classList.add('d-none');
      leadResults.innerHTML = '';
    }

    // 코드 미리보기 초기화
    const codeElement = document.getElementById('modern-generated-code');
    if (codeElement) {
      codeElement.textContent = '사업자와 담당자를 선택하세요';
      codeElement.className = 'generated-code';
      codeElement.style.fontFamily = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif";
      codeElement.style.letterSpacing = 'normal';
    }

    // 금액 입력 필드 초기화
    const totalAmountDisplay = document.getElementById('modern-total-amount-display');
    const totalAmountHidden = document.getElementById('modern-total-amount');
    if (totalAmountDisplay) totalAmountDisplay.value = '';
    if (totalAmountHidden) totalAmountHidden.value = '0';

    // 금액 표시 초기화
    const displayElement = document.getElementById('modern-display-amount');
    const koreanElement = document.getElementById('modern-amount-korean');
    if (displayElement) displayElement.textContent = '0원';
    if (koreanElement) koreanElement.textContent = '(금 영원정)';

    // 체크박스 초기화
    this.modalElement.querySelectorAll('.modern-work-category-checkbox').forEach(cb => cb.checked = false);
    this.modalElement.querySelectorAll('.modern-machine-category-checkbox').forEach(cb => cb.checked = false);
    this.modalElement.querySelectorAll('.modern-brand-category-checkbox').forEach(cb => cb.checked = false);
    this.modalElement.querySelectorAll('.modern-contract-category-checkbox').forEach(cb => cb.checked = false);
    this.modalElement.querySelectorAll('.modern-constructor-checkbox').forEach(cb => cb.checked = false);

    // 시공자 기타 체크박스 및 입력창 초기화
    const constructorOtherCheck = document.getElementById('modern-constructor-other-check');
    const constructorOtherInput = document.getElementById('modern-constructor-other-input');
    if (constructorOtherCheck) constructorOtherCheck.checked = false;
    if (constructorOtherInput) {
      constructorOtherInput.value = '';
      constructorOtherInput.classList.add('d-none');
    }

    // 날짜 동기화 체크박스 초기화
    const sameDateCheck = document.getElementById('modern-same-date-check');
    const endDateInput = document.getElementById('modern-end-date');
    if (sameDateCheck) sameDateCheck.checked = false;
    if (endDateInput) endDateInput.disabled = false;

    // 이메일 확인 불가 체크박스 초기화
    const emailUnknownCheck = document.getElementById('modern-email-unknown-check');
    const emailInput = document.getElementById('modern-manager-email');
    if (emailUnknownCheck) emailUnknownCheck.checked = false;
    if (emailInput) {
      emailInput.readOnly = false;
      emailInput.style.backgroundColor = '';
      emailInput.style.cursor = '';
    }

    // 폴더 경로 입력 초기화
    const documentPathInput = document.getElementById('modern-document-path');
    if (documentPathInput) {
      documentPathInput.style.borderColor = 'var(--gray-300)';
    }
    const folderPathError = document.getElementById('modern-folder-path-error');
    const folderPathSuccess = document.getElementById('modern-folder-path-success');
    if (folderPathError) folderPathError.remove();
    if (folderPathSuccess) folderPathSuccess.remove();

    // 에러 메시지 숨기기
    this.hideEmailError();
    const addressInfo = document.getElementById('modern-address-info');
    const addressWarning = document.getElementById('modern-address-warning');
    if (addressInfo) addressInfo.classList.add('d-none');
    if (addressWarning) addressWarning.classList.add('d-none');
  }

  /**
   * 알림 표시
   */
  showAlert(message, type = 'info') {
    const container = document.getElementById('modernModalAlertContainer');
    if (!container) {
      window.alert(message);
      return;
    }

    const alertElement = document.createElement('div');
    alertElement.className = `alert alert-${type} alert-dismissible fade show`;
    alertElement.innerHTML = `
      ${message}
      <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
    `;

    container.innerHTML = '';
    container.appendChild(alertElement);

    setTimeout(() => alertElement.remove(), 5000);
  }

  /**
   * 진행 상황 배너 표시/숨김
   */
  showProgressBanner(message) {
    const banner = document.getElementById('modernModalProgressBanner');
    const messageElement = document.getElementById('modernModalProgressMessage');

    if (banner && messageElement) {
      messageElement.textContent = message;
      banner.classList.remove('d-none');
    }
  }

  hideProgressBanner() {
    const banner = document.getElementById('modernModalProgressBanner');
    if (banner) {
      banner.classList.add('d-none');
    }
  }

  /**
   * Toast 컴포넌트 가져오기
   */
  async getToastComponent() {
    if (!this.toastComponent) {
      const { default: Toast } = await import('./Toast.js');
      this.toastComponent = new Toast();
    }
    return this.toastComponent;
  }

  /**
   * 성공 메시지
   */
  async showSuccessMessage(message) {
    const toast = await this.getToastComponent();
    toast.show(message, 'success');
  }

  /**
   * 에러 메시지
   */
  async showErrorMessage(message) {
    const toast = await this.getToastComponent();
    toast.show(message, 'error');
  }

  /**
   * 모달 닫기
   */
  close() {
    if (this.modal) {
      this.modal.hide();
    }
  }


  /**
   * 정리
   */
  destroy() {
    this.codeGenerator.cleanup();

    if (this.modal) {
      this.modal.dispose();
      this.modal = null;
    }
  }
}
