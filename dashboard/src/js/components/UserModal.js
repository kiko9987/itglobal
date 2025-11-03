import logger from '../utils/logger.js';

/**
 * User Management Modal 컴포넌트 (완전한 레거시 포팅)
 * 사용자 관리 모달 - 권한 변경, 상태 토글 등
 */
export default class UserModal {
  constructor() {
    this.modal = null;
    this.modalElement = null;
    this.toastComponent = null;
    this.eventDelegationSetup = false;
    this.refreshBtnBound = false;
    this.modalHideBound = false;
    this.currentPage = 1;
    this.totalPages = 1;
    this.perPage = 15; // 사용자 목록은 15개씩 (작업 로그와 동일)
    this.init();
  }

  /**
   * 컴포넌트 초기화
   */
  init() {
    this.modalElement = document.getElementById('userManagementModal');
    if (!this.modalElement) {
      logger.debug('[USER_MODAL] userManagementModal 요소가 없어서 생성합니다.');
      this.createModalHTML();
      this.modalElement = document.getElementById('userManagementModal');

      // 모달 생성 후 약간의 지연을 두고 이벤트 설정 (DOM 반영 대기)
      setTimeout(() => {
        this.setupEventListeners();
        this.setupUserManagementEventDelegation();
      }, 0);
    } else {
      // 기존 모달이 있는 경우 즉시 이벤트 설정
      this.setupEventListeners();
      this.setupUserManagementEventDelegation();
    }

    // table-striped 클래스 제거 (Bootstrap 자동 줄무늬 방지)
    const table = this.modalElement.querySelector('#usersTable');
    if (table) {
      table.classList.remove('table-striped');
    }

    // Bootstrap 모달 인스턴스 생성
    this.modal = new bootstrap.Modal(this.modalElement, {
      backdrop: true,
      keyboard: true
    });
  }

  /**
   * 이벤트 리스너 설정
   */
  setupEventListeners() {
    // 새로고침 버튼 (createModalHTML에서 이미 연결했으므로 건너뜀)
    // 기존 모달인 경우에만 연결
    if (!this.refreshBtnBound) {
      const refreshBtn = document.getElementById('refreshUsersBtn');
      if (refreshBtn) {
        refreshBtn.addEventListener('click', () => this.loadUsers(this.currentPage, true)); // showFeedback=true
        this.refreshBtnBound = true;
      }
    }

    // 모달 숨김 이벤트
    if (!this.modalHideBound) {
      this.modalElement.addEventListener('hidden.bs.modal', () => {
        this.clearAlert();
      });
      this.modalHideBound = true;
    }

    // 페이지네이션 클릭 이벤트 위임
    const paginationContainer = document.getElementById('usersPagination');
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
   * 사용자 관리 이벤트 위임 설정 (한 번만 실행)
   */
  setupUserManagementEventDelegation() {
    if (this.eventDelegationSetup) return;

    const userTableBody = document.getElementById('usersTableBody');
    if (!userTableBody) return;

    // 드롭다운 변경 이벤트 위임
    userTableBody.addEventListener('change', (e) => {
      // 권한 드롭다운
      if (e.target.matches('.user-permission-select')) {
        const selectElement = e.target;
        const userEmail = selectElement.dataset.userEmail;
        const newPermission = selectElement.value;
        const oldPermission = selectElement.dataset.originalValue || selectElement.value;

        // 값이 변경되지 않았으면 무시
        if (newPermission === oldPermission) {
          return;
        }

        // 사용자 이름 찾기
        const userRow = selectElement.closest('tr');
        const userName = userRow ? userRow.cells[0].textContent.trim() : '사용자';

        // 권한 텍스트 매핑
        const permissionText = {
          'admin': 'Admin',
          'editor': 'Editor',
          'viewer': 'Viewer'
        }[newPermission] || newPermission;

        // 확인 다이얼로그
        if (confirm(`"${userName}"님의 권한을 "${permissionText}"(으)로 변경하시겠습니까?`)) {
          if (!selectElement.disabled) {
            // 원본 값 저장 (취소 시 되돌리기 위해)
            selectElement.dataset.originalValue = newPermission;
            this.updateUserPermission(userEmail, newPermission);
          }
        } else {
          // 취소: 이전 값으로 되돌림
          selectElement.value = oldPermission;
        }
      }

      // 상태 드롭다운
      if (e.target.matches('.user-status-select')) {
        const selectElement = e.target;
        const userEmail = selectElement.dataset.userEmail;
        const newStatus = selectElement.value;
        const oldStatus = selectElement.dataset.originalValue || selectElement.value;

        // 값이 변경되지 않았으면 무시
        if (newStatus === oldStatus) {
          return;
        }

        // 사용자 이름 찾기
        const userRow = selectElement.closest('tr');
        const userName = userRow ? userRow.cells[0].textContent.trim() : '사용자';

        // 상태 텍스트 매핑
        const statusText = {
          'active': '활성',
          'inactive': '비활성',
          'resigned': '퇴사'
        }[newStatus] || newStatus;

        // 확인 다이얼로그
        if (confirm(`"${userName}"님의 상태를 "${statusText}"(으)로 변경하시겠습니까?`)) {
          if (!selectElement.disabled) {
            // 원본 값 저장 (취소 시 되돌리기 위해)
            selectElement.dataset.originalValue = newStatus;
            this.updateUserStatus(userEmail, newStatus);
          }
        } else {
          // 취소: 이전 값으로 되돌림
          selectElement.value = oldStatus;
        }
      }
    });

    // 프로젝트 코드 접미사 수정 버튼 클릭
    userTableBody.addEventListener('click', (e) => {
      if (e.target.closest('.edit-project-code-btn')) {
        const btn = e.target.closest('.edit-project-code-btn');
        const userEmail = btn.dataset.userEmail;
        this.enterEditMode(userEmail);
      }
    });

    // 프로젝트 코드 접미사 저장 버튼 클릭
    userTableBody.addEventListener('click', (e) => {
      if (e.target.closest('.save-project-code-btn')) {
        const btn = e.target.closest('.save-project-code-btn');
        const userEmail = btn.dataset.userEmail;
        this.saveProjectCodeSuffix(userEmail);
      }
    });

    // 프로젝트 코드 접미사 취소 버튼 클릭
    userTableBody.addEventListener('click', (e) => {
      if (e.target.closest('.cancel-project-code-btn')) {
        const btn = e.target.closest('.cancel-project-code-btn');
        const userEmail = btn.dataset.userEmail;
        this.cancelEditMode(userEmail);
      }
    });

    // 프로젝트 코드 접미사 입력 필드 Enter 키 처리
    userTableBody.addEventListener('keydown', (e) => {
      if (e.target.matches('.project-code-suffix-input') && e.key === 'Enter') {
        e.preventDefault();
        const userEmail = e.target.dataset.userEmail;
        this.saveProjectCodeSuffix(userEmail);
      }
    });

    this.eventDelegationSetup = true;
  }

  /**
   * 사용자 관리 모달 열기
   */
  async open() {
    try {
      // 페이지네이션 초기화
      this.currentPage = 1;
      this.totalPages = 1;

      this.modal.show();
      await this.loadUsers(1); // 첫 페이지부터 시작
    } catch (error) {
      logger.error('사용자 모달 열기 실패:', error);
      this.showAlert('모달을 여는 중 오류가 발생했습니다.', 'error');
    }
  }

  /**
   * 사용자 목록 로드
   * @param {number} page - 페이지 번호
   * @param {boolean} showFeedback - 성공/실패 피드백 표시 여부 (새로고침 버튼 클릭 시만 true)
   */
  async loadUsers(page = 1, showFeedback = false) {
    this.currentPage = page;
    logger.debug('[USER_MODAL] loadUsers() 시작 - page:', page, 'showFeedback:', showFeedback);

    const tbody = document.getElementById('usersTableBody');
    const countElement = document.getElementById('usersCount');
    const pageInfoElement = document.getElementById('usersPageInfo');
    const refreshBtn = document.getElementById('refreshUsersBtn');

    if (!tbody) {
      logger.error('[USER_MODAL] usersTableBody를 찾을 수 없습니다');
      return;
    }

    // 새로고침 버튼 로딩 효과 (showFeedback이 true일 때만)
    if (refreshBtn && showFeedback) {
      refreshBtn.disabled = true;
      const icon = refreshBtn.querySelector('i');
      if (icon) {
        icon.className = 'fas fa-sync-alt fa-spin me-1';
      }
    }

    // 로딩 상태 표시
    tbody.innerHTML = `
      <tr>
        <td colspan="7" class="text-center">
          <i class="fas fa-spinner fa-spin"></i> 로딩 중...
        </td>
      </tr>
    `;

    try {
      logger.debug('[USER_MODAL] API 호출: /api/users?page=' + this.currentPage + '&per_page=' + this.perPage);
      const response = await fetch(`/api/users?page=${this.currentPage}&per_page=${this.perPage}`, {
        credentials: 'include'
      });
      logger.debug('[USER_MODAL] API 응답 상태:', response.status);

      if (!response.ok) {
        throw new Error('Failed to load users - status: ' + response.status);
      }

      const data = await response.json();
      logger.debug('[USER_MODAL] API 응답 데이터:', data);

      if (data.success) {
        logger.debug('[USER_MODAL] 사용자 수:', data.users.length);
        this.renderUsersTable(data.users);
        logger.debug('[USER_MODAL] 테이블 렌더링 완료');

        // 페이지네이션 정보 업데이트
        const pagination = data.pagination;
        this.totalPages = pagination.total_pages;

        if (countElement) {
          countElement.textContent = `사용자 수: ${pagination.total_count}명`;
        }
        if (pageInfoElement) {
          pageInfoElement.textContent = `페이지: ${this.currentPage}/${this.totalPages}`;
        }

        // 페이지네이션 UI 업데이트
        this.updatePaginationUI(pagination);
      } else {
        throw new Error(data.message || 'Failed to load users');
      }
    } catch (error) {
      logger.error('[USER_MODAL] 사용자 목록 로드 실패:', error);
      tbody.innerHTML = `
        <tr>
          <td colspan="7" class="text-center text-danger">
            <i class="fas fa-exclamation-triangle me-2"></i>사용자 목록을 불러올 수 없습니다.
          </td>
        </tr>
      `;
      this.resetPaginationUI();

      // 실패 시 에러 아이콘 표시 (showFeedback이 true일 때만)
      if (refreshBtn && showFeedback) {
        const icon = refreshBtn.querySelector('i');
        if (icon) {
          icon.className = 'fas fa-times-circle me-1 text-danger';
        }
        // 1초 후 원래대로 복원
        setTimeout(() => {
          refreshBtn.disabled = false;
          if (icon) {
            icon.className = 'fas fa-sync-alt me-1';
          }
        }, 1000);
      }
      return;
    }

    // 성공 시 체크 아이콘 표시 (showFeedback이 true일 때만)
    if (refreshBtn && showFeedback) {
      const icon = refreshBtn.querySelector('i');
      if (icon) {
        icon.className = 'fas fa-check-circle me-1 text-success';
      }
      // 800ms 후 원래대로 복원
      setTimeout(() => {
        refreshBtn.disabled = false;
        if (icon) {
          icon.className = 'fas fa-sync-alt me-1';
        }
      }, 800);
    }
  }

  /**
   * 사용자 테이블 렌더링
   */
  renderUsersTable(users) {
    const tbody = document.getElementById('usersTableBody');

    if (!users || users.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="7" class="text-center text-muted">
            등록된 사용자가 없습니다.
          </td>
        </tr>
      `;
      return;
    }

    // 현재 세션 사용자 이메일 가져오기 (템플릿에서 전역 변수로 설정)
    const currentUserEmail = window.currentUserEmail || '';

    const htmlContent = users.map(user => {
      const isCurrentUser = user.email === currentUserEmail;
      const userEmailEscaped = user.email.replace(/['"]/g, '\\$&');

      let html = '<tr>';

      // 사용자 이름
      html += `<td class="text-center">${user.name}</td>`;

      // 이메일
      html += `<td class="text-center">${user.email}</td>`;

      // 권한 드롭다운
      html += '<td class="text-center">';
      const isAdmin = user.permission_level === 'admin' || user.permission_level === 'super_admin';
      const currentPermission = isAdmin ? 'admin' : user.permission_level;
      html += `
        <select class="form-select form-select-sm user-permission-select"
                data-user-email="${userEmailEscaped}"
                data-original-value="${currentPermission}">
          <option value="admin" ${isAdmin ? 'selected' : ''}>Admin</option>
          <option value="editor" ${user.permission_level === 'editor' ? 'selected' : ''}>Editor</option>
          <option value="viewer" ${user.permission_level === 'viewer' ? 'selected' : ''}>Viewer</option>
        </select>
      `;
      html += '</td>';

      // 프로젝트 코드 접미사 (관리자만 수정 가능)
      html += '<td class="text-center project-code-cell" data-user-email="${userEmailEscaped}">';
      html += `
        <div class="d-flex align-items-center justify-content-center gap-1">
          <!-- 읽기 모드 (기본) - 비활성화된 입력창 -->
          <input type="text"
                 class="form-control form-control-sm project-code-display-input"
                 data-user-email="${userEmailEscaped}"
                 placeholder="-"
                 maxlength="3"
                 style="max-width: 50px; font-size: 0.8rem; padding: 0.2rem 0.4rem; text-transform: uppercase; background-color: #e9ecef; color: #6c757d; border-color: #ced4da;"
                 disabled
                 readonly>
          <button class="btn btn-sm btn-outline-primary edit-project-code-btn"
                  data-user-email="${userEmailEscaped}"
                  title="접미사 수정"
                  style="display: none; padding: 0.15rem 0.4rem;">
            <i class="fas fa-edit fa-xs"></i>
          </button>

          <!-- 편집 모드 (숨김) -->
          <input type="text"
                 class="form-control form-control-sm project-code-suffix-input"
                 data-user-email="${userEmailEscaped}"
                 placeholder="접미사"
                 maxlength="3"
                 style="max-width: 50px; font-size: 0.8rem; padding: 0.2rem 0.4rem; text-transform: uppercase; display: none;"
                 disabled>
          <button class="btn btn-sm btn-success save-project-code-btn"
                  data-user-email="${userEmailEscaped}"
                  title="저장"
                  style="display: none; padding: 0.15rem 0.4rem;">
            <i class="fas fa-check fa-xs"></i>
          </button>
          <button class="btn btn-sm btn-secondary cancel-project-code-btn"
                  data-user-email="${userEmailEscaped}"
                  title="취소"
                  style="display: none; padding: 0.15rem 0.4rem;">
            <i class="fas fa-times fa-xs"></i>
          </button>
        </div>
      `;
      html += '</td>';

      // 상태 뱃지
      html += '<td class="text-center">';
      html += this.getStatusBadge(user);
      html += '</td>';

      // 마지막 로그인
      html += '<td class="text-center">';
      if (user.last_login) {
        const lastLogin = new Date(user.last_login);
        html += `${lastLogin.toLocaleString('ko-KR')}`;
      } else {
        html += '<span class="text-muted">로그인 기록 없음</span>';
      }
      html += '</td>';

      // 상태 변경 드롭다운
      html += '<td class="text-center">';
      const currentStatus = user.is_resigned ? 'resigned' : (user.is_active ? 'active' : 'inactive');
      html += `
        <select class="form-select form-select-sm user-status-select"
                data-user-email="${userEmailEscaped}"
                data-original-value="${currentStatus}">
          <option value="active" ${user.is_active && !user.is_resigned ? 'selected' : ''}>활성</option>
          <option value="inactive" ${!user.is_active && !user.is_resigned ? 'selected' : ''}>비활성</option>
          <option value="resigned" ${user.is_resigned ? 'selected' : ''}>퇴사</option>
        </select>
      `;
      html += '</td>';

      html += '</tr>';
      return html;
    }).join('');

    tbody.innerHTML = htmlContent;

    // 프로젝트 코드 접미사 로드
    this.loadProjectCodeSuffixes(users);
  }

  /**
   * 프로젝트 코드 접미사 로드
   */
  async loadProjectCodeSuffixes(users) {
    // 관리자 여부 확인
    const isAdmin = window.userRole === 'admin' || window.userRole === 'super_admin';

    for (const user of users) {
      try {
        const response = await fetch(`/api/users/${encodeURIComponent(user.email)}/project-code-suffix`);

        const emailEscaped = user.email.replace(/['"]/g, '\\$&');
        const displayInput = document.querySelector(`.project-code-display-input[data-user-email="${emailEscaped}"]`);
        const editInput = document.querySelector(`.project-code-suffix-input[data-user-email="${emailEscaped}"]`);
        const editBtn = document.querySelector(`.edit-project-code-btn[data-user-email="${emailEscaped}"]`);

        if (!displayInput || !editInput) continue;

        // HTTP 상태 코드 체크
        if (!response.ok) {
          logger.error(`[USER_MODAL] 프로젝트 코드 접미사 API 오류 (${user.email}): HTTP ${response.status}`);
          displayInput.value = '오류';
          displayInput.style.color = '#dc3545'; // 빨간색
          editInput.dataset.originalValue = '';
          continue; // 다음 사용자 처리
        }

        const data = await response.json();

        if (data.success && data.suffix) {
          displayInput.value = data.suffix;
          displayInput.style.color = '#6c757d'; // 회색
          editInput.value = data.suffix;
          // 초기 originalValue 설정 (취소 시 복원용)
          editInput.dataset.originalValue = data.suffix;
        } else {
          displayInput.value = '';
          displayInput.style.color = '#6c757d'; // 회색
          editInput.value = '';
          // 초기 originalValue 설정 (빈 값)
          editInput.dataset.originalValue = '';
        }

        // 관리자만 수정 버튼 표시
        if (isAdmin && editBtn) {
          editBtn.style.display = 'inline-block';
        }

      } catch (error) {
        logger.error(`[USER_MODAL] 프로젝트 코드 접미사 로드 오류 (${user.email}):`, error);
        const emailEscaped = user.email.replace(/['"]/g, '\\$&');
        const displayInput = document.querySelector(`.project-code-display-input[data-user-email="${emailEscaped}"]`);
        if (displayInput) {
          displayInput.value = '오류';
          displayInput.style.color = '#dc3545'; // 빨간색
        }
        // 에러가 발생해도 다음 사용자 처리 계속
      }
    }
  }

  /**
   * 사용자 권한 업데이트
   */
  async updateUserPermission(email, permission) {
    try {
      // 사용자 이름 찾기
      const userRow = document.querySelector(`select[data-user-email="${email}"]`)?.closest('tr');
      const userName = userRow ? userRow.cells[0].textContent.trim() : '사용자';

      const response = await fetch('/api/users/permission', {
        method: 'POST',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          email: email,
          permission: permission
        })
      });

      if (response.ok) {
        const result = await response.json();
        const permissionText = permission === 'admin' ? 'Admin' :
                               permission === 'editor' ? 'Editor' : 'Viewer';
        this.showAlert(`"${userName}"님의 권한이 "${permissionText}"로 변경되었습니다.`, 'success');
      } else {
        const error = await response.json();
        logger.error('권한 업데이트 실패 (응답):', error);
        throw new Error(error.message || 'Failed to update permission');
      }
    } catch (error) {
      logger.error('권한 업데이트 실패 (예외):', error);
      this.showAlert('권한 변경에 실패했습니다.', 'error');
      this.loadUsers(); // 실패 시 원래 상태로 복원
    }
  }

  /**
   * 사용자 상태 변경 (활성/비활성/퇴사)
   */
  async updateUserStatus(email, status) {
    try {
      // 사용자 이름 찾기
      const userRow = document.querySelector(`select.user-status-select[data-user-email="${email}"]`)?.closest('tr');
      const userName = userRow ? userRow.cells[0].textContent.trim() : '사용자';

      const response = await fetch('/api/users/status', {
        method: 'POST',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          email: email,
          status: status
        })
      });

      if (response.ok) {
        const result = await response.json();
        this.showAlert(result.message || `"${userName}"님의 상태가 변경되었습니다.`, 'success');

        // 상태 변경 후 사용자 목록 새로고침 (모달은 유지)
        setTimeout(() => {
          this.loadUsers();
        }, 500);

        // 메인 테이블의 담당자 뱃지를 실시간 업데이트 (새로고침 없음)
        const isResigned = status === 'resigned';
        this.updateMainTableManagerBadges(userName, email, isResigned);
      } else {
        const error = await response.json();
        logger.error('상태 업데이트 실패 (응답):', error);
        throw new Error(error.message || 'Failed to update status');
      }
    } catch (error) {
      logger.error('상태 업데이트 실패 (예외):', error);
      this.showAlert('사용자 상태 변경에 실패했습니다.', 'error');
      this.loadUsers(); // 실패 시 원래 상태로 복원
    }
  }

  /**
   * 프로젝트 코드 접미사 업데이트
   */
  async updateProjectCodeSuffix(email, suffix, inputElement, oldSuffix) {
    try {
      const response = await fetch(`/api/users/${encodeURIComponent(email)}/project-code-suffix`, {
        method: 'PUT',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          suffix: suffix
        })
      });

      const result = await response.json();

      if (response.ok && result.success) {
        this.showToast(result.message || '프로젝트 코드 접미사가 업데이트되었습니다.', 'success');
        // 성공 시에만 originalValue 업데이트
        inputElement.dataset.originalValue = suffix;
        // placeholder 업데이트
        if (suffix) {
          inputElement.placeholder = '';
        } else {
          inputElement.placeholder = '-';
        }
      } else {
        logger.error('프로젝트 코드 접미사 업데이트 실패 (응답):', result);
        this.showToast(result.message || '프로젝트 코드 접미사 업데이트에 실패했습니다.', 'error');
        // 실패 시 이전 값으로 되돌림
        inputElement.value = oldSuffix;
      }
    } catch (error) {
      logger.error('프로젝트 코드 접미사 업데이트 실패 (예외):', error);
      this.showToast('프로젝트 코드 접미사 업데이트 중 오류가 발생했습니다.', 'error');
      // 실패 시 이전 값으로 되돌림
      inputElement.value = oldSuffix;
    }
  }

  /**
   * 편집 모드 진입
   */
  enterEditMode(userEmail) {
    const emailEscaped = userEmail.replace(/['"]/g, '\\$&');
    const displayInput = document.querySelector(`.project-code-display-input[data-user-email="${emailEscaped}"]`);
    const editBtn = document.querySelector(`.edit-project-code-btn[data-user-email="${emailEscaped}"]`);
    const editInput = document.querySelector(`.project-code-suffix-input[data-user-email="${emailEscaped}"]`);
    const saveBtn = document.querySelector(`.save-project-code-btn[data-user-email="${emailEscaped}"]`);
    const cancelBtn = document.querySelector(`.cancel-project-code-btn[data-user-email="${emailEscaped}"]`);

    if (!displayInput || !editBtn || !editInput || !saveBtn || !cancelBtn) return;

    // 읽기 모드 숨김
    displayInput.style.display = 'none';
    editBtn.style.display = 'none';

    // 편집 모드 표시
    editInput.style.display = 'inline-block';
    editInput.disabled = false;
    saveBtn.style.display = 'inline-block';
    cancelBtn.style.display = 'inline-block';

    // 입력 필드에 포커스
    editInput.focus();
    editInput.select();
  }

  /**
   * 편집 모드 취소
   */
  cancelEditMode(userEmail) {
    const emailEscaped = userEmail.replace(/['"]/g, '\\$&');
    const displayInput = document.querySelector(`.project-code-display-input[data-user-email="${emailEscaped}"]`);
    const editBtn = document.querySelector(`.edit-project-code-btn[data-user-email="${emailEscaped}"]`);
    const editInput = document.querySelector(`.project-code-suffix-input[data-user-email="${emailEscaped}"]`);
    const saveBtn = document.querySelector(`.save-project-code-btn[data-user-email="${emailEscaped}"]`);
    const cancelBtn = document.querySelector(`.cancel-project-code-btn[data-user-email="${emailEscaped}"]`);

    if (!displayInput || !editBtn || !editInput || !saveBtn || !cancelBtn) return;

    // 원래 값으로 복원
    const originalValue = editInput.dataset.originalValue || '';
    editInput.value = originalValue;

    // 편집 모드 숨김
    editInput.style.display = 'none';
    saveBtn.style.display = 'none';
    cancelBtn.style.display = 'none';

    // 읽기 모드 표시
    displayInput.style.display = 'inline-block';
    editBtn.style.display = 'inline-block';
  }

  /**
   * 프로젝트 코드 접미사 저장
   */
  async saveProjectCodeSuffix(userEmail) {
    const emailEscaped = userEmail.replace(/['"]/g, '\\$&');
    const editInput = document.querySelector(`.project-code-suffix-input[data-user-email="${emailEscaped}"]`);
    const displayInput = document.querySelector(`.project-code-display-input[data-user-email="${emailEscaped}"]`);

    if (!editInput || !displayInput) return;

    const newSuffix = editInput.value.trim().toUpperCase();
    const oldSuffix = editInput.dataset.originalValue || '';

    // 값이 변경되지 않았으면 취소 처리
    if (newSuffix === oldSuffix) {
      this.cancelEditMode(userEmail);
      return;
    }

    // 클라이언트 측 유효성 검사
    if (newSuffix && (newSuffix.length < 1 || newSuffix.length > 3)) {
      this.showToast('접미사는 1~3자 사이여야 합니다.', 'error');
      editInput.value = oldSuffix;
      return;
    }

    if (newSuffix && !/^[A-Z]+$/.test(newSuffix)) {
      this.showToast('접미사는 영문자만 입력 가능합니다.', 'error');
      editInput.value = oldSuffix;
      return;
    }

    // 확인 다이얼로그
    const confirmMessage = newSuffix
      ? `"${newSuffix}"로 프로젝트 코드 접미사를 저장하시겠습니까?`
      : '프로젝트 코드 접미사를 삭제하시겠습니까?';

    if (!confirm(confirmMessage)) {
      return;
    }

    try {
      const response = await fetch(`/api/users/${encodeURIComponent(userEmail)}/project-code-suffix`, {
        method: 'PUT',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          suffix: newSuffix
        })
      });

      const result = await response.json();

      if (response.ok && result.success) {
        this.showToast(result.message || '프로젝트 코드 접미사가 업데이트되었습니다.', 'success');

        // 성공 시 originalValue와 displayInput 업데이트
        editInput.dataset.originalValue = newSuffix;
        displayInput.value = newSuffix;
        displayInput.style.color = '#6c757d'; // 회색

        // 읽기 모드로 전환
        this.cancelEditMode(userEmail);
      } else {
        logger.error('프로젝트 코드 접미사 업데이트 실패 (응답):', result);
        this.showToast(result.message || '프로젝트 코드 접미사 업데이트에 실패했습니다.', 'error');
      }
    } catch (error) {
      logger.error('프로젝트 코드 접미사 업데이트 실패 (예외):', error);
      this.showToast('프로젝트 코드 접미사 업데이트 중 오류가 발생했습니다.', 'error');
    }
  }

  /**
   * 사용자 상태 토글 (활성화/비활성화) - 하위 호환성
   */
  async toggleUserStatus(email, targetStatus) {
    try {
      // 사용자 이름 찾기
      const userRow = document.querySelector(`button[data-user-email="${email}"]`)?.closest('tr');
      const userName = userRow ? userRow.cells[0].textContent.trim() : '사용자';

      const response = await fetch('/api/users/status', {
        method: 'POST',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          email: email,
          is_active: targetStatus
        })
      });

      if (response.ok) {
        const result = await response.json();
        const statusText = targetStatus ? '활성화' : '비활성화';
        this.showAlert(`"${userName}"님이 ${statusText}되었습니다.`, 'success');

        // 상태 변경 후 사용자 목록 새로고침
        setTimeout(() => {
          this.loadUsers();
        }, 500);
      } else {
        const error = await response.json();
        logger.error('상태 업데이트 실패 (응답):', error);
        throw new Error(error.message || 'Failed to update status');
      }
    } catch (error) {
      logger.error('상태 업데이트 실패 (예외):', error);
      this.showAlert('사용자 상태 변경에 실패했습니다.', 'error');
      this.loadUsers(); // 실패 시 원래 상태로 복원
    }
  }

  /**
   * 알림 메시지 표시 (사용자 관리 모달 전용)
   */
  showAlert(message, type = 'info') {
    const alertContainer = document.getElementById('userManagementAlert');
    if (!alertContainer) return;

    // 기존 알림 제거
    this.clearAlert();

    // 타입별 스타일 설정
    const alertTypes = {
      success: { class: 'alert-success', icon: 'fa-check-circle' },
      error: { class: 'alert-danger', icon: 'fa-exclamation-triangle' },
      warning: { class: 'alert-warning', icon: 'fa-exclamation-circle' },
      info: { class: 'alert-info', icon: 'fa-info-circle' }
    };

    const alertType = alertTypes[type] || alertTypes.info;

    const alertElement = document.createElement('div');
    alertElement.className = `alert ${alertType.class} alert-dismissible fade show py-2 px-3 mb-0`;
    alertElement.style.fontSize = '0.85rem';
    alertElement.innerHTML = `
      <i class="fas ${alertType.icon} me-2"></i>${message}
      <button type="button" class="btn-close btn-close-sm" data-bs-dismiss="alert"></button>
    `;

    alertContainer.appendChild(alertElement);

    // 3초 후 자동 제거 (성공 메시지만)
    if (type === 'success') {
      setTimeout(() => {
        this.clearAlert();
      }, 3000);
    }
  }

  /**
   * 알림 메시지 제거
   */
  clearAlert() {
    const alertContainer = document.getElementById('userManagementAlert');
    if (alertContainer) {
      alertContainer.innerHTML = '';
    }
  }

  /**
   * Toast 메시지 표시 (showAlert의 별칭)
   */
  showToast(message, type = 'info') {
    this.showAlert(message, type);
  }

  /**
   * 권한 뱃지 생성 (작업 로그와 동일한 스타일)
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
   * 상태 뱃지 생성 (활성/비활성/퇴사)
   */
  getStatusBadge(user) {
    if (user.is_resigned) {
      return '<span class="badge bg-secondary" style="font-size: 0.7rem;">퇴사</span>';
    } else if (user.is_active) {
      return '<span class="badge bg-success" style="font-size: 0.7rem;">활성</span>';
    } else {
      return '<span class="badge bg-warning text-dark" style="font-size: 0.7rem;">비활성</span>';
    }
  }

  /**
   * 페이지네이션 UI 업데이트
   */
  updatePaginationUI(pagination) {
    const paginationContainer = document.getElementById('usersPagination');
    const prevButton = document.getElementById('usersPrevPage');
    const nextButton = document.getElementById('usersNextPage');

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
    const paginationContainer = document.getElementById('usersPagination');
    const prevButton = document.getElementById('usersPrevPage');
    const nextButton = document.getElementById('usersNextPage');
    const pageInfoElement = document.getElementById('usersPageInfo');
    const countElement = document.getElementById('usersCount');

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
    if (countElement) {
      countElement.textContent = '사용자 수: 0명';
    }
  }

  /**
   * 페이지 변경
   */
  changePage(page) {
    if (page < 1 || page > this.totalPages || page === this.currentPage) {
      return;
    }
    this.loadUsers(page);
  }

  /**
   * 기존 testDropdown 함수 (호환성 유지)
   */
  testDropdown(selectElement, email) {
    const newPermission = selectElement.value;

    if (selectElement.disabled) {
      return;
    }

    // 권한 업데이트 함수 호출
    this.updateUserPermission(email, newPermission);
  }

  /**
   * 모달 HTML 생성
   */
  createModalHTML() {
    const modalHTML = `
      <!-- User Management Modal -->
      <div class="modal fade" id="userManagementModal" tabindex="-1" aria-labelledby="userManagementModalLabel" data-bs-backdrop="true" data-bs-keyboard="true">
        <div class="modal-dialog user-management-modal">
          <div class="modal-content">
            <div class="modal-header">
              <div class="d-flex align-items-center flex-grow-1">
                <h5 class="modal-title mb-0" id="userManagementModalLabel">
                  <i class="fas fa-users-cog me-2"></i>사용자 관리
                </h5>
                <div id="userManagementAlert" class="ms-3 user-management-alert"></div>
              </div>
              <div class="d-flex align-items-center">
                <button class="btn btn-sm me-2 user-refresh-btn" id="refreshUsersBtn">
                  <i class="fas fa-sync-alt me-1"></i>새로고침
                </button>
              </div>
            </div>
            <div class="modal-body">
              <div class="table-responsive">
                <table class="table table-striped" id="usersTable">
                  <thead class="table-light">
                    <tr>
                      <th>이름</th>
                      <th>이메일</th>
                      <th>권한</th>
                      <th style="width: 14%;">프로젝트 코드</th>
                      <th style="width: 8%;">상태</th>
                      <th>마지막 로그인</th>
                      <th>상태 변경</th>
                    </tr>
                  </thead>
                  <tbody id="usersTableBody">
                    <tr>
                      <td colspan="7" class="text-center">
                        <i class="fas fa-spinner fa-spin"></i> 로딩 중...
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
            <div class="modal-footer">
              <!-- 한 줄에 좌측(사용자 정보) - 중앙(페이지네이션) - 우측(닫기 버튼) -->
              <div class="d-flex align-items-center w-100 user-management-footer">
                <!-- 좌측: 사용자 정보 -->
                <div class="d-flex align-items-center user-management-info">
                  <small class="text-muted me-3" id="usersCount">사용자 수: 0명</small>
                  <small class="text-muted" id="usersPageInfo">페이지 정보</small>
                </div>

                <!-- 중앙: 페이지네이션 (절대 중앙 배치) -->
                <div class="user-management-pagination-center">
                  <nav aria-label="사용자 페이지네이션">
                    <ul class="pagination pagination-sm mb-0" id="usersPagination">
                      <li class="page-item disabled" id="usersPrevPage">
                        <a class="page-link" href="#" tabindex="-1" aria-disabled="true">
                          <i class="fas fa-chevron-left"></i>
                        </a>
                      </li>
                      <!-- 페이지 번호들이 여기에 동적으로 추가됩니다 -->
                      <li class="page-item disabled" id="usersNextPage">
                        <a class="page-link" href="#" tabindex="-1" aria-disabled="true">
                          <i class="fas fa-chevron-right"></i>
                        </a>
                      </li>
                    </ul>
                  </nav>
                </div>

                <!-- 우측: 닫기 버튼 -->
                <div class="user-management-close">
                  <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">닫기</button>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;

    // DOM에 모달 HTML 추가
    document.body.insertAdjacentHTML('beforeend', modalHTML);

    // 모달 생성 직후 새로고침 버튼 이벤트 바로 연결
    const refreshBtn = document.getElementById('refreshUsersBtn');
    if (refreshBtn) {
      logger.debug('[USER_MODAL] 새로고침 버튼 이벤트 연결');
      refreshBtn.addEventListener('click', () => {
        logger.debug('[USER_MODAL] 새로고침 버튼 클릭됨');
        this.loadUsers(this.currentPage, true); // showFeedback=true
      });
      this.refreshBtnBound = true; // 이벤트 연결 플래그 설정
    } else {
      logger.warn('[USER_MODAL] 새로고침 버튼을 찾을 수 없습니다');
    }
  }

  /**
   * 메인 테이블의 담당자 뱃지를 실시간 업데이트 (새로고침 없음)
   * @param {string} userName - 변경된 사용자 이름
   * @param {string} userEmail - 변경된 사용자 이메일 (고유 식별자)
   * @param {boolean} isResigned - 퇴사 여부
   */
  async updateMainTableManagerBadges(userName, userEmail, isResigned) {
    try {
      logger.debug(`[UserModal] 담당자 뱃지 업데이트 시작: ${userName} (${userEmail}), 퇴사=${isResigned}`);

      // 1. UnifiedBadgeSystem 가져오기
      const badgeSystem = window.projectListApp?.components?.table?.badgeSystem;
      if (!badgeSystem || !badgeSystem.managerBadge) {
        logger.warn('[UserModal] UnifiedBadgeSystem을 찾을 수 없습니다.');
        return;
      }

      // 2. ManagerBadge의 resignedTeam 배열 업데이트 (이메일 기반)
      const managerBadge = badgeSystem.managerBadge;
      if (isResigned) {
        // 퇴사자 목록에 추가 (중복 방지 - 이메일로 확인)
        const existingIndex = managerBadge.resignedTeam.findIndex(user => user.email === userEmail);
        if (existingIndex === -1) {
          managerBadge.resignedTeam.push({ name: userName, email: userEmail });
          logger.debug(`[UserModal] ${userName} (${userEmail})을 퇴사자 목록에 추가`);
        }
      } else {
        // 퇴사자 목록에서 제거 (이메일로 찾기)
        const index = managerBadge.resignedTeam.findIndex(user => user.email === userEmail);
        if (index > -1) {
          managerBadge.resignedTeam.splice(index, 1);
          logger.debug(`[UserModal] ${userName} (${userEmail})을 퇴사자 목록에서 제거`);
        }
      }

      // 3. DataTable에서 해당 담당자가 있는 모든 행 찾기
      const table = window.projectListApp?.components?.table?.table;
      if (!table) {
        logger.warn('[UserModal] DataTable을 찾을 수 없습니다.');
        return;
      }

      let updatedCount = 0;

      // 모든 행을 순회하면서 담당자 컬럼(인덱스 1) 확인
      // ✅ IMPROVED: 이메일 우선 매칭 → 이름 폴백 방식으로 동명이인 문제 완화
      //    - 담당자 이메일이 있으면: 이메일로 정확히 매칭 (동명이인 문제 해결)
      //    - 담당자 이메일이 없으면: 이름으로 매칭 (기존 방식 유지)
      table.rows().every(function() {
        const rowData = this.data();
        const managerName = rowData['담당자'];
        const managerEmail = rowData['담당자 이메일'];

        // 우선순위 매칭: 이메일 → 이름
        let isMatch = false;
        if (managerEmail && managerEmail.trim() && managerEmail !== '-') {
          // 이메일 데이터가 있으면 이메일로 정확히 매칭
          isMatch = (managerEmail === userEmail);
        } else {
          // 이메일 데이터가 없으면 이름으로 매칭 (fallback)
          isMatch = (managerName === userName);
        }

        // 해당 담당자가 변경된 사용자인 경우
        if (isMatch) {
          // 새로운 뱃지 HTML 생성
          const newBadgeHtml = badgeSystem.createManagerBadge(userName);

          // 행의 DOM 노드 가져오기
          const rowNode = this.node();
          if (rowNode) {
            // 담당자 셀 찾기 (인덱스 1)
            const managerCell = rowNode.cells[1];
            if (managerCell) {
              // 뱃지 HTML만 교체 (깜빡임 없음)
              managerCell.innerHTML = newBadgeHtml;
              updatedCount++;
            }
          }
        }
      });

      logger.debug(`[UserModal] ${updatedCount}개 행의 담당자 뱃지 업데이트 완료`);

      // 4. ModernProjectFilters의 resignedManagers 배열도 실시간 업데이트
      const filters = window.modernFilters;
      if (filters) {
        // 퇴사자 목록 재로드
        await filters.fetchResignedManagers();

        // 필터 드롭다운 재생성 (현재 데이터 기준)
        if (filters.currentData) {
          filters.populateManagerFilter(filters.currentData);
        }

        logger.debug(`[UserModal] 필터 드롭다운 업데이트 완료`);
      } else {
        logger.warn('[UserModal] ModernProjectFilters를 찾을 수 없습니다.');
      }

    } catch (error) {
      logger.error('[UserModal] 담당자 뱃지 업데이트 실패:', error);
    }
  }

  /**
   * 컴포넌트 정리
   */
  destroy() {
    if (this.modal) {
      this.modal.dispose();
    }
    this.eventDelegationSetup = false;
  }
}