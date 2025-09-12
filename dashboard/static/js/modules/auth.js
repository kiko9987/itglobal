// 권한 관리 모듈
class AuthManager {
    constructor() {
        // 사용자 권한 정보 (서버에서 전달받음)
        this.userPermission = '{{ session.user.permission_level if session.user else "Admin" }}';
        this.userName = '{{ session.user.name if session.user else "관리자" }}';
        this.userRole = '{{ user_role }}';
        window.currentUserEmail = '{{ user_email }}';
        
        // 편집자가 수정 가능한 필드 정의
        this.editorAllowedFields = {
            'basic': ['현장 담당자', '도급 구분', '담당자 연락처', '시공자', '담당자 이메일'],
            'construction': 'all',
            'financial': [],
            'payment': [],
            'profit': ['제품대', '도급비', '자재비', '기타비'],
            'documents': ['견적서 및 계약서 폴더 경로']
        };
    }
    
    // 개발자용 임시 권한 변경 함수 (테스트용)
    changeTestPermission(role) {
        if (confirm(`테스트를 위해 임시로 권한을 ${role}로 변경하시겠습니다?`)) {
            window.testUserRole = role;
            
            // 권한 뱃지 즉시 업데이트
            const roleBadge = document.getElementById('user-role-badge');
            if (roleBadge) {
                roleBadge.textContent = `${role} (테스트)`;
                roleBadge.style.backgroundColor = '#f59e0b';
                roleBadge.style.color = 'white';
                roleBadge.style.animation = 'pulse 2s infinite';
            }
            
            // 관리자 메뉴 표시/숨김
            const adminMenuItem = document.getElementById('admin-menu-item');
            if (adminMenuItem) {
                adminMenuItem.style.display = role === 'Admin' ? 'block' : 'none';
            }
            
            // 테스트 모드 활성화됨
            
            // 페이지의 모든 편집 버튼 다시 생성
            setTimeout(() => {
                if (dataTable) {
                    dataTable.draw(false);
                }
            }, 100);
        }
    }
    
    // 테스트 모드 초기화
    resetTestMode() {
        if (window.testUserRole) {
            delete window.testUserRole;
            location.reload();
        }
    }
    
    // 현재 사용자 역할 반환
    getUserRole() {
        return window.testUserRole || this.userRole;
    }
    
    // 현재 사용자가 특정 카드를 수정할 수 있는지 확인
    canEditCard(cardType) {
        const currentRole = this.getUserRole();
        
        if (currentRole === 'Admin') {
            return true;
        } else if (currentRole === 'Editor') {
            const canEdit = this.editorAllowedFields.hasOwnProperty(cardType);
            return canEdit;
        }
        return false;
    }
    
    // 편집자가 특정 필드를 수정할 수 있는지 확인
    canEditField(cardType, fieldName) {
        const currentRole = this.getUserRole();
        if (currentRole === 'Admin') {
            return true;
        } else if (currentRole === 'Editor') {
            const allowedFields = this.editorAllowedFields[cardType];
            let canEdit = false;
            
            if (allowedFields === 'all') {
                canEdit = true;
            } else if (Array.isArray(allowedFields)) {
                canEdit = allowedFields.includes(fieldName);
            }
            
            return canEdit;
        }
        return false;
    }
    
    // 권한에 따른 수정 버튼 HTML 생성
    generateEditButtons(projectCode, cardType) {
        if (!this.canEditCard(cardType)) {
            return '';
        }
        
        const lockId = `lock-${cardType}-${projectCode}`;
        
        return `
            <div class="card-controls">
                <button class="btn btn-sm btn-outline-primary edit-btn" 
                        onclick="handleCardEdit('${cardType}', '${projectCode}')" 
                        data-project="${projectCode}" 
                        data-card="${cardType}"
                        title="카드 편집">
                    <i class="fas fa-edit"></i> 편집
                </button>
                <div id="${lockId}" class="lock-indicator" style="display: none;">
                    <i class="fas fa-lock text-warning"></i>
                    <span class="lock-user"></span>
                </div>
            </div>
        `;
    }
}

// 전역 AuthManager 인스턴스 생성
window.authManager = new AuthManager();

// 전역 함수들을 위한 래퍼
function getUserRole() {
    return window.authManager.getUserRole();
}

function canEditCard(cardType) {
    return window.authManager.canEditCard(cardType);
}

function canEditField(cardType, fieldName) {
    return window.authManager.canEditField(cardType, fieldName);
}

function generateEditButtons(projectCode, cardType) {
    return window.authManager.generateEditButtons(projectCode, cardType);
}

function changeTestPermission(role) {
    window.authManager.changeTestPermission(role);
}

function resetTestMode() {
    window.authManager.resetTestMode();
}