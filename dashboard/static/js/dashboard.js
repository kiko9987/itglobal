// 메인 대시보드 스크립트 - 모듈화된 버전
// 전역 변수 최소화
let projectsData = [];
let dataTable = null;

// 감사 로그 페이지네이션 변수
let currentPage = 1;
let totalPages = 1;
let perPage = 20;

// 사용자 권한 정보 (서버에서 전달받음)
const userPermission = '{{ session.user.permission_level if session.user else "Admin" }}';
const userName = '{{ session.user.name if session.user else "관리자" }}';
const userRole = '{{ user_role }}';
window.currentUserEmail = '{{ user_email }}';

// 페이지 로드 시 초기화
document.addEventListener('DOMContentLoaded', function() {
    initializePage();
    setupGlobalEventListeners();
});

// 전역 이벤트 리스너 설정
function setupGlobalEventListeners() {
    // 검색 입력 이벤트
    $('#searchInput').on('input', function() {
        applyFilters();
    });
    
    // 필터 변경 이벤트
    $('#statusFilter, #managerFilter').on('change', function() {
        applyFilters();
    });
    
    // 날짜 필터 이벤트
    $('#startDateFilter, #endDateFilter').on('change', function() {
        applyFilters();
    });
    
    // 필터 초기화 버튼
    $('#clearFilters').on('click', function() {
        clearFilters();
    });
    
    // 새 프로젝트 버튼
    $('#newProjectBtn').on('click', function() {
        openNewProjectModal();
    });
}

// 뷰 프로젝트 함수 (모듈에서 호출됨)
function viewProject(projectCode) {
    try {
        // 프로젝트 상세 보기 로직
        window.location.href = `/project/${projectCode}`;
    } catch (error) {
        alert('프로젝트를 열 수 없습니다.');
    }
}

// 프로젝트 편집 함수
function editProject(projectCode) {
    try {
        // 프로젝트 편집 로직
        window.location.href = `/project/${projectCode}/edit`;
    } catch (error) {
        alert('프로젝트 편집을 열 수 없습니다.');
    }
}

// 새 프로젝트 모달 열기
async function openNewProjectModal() {
    try {
        // 모달 로직 구현
        const modal = new bootstrap.Modal(document.getElementById('newProjectModal'));
        modal.show();
    } catch (error) {
        alert('새 프로젝트 모달을 열 수 없습니다.');
    }
}

// 카드 편집 핸들러
function handleCardEdit(cardType, projectCode) {
    try {
        if (window.authManager && !window.authManager.canEditCard(cardType)) {
            alert('이 카드를 편집할 권한이 없습니다.');
            return;
        }
        
        // 편집 로직 구현
    } catch (error) {
        alert('편집 중 오류가 발생했습니다.');
    }
}

// 잠금 상태 복원 (임시 함수)
function restoreAllLockStates() {
    // 잠금 상태 복원 로직
}

// 모듈 로딩 상태 확인
function checkModulesLoaded() {
    const requiredModules = [
        'authManager',
        'dataTableManager', 
        'apiManager',
        'uiManager'
    ];
    
    const missingModules = requiredModules.filter(module => !window[module]);
    
    return missingModules.length === 0;
}

// 모듈 로딩 완료 후 초기화
window.addEventListener('load', function() {
    setTimeout(() => {
        if (checkModulesLoaded()) {
            console.log('모든 모듈 로딩 완료 - 페이지 초기화 시작');
            // UI Manager를 통해 페이지 초기화
            if (window.uiManager && typeof window.uiManager.initializePage === 'function') {
                window.uiManager.initializePage();
            } else {
                console.error('UIManager 또는 initializePage 함수를 찾을 수 없습니다');
            }
        } else {
            console.error('일부 모듈이 로드되지 않았습니다');
        }
    }, 100);
});