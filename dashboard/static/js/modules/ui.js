// UI 및 필터 관리 모듈
class UIManager {
    constructor() {
        this.setupEventListeners();
    }

    // 이벤트 리스너 설정
    setupEventListeners() {
        // 브라우저 크기 변경 시 DataTable 반응형 조정
        $(window).on('resize', this.debounce(() => {
            this.adjustResponsiveLayout();
        }, 250));
        
        // 새 프로젝트 확인 타이머 설정
        this.setupNewProjectCheckTimer();
    }

    // 반응형 레이아웃 조정
    adjustResponsiveLayout() {
        try {
            if (window.dataTableManager && window.dataTableManager.getDataTable()) {
                const dataTable = window.dataTableManager.getDataTable();
                
                // DataTable 컬럼 조정
                dataTable.columns.adjust();
                
                // 컬럼 너비 동기화 (스크롤 비활성화 시에도 안전하게 처리)
                window.dataTableManager.syncTableColumnWidths();
                
                // 모바일 화면에서는 DataTable 숨김
                const isMobile = $(window).width() < 768;
                $('#desktop-table-container').toggle(!isMobile);
                $('#mobile-cards-container').toggle(isMobile);
            }
        } catch (error) {
            // 반응형 조정 실패
        }
    }

    // 새 프로젝트 확인 타이머
    setupNewProjectCheckTimer() {
        function checkForNewProject() {
            const now = new Date();
            const hours = now.getHours();
            
            // 업무 시간 (9-18시)에만 30초마다 확인
            if (hours >= 9 && hours < 18) {
                setTimeout(() => {
                    if (window.apiManager) {
                        window.apiManager.loadProjectsData().catch(() => {});
                    }
                    checkForNewProject();
                }, 30000);
            } else {
                // 업무 외 시간엔 5분마다 확인
                setTimeout(() => {
                    if (window.apiManager) {
                        window.apiManager.loadProjectsData().catch(() => {});
                    }
                    checkForNewProject();
                }, 300000);
            }
        }
        
        // 초기 실행
        setTimeout(checkForNewProject, 30000);
    }

    // 필터 하이라이트 업데이트
    updateFilterHighlight(element) {
        $('.filter-control').removeClass('filter-active');
        if (element.value && element.value.trim() !== '') {
            $(element).addClass('filter-active');
        }
    }

    // 필터 적용
    applyFilters() {
        try {
            if (!window.dataTableManager || !window.dataTableManager.getDataTable()) {
                return;
            }

            const dataTable = window.dataTableManager.getDataTable();
            const searchValue = $('#searchInput').val().toLowerCase();
            const statusFilter = $('#statusFilter').val();
            const managerFilter = $('#managerFilter').val();
            const startDateFilter = $('#startDateFilter').val();
            const endDateFilter = $('#endDateFilter').val();

            // 전체 데이터에서 필터링
            const filteredData = window.apiManager.getProjectsData().filter(row => {
                // 검색어 필터
                if (searchValue && !this.matchesSearch(row, searchValue)) {
                    return false;
                }

                // 상태 필터
                if (statusFilter && !this.matchesStatus(row, statusFilter)) {
                    return false;
                }

                // 담당자 필터
                if (managerFilter && row['담당자'] !== managerFilter) {
                    return false;
                }

                // 날짜 필터
                if (startDateFilter || endDateFilter) {
                    if (!this.matchesDateRange(row, startDateFilter, endDateFilter)) {
                        return false;
                    }
                }

                return true;
            });

            // DataTable 업데이트
            dataTable.clear().rows.add(filteredData).draw();

            // 필터 하이라이트 업데이트
            $('.filter-control').each((index, element) => {
                this.updateFilterHighlight(element);
            });

        } catch (error) {
            // 필터 적용 실패
        }
    }

    // 검색어 매칭
    matchesSearch(row, searchValue) {
        const searchableFields = [
            '프로젝트 코드', '담당자', '거래처', '현장 주소', '공사 내용'
        ];
        
        return searchableFields.some(field => {
            const value = row[field];
            return value && value.toString().toLowerCase().includes(searchValue);
        });
    }

    // 상태 매칭
    matchesStatus(row, status) {
        const currentStatus = this.getCurrentStatus(row);
        return currentStatus === status;
    }

    // 현재 상태 확인
    getCurrentStatus(row) {
        const startDate = row['공사 시작'];
        const endDate = row['공사 종료'];
        const totalAmount = parseFloat(row['총액 2'] || row['총액2'] || row['S'] || 0);
        const outstandingAmount = parseFloat(row['미수금'] || row['미수금W'] || row['미수금 W'] || row['W'] || 0);
        
        const today = new Date();
        const start = startDate ? new Date(startDate) : null;
        const end = endDate ? new Date(endDate) : null;
        
        if (!start || !end) return 'unknown';
        
        if (outstandingAmount <= 0 && totalAmount > 0) return 'completed';
        if (today < start) return 'scheduled';
        if (today >= start && today <= end) return 'in_progress';
        if (today > end && outstandingAmount > 0) return 'outstanding';
        if (today > end) return 'completed';
        
        return 'unknown';
    }

    // 날짜 범위 매칭
    matchesDateRange(row, startDateFilter, endDateFilter) {
        const projectStart = row['공사 시작'];
        const projectEnd = row['공사 종료'];
        
        if (!projectStart || !projectEnd) return false;
        
        const projectStartDate = new Date(projectStart);
        const projectEndDate = new Date(projectEnd);
        
        if (startDateFilter) {
            const filterStartDate = new Date(startDateFilter);
            if (projectStartDate < filterStartDate) return false;
        }
        
        if (endDateFilter) {
            const filterEndDate = new Date(endDateFilter);
            if (projectEndDate > filterEndDate) return false;
        }
        
        return true;
    }

    // 필터 초기화
    clearFilters() {
        try {
            $('#searchInput').val('');
            $('#statusFilter').val('');
            $('#managerFilter').val('');
            $('#startDateFilter').val('');
            $('#endDateFilter').val('');
            
            $('.filter-control').removeClass('filter-active');
            
            // DataTable 초기화
            if (window.dataTableManager && window.dataTableManager.getDataTable()) {
                const dataTable = window.dataTableManager.getDataTable();
                dataTable.clear().rows.add(window.apiManager.getProjectsData()).draw();
            }
            
            // 모바일 카드도 초기화
            if (window.apiManager) {
                window.apiManager.generateMobileCards();
            }
            
        } catch (error) {
            // 필터 초기화 실패
        }
    }

    // 디바운스 함수
    debounce(func, wait, immediate) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                timeout = null;
                if (!immediate) func(...args);
            };
            const callNow = immediate && !timeout;
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
            if (callNow) func(...args);
        };
    }

    // 페이지 초기화
    async initializePage() {
        try {
            // 로딩 표시
            this.showLoading();
            
            // 데이터 로드
            await window.apiManager.loadProjectsData();
            
            // 초기 레이아웃 설정
            this.adjustResponsiveLayout();
            
            // 로딩 숨김
            this.hideLoading();
            
        } catch (error) {
            console.error('페이지 초기화 오류:', error);
            const errorMsg = error.message || '알 수 없는 오류가 발생했습니다';
            this.showError(`데이터 로드 실패: ${errorMsg}\n페이지를 새로고침해주세요.`);
        }
    }

    // 로딩 표시
    showLoading() {
        $('#loadingIndicator').show();
    }

    // 로딩 숨김
    hideLoading() {
        $('#loadingIndicator').hide();
    }

    // 오류 표시
    showError(message) {
        // 사용자에게 오류 메시지 표시
        alert(message);
    }
}

// 전역 UIManager 인스턴스 생성
window.uiManager = new UIManager();

// 전역 함수들을 위한 래퍼
function updateFilterHighlight(element) {
    return window.uiManager.updateFilterHighlight(element);
}

function applyFilters() {
    return window.uiManager.applyFilters();
}

function clearFilters() {
    return window.uiManager.clearFilters();
}

async function initializePage() {
    return await window.uiManager.initializePage();
}