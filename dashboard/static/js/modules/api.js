// API 통신 및 데이터 관리 모듈
class ApiManager {
    constructor() {
        this.projectsData = [];
        this.currentPage = 1;
        this.totalPages = 1;
        this.perPage = 20;
    }

    // 프로젝트 데이터 로드
    async loadProjectsData() {
        try {
            const response = await fetch('/api/get-projects-data', {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                },
                credentials: 'same-origin'
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const result = await response.json();
            
            if (result.success && Array.isArray(result.data)) {
                this.projectsData = result.data;
                
                // 데이터 로드 후 UI 업데이트
                if (window.dataTableManager) {
                    window.dataTableManager.initializeDataTable(this.projectsData);
                }
                
                this.loadManagerFilter();
                this.generateMobileCards();
                
                return this.projectsData;
            } else {
                throw new Error(result.message || '데이터 로드 실패');
            }
        } catch (error) {
            throw error;
        }
    }

    // 담당자 필터 로드
    loadManagerFilter() {
        try {
            const managers = [...new Set(this.projectsData
                .map(item => item['담당자'])
                .filter(manager => manager && manager.trim() !== '')
            )].sort();

            const managerSelect = document.getElementById('managerFilter');
            if (managerSelect) {
                managerSelect.innerHTML = '<option value="">모든 담당자</option>';
                managers.forEach(manager => {
                    const option = document.createElement('option');
                    option.value = manager;
                    option.textContent = manager;
                    managerSelect.appendChild(option);
                });
            }
        } catch (error) {
            // 담당자 필터 로드 실패
        }
    }

    // 모바일 카드 생성
    generateMobileCards() {
        const mobileContainer = document.getElementById('mobile-cards-container');
        if (!mobileContainer) return;

        if (this.projectsData.length === 0) {
            mobileContainer.innerHTML = '<div class="text-center p-4"><p class="text-muted">프로젝트 데이터가 없습니다.</p></div>';
            return;
        }

        let cardsHTML = '';
        this.projectsData.forEach(project => {
            const statusBadge = this.getStatusBadge(project);
            const completeness = this.checkDataCompleteness(project);
            
            cardsHTML += `
                <div class="col-12 col-md-6 col-lg-4 mb-3">
                    <div class="card h-100 mobile-project-card" onclick="viewProject('${project['프로젝트 코드']}')">
                        <div class="card-body">
                            <div class="d-flex justify-content-between align-items-start mb-2">
                                <h6 class="card-title mb-0">${this.createProjectBadge(project['프로젝트 코드'])}</h6>
                                <div class="d-flex align-items-center">
                                    ${completeness.missingCount === 0 ? 
                                        '<i class="fas fa-check-circle text-success me-2"></i>' : 
                                        '<i class="fas fa-exclamation-triangle text-warning me-2"></i>'
                                    }
                                    ${statusBadge}
                                </div>
                            </div>
                            <p class="card-text text-muted small mb-1">
                                <i class="fas fa-user me-1"></i>${project['담당자'] || '-'}
                                <span class="mx-2">|</span>
                                <i class="fas fa-building me-1"></i>${project['거래처'] || '-'}
                            </p>
                            <p class="card-text mb-2">
                                <i class="fas fa-map-marker-alt me-1"></i>
                                ${project['현장 주소'] || '주소 정보 없음'}
                            </p>
                            <p class="card-text text-muted small">
                                <i class="fas fa-tools me-1"></i>
                                ${project['공사 내용'] || '공사내용 없음'}
                            </p>
                            <div class="row text-center mt-3">
                                <div class="col-6">
                                    <div class="border-end">
                                        <div class="text-muted small">총액</div>
                                        <div class="fw-bold">
                                            ${this.formatCurrency(project['총액 2'] || project['총액2'] || project['S'] || 0)}
                                        </div>
                                    </div>
                                </div>
                                <div class="col-6">
                                    <div class="text-muted small">미수금</div>
                                    <div class="fw-bold text-danger">
                                        ${this.formatCurrency(project['미수금'] || project['미수금W'] || project['W'] || 0)}
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            `;
        });

        mobileContainer.innerHTML = cardsHTML;
    }

    // 프로젝트 상태 뱃지 생성
    getStatusBadge(row) {
        const startDate = row['공사 시작'];
        const endDate = row['공사 종료'];
        const totalAmount = parseFloat(row['총액 2'] || row['총액2'] || row['S'] || 0);
        const outstandingAmount = parseFloat(row['미수금'] || row['미수금W'] || row['미수금 W'] || row['W'] || 0);
        
        const today = new Date();
        const start = startDate ? new Date(startDate) : null;
        const end = endDate ? new Date(endDate) : null;
        
        if (!start || !end) {
            return '<span class="badge bg-secondary">일정 미정</span>';
        }
        
        if (outstandingAmount <= 0 && totalAmount > 0) {
            return '<span class="badge bg-success">완료</span>';
        }
        
        if (today < start) {
            return '<span class="badge bg-info">시공 예정</span>';
        } else if (today >= start && today <= end) {
            return '<span class="badge bg-warning">시공 중</span>';
        } else if (today > end) {
            if (outstandingAmount > 0) {
                return '<span class="badge bg-danger">미수금</span>';
            } else {
                return '<span class="badge bg-success">완료</span>';
            }
        }
        
        return '<span class="badge bg-light text-dark">상태 미정</span>';
    }

    // 데이터 완성도 확인
    checkDataCompleteness(row) {
        const requiredFields = [
            '프로젝트 코드', '담당자', '거래처', '현장 주소', '공사 내용',
            '공사 시작', '공사 종료', '총액 2', '미수금'
        ];
        
        const missingFields = requiredFields.filter(field => {
            const value = row[field] || row[field.replace(' ', '')] || row[field.replace(' 2', '2')];
            return !value || (typeof value === 'string' && value.trim() === '');
        });
        
        return {
            missingCount: missingFields.length,
            missingFields: missingFields,
            completeness: ((requiredFields.length - missingFields.length) / requiredFields.length) * 100
        };
    }

    // 통화 포맷팅
    formatCurrency(amount) {
        const numAmount = parseFloat(amount) || 0;
        if (numAmount === 0) return '-';
        
        const formattedAmount = Math.abs(numAmount).toLocaleString('ko-KR');
        const unit = numAmount >= 100000000 ? '억' : numAmount >= 10000 ? '만원' : '원';
        
        if (unit === '억') {
            const eok = Math.floor(numAmount / 100000000);
            const remainder = numAmount % 100000000;
            if (remainder === 0) {
                return `${eok}억원`;
            } else {
                const man = Math.floor(remainder / 10000);
                return man > 0 ? `${eok}억 ${man}만원` : `${eok}억원`;
            }
        } else if (unit === '만원') {
            const man = Math.floor(numAmount / 10000);
            const remainder = numAmount % 10000;
            return remainder === 0 ? `${man}만원` : `${formattedAmount}원`;
        } else {
            return `${formattedAmount}원`;
        }
    }

    // 프로젝트 뱃지 생성
    createProjectBadge(projectCode) {
        if (!projectCode) return '';
        return `<span class="badge bg-primary">${projectCode}</span>`;
    }

    // 담당자 뱃지 생성  
    createManagerBadge(manager) {
        if (!manager) return '<span class="text-muted">-</span>';
        return `<span class="badge bg-info">${manager}</span>`;
    }

    // 업체 뱃지 생성
    createCompanyBadge(company) {
        if (!company) return '<span class="text-muted">-</span>';
        return `<span class="badge bg-secondary">${company}</span>`;
    }

    // 날짜 포맷팅
    formatDate(dateString) {
        if (!dateString) return '';
        const date = new Date(dateString);
        return date.toLocaleDateString('ko-KR', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        });
    }

    // 프로젝트 데이터 반환
    getProjectsData() {
        return this.projectsData;
    }
}

// 전역 ApiManager 인스턴스 생성
window.apiManager = new ApiManager();

// 전역 함수들을 위한 래퍼
async function loadProjectsData() {
    return await window.apiManager.loadProjectsData();
}

function loadManagerFilter() {
    return window.apiManager.loadManagerFilter();
}

function generateMobileCards() {
    return window.apiManager.generateMobileCards();
}

function getStatusBadge(row) {
    return window.apiManager.getStatusBadge(row);
}

function checkDataCompleteness(row) {
    return window.apiManager.checkDataCompleteness(row);
}

function formatCurrency(amount) {
    return window.apiManager.formatCurrency(amount);
}

function createProjectBadge(projectCode) {
    return window.apiManager.createProjectBadge(projectCode);
}

function createManagerBadge(manager) {
    return window.apiManager.createManagerBadge(manager);
}

function createCompanyBadge(company) {
    return window.apiManager.createCompanyBadge(company);
}

function formatDate(dateString) {
    return window.apiManager.formatDate(dateString);
}