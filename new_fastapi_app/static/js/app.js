/**
 * ITG 대시보드 메인 JavaScript
 */

// 전역 변수
window.ITG = window.ITG || {};
ITG.config = {
    apiBaseUrl: '/api',
    maxRetries: 3,
    retryDelay: 1000,
    animationDuration: 300
};

/**
 * 유틸리티 함수들
 */
ITG.utils = {
    /**
     * 숫자를 통화 형식으로 포맷
     */
    formatCurrency: function(amount) {
        if (!amount) return '0원';
        return new Intl.NumberFormat('ko-KR', {
            style: 'currency',
            currency: 'KRW'
        }).format(amount);
    },

    /**
     * 날짜 포맷팅
     */
    formatDate: function(date, format = 'YYYY-MM-DD') {
        if (!date) return '-';
        const d = new Date(date);
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');

        switch (format) {
            case 'YYYY-MM-DD':
                return `${year}-${month}-${day}`;
            case 'YYYY.MM.DD':
                return `${year}.${month}.${day}`;
            case 'MM/DD':
                return `${month}/${day}`;
            default:
                return d.toLocaleDateString('ko-KR');
        }
    },

    /**
     * 디바운스 함수
     */
    debounce: function(func, wait, immediate) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                timeout = null;
                if (!immediate) func.apply(this, args);
            };
            const callNow = immediate && !timeout;
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
            if (callNow) func.apply(this, args);
        };
    },

    /**
     * 스로틀 함수
     */
    throttle: function(func, limit) {
        let inThrottle;
        return function(...args) {
            if (!inThrottle) {
                func.apply(this, args);
                inThrottle = true;
                setTimeout(() => inThrottle = false, limit);
            }
        };
    },

    /**
     * URL 파라미터 가져오기
     */
    getUrlParameter: function(name) {
        const urlParams = new URLSearchParams(window.location.search);
        return urlParams.get(name);
    },

    /**
     * 랜덤 ID 생성
     */
    generateId: function(prefix = 'id') {
        return prefix + '_' + Math.random().toString(36).substr(2, 9);
    },

    /**
     * 클립보드에 텍스트 복사
     */
    copyToClipboard: function(text) {
        return navigator.clipboard.writeText(text).then(() => {
            ITG.ui.showToast('클립보드에 복사되었습니다', 'success');
            return true;
        }).catch(err => {
            console.error('클립보드 복사 실패:', err);
            ITG.ui.showToast('클립보드 복사에 실패했습니다', 'error');
            return false;
        });
    },

    /**
     * 로컬 스토리지 관리
     */
    storage: {
        set: function(key, value) {
            try {
                localStorage.setItem(key, JSON.stringify(value));
                return true;
            } catch (e) {
                console.error('로컬 스토리지 저장 실패:', e);
                return false;
            }
        },
        get: function(key, defaultValue = null) {
            try {
                const item = localStorage.getItem(key);
                return item ? JSON.parse(item) : defaultValue;
            } catch (e) {
                console.error('로컬 스토리지 읽기 실패:', e);
                return defaultValue;
            }
        },
        remove: function(key) {
            try {
                localStorage.removeItem(key);
                return true;
            } catch (e) {
                console.error('로컬 스토리지 삭제 실패:', e);
                return false;
            }
        }
    }
};

/**
 * UI 관련 함수들
 */
ITG.ui = {
    /**
     * 토스트 메시지 표시
     */
    showToast: function(message, type = 'info', duration = 3000) {
        const toastContainer = document.getElementById('toast-container') || this.createToastContainer();
        const toast = this.createToast(message, type);

        toastContainer.appendChild(toast);

        // 애니메이션으로 표시
        setTimeout(() => toast.classList.add('show'), 100);

        // 자동 제거
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, duration);
    },

    /**
     * 토스트 컨테이너 생성
     */
    createToastContainer: function() {
        const container = document.createElement('div');
        container.id = 'toast-container';
        container.className = 'position-fixed top-0 end-0 p-3';
        container.style.zIndex = '1055';
        document.body.appendChild(container);
        return container;
    },

    /**
     * 토스트 요소 생성
     */
    createToast: function(message, type) {
        const typeClass = {
            success: 'text-bg-success',
            error: 'text-bg-danger',
            warning: 'text-bg-warning',
            info: 'text-bg-info'
        }[type] || 'text-bg-info';

        const toast = document.createElement('div');
        toast.className = `toast ${typeClass}`;
        toast.innerHTML = `
            <div class="toast-body d-flex justify-content-between align-items-center">
                <span>${message}</span>
                <button type="button" class="btn-close btn-close-white" onclick="this.closest('.toast').remove()"></button>
            </div>
        `;
        return toast;
    },

    /**
     * 확인 다이얼로그
     */
    confirm: function(message, title = '확인') {
        return new Promise((resolve) => {
            if (window.bootstrap) {
                // Bootstrap 모달 사용
                const modal = this.createConfirmModal(message, title, resolve);
                document.body.appendChild(modal);
                const bsModal = new bootstrap.Modal(modal);
                bsModal.show();
            } else {
                // 기본 confirm 사용
                resolve(confirm(`${title}\n\n${message}`));
            }
        });
    },

    /**
     * 확인 모달 생성
     */
    createConfirmModal: function(message, title, callback) {
        const modalId = ITG.utils.generateId('confirm-modal');
        const modal = document.createElement('div');
        modal.className = 'modal fade';
        modal.id = modalId;
        modal.innerHTML = `
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">${title}</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <p>${message}</p>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal" onclick="ITG.ui.handleConfirm('${modalId}', false)">취소</button>
                        <button type="button" class="btn btn-primary" onclick="ITG.ui.handleConfirm('${modalId}', true)">확인</button>
                    </div>
                </div>
            </div>
        `;

        modal.callback = callback;
        modal.addEventListener('hidden.bs.modal', () => modal.remove());
        return modal;
    },

    /**
     * 확인 모달 결과 처리
     */
    handleConfirm: function(modalId, result) {
        const modal = document.getElementById(modalId);
        if (modal && modal.callback) {
            modal.callback(result);
            bootstrap.Modal.getInstance(modal).hide();
        }
    },

    /**
     * 로딩 스피너 표시
     */
    showLoading: function(element) {
        if (typeof element === 'string') {
            element = document.querySelector(element);
        }
        if (element) {
            element.style.position = 'relative';
            const spinner = document.createElement('div');
            spinner.className = 'loading-overlay d-flex justify-content-center align-items-center';
            spinner.innerHTML = '<div class="spinner-border text-primary" role="status"><span class="visually-hidden">로딩 중...</span></div>';
            spinner.style.cssText = `
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: rgba(255, 255, 255, 0.8);
                z-index: 1000;
            `;
            element.appendChild(spinner);
        }
    },

    /**
     * 로딩 스피너 제거
     */
    hideLoading: function(element) {
        if (typeof element === 'string') {
            element = document.querySelector(element);
        }
        if (element) {
            const spinner = element.querySelector('.loading-overlay');
            if (spinner) {
                spinner.remove();
            }
        }
    }
};

/**
 * API 통신 관련 함수들
 */
ITG.api = {
    /**
     * HTTP 요청 공통 함수
     */
    request: async function(url, options = {}) {
        const defaultOptions = {
            headers: {
                'Content-Type': 'application/json',
            },
            credentials: 'same-origin'
        };

        const mergedOptions = { ...defaultOptions, ...options };

        try {
            const response = await fetch(url, mergedOptions);

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const contentType = response.headers.get('content-type');
            if (contentType && contentType.includes('application/json')) {
                return await response.json();
            } else {
                return await response.text();
            }
        } catch (error) {
            console.error('API 요청 실패:', error);
            throw error;
        }
    },

    /**
     * GET 요청
     */
    get: function(url, params = {}) {
        const urlParams = new URLSearchParams(params);
        const fullUrl = urlParams.toString() ? `${url}?${urlParams}` : url;
        return this.request(fullUrl, { method: 'GET' });
    },

    /**
     * POST 요청
     */
    post: function(url, data = {}) {
        return this.request(url, {
            method: 'POST',
            body: JSON.stringify(data)
        });
    },

    /**
     * PUT 요청
     */
    put: function(url, data = {}) {
        return this.request(url, {
            method: 'PUT',
            body: JSON.stringify(data)
        });
    },

    /**
     * DELETE 요청
     */
    delete: function(url) {
        return this.request(url, { method: 'DELETE' });
    }
};

/**
 * 폼 관련 함수들
 */
ITG.form = {
    /**
     * 폼 유효성 검사
     */
    validate: function(form) {
        if (typeof form === 'string') {
            form = document.querySelector(form);
        }

        if (!form) return false;

        form.classList.add('was-validated');
        return form.checkValidity();
    },

    /**
     * 폼 데이터를 객체로 변환
     */
    serialize: function(form) {
        if (typeof form === 'string') {
            form = document.querySelector(form);
        }

        const formData = new FormData(form);
        const object = {};

        formData.forEach((value, key) => {
            if (object[key]) {
                if (!Array.isArray(object[key])) {
                    object[key] = [object[key]];
                }
                object[key].push(value);
            } else {
                object[key] = value;
            }
        });

        return object;
    },

    /**
     * 필드 값 설정
     */
    setFieldValue: function(form, fieldName, value) {
        if (typeof form === 'string') {
            form = document.querySelector(form);
        }

        const field = form.querySelector(`[name="${fieldName}"]`);
        if (field) {
            field.value = value;
            field.dispatchEvent(new Event('change', { bubbles: true }));
        }
    },

    /**
     * 필드 값 가져오기
     */
    getFieldValue: function(form, fieldName) {
        if (typeof form === 'string') {
            form = document.querySelector(form);
        }

        const field = form.querySelector(`[name="${fieldName}"]`);
        return field ? field.value : null;
    }
};

/**
 * 테이블 관련 함수들
 */
ITG.table = {
    /**
     * 테이블 정렬
     */
    sort: function(table, columnIndex, direction = 'asc') {
        if (typeof table === 'string') {
            table = document.querySelector(table);
        }

        const tbody = table.querySelector('tbody');
        const rows = Array.from(tbody.querySelectorAll('tr'));

        rows.sort((a, b) => {
            const aValue = a.cells[columnIndex].textContent.trim();
            const bValue = b.cells[columnIndex].textContent.trim();

            // 숫자인지 확인
            const aNum = parseFloat(aValue.replace(/[^\d.-]/g, ''));
            const bNum = parseFloat(bValue.replace(/[^\d.-]/g, ''));

            if (!isNaN(aNum) && !isNaN(bNum)) {
                return direction === 'asc' ? aNum - bNum : bNum - aNum;
            } else {
                return direction === 'asc'
                    ? aValue.localeCompare(bValue, 'ko-KR')
                    : bValue.localeCompare(aValue, 'ko-KR');
            }
        });

        // 정렬된 행들을 다시 추가
        rows.forEach(row => tbody.appendChild(row));
    },

    /**
     * 테이블 필터링
     */
    filter: function(table, searchText) {
        if (typeof table === 'string') {
            table = document.querySelector(table);
        }

        const tbody = table.querySelector('tbody');
        const rows = tbody.querySelectorAll('tr');

        rows.forEach(row => {
            const text = row.textContent.toLowerCase();
            const matches = text.includes(searchText.toLowerCase());
            row.style.display = matches ? '' : 'none';
        });
    }
};

/**
 * 프로젝트 관련 함수들
 */
ITG.project = {
    /**
     * 프로젝트 목록 새로고침
     */
    refreshList: function() {
        window.location.reload();
    },

    /**
     * 프로젝트 취소
     */
    cancel: async function(projectCode) {
        const confirmed = await ITG.ui.confirm(
            `프로젝트 ${projectCode}를 취소하시겠습니까?`,
            '프로젝트 취소'
        );

        if (confirmed) {
            try {
                const response = await fetch(`/projects/${projectCode}/cancel`, {
                    method: 'POST',
                    credentials: 'same-origin'
                });

                if (response.ok) {
                    ITG.ui.showToast('프로젝트가 취소되었습니다', 'success');
                    setTimeout(() => window.location.reload(), 1000);
                } else {
                    throw new Error('취소 요청이 실패했습니다');
                }
            } catch (error) {
                console.error('프로젝트 취소 실패:', error);
                ITG.ui.showToast('프로젝트 취소에 실패했습니다', 'error');
            }
        }
    },

    /**
     * 프로젝트 재개
     */
    resume: async function(projectCode) {
        const confirmed = await ITG.ui.confirm(
            `프로젝트 ${projectCode}를 재개하시겠습니까?`,
            '프로젝트 재개'
        );

        if (confirmed) {
            try {
                const response = await fetch(`/projects/${projectCode}/resume`, {
                    method: 'POST',
                    credentials: 'same-origin'
                });

                if (response.ok) {
                    ITG.ui.showToast('프로젝트가 재개되었습니다', 'success');
                    setTimeout(() => window.location.reload(), 1000);
                } else {
                    throw new Error('재개 요청이 실패했습니다');
                }
            } catch (error) {
                console.error('프로젝트 재개 실패:', error);
                ITG.ui.showToast('프로젝트 재개에 실패했습니다', 'error');
            }
        }
    }
};

/**
 * 초기화 함수
 */
ITG.init = function() {
    console.log('ITG 대시보드 초기화 중...');

    // 툴팁 초기화
    if (window.bootstrap) {
        const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
        tooltipTriggerList.map(function (tooltipTriggerEl) {
            return new bootstrap.Tooltip(tooltipTriggerEl);
        });
    }

    // 현재 페이지에 맞는 네비게이션 활성화
    const currentPath = window.location.pathname;
    const navLinks = document.querySelectorAll('.nav-link');
    navLinks.forEach(link => {
        if (link.getAttribute('href') === currentPath) {
            link.classList.add('active');
        }
    });

    // 숫자 입력 필드에 천 단위 구분자 추가
    const numberInputs = document.querySelectorAll('input[type="number"]');
    numberInputs.forEach(input => {
        input.addEventListener('blur', function() {
            if (this.value) {
                this.setAttribute('data-raw-value', this.value);
                this.value = ITG.utils.formatCurrency(parseInt(this.value));
            }
        });

        input.addEventListener('focus', function() {
            const rawValue = this.getAttribute('data-raw-value');
            if (rawValue) {
                this.value = rawValue;
            }
        });
    });

    // 검색 필드 디바운스 적용
    const searchInputs = document.querySelectorAll('input[type="search"], input[name="search"]');
    searchInputs.forEach(input => {
        input.addEventListener('input', ITG.utils.debounce(function() {
            // 검색 로직은 각 페이지에서 구현
            console.log('검색어:', this.value);
        }, 500));
    });

    // 자동 저장 기능 (편집 페이지용)
    const editForms = document.querySelectorAll('form[data-auto-save="true"]');
    editForms.forEach(form => {
        let lastSaveData = ITG.form.serialize(form);

        const autoSave = ITG.utils.debounce(function() {
            const currentData = ITG.form.serialize(form);
            if (JSON.stringify(currentData) !== JSON.stringify(lastSaveData)) {
                // 자동 저장 로직
                console.log('자동 저장 실행');
                lastSaveData = currentData;
            }
        }, 5000);

        form.addEventListener('input', autoSave);
    });

    console.log('ITG 대시보드 초기화 완료');
};

// DOM 로드 완료 시 초기화
document.addEventListener('DOMContentLoaded', ITG.init);

// 브라우저 뒤로가기/앞으로가기 처리
window.addEventListener('popstate', function(event) {
    if (event.state && event.state.page) {
        // 필요한 경우 상태 복원
        console.log('페이지 상태 복원:', event.state);
    }
});

// 전역 에러 핸들러
window.addEventListener('error', function(event) {
    console.error('전역 오류:', event.error);
    ITG.ui.showToast('예상치 못한 오류가 발생했습니다', 'error');
});

// 전역 객체에 ITG 등록
window.ITG = ITG;