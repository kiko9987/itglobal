/**
 * Lead Autocomplete Utility
 * 프로젝트 등록 시 리드 데이터로 자동완성
 */

let autocompleteTimeout = null;
let selectedLeadData = null;

/**
 * 리드 자동완성 초기화
 */
export function initLeadAutocomplete() {
    const customerNameInput = document.getElementById('modern-site-manager');
    const autocompleteDropdown = document.getElementById('lead-autocomplete-dropdown');
    const leadNoInput = document.getElementById('modern-lead-no');
    const linkedBadge = document.getElementById('lead-linked-badge');
    const linkedText = document.getElementById('lead-linked-text');

    if (!customerNameInput || !autocompleteDropdown) {
        console.warn('[LEAD_AUTOCOMPLETE] Required elements not found');
        return;
    }

    // 입력 이벤트: 디바운싱 적용
    customerNameInput.addEventListener('input', async (e) => {
        const query = e.target.value.trim();

        // 입력값이 없으면 드롭다운 숨기기 및 리드 연동 해제
        if (!query) {
            hideAutocompleteDropdown();
            clearLeadSelection();
            return;
        }

        // 기존 타이머 취소
        if (autocompleteTimeout) {
            clearTimeout(autocompleteTimeout);
        }

        // 300ms 디바운싱
        autocompleteTimeout = setTimeout(async () => {
            await searchLeads(query);
        }, 300);
    });

    // 포커스 아웃 시 드롭다운 숨기기 (클릭 이벤트와 충돌 방지를 위해 지연)
    customerNameInput.addEventListener('blur', () => {
        setTimeout(() => {
            hideAutocompleteDropdown();
        }, 200);
    });

    // 모달이 닫힐 때 리드 선택 초기화
    const modal = document.getElementById('modernProjectModal');
    if (modal) {
        modal.addEventListener('hidden.bs.modal', () => {
            clearLeadSelection();
            hideAutocompleteDropdown();
        });
    }

    console.log('[LEAD_AUTOCOMPLETE] Initialized');
}

/**
 * 리드 검색 API 호출
 */
async function searchLeads(customerName) {
    const autocompleteDropdown = document.getElementById('lead-autocomplete-dropdown');

    try {
        const response = await fetch(`/leads/api/search?customer_name=${encodeURIComponent(customerName)}`, {
            method: 'GET',
            credentials: 'same-origin'
        });

        if (!response.ok) {
            console.error('[LEAD_AUTOCOMPLETE] Search failed:', response.statusText);
            hideAutocompleteDropdown();
            return;
        }

        const result = await response.json();
        const leads = result.leads || [];

        if (leads.length === 0) {
            showNoResultsMessage();
            return;
        }

        // 드롭다운 렌더링
        renderAutocompleteDropdown(leads);

    } catch (error) {
        console.error('[LEAD_AUTOCOMPLETE] Search error:', error);
        hideAutocompleteDropdown();
    }
}

/**
 * 자동완성 드롭다운 렌더링
 */
function renderAutocompleteDropdown(leads) {
    const dropdown = document.getElementById('lead-autocomplete-dropdown');

    // 드롭다운 HTML 생성
    const html = leads.map(lead => {
        const leadNo = lead['리드 No'] || '';
        const customerName = lead['고객명'] || '-';
        const phone = lead['연락처'] || '-';
        const address = lead['방문 주소'] || '-';
        const company = lead['거래처'] || '-';
        const status = lead['상태'] || '-';

        // 상태별 배지 색상
        const statusColors = {
            '유선 상담': 'secondary',
            '방문 예약': 'primary',
            '견적 제출': 'info',
            '방문 취소': 'warning',
            '공사 확정': 'success',
            '공사 드랍': 'danger'
        };
        const badgeColor = statusColors[status] || 'secondary';

        return `
            <a href="#" class="list-group-item list-group-item-action lead-autocomplete-item"
               data-lead='${JSON.stringify(lead).replace(/'/g, "&apos;")}'>
                <div class="d-flex w-100 justify-content-between align-items-start">
                    <div>
                        <h6 class="mb-1" style="font-size: 0.9rem; font-weight: 600;">
                            ${customerName}
                            <span class="badge bg-${badgeColor} ms-2" style="font-size: 0.7rem;">${status}</span>
                        </h6>
                        <p class="mb-0" style="font-size: 0.8rem; color: #6c757d;">
                            <i class="fas fa-phone me-1"></i>${phone}
                            ${company !== '-' ? `<span class="ms-2"><i class="fas fa-building me-1"></i>${company}</span>` : ''}
                        </p>
                        ${address !== '-' ? `<p class="mb-0 mt-1" style="font-size: 0.75rem; color: #999;"><i class="fas fa-map-marker-alt me-1"></i>${address}</p>` : ''}
                    </div>
                    <small class="text-muted" style="font-size: 0.7rem;">${leadNo}</small>
                </div>
            </a>
        `;
    }).join('');

    dropdown.innerHTML = html;
    dropdown.style.display = 'block';

    // 클릭 이벤트 등록
    dropdown.querySelectorAll('.lead-autocomplete-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const leadData = JSON.parse(item.getAttribute('data-lead'));
            selectLead(leadData);
        });
    });
}

/**
 * 리드 선택 처리
 */
function selectLead(leadData) {
    selectedLeadData = leadData;

    // 폼 필드 자동 채우기
    const customerNameInput = document.getElementById('modern-site-manager');
    const phoneInput = document.getElementById('modern-manager-phone');
    const addressInput = document.getElementById('modern-address');
    const clientInput = document.getElementById('modern-client');
    const leadNoInput = document.getElementById('modern-lead-no');
    const linkedBadge = document.getElementById('lead-linked-badge');
    const linkedText = document.getElementById('lead-linked-text');

    if (customerNameInput) customerNameInput.value = leadData['고객명'] || '';
    if (phoneInput) phoneInput.value = leadData['연락처'] || '';
    if (addressInput) addressInput.value = leadData['방문 주소'] || '';
    if (clientInput) clientInput.value = leadData['거래처'] || '';
    if (leadNoInput) leadNoInput.value = leadData['리드 No'] || '';

    // 리드 연동 배지 표시
    if (linkedBadge && linkedText) {
        linkedText.textContent = `${leadData['리드 No']} (${leadData['고객명']})`;
        linkedBadge.classList.remove('d-none');
    }

    // 드롭다운 숨기기
    hideAutocompleteDropdown();

    console.log('[LEAD_AUTOCOMPLETE] Lead selected:', leadData['리드 No']);
}

/**
 * 리드 선택 초기화
 */
function clearLeadSelection() {
    selectedLeadData = null;

    const leadNoInput = document.getElementById('modern-lead-no');
    const linkedBadge = document.getElementById('lead-linked-badge');

    if (leadNoInput) leadNoInput.value = '';
    if (linkedBadge) linkedBadge.classList.add('d-none');

    console.log('[LEAD_AUTOCOMPLETE] Lead selection cleared');
}

/**
 * 결과 없음 메시지 표시
 */
function showNoResultsMessage() {
    const dropdown = document.getElementById('lead-autocomplete-dropdown');
    dropdown.innerHTML = `
        <div class="list-group-item text-center text-muted" style="padding: 1rem;">
            <i class="fas fa-search me-2"></i>검색 결과가 없습니다
        </div>
    `;
    dropdown.style.display = 'block';
}

/**
 * 드롭다운 숨기기
 */
function hideAutocompleteDropdown() {
    const dropdown = document.getElementById('lead-autocomplete-dropdown');
    if (dropdown) {
        dropdown.style.display = 'none';
        dropdown.innerHTML = '';
    }
}

/**
 * 선택된 리드 데이터 반환
 */
export function getSelectedLead() {
    return selectedLeadData;
}
