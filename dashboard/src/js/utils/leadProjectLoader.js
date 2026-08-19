/**
 * Lead → Project Loader
 * 기존 리드(방문 예약/견적 제출)를 검색해서 프로젝트 폼에 자동 채움.
 * /leads/api/search-for-project 엔드포인트 사용.
 */

const STATUS_COLORS = {
  '방문 예약': 'primary',
  '견적 제출': 'info',
  '유선 상담': 'secondary',
  '공사 확정': 'success',
};

// 리드 플랫폼 → 프로젝트 유입 구분 매핑 (플랫폼별 1:1)
// 슬랙은 별도로 온라인(플랫폼) 형식으로 그룹핑하지만, 프로젝트 시트는 플랫폼별로 구분해 기록
const PLATFORM_TO_INFLOW = {
  '홈페이지': '홈페이지',
  '카카오톡': '카카오톡',
  '채널톡': '채널톡',
  '전화': '전화',
  '메일': '메일',
  '당근': '당근',
  '당근마켓': '당근',
  '숨고': '숨고',
  '큐플레이스': '큐플레이스',
  '거래처': '거래처',
  '소개': '소개',
};

function mapPlatformToInflow(platform) {
  if (!platform) return '';
  return PLATFORM_TO_INFLOW[platform.trim()] || platform.trim();
}

// 카카오/채널톡 랜덤 닉네임 패턴 감지 — "햇님 310", "초콜릿 699", "스노우맨 257" 등
// 매니저가 실명 확인 안 한 채로 프로젝트 등록되지 않도록 자동 채우기 skip
function looksLikeChatNickname(name) {
  const s = (name || '').trim();
  if (!s) return false;
  // 한글 단어 + 공백 + 2~4자리 숫자 (카카오 오픈채팅 랜덤 닉 표준)
  if (/^[가-힣]{2,10}\s\d{2,4}$/.test(s)) return true;
  // generic 대체 값
  if (['익명', '익명 고객', '카카오톡 사용자', '채널톡 사용자'].includes(s)) return true;
  return false;
}

let debounceTimer = null;
let currentLead = null;
let moduleOpts = null;

/**
 * 리드→프로젝트 검색 로더 초기화.
 * @param {object} opts
 * @param {string} opts.inputId          - 검색 input id
 * @param {string} opts.resultsId        - 결과 드롭다운 id
 * @param {string} opts.leadNoInputId    - hidden Lead No input id
 * @param {string} opts.linkedInfoId     - 연결 표시 컨테이너 id
 * @param {string} opts.linkedInfoTextId - 연결 표시 텍스트 id
 * @param {string} opts.unlinkBtnId      - 연결 해제 버튼 id
 * @param {string} opts.badgeId          - 상단 뱃지 id (없으면 무시)
 * @param {string} opts.badgeTextId      - 상단 뱃지 텍스트 id
 * @param {string} opts.modalId          - 모달 id (닫힘 시 초기화)
 * @param {object} opts.targets          - 자동 채움 대상 { client, siteManager, phone, email, address }
 */
export function initLeadProjectLoader(opts) {
  const input = document.getElementById(opts.inputId);
  const results = document.getElementById(opts.resultsId);
  if (!input || !results) {
    console.warn('[LEAD_PROJECT_LOADER] Required elements missing');
    return;
  }
  moduleOpts = opts;

  // "내 방문 현장만 불러오기" 체크박스 — 해제 시 전체 팀 현장 검색(?all=1)
  const myOnlyCheckbox = opts.myOnlyCheckboxId
    ? document.getElementById(opts.myOnlyCheckboxId)
    : null;

  const runSearch = async (query) => {
    try {
      const allParam = (myOnlyCheckbox && !myOnlyCheckbox.checked) ? '&all=1' : '';
      const url = `/leads/api/search-for-project?q=${encodeURIComponent(query)}${allParam}`;
      const res = await fetch(url, { credentials: 'same-origin' });
      if (!res.ok) {
        renderNoResults(results, `검색 실패 (${res.status})`);
        return;
      }
      const json = await res.json();
      const leads = (json.data && json.data.leads) || json.leads || [];
      if (!leads.length) {
        renderNoResults(results, query ? '검색 결과가 없습니다' : '표시할 리드가 없습니다');
        return;
      }
      renderResults(results, leads, (lead) => applyLead(lead, opts));
    } catch (err) {
      console.error('[LEAD_PROJECT_LOADER] search error', err);
      renderNoResults(results, '검색 중 오류가 발생했습니다');
    }
  };

  input.addEventListener('focus', () => {
    if (debounceTimer) clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => runSearch(input.value.trim()), 100);
  });

  input.addEventListener('input', () => {
    if (debounceTimer) clearTimeout(debounceTimer);
    const query = input.value.trim();
    debounceTimer = setTimeout(() => runSearch(query), 250);
  });

  input.addEventListener('blur', () => {
    setTimeout(() => hideResults(results), 200);
  });

  // 체크박스 토글 시 현재 검색어로 즉시 재검색 (본인/전체 전환 반영)
  if (myOnlyCheckbox) {
    myOnlyCheckbox.addEventListener('change', () => {
      if (debounceTimer) clearTimeout(debounceTimer);
      runSearch(input.value.trim());
    });
  }

  const unlinkBtn = document.getElementById(opts.unlinkBtnId);
  if (unlinkBtn) {
    unlinkBtn.addEventListener('click', () => clearLead(opts));
  }

  const modal = document.getElementById(opts.modalId);
  if (modal) {
    modal.addEventListener('hidden.bs.modal', () => clearLead(opts));
  }
}

function renderResults(container, leads, onSelect) {
  const html = leads.map((lead) => {
    const status = escapeHtml(lead.status || '-');
    const badge = STATUS_COLORS[lead.status] || 'secondary';
    const name = escapeHtml(lead.name || '-');
    const phone = escapeHtml(lead.phone || '-');
    const addr = escapeHtml(lead.address || '');
    const platform = escapeHtml(lead.platform || '');
    const leadNo = escapeHtml(lead.lead_no || '');
    return `
      <a href="#" class="list-group-item list-group-item-action lead-project-item"
         style="padding: 0.6rem 0.9rem; border-bottom: 1px solid var(--gray-200); text-decoration: none; color: inherit; display: block;"
         data-lead-no="${leadNo}">
        <div style="display: flex; justify-content: space-between; align-items: flex-start; gap: 0.5rem;">
          <div style="flex: 1; min-width: 0;">
            <div style="font-weight: 600; font-size: 0.9rem;">
              ${name}
              <span class="badge bg-${badge} ms-1" style="font-size: 0.65rem; font-weight: 500;">${status}</span>
              ${platform ? `<span class="text-muted ms-1" style="font-size: 0.7rem;">${platform}</span>` : ''}
            </div>
            <div style="font-size: 0.78rem; color: #666; margin-top: 2px;">
              <i class="fas fa-phone" style="font-size: 0.7rem;"></i> ${phone}
            </div>
            ${addr ? `<div style="font-size: 0.72rem; color: #999; margin-top: 2px;"><i class="fas fa-map-marker-alt" style="font-size: 0.7rem;"></i> ${addr}</div>` : ''}
          </div>
          <small class="text-muted" style="font-size: 0.7rem; white-space: nowrap;">${leadNo}</small>
        </div>
      </a>
    `;
  }).join('');
  container.innerHTML = html;
  container.classList.remove('d-none');
  container.querySelectorAll('.lead-project-item').forEach((item, idx) => {
    item.addEventListener('mousedown', (e) => {
      e.preventDefault();
      onSelect(leads[idx]);
    });
  });
}

function renderNoResults(container, msg) {
  container.innerHTML = `
    <div class="text-center text-muted" style="padding: 1rem; font-size: 0.85rem;">
      <i class="fas fa-info-circle me-1"></i>${escapeHtml(msg)}
    </div>
  `;
  container.classList.remove('d-none');
}

function hideResults(container) {
  container.classList.add('d-none');
  container.innerHTML = '';
}

function applyLead(lead, opts) {
  const { targets } = opts;
  const inflow = mapPlatformToInflow(lead.platform);

  // 사업자 ↔ 유입구분 정합성 가드.
  // #modern-client 셀렉트는 선택된 사업자의 허용 유입구분만 담고 있으므로,
  // 이 현장의 유입구분이 옵션에 없으면 = 사업자와 안 맞는 현장 → 불러오기 차단.
  // (사업자 미선택 상태면 전체 옵션이라 통과 → 이후 사업자 변경 시 그쪽에서 경고)
  const clientEl = targets.client ? document.getElementById(targets.client) : null;
  const companyEl = document.getElementById('modern-company');
  const company = companyEl ? (companyEl.value || '').trim() : '';
  if (company && inflow && clientEl && clientEl.tagName === 'SELECT') {
    const realOptions = Array.from(clientEl.options).filter((o) => o.value);
    // 옵션이 아직 로드 안 된 상태(placeholder만)면 판정 보류 → 오차단 방지
    const allowed = realOptions.length === 0 || realOptions.some((o) => o.value === inflow);
    if (!allowed) {
      showMismatchAlert(
        opts,
        `${company} 프로젝트에는 이 현장(유입구분: ${inflow})을 지정할 수 없습니다. 사업자를 확인하세요.`,
      );
      return; // 어떤 필드도 채우지 않고 리드 링크도 하지 않음
    }
  }

  currentLead = lead;

  const setVal = (id, val) => {
    if (!id) return;
    const el = document.getElementById(id);
    if (!el || val == null || val === '') return;
    el.value = val;
    flashHighlight(el);
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new Event('input', { bubbles: true }));
  };

  setVal(targets.client, inflow);
  // 채팅 인입 리드의 랜덤 닉네임(예: "햇님 310")은 auto-fill skip — 실명 확인용 placeholder 강조
  const siteMgrEl = targets.siteManager ? document.getElementById(targets.siteManager) : null;
  if (siteMgrEl && looksLikeChatNickname(lead.name)) {
    siteMgrEl.value = '';
    siteMgrEl.placeholder = `채팅 닉네임(${lead.name}) — 실제 담당자명을 입력하세요`;
    flashHighlight(siteMgrEl);
  } else {
    setVal(targets.siteManager, lead.name);
  }
  setVal(targets.phone, lead.phone);
  setVal(targets.address, lead.address);

  // 방문 사진 업로드된 리드는 폴더 링크 자동 채움 (없으면 스킵)
  if (targets.documentPath && lead.folder_link) {
    setVal(targets.documentPath, lead.folder_link);
  }

  // 이메일: 값이 있으면 채우고, 없으면 "이메일 확인 불가" 체크박스를 자동 활성화
  const emailInput = targets.email ? document.getElementById(targets.email) : null;
  const emailUnknownCheck = opts.emailUnknownCheckId
    ? document.getElementById(opts.emailUnknownCheckId)
    : null;
  const emailValue = (lead.email || '').trim();
  if (emailInput) {
    if (emailValue && emailValue !== '-') {
      if (emailUnknownCheck && emailUnknownCheck.checked) {
        emailUnknownCheck.checked = false;
        emailUnknownCheck.dispatchEvent(new Event('change', { bubbles: true }));
      }
      emailInput.value = emailValue;
      flashHighlight(emailInput);
      emailInput.dispatchEvent(new Event('change', { bubbles: true }));
      emailInput.dispatchEvent(new Event('input', { bubbles: true }));
    } else if (emailUnknownCheck) {
      if (!emailUnknownCheck.checked) {
        emailUnknownCheck.checked = true;
        emailUnknownCheck.dispatchEvent(new Event('change', { bubbles: true }));
      }
      flashHighlight(emailInput);
    }
  }

  const leadNoInput = document.getElementById(opts.leadNoInputId);
  if (leadNoInput) leadNoInput.value = lead.lead_no || '';

  const searchInput = document.getElementById(opts.inputId);
  if (searchInput) searchInput.value = '';

  const linkedInfo = document.getElementById(opts.linkedInfoId);
  const linkedText = document.getElementById(opts.linkedInfoTextId);
  if (linkedInfo && linkedText) {
    linkedText.textContent = `${lead.lead_no} · ${lead.name || '이름 미상'} (${lead.status || '-'})`;
    linkedInfo.classList.remove('d-none');
  }

  const badge = document.getElementById(opts.badgeId);
  const badgeText = document.getElementById(opts.badgeTextId);
  if (badge && badgeText) {
    badgeText.textContent = `${lead.lead_no} 연결됨`;
    badge.classList.remove('d-none');
  }

  const results = document.getElementById(opts.resultsId);
  if (results) hideResults(results);

  console.log('[LEAD_PROJECT_LOADER] linked lead:', lead.lead_no);
}

function clearLead(opts) {
  currentLead = null;
  const leadNoInput = document.getElementById(opts.leadNoInputId);
  if (leadNoInput) leadNoInput.value = '';
  const linkedInfo = document.getElementById(opts.linkedInfoId);
  if (linkedInfo) linkedInfo.classList.add('d-none');
  const badge = document.getElementById(opts.badgeId);
  if (badge) badge.classList.add('d-none');
  const results = document.getElementById(opts.resultsId);
  if (results) hideResults(results);
  const searchInput = document.getElementById(opts.inputId);
  if (searchInput) searchInput.value = '';

  // "내 방문 현장만" 체크박스는 클리어(모달 닫힘·연결해제) 시 기본 ON 으로 복귀
  const myOnlyCheckbox = opts.myOnlyCheckboxId
    ? document.getElementById(opts.myOnlyCheckboxId)
    : null;
  if (myOnlyCheckbox) myOnlyCheckbox.checked = true;

  // 자동 기입된 필드 비우기 (5개: 유입/담당자/연락처/이메일/주소)
  const targets = opts.targets || {};
  const clearField = (id) => {
    if (!id) return;
    const el = document.getElementById(id);
    if (!el) return;
    el.value = '';
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new Event('input', { bubbles: true }));
  };
  clearField(targets.client);
  clearField(targets.siteManager);
  clearField(targets.phone);
  clearField(targets.address);
  clearField(targets.documentPath);

  // 이메일: 확인 불가 체크가 켜져 있으면 해제하고 값 비우기
  const emailUnknownCheck = opts.emailUnknownCheckId
    ? document.getElementById(opts.emailUnknownCheckId)
    : null;
  if (emailUnknownCheck && emailUnknownCheck.checked) {
    emailUnknownCheck.checked = false;
    emailUnknownCheck.dispatchEvent(new Event('change', { bubbles: true }));
  }
  clearField(targets.email);
}

// 사업자-유입구분 불일치로 불러오기가 차단됐을 때 경고 표시.
// 검색 결과 영역(클릭 지점 바로 아래)에만 인라인 경고 — 상단 배너와 중복 방지.
function showMismatchAlert(opts, message) {
  const results = document.getElementById(opts.resultsId);
  if (results) {
    results.innerHTML = `
      <div class="text-danger" style="padding: 0.8rem; font-size: 0.85rem;">
        <i class="fas fa-exclamation-triangle me-1"></i>${escapeHtml(message)}
      </div>
    `;
    results.classList.remove('d-none');
  }
}

function flashHighlight(el) {
  const prevBg = el.style.backgroundColor;
  const prevTr = el.style.transition;
  el.style.transition = 'background-color 0.35s ease';
  el.style.backgroundColor = '#fff3cd';
  setTimeout(() => {
    el.style.backgroundColor = prevBg || '';
    setTimeout(() => { el.style.transition = prevTr || ''; }, 400);
  }, 700);
}

function escapeHtml(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function getLinkedLead() {
  return currentLead;
}

// 외부(모달)에서 연결된 현장을 해제할 때 사용 — 로드된 모든 필드+뱃지 초기화.
// 사업자를 현장과 안 맞는 값으로 바꿨을 때 현장 연결 자체를 끊는 용도.
export function clearLinkedLead() {
  if (currentLead && moduleOpts) {
    clearLead(moduleOpts);
  }
}
