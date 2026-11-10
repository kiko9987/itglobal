/* 그룹웨어 인사 화면 — 구성원 목록·인사카드·수정 창, 내 정보.
   window.GW_HR = { mode: 'people' } 또는 { mode: 'me', personId } */
(function () {
  'use strict';

  var CFG = window.GW_HR || {};
  var cardEl = document.getElementById('hrCard');
  var state = { people: [], selectedId: null, showResigned: false, query: '', current: null };

  // ── 공통 ──────────────────────────────────────────
  function esc(v) {
    return String(v == null ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function api(url, opts) {
    opts = opts || {};
    return fetch(url, {
      method: opts.method || 'GET',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: opts.body ? JSON.stringify(opts.body) : undefined
    }).then(function (r) {
      return r.json().catch(function () { return {}; }).then(function (j) {
        if (!r.ok || j.success === false) {
          throw new Error(j.message || ('요청을 처리하지 못했습니다 (' + r.status + ')'));
        }
        return j;
      });
    });
  }

  function phone(v) {
    var d = String(v || '').replace(/\D/g, '');
    if (d.length === 11) return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
    if (d.length === 10) return d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6);
    return v || '';
  }

  function tenure(from, to) {
    if (!from) return '';
    var a = new Date(from + 'T00:00:00');
    var b = to ? new Date(to + 'T00:00:00') : new Date();
    if (isNaN(a) || isNaN(b) || b < a) return '';
    var months = (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth());
    if (b.getDate() < a.getDate()) months -= 1;
    var y = Math.floor(months / 12), m = months % 12;
    if (y === 0 && m === 0) return '1개월 미만';
    return (y ? y + '년 ' : '') + (m ? m + '개월' : '').trim();
  }

  function initial(name) {
    return esc(String(name || '?').slice(0, 1));
  }

  function statusBadge(s) {
    var cls = s === '재직' ? 'gw-badge-active' : (s === '휴직' ? 'gw-badge-leave' : 'gw-badge-resigned');
    return '<span class="gw-badge ' + cls + '">' + esc(s) + '</span>';
  }

  function org(p) {
    return [p.department, p.job_rank, p.job_title].filter(Boolean).map(esc).join(' · ');
  }

  function dd(label, value, opts) {
    opts = opts || {};
    var empty = value == null || value === '';
    return '<div' + (opts.wide ? ' class="gw-dl-wide"' : '') + '><dt>' + esc(label) + '</dt>' +
      '<dd' + (empty ? ' class="empty"' : '') + (opts.id ? ' id="' + opts.id + '"' : '') + '>' +
      (empty ? '-' : (opts.html ? value : esc(value))) + '</dd></div>';
  }

  // ── 인사카드 ──────────────────────────────────────
  function renderCard(p) {
    state.current = p;
    var resigned = p.status === '퇴사';
    var h = '<div class="gw-card">';
    h += '<div class="gw-card-head"><div class="gw-avatar gw-avatar-lg" aria-hidden="true">' + initial(p.name) + '</div>' +
      '<div><h2 class="gw-card-name">' + esc(p.name) + ' ' + statusBadge(p.status) +
      (p.access === 'field' ? ' <span class="gw-badge gw-badge-field">현장·미화</span>' : '') + '</h2>' +
      '<div class="gw-card-org">' + (org(p) || '<span class="text-muted">소속 없음</span>') + '</div></div>';
    if (p.can_edit && CFG.mode !== 'me') {   // 수정은 구성원 화면에서(내 정보엔 수정 창이 없음)
      h += '<div class="gw-card-actions"><button type="button" class="btn btn-sm btn-outline-secondary" id="hrEditBtn">' +
        '<i class="fas fa-pen me-1" aria-hidden="true"></i>수정</button></div>';
    }
    h += '</div>';

    h += '<div class="gw-sec-title">근무</div><dl class="gw-dl">' +
      dd('입사일', p.hire_date) +
      dd('근속', tenure(p.hire_date, resigned ? (p.resign_date || p.last_work_date) : null)) +
      dd('그룹 입사일', p.group_hire_date) +
      dd('계약 형태', p.contract_type) +
      dd('직무', p.job_role) +
      '</dl>';

    h += '<div class="gw-sec-title">연락처</div><dl class="gw-dl">' +
      dd('회사 메일', p.work_email) + dd('개인 메일', p.personal_email) + dd('전화', phone(p.phone)) +
      dd('주소', p.address, { wide: true }) + '</dl>';

    var ssnHtml = p.has_ssn ? esc(p.ssn_masked) : '';
    if (p.has_ssn && p.can_reveal_ssn) {
      ssnHtml += '<button type="button" class="btn btn-sm btn-outline-secondary gw-ssn-btn" id="hrSsnBtn">전체 보기</button>';
    }
    h += '<div class="gw-sec-title">개인 정보</div><dl class="gw-dl">' +
      dd('성별', p.gender) + dd('생년월일', p.birth_date) +
      dd('주민등록번호', p.has_ssn ? ssnHtml : '', { html: true, id: 'hrSsnCell' }) + '</dl>';

    if (p.can_edit) {
      h += '<div class="gw-sec-title">그룹웨어</div><dl class="gw-dl">' +
        dd('로그인 메일', p.login_email) +
        dd('접근', p.access === 'field' ? '현장·미화(내 정보·휴가만)' : '사무실(전체 메뉴)') +
        dd('슬랙 ID', p.slack_id) + '</dl>';
    }

    if (resigned || p.resign_date) {
      h += '<div class="gw-sec-title">퇴사</div><dl class="gw-dl">' +
        dd('퇴사일', p.resign_date) + dd('마지막 근무일', p.last_work_date) +
        dd('퇴사 사유', p.resign_reason, { wide: true }) + '</dl>';
    }

    var appts = p.appointments || [];
    h += '<div class="gw-sec-title">발령 이력</div>';
    if (appts.length) {
      h += '<table class="gw-table"><thead><tr><th>적용일</th><th>부서</th><th>직무</th><th>직급</th><th>직책</th><th>메모</th></tr></thead><tbody>';
      appts.forEach(function (a) {
        h += '<tr><td>' + esc(a.apply_date || '-') + '</td><td>' + esc(a.department || '') + '</td><td>' +
          esc(a.job_role || '') + '</td><td>' + esc(a.job_rank || '') + '</td><td>' + esc(a.job_title || '') +
          '</td><td>' + esc(a.memo || '') + '</td></tr>';
      });
      h += '</tbody></table>';
    } else {
      h += '<div class="gw-log">기록 없음</div>';
    }

    if (p.can_edit && p.memo) {
      h += '<div class="gw-sec-title">메모</div><div style="white-space:pre-wrap">' + esc(p.memo) + '</div>';
    }

    if (p.access_log && p.access_log.length) {
      var labels = { ssn_view: '주민번호 전체 보기', edit: '수정', create: '추가' };
      h += '<details class="mt-3"><summary class="gw-log">열람·수정 기록 ' + p.access_log.length + '건</summary><ul class="gw-log mt-1">';
      p.access_log.forEach(function (l) {
        h += '<li>' + esc(l.at) + ' · ' + esc(l.viewer_email) + ' · ' + esc(labels[l.action] || l.action) +
          (l.detail && l.action === 'edit' ? ' (' + esc(l.detail) + ')' : '') + '</li>';
      });
      h += '</ul></details>';
    }

    h += '</div>';
    cardEl.innerHTML = h;

    var editBtn = document.getElementById('hrEditBtn');
    if (editBtn) editBtn.addEventListener('click', function () { openEdit(p); });
    var ssnBtn = document.getElementById('hrSsnBtn');
    if (ssnBtn) ssnBtn.addEventListener('click', function () { revealSsn(p, ssnBtn); });
  }

  function loadCard(id) {
    cardEl.innerHTML = '<div class="gw-empty">불러오는 중…</div>';
    return api('/api/hr/people/' + id).then(function (j) {
      renderCard(j.person);
    }).catch(function (e) {
      cardEl.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>';
    });
  }

  var ssnTimer = null;
  function revealSsn(p, btn) {
    if (!window.confirm(p.name + '님의 주민등록번호 전체를 봅니다.\n본 기록(누가·언제)이 남습니다.')) return;
    btn.disabled = true;
    api('/api/hr/people/' + p.id + '/ssn', { method: 'POST' }).then(function (j) {
      var cell = document.getElementById('hrSsnCell');
      if (!cell) return;
      cell.textContent = j.ssn;
      clearTimeout(ssnTimer);
      // 30초 뒤 다시 가림(화면을 열어 둔 채 자리를 비울 때 대비)
      ssnTimer = setTimeout(function () { if (state.current && state.current.id === p.id) loadCard(p.id); }, 30000);
    }).catch(function (e) {
      btn.disabled = false;
      window.alert(e.message);
    });
  }

  // ── 구성원 목록 ────────────────────────────────────
  var listEl = document.getElementById('hrList');
  var countEl = document.getElementById('hrCount');

  function renderList() {
    if (!listEl) return;
    var q = state.query.trim().toLowerCase();
    var rows = state.people.filter(function (p) {
      if (!q) return true;
      return [p.name, p.department, p.job_role, p.job_rank, p.job_title].join(' ').toLowerCase().indexOf(q) !== -1;
    });
    countEl.textContent = rows.length + '명';
    listEl.innerHTML = rows.map(function (p) {
      var sub = [p.department, p.job_rank].filter(Boolean).join(' · ');
      return '<li><button type="button" role="option" class="gw-list-item' + (p.status === '퇴사' ? ' resigned' : '') +
        '" data-id="' + p.id + '" aria-selected="' + (p.id === state.selectedId) + '">' +
        '<span class="gw-avatar" aria-hidden="true">' + initial(p.name) + '</span>' +
        '<span><span class="gw-list-name">' + esc(p.name) + '</span>' +
        (p.status !== '재직' ? ' ' + statusBadge(p.status) : '') +
        (p.access === 'field' && p.status !== '퇴사' ? ' <span class="gw-badge gw-badge-field">현장</span>' : '') +
        '<br><span class="gw-list-sub">' + esc(sub || '-') + '</span></span></button></li>';
    }).join('') || '<li class="gw-empty">맞는 구성원이 없습니다.</li>';
  }

  function select(id) {
    state.selectedId = id;
    renderList();
    try { history.replaceState(null, '', '#' + id); } catch (e) { /* 무시 */ }
    loadCard(id);
  }

  function loadList(keepId) {
    return api('/api/hr/people' + (state.showResigned ? '?resigned=1' : '')).then(function (j) {
      state.people = j.people || [];
      renderList();
      var want = keepId || parseInt((location.hash || '').slice(1), 10);
      if (want && state.people.some(function (p) { return p.id === want; })) select(want);
    }).catch(function (e) {
      listEl.innerHTML = '<li class="gw-empty">' + esc(e.message) + '</li>';
    });
  }

  // ── 수정·추가 창 ───────────────────────────────────
  var modalEl = document.getElementById('hrEditModal');
  var formEl = document.getElementById('hrEditForm');
  var errEl = document.getElementById('hrEditError');
  var editingId = null;
  var FIELDS = ['name', 'status', 'contract_type', 'hire_date', 'group_hire_date', 'department', 'job_role',
    'job_rank', 'job_title', 'work_email', 'personal_email', 'phone', 'address', 'gender', 'birth_date',
    'login_email', 'access', 'slack_id', 'resign_date', 'last_work_date', 'resign_reason', 'memo'];

  function modal() {
    return window.bootstrap ? window.bootstrap.Modal.getOrCreateInstance(modalEl) : null;
  }

  function openEdit(p) {
    if (!formEl) return;
    editingId = p ? p.id : null;
    document.getElementById('hrEditTitle').textContent = p ? p.name + ' 인사카드 수정' : '구성원 추가';
    FIELDS.forEach(function (k) {
      var el = formEl.elements[k];
      if (!el) return;
      var v = p ? p[k] : '';
      if (!p && k === 'status') v = '재직';
      if (!p && k === 'access') v = 'staff';
      el.value = v == null ? '' : v;
    });
    formEl.elements.ssn.value = '';
    formEl.elements.ssn.placeholder = p && p.has_ssn ? '저장됨 — 바꿀 때만 입력' : '숫자 13자리';
    errEl.textContent = '';
    var m = modal();
    if (m) m.show();
  }

  if (formEl) {
    formEl.addEventListener('submit', function (e) {
      e.preventDefault();
      errEl.textContent = '';
      var body = {};
      FIELDS.forEach(function (k) { if (formEl.elements[k]) body[k] = formEl.elements[k].value; });
      if (!body.name.trim()) { errEl.textContent = '이름을 입력하세요.'; formEl.elements.name.focus(); return; }
      var ssn = formEl.elements.ssn.value.trim();
      if (ssn) {
        if (ssn.replace(/\D/g, '').length !== 13) { errEl.textContent = '주민등록번호는 숫자 13자리여야 합니다.'; formEl.elements.ssn.focus(); return; }
        body.ssn = ssn;
      }
      var saveBtn = document.getElementById('hrEditSave');
      saveBtn.disabled = true;
      var req = editingId
        ? api('/api/hr/people/' + editingId, { method: 'PUT', body: body })
        : api('/api/hr/people', { method: 'POST', body: body });
      req.then(function (j) {
        var m = modal();
        if (m) m.hide();
        state.selectedId = j.person.id;
        return loadList(j.person.id);
      }).catch(function (err) {
        errEl.textContent = err.message;
      }).then(function () { saveBtn.disabled = false; });
    });
    formEl.elements.ssn.addEventListener('input', function () { errEl.textContent = ''; });
  }

  // ── 시작 ───────────────────────────────────────────
  if (CFG.mode === 'me' && CFG.personId) {
    loadCard(CFG.personId);
    return;
  }

  if (CFG.mode === 'people') {
    listEl.addEventListener('click', function (e) {
      var btn = e.target.closest('.gw-list-item');
      if (btn) select(parseInt(btn.getAttribute('data-id'), 10));
    });
    document.getElementById('hrSearch').addEventListener('input', function (e) {
      state.query = e.target.value;
      renderList();
    });
    document.getElementById('hrShowResigned').addEventListener('change', function (e) {
      state.showResigned = e.target.checked;
      loadList(state.selectedId);
    });
    var addBtn = document.getElementById('hrAddBtn');
    if (addBtn) addBtn.addEventListener('click', function () { openEdit(null); });
    loadList();
  }
})();
