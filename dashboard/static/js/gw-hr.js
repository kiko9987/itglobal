/* 그룹웨어 인사 화면 — 구성원 목록·인사카드·수정 창, 내 정보.
   window.GW_HR = { mode: 'people' } 또는 { mode: 'me', personId } */
(function () {
  'use strict';

  var CFG = window.GW_HR || {};
  // 로그인 방식(인사카드 access 칸) — 카드와 수정 창이 같은 말을 쓰게 여기 한곳에
  var ACCESS_LABEL = {
    staff: '회사 구글 계정',
    field: '슬랙 휴가 앱 링크 (현장·미화 · 내 정보·휴가만)'
  };
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

  // 현장·미화 직원 배지 = 직무(시공·미화 등). 직무가 비었으면 '현장'
  function fieldBadge(p) {
    if (p.access !== 'field') return '';
    return ' <span class="gw-badge gw-badge-field">' + esc(p.job_role || '현장') + '</span>';
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
      fieldBadge(p) + '</h2>' +
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
        dd('로그인 방식', ACCESS_LABEL[p.access] || ACCESS_LABEL.staff) +
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
      h += '<table class="gw-table gw-appt-table"><thead><tr><th>적용일</th><th>부서</th><th>직무</th><th>직급</th><th>직책</th><th>메모</th></tr></thead><tbody>';
      appts.forEach(function (a) {
        h += '<tr><td>' + esc(a.apply_date || '-') + '</td><td>' + esc(a.department || '') + '</td><td>' +
          esc(a.job_role || '') + '</td><td>' + esc(a.job_rank || '') + '</td><td>' + esc(a.job_title || '') +
          '</td><td>' + esc(a.memo || '') + '</td></tr>';
      });
      h += '</tbody></table>';
    } else {
      h += '<div class="gw-log">기록 없음</div>';
    }

    h += renderPrivate(p);

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
    bindPrivate(p);
  }

  // ── 경력·학력·가족·서류(본인·인사카드 관리자만 서버가 보내 줌) ─────────────
  function period2(a, b) {
    if (!a && !b) return '';
    return esc(a || '') + ' ~ ' + esc(b || '');
  }

  function table(head, rows) {
    return '<div class="gw-table-wrap"><table class="gw-table gw-appt-table"><thead><tr>' +
      head.map(function (x) { return '<th>' + esc(x) + '</th>'; }).join('') + '</tr></thead><tbody>' +
      rows.join('') + '</tbody></table></div>';
  }

  function kb(n) {
    if (!n) return '';
    return n >= 1048576 ? (Math.round(n / 104857.6) / 10) + 'MB' : Math.max(1, Math.round(n / 1024)) + 'KB';
  }

  function renderPrivate(p) {
    if (!p.careers && !p.documents) return '';
    var h = '';
    if (p.careers && p.careers.length) {
      h += '<div class="gw-sec-title">경력</div>' + table(['회사', '부서', '직무', '직급', '계약', '기간'],
        p.careers.map(function (c) {
          return '<tr><td>' + esc(c.company) + '</td><td>' + esc(c.department) + '</td><td>' + esc(c.job_role) +
            '</td><td>' + esc(c.job_rank) + '</td><td>' + esc(c.contract_type) + '</td><td>' + period2(c.begin_date, c.end_date) + '</td></tr>';
        }));
    }
    if (p.educations && p.educations.length) {
      h += '<div class="gw-sec-title">학력</div>' + table(['학교', '전공', '구분', '졸업', '기간'],
        p.educations.map(function (e) {
          return '<tr><td>' + esc(e.school) + '</td><td>' + esc(e.major) + '</td><td>' + esc(e.level) +
            '</td><td>' + esc(e.graduate_type) + '</td><td>' + period2(e.begin_date, e.end_date) + '</td></tr>';
        }));
    }
    if (p.families && p.families.length) {
      h += '<div class="gw-sec-title">가족 · 공제 대상</div>' + table(['이름', '관계', '연락처', '주민등록번호', '공제 관계', '기본공제', '자녀공제', '건보 피부양'],
        p.families.map(function (f) {
          var ssn = f.has_ssn ? esc(f.ssn_masked) + (p.can_reveal_ssn ? ' <button type="button" class="btn btn-sm btn-outline-secondary gw-ssn-btn" data-fam-ssn="' + f.id + '">전체 보기</button>' : '') : '';
          var yn = function (b) { return b ? '<span class="gw-badge gw-badge-active">예</span>' : '<span class="text-muted">-</span>'; };
          return '<tr><td>' + esc(f.name) + '</td><td>' + esc(f.relation) + '</td><td>' + esc(phone(f.phone)) +
            '</td><td data-fam-cell="' + f.id + '">' + ssn + '</td><td>' + esc(f.deduction_relation) + '</td><td>' + yn(f.default_deduction) +
            '</td><td>' + yn(f.child_deduction) + '</td><td>' + yn(f.health_dependent) + '</td></tr>';
        }));
    }
    var docs = p.documents || [];
    if (docs.length || p.can_edit) {
      h += '<div class="gw-sec-title">서류</div>';
      if (docs.length) {
        h += table(['종류', '파일', '크기', '올린 날', ''], docs.map(function (d) {
          return '<tr><td>' + esc(d.doc_type || '기타') + '</td><td class="gw-doc-name">' + esc(d.file_name) + '</td><td>' + kb(d.size) +
            '</td><td>' + (d.uploaded_by === 'flex-import' ? '<span class="text-muted">플렉스에서 옮김</span>' : esc((d.created_at || '').slice(0, 10))) +
            '</td><td class="gw-doc-act">' +
            '<a class="btn btn-sm btn-outline-secondary" href="/api/hr/documents/' + d.id + '?view=1" target="_blank" rel="noopener">보기</a> ' +
            '<a class="btn btn-sm btn-outline-secondary" href="/api/hr/documents/' + d.id + '">받기</a>' +
            (p.can_edit && CFG.mode !== 'me' ? ' <button type="button" class="btn btn-sm btn-outline-danger" data-doc-del="' + d.id + '">삭제</button>' : '') +
            '</td></tr>';
        }));
      } else {
        h += '<div class="gw-log">올린 서류가 없습니다.</div>';
      }
      if (p.can_edit && CFG.mode !== 'me') {
        h += '<form class="gw-doc-upload" id="hrDocForm">' +
          '<select class="form-select form-select-sm" name="doc_type" aria-label="서류 종류">' +
          (p.doc_types || []).map(function (t) { return '<option>' + esc(t) + '</option>'; }).join('') + '</select>' +
          '<input type="file" class="form-control form-control-sm" name="file" aria-label="파일" required>' +
          '<button type="submit" class="btn btn-sm btn-outline-primary">올리기</button>' +
          '<span class="gw-form-error" id="hrDocError" role="alert"></span></form>';
      }
      h += '<div class="gw-log mt-1">서류를 열거나 받으면 누가 언제 봤는지 기록이 남습니다.</div>';
    }
    return h;
  }

  function bindPrivate(p) {
    cardEl.querySelectorAll('[data-fam-ssn]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        if (!window.confirm('가족 주민등록번호 전체를 봅니다.\n본 기록(누가·언제)이 남습니다.')) return;
        var id = btn.getAttribute('data-fam-ssn');
        btn.disabled = true;
        api('/api/hr/families/' + id + '/ssn', { method: 'POST' }).then(function (j) {
          var cell = cardEl.querySelector('[data-fam-cell="' + id + '"]');
          if (cell) cell.textContent = j.ssn;
          setTimeout(function () { if (state.current && state.current.id === p.id) loadCard(p.id); }, 30000);
        }).catch(function (e) { btn.disabled = false; window.alert(e.message); });
      });
    });
    cardEl.querySelectorAll('[data-doc-del]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        if (!window.confirm('이 서류를 지웁니다. 되돌릴 수 없습니다.')) return;
        api('/api/hr/documents/' + btn.getAttribute('data-doc-del'), { method: 'DELETE' })
          .then(function () { loadCard(p.id); })
          .catch(function (e) { window.alert(e.message); });
      });
    });
    var form = document.getElementById('hrDocForm');
    if (form) {
      form.addEventListener('submit', function (e) {
        e.preventDefault();
        var err = document.getElementById('hrDocError');
        var file = form.elements.file.files[0];
        if (!file) { err.textContent = '파일을 골라 주세요.'; return; }
        if (file.size > 8 * 1024 * 1024) { err.textContent = '8MB까지 올릴 수 있습니다.'; return; }
        var fd = new FormData(form);
        var btn = form.querySelector('button[type=submit]');
        btn.disabled = true;
        err.textContent = '';
        fetch('/api/hr/people/' + p.id + '/documents', { method: 'POST', body: fd, credentials: 'same-origin' })
          .then(function (r) { return r.json().then(function (j) { if (!r.ok || !j.success) throw new Error(j.message || '올리지 못했습니다.'); return j; }); })
          .then(function () { loadCard(p.id); })
          .catch(function (x) { err.textContent = x.message; btn.disabled = false; });
      });
    }
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
        (p.status !== '퇴사' ? fieldBadge(p) : '') +
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

  // 수정 창의 로그인 방식 문구를 카드와 같게(템플릿이 옛 문구여도 맞춰짐)
  if (formEl && formEl.elements.access) {
    Array.prototype.forEach.call(formEl.elements.access.options, function (o) {
      if (o.value === 'staff') o.textContent = ACCESS_LABEL.staff;
      if (o.value === 'field') o.textContent = '슬랙 휴가 앱 링크 (현장·미화)';
    });
    var accLabel = formEl.querySelector('label[for="f_access"]');
    if (accLabel) accLabel.textContent = '로그인 방식';
    var hint = formEl.elements.login_email && formEl.elements.login_email.parentNode.querySelector('.form-text');
    if (hint) hint.textContent = '회사 구글 계정으로 들어오는 메일이에요. 슬랙 링크로 들어오는 현장·미화 직원은 비워 둬도 돼요.';
  }

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
