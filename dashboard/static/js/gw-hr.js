/* 그룹웨어 인사 화면 — 구성원 목록·인사카드·수정 창, 내 정보.
   window.GW_HR = { mode: 'people' } 또는 { mode: 'me', personId } */
(function () {
  'use strict';

  var CFG = window.GW_HR || {};
  // 로그인 방식(인사카드 access 칸) — 카드와 수정 창이 같은 말을 쓰게 여기 한곳에
  var ACCESS_LABEL = {
    staff: '회사 구글 계정',
    field: '로그인 번호·슬랙 휴가 앱 링크 (현장·미화 · 내 정보·휴가만)'
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
    // 본문이 있을 때만 JSON 머리글을 붙인다 — 본문 없는 POST/DELETE(주민번호 전체 보기·서류 삭제)에
    // JSON 머리글을 붙이면 운영 보안 미들웨어가 빈 본문을 JSON 으로 읽다 400 을 낸다(11-10 운영에서 발견)
    var init = { method: opts.method || 'GET', credentials: 'same-origin' };
    if (opts.body) {
      init.headers = { 'Content-Type': 'application/json' };
      init.body = JSON.stringify(opts.body);
    }
    return fetch(url, init).then(function (r) {
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
    var sr = document.getElementById('hrCardSr');
    if (sr) sr.textContent = (p.name || '') + ' 인사카드를 불러왔어요.';
    var resigned = p.status === '퇴사';
    var h = '<div class="gw-card">';
    h += '<div class="gw-card-head"><div class="gw-avatar gw-avatar-lg" aria-hidden="true">' + initial(p.name) + '</div>' +
      '<div><h2 class="gw-card-name">' + esc(p.name) + ' ' + statusBadge(p.status) +
      fieldBadge(p) + '</h2>' +
      '<div class="gw-card-org">' + (org(p) || '<span class="text-muted">소속 없음</span>') + '</div></div>';
    if (p.can_edit && CFG.mode !== 'me') {   // 수정은 구성원 화면에서(내 정보엔 수정 창이 없음)
      h += '<div class="gw-card-actions">' +
        '<button type="button" class="btn btn-sm btn-outline-secondary" id="hrCertBtn" aria-expanded="false" aria-controls="hrCertPanel">' +
        '<i class="fas fa-file-signature me-1" aria-hidden="true"></i>증명서 발급</button>' +
        '<button type="button" class="btn btn-sm btn-outline-secondary" id="hrEditBtn">' +
        '<i class="fas fa-pen me-1" aria-hidden="true"></i>수정</button></div>';
    }
    h += '</div>';
    if (p.can_edit && CFG.mode !== 'me') h += '<div class="gw-cert" id="hrCertPanel" hidden></div>';
    h += tasksHtml(p);     // 입사·퇴사 할 일은 맨 위(11-20)

    h += '<div class="gw-sec-title">근무</div><dl class="gw-dl">' +
      dd('입사일', p.hire_date) +
      dd('근속', tenure(p.hire_date, resigned ? (p.resign_date || p.last_work_date) : null)) +
      dd('그룹 입사일', p.group_hire_date) +
      dd('계약 형태', p.contract_type) +
      dd('직무', p.job_role) +
      '</dl>';

    var selfMode = CFG.mode === 'me' && p.can_self_edit;
    h += (selfMode
      ? '<div class="gw-sec-head"><div class="gw-sec-title">연락처</div>' + reqBtn(p, 'info', '고치기') + '</div><div id="hrReq-info"></div>'
      : '<div class="gw-sec-title">연락처</div>') + '<dl class="gw-dl">' +
      dd('회사 메일', p.work_email) + dd('개인 메일', p.personal_email) + dd('전화', phone(p.phone)) +
      dd('주소', p.address, { wide: true }) + '</dl>';

    // 전체 보기 = 그 자리에서 * 만 풀림(다시 누르면 가림). 열람 기록은 남기지 않음(11-10 KiKO)
    var ssnHtml = p.has_ssn ? '<span class="gw-ssn-val">' + esc(p.ssn_masked) + '</span>' : '';
    if (p.has_ssn && p.can_reveal_ssn) {
      ssnHtml += '<button type="button" class="btn btn-sm btn-outline-secondary gw-ssn-btn" data-ssn-url="/api/hr/people/' + p.id + '/ssn" data-what="' + esc(p.name) + ' 주민등록번호" aria-label="' + esc(p.name) + ' 주민등록번호 전체 보기">전체 보기</button>';
    }
    h += '<div class="gw-sec-title">개인 정보</div><dl class="gw-dl">' +
      dd('성별', p.gender) + dd('생년월일', p.birth_date) +
      dd('주민등록번호', p.has_ssn ? ssnHtml : '', { html: true, id: 'hrSsnCell' }) + '</dl>';

    if (p.can_edit) {
      h += '<div class="gw-sec-title">그룹웨어</div><dl class="gw-dl">' +
        dd('로그인 메일', p.login_email) +
        dd('로그인 방식', ACCESS_LABEL[p.access] || ACCESS_LABEL.staff) +
        dd('슬랙 ID', p.slack_id) + '</dl>';
      if (p.access === 'field' && !resigned && CFG.mode !== 'me') {
        // 현장 직원 로그인 번호(11-15) — 번호를 만들어 직원 휴대폰 로그인 화면에 넣으면 그 휴대폰은 계속 로그인
        h += '<div class="gw-dev" id="hrDevices"><div class="gw-log">불러오는 중…</div></div>';
      }
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
      var labels = { edit: '수정', create: '추가', doc_upload: '서류 업로드', doc_delete: '서류 삭제', cert_issue: '증명서 발급', self_edit: '본인 수정' };
      h += '<details class="mt-3"><summary class="gw-log">수정 기록 ' + p.access_log.length + '건</summary><ul class="gw-log mt-1">';
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
    bindRequests(p);
    cardEl.querySelectorAll('[data-task-done]').forEach(function (b) {
      b.addEventListener('click', function () {
        b.disabled = true;
        api('/api/hr/tasks/' + b.getAttribute('data-task-done') + '/done', { method: 'POST' })
          .then(function () { return loadCard(p.id); })
          .catch(function (e) { b.disabled = false; window.alert(e.message); });
      });
    });
    var certBtn = document.getElementById('hrCertBtn');
    if (certBtn) certBtn.addEventListener('click', function () { toggleCert(p, certBtn); });
    bindSsnToggles();
    bindPrivate(p);
    if (document.getElementById('hrDevices')) loadDevices(p);
  }

  function tasksHtml(p) {
    var h = '';
    // 입사·퇴사 할 일(11-20) — 경영지원·대표. 슬랙 DM 카드와 같은 일, 여기서 닫으면 슬랙도 닫힘
    var tasks = p.tasks || [];
    if (tasks.length && CFG.mode !== 'me') {
      h += '<div class="gw-sec-title">입사·퇴사 할 일</div><ul class="gw-tasks">';
      tasks.forEach(function (t) {
        var done = t.status === 'done';
        h += '<li class="gw-task' + (done ? ' done' : (/지남/.test(t.due_label || '') ? ' late' : '')) + '">' +
          '<div class="gw-task-main"><b>' + esc(t.title) + '</b> <span class="gw-task-due">' +
          (done ? '완료 · ' + esc(t.done_by || '') + ' · ' + esc((t.done_at || '').slice(0, 10)) : esc(t.due_label || '')) + '</span>' +
          (done ? '' : '<div class="gw-task-guide">' + esc(t.guide || '') + '</div>') + '</div>' +
          (done || !p.can_edit ? '' : '<button type="button" class="btn btn-sm btn-outline-primary" data-task-done="' + t.id + '" aria-label="' + esc(t.title) + ' 완료">완료</button>') +
          '</li>';
      });
      h += '</ul>';
    }
    return h;
  }

  // ── 내 정보 직접 고치기(11-20) — 연락처·비상연락망·서류 올리기는 본인이 바로(부장님: 본인 정보는 본인이) ──
  function reqBtn(p, kind, label) {
    return '<button type="button" class="btn btn-sm btn-outline-secondary" data-req-open="' + kind + '">' + label + '</button>';
  }
  function reqPost(url, body) {
    return api(url, { method: 'POST', body: body });
  }
  function openInfoForm(p, box, btn) {
    box.innerHTML = '<form class="gw-req-form" id="hrReqInfoForm">' +
      '<label>전화<input type="tel" class="form-control form-control-sm" name="phone" maxlength="20" value="' + esc(p.phone || '') + '"></label>' +
      '<label>개인 메일<input type="email" class="form-control form-control-sm" name="personal_email" maxlength="120" value="' + esc(p.personal_email || '') + '"></label>' +
      '<label class="gw-req-wide">주소<input type="text" class="form-control form-control-sm" name="address" maxlength="200" value="' + esc(p.address || '') + '"></label>' +
      '<div class="gw-req-btns"><button type="submit" class="btn btn-sm btn-primary">저장</button><button type="button" class="btn btn-sm btn-outline-secondary" data-req-close>닫기</button>' +
      '<span class="gw-form-error" role="alert"></span></div></form>';
    var f = box.querySelector('form');
    f.elements.phone.focus();
    f.querySelector('[data-req-close]').addEventListener('click', function () { box.innerHTML = ''; btn.focus(); });
    f.addEventListener('submit', function (e) {
      e.preventDefault();
      var err = f.querySelector('.gw-form-error'), sb = f.querySelector('button[type=submit]');
      sb.disabled = true; err.textContent = '';
      reqPost('/api/hr/me/contact', { phone: f.elements.phone.value, personal_email: f.elements.personal_email.value, address: f.elements.address.value })
        .then(function () { return loadCard(p.id); })
        .catch(function (x) { err.textContent = x.message; sb.disabled = false; });
    });
  }
  function famRow(x) {
    return '<div class="gw-req-fam"' + (x.id ? ' data-id="' + x.id + '"' : '') + '>' +
      '<input type="text" class="form-control form-control-sm" placeholder="이름" aria-label="이름" maxlength="30" value="' + esc(x.name || '') + '">' +
      '<input type="text" class="form-control form-control-sm" placeholder="관계 (예: 배우자)" aria-label="관계" maxlength="20" value="' + esc(x.relation || '') + '">' +
      '<input type="tel" class="form-control form-control-sm" placeholder="연락처" aria-label="연락처" maxlength="20" value="' + esc(x.phone || '') + '">' +
      '<button type="button" class="btn btn-sm btn-outline-danger" data-fam-del aria-label="이 줄 빼기">빼기</button></div>';
  }
  function openFamilyForm(p, box, btn) {
    var rows = (p.families || []).map(function (f) { return { id: f.id, name: f.name, relation: f.relation, phone: f.phone }; });
    if (!rows.length) rows.push({});
    box.innerHTML = '<form class="gw-req-form" id="hrReqFamForm"><div class="gw-req-fams">' + rows.map(famRow).join('') + '</div>' +
      '<button type="button" class="btn btn-sm btn-outline-secondary" data-fam-add>+ 사람 추가</button>' +
      '<div class="gw-req-btns"><button type="submit" class="btn btn-sm btn-primary">저장</button><button type="button" class="btn btn-sm btn-outline-secondary" data-req-close>닫기</button>' +
      '<span class="gw-form-error" role="alert"></span></div></form>';
    var f = box.querySelector('form'), list = f.querySelector('.gw-req-fams');
    function bindDel() {
      list.querySelectorAll('[data-fam-del]').forEach(function (b) { b.onclick = function () { b.parentElement.remove(); }; });
    }
    bindDel();
    list.querySelector('input').focus();
    f.querySelector('[data-fam-add]').addEventListener('click', function () {
      list.insertAdjacentHTML('beforeend', famRow({})); bindDel();
      list.lastElementChild.querySelector('input').focus();
    });
    f.querySelector('[data-req-close]').addEventListener('click', function () { box.innerHTML = ''; btn.focus(); });
    f.addEventListener('submit', function (e) {
      e.preventDefault();
      var err = f.querySelector('.gw-form-error'), sb = f.querySelector('button[type=submit]');
      var out = Array.prototype.map.call(list.querySelectorAll('.gw-req-fam'), function (r) {
        var i = r.querySelectorAll('input');
        var o = { name: i[0].value, relation: i[1].value, phone: i[2].value };
        if (r.getAttribute('data-id')) o.id = parseInt(r.getAttribute('data-id'), 10);
        return o;
      });
      sb.disabled = true; err.textContent = '';
      reqPost('/api/hr/me/family', { rows: out })
        .then(function () { return loadCard(p.id); })
        .catch(function (x) { err.textContent = x.message; sb.disabled = false; });
    });
  }
  function bindRequests(p) {
    cardEl.querySelectorAll('[data-req-open]').forEach(function (b) {
      b.addEventListener('click', function () {
        var kind = b.getAttribute('data-req-open'), box = document.getElementById('hrReq-' + kind);
        if (!box) return;
        if (box.innerHTML) { box.innerHTML = ''; return; }
        (kind === 'info' ? openInfoForm : openFamilyForm)(p, box, b);
      });
    });
    var sf = document.getElementById('hrSelfDocForm');
    if (sf) {
      sf.addEventListener('submit', function (e) {
        e.preventDefault();
        var err = document.getElementById('hrSelfDocError'), file = sf.elements.file.files[0];
        if (!file) { err.textContent = '파일을 골라 주세요.'; return; }
        if (file.size > 8 * 1024 * 1024) { err.textContent = '8MB까지 올릴 수 있어요.'; return; }
        var sb = sf.querySelector('button[type=submit]');
        sb.disabled = true; err.textContent = '';
        fetch('/api/hr/me/documents', { method: 'POST', body: new FormData(sf), credentials: 'same-origin' })
          .then(function (r) { return r.json().then(function (j) { if (!r.ok || !j.success) throw new Error(j.message || '올리지 못했어요.'); return j; }); })
          .then(function () { return loadCard(p.id); })
          .catch(function (x) { err.textContent = x.message; sb.disabled = false; });
      });
    }
  }

  // ── 증명서 바로 발급(11-20) — 인사카드 관리자. 퇴직자·창구 요청용(직원 신청은 슬랙 전자결재 [📜 증명서 신청]) ──
  var CERT_PURPOSES = ['금융기관 제출', '관공서 제출', '회사·학교 제출', '기타'];
  function toggleCert(p, btn) {
    var panel = document.getElementById('hrCertPanel');
    if (!panel) return;
    if (!panel.hidden) { panel.hidden = true; btn.setAttribute('aria-expanded', 'false'); return; }
    var resigned = p.status === '퇴사';
    var groupDiff = p.group_hire_date && p.group_hire_date !== p.hire_date;
    panel.innerHTML =
      '<div class="gw-cert-title">증명서 발급 <span class="text-muted">직인이 찍힌 PDF를 만들고 아래 서류에도 남겨요</span></div>' +
      '<div class="gw-cert-grid">' +
      '<label>발급 회사<select class="form-select form-select-sm" id="hrCertCompany">' +
      '<option value="">고르기</option><option value="itg">(주)아이티글로벌</option><option value="group">(주)아이티글로벌그룹</option></select></label>' +
      '<label>종류<select class="form-select form-select-sm" id="hrCertType">' +
      (resigned ? '' : '<option>재직증명서</option>') + '<option>경력증명서</option></select></label>' +
      '<label>용도<select class="form-select form-select-sm" id="hrCertPurpose">' +
      CERT_PURPOSES.map(function (x) { return '<option>' + x + '</option>'; }).join('') + '</select></label>' +
      '<label class="gw-cert-wide"><span>제출처 <span class="text-muted">(선택)</span></span><input type="text" class="form-control form-control-sm" id="hrCertDest" maxlength="60" placeholder="예: ○○은행 대출, ○○구청"></label>' +
      '</div>' +
      '<label class="gw-cert-check"><input type="checkbox" id="hrCertSsn"' + (p.has_ssn ? '' : ' disabled') + '> 주민등록번호 전체 넣기' +
      (p.has_ssn ? '' : ' <span class="text-muted">(인사카드에 없음)</span>') + '</label>' +
      (groupDiff ? '<label class="gw-cert-check"><input type="checkbox" id="hrCertGroup"> 재직 기간을 그룹 입사일(' + esc(p.group_hire_date) + ')부터</label>' : '') +
      '<div class="gw-cert-actions"><button type="button" class="btn btn-sm btn-primary" id="hrCertGo">PDF 만들기</button>' +
      '<button type="button" class="btn btn-sm btn-outline-secondary" id="hrCertClose">닫기</button></div>' +
      '<div class="gw-cert-msg" id="hrCertMsg" role="status"></div>';
    panel.hidden = false;
    btn.setAttribute('aria-expanded', 'true');
    document.getElementById('hrCertCompany').focus();
    document.getElementById('hrCertClose').addEventListener('click', function () {
      panel.hidden = true; btn.setAttribute('aria-expanded', 'false'); btn.focus();
    });
    document.getElementById('hrCertGo').addEventListener('click', function () {
      var go = this, msg = document.getElementById('hrCertMsg');
      var dest = document.getElementById('hrCertDest').value.trim();
      var purpose = document.getElementById('hrCertPurpose').value + (dest ? ' (' + dest + ')' : '');
      var g = document.getElementById('hrCertGroup');
      var company = document.getElementById('hrCertCompany').value;
      if (!company) {
        msg.className = 'gw-cert-msg text-danger';
        msg.textContent = '발급 회사를 골라 주세요.';
        document.getElementById('hrCertCompany').focus();
        return;
      }
      go.disabled = true;
      msg.className = 'gw-cert-msg';
      msg.textContent = '만드는 중…';
      fetch('/api/hr/people/' + p.id + '/certificate', {
        method: 'POST', credentials: 'same-origin', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ company: company, ctype: document.getElementById('hrCertType').value, purpose: purpose,
                               ssn: document.getElementById('hrCertSsn').checked, group_date: !!(g && g.checked) })
      }).then(function (r) {
        if (!r.ok) return r.json().catch(function () { return {}; }).then(function (j) { throw new Error(j.message || ('발급하지 못했어요 (' + r.status + ')')); });
        var no = r.headers.get('X-Doc-No') || '';
        var cd = r.headers.get('Content-Disposition') || '';
        var m = /filename\*=UTF-8''([^;]+)/i.exec(cd);
        var name = m ? decodeURIComponent(m[1]) : '증명서.pdf';
        return r.blob().then(function (b) {
          var a = document.createElement('a');
          a.href = URL.createObjectURL(b); a.download = name;
          document.body.appendChild(a); a.click(); a.remove();
          setTimeout(function () { URL.revokeObjectURL(a.href); }, 10000);
          return no;
        });
      }).then(function (no) {
        return loadCard(p.id).then(function () {
          var again = document.getElementById('hrCertBtn');
          if (again) again.focus();
          var sr = document.getElementById('hrCardSr');
          if (sr) sr.textContent = '증명서를 발급했어요. 발급번호 ' + no + ', 아래 서류에 남겼어요.';
        });
      }).catch(function (e) {
        go.disabled = false;
        msg.className = 'gw-cert-msg text-danger';
        msg.textContent = e.message;
      });
    });
  }

  // 창을 JS 로 열면 닫힌 뒤 포커스가 여는 단추로 돌아가지 않음(부트스트랩은 data-bs 로 열 때만) → 직접 되돌림
  function returnFocusOnHide(modalEl) {
    var opener = document.activeElement;
    modalEl.addEventListener('hidden.bs.modal', function h() {
      modalEl.removeEventListener('hidden.bs.modal', h);
      if (opener && opener.focus && document.contains(opener)) opener.focus();
    });
  }

  // ── 현장 직원 로그인 번호 ─────────────────────────────
  var codeTimer = null;

  function renderDevices(p, j, code) {
    var box = document.getElementById('hrDevices');
    if (!box || !state.current || state.current.id !== p.id) return;
    var h = '<div class="gw-dev-head"><strong>로그인된 휴대폰</strong>' +
      '<button type="button" class="btn btn-sm btn-primary" id="hrDevCodeBtn"><i class="fas fa-mobile-alt me-1" aria-hidden="true"></i>로그인 번호 만들기</button></div>';
    if (code) {
      h += '<div class="visually-hidden" role="status" id="hrDevSr">로그인 번호 ' + esc(code.code.split('').join(' ')) + ' 을 만들었어요.</div>' +
        '<div class="gw-dev-code"><div class="gw-dev-code-num" id="hrDevNum">' + esc(code.code.replace(/(\d{3})(\d{3})/, '$1 $2')) + '</div>' +
        '<div class="gw-dev-code-help" id="hrDevHelp">' + esc(p.name) + ' 님 휴대폰 로그인 화면의 <strong>「번호로 로그인」</strong>에 이 번호를 넣어 주세요. ' +
        '<span id="hrDevLeft">' + code.minutes + '분</span> 동안 한 번만 쓸 수 있어요.</div></div>';
    }
    var list = j.devices || [];
    if (list.length) {
      h += '<ul class="gw-dev-list">' + list.map(function (d) {
        return '<li><div><strong>' + esc(d.label || '기기') + '</strong><div class="gw-log">등록 ' + esc((d.created_at || '').slice(0, 10)) +
          ' · 마지막 사용 ' + esc((d.last_used_at || '').slice(0, 16)) + '</div></div>' +
          '<button type="button" class="btn btn-sm btn-outline-danger" data-dev-revoke="' + d.id + '">로그아웃</button></li>';
      }).join('') + '</ul>';
    } else {
      h += '<div class="gw-log">아직 로그인된 휴대폰이 없어요.</div>';
    }
    h += '<div class="gw-log mt-1">한 번 로그인한 휴대폰은 계속 열려요. ' + (j.idle_days || 90) + '일 동안 안 쓰거나 퇴사·로그인 방식 변경이 있으면, 또는 여기나 그 휴대폰에서 로그아웃하면 끝나요.</div>';
    box.innerHTML = h;

    document.getElementById('hrDevCodeBtn').addEventListener('click', function () {
      var b = this;
      b.disabled = true;
      api('/api/hr/people/' + p.id + '/device-code', { method: 'POST' }).then(function (c) {
        renderDevices(p, j, c);
        startCountdown(c.expires_at);
      }).catch(function (e) { b.disabled = false; window.alert(e.message); });
    });
    box.querySelectorAll('[data-dev-revoke]').forEach(function (b) {
      b.addEventListener('click', function () {
        if (!window.confirm('이 휴대폰을 로그아웃할까요? 다시 쓰려면 새 번호로 로그인해야 해요.')) return;
        b.disabled = true;
        api('/api/hr/people/' + p.id + '/devices/' + b.getAttribute('data-dev-revoke') + '/revoke', { method: 'POST' })
          .then(function (r) { j.devices = r.devices; renderDevices(p, j, null); })
          .catch(function (e) { b.disabled = false; window.alert(e.message); });
      });
    });
  }

  function startCountdown(expiresAt) {
    clearInterval(codeTimer);
    var end = new Date(expiresAt.replace(' ', 'T')).getTime();
    codeTimer = setInterval(function () {
      var el = document.getElementById('hrDevLeft');
      if (!el) { clearInterval(codeTimer); return; }
      var s = Math.max(0, Math.round((end - Date.now()) / 1000));
      if (s) { el.textContent = Math.floor(s / 60) + '분 ' + (s % 60) + '초'; return; }
      clearInterval(codeTimer);
      // 만료: 안내 문장을 통째로 바꿔 ' 동안 한 번만 쓸 수 있어요.'가 뒤에 붙지 않게, 지난 번호는 흐리게
      var msg = '시간이 지나 이 번호는 쓸 수 없어요. 「로그인 번호 만들기」로 새 번호를 만들어 주세요.';
      var help = document.getElementById('hrDevHelp');
      if (help) help.textContent = msg; else el.textContent = msg;
      var num = document.getElementById('hrDevNum');
      if (num) { num.style.opacity = '0.4'; num.style.textDecoration = 'line-through'; }
      var sr = document.getElementById('hrDevSr');
      if (sr) sr.textContent = msg;
    }, 1000);
  }

  function loadDevices(p) {
    api('/api/hr/people/' + p.id + '/devices').then(function (j) {
      if (!j.can_register) { var b = document.getElementById('hrDevices'); if (b) b.remove(); return; }
      renderDevices(p, j, null);
    }).catch(function (e) {
      var b = document.getElementById('hrDevices');
      if (b) b.innerHTML = '<div class="gw-log">' + esc(e.message) + '</div>';
    });
  }

  // ── 경력·학력·비상연락망·서류(본인·인사카드 관리자만 서버가 보내 줌) ─────────────
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
    // 가족 = 비상연락망(11-19 부장님: 플렉스에 연락용으로 올린 것 — 공제 대상 아님). 공제는 연말정산 화면에서 따로 받는다.
    // 주민번호는 플렉스에 넣어 둔 사람이 있을 때만 칸을 보인다.
    var selfFam = CFG.mode === 'me' && p.can_self_edit;
    if (selfFam && !(p.families && p.families.length)) {
      h += '<div class="gw-sec-head"><div class="gw-sec-title">비상연락망</div>' + reqBtn(p, 'family', '등록하기') + '</div>' +
        '<div id="hrReq-family"></div><div class="gw-log">등록된 비상연락망이 없어요. 급할 때 연락할 가족을 알려 주세요.</div>';
    }
    if (p.families && p.families.length) {
      var famSsn = p.families.some(function (f) { return f.has_ssn; });
      h += (selfFam ? '<div class="gw-sec-head"><div class="gw-sec-title">비상연락망</div>' + reqBtn(p, 'family', '고치기') + '</div><div id="hrReq-family"></div>'
                    : '<div class="gw-sec-title">비상연락망</div>') +
        '<p class="text-muted small mb-2">급할 때 연락할 가족이에요. 연말정산 부양가족 공제는 1월 연말정산 화면에서 따로 받아요.</p>' +
        table(['이름', '관계', '연락처'].concat(famSsn ? ['주민등록번호'] : []),
        p.families.map(function (f) {
          var ssn = f.has_ssn ? '<span class="gw-ssn-val">' + esc(f.ssn_masked) + '</span>' + (p.can_reveal_ssn ? ' <button type="button" class="btn btn-sm btn-outline-secondary gw-ssn-btn" data-ssn-url="/api/hr/families/' + f.id + '/ssn" data-what="' + esc(f.name) + ' 주민등록번호" aria-label="' + esc(f.name) + ' 주민등록번호 전체 보기">전체 보기</button>' : '') : '';
          return '<tr><td>' + esc(f.name) + '</td><td>' + esc(f.relation) + '</td><td>' + esc(phone(f.phone)) + '</td>' +
            (famSsn ? '<td data-fam-cell="' + f.id + '">' + ssn + '</td>' : '') + '</tr>';
        }));
    }
    var docs = p.documents || [];
    if (docs.length || p.can_edit) {
      h += '<div class="gw-sec-title">서류</div>';
      if (docs.length) {
        h += table(['종류', '파일', '크기', '업로드일', ''], docs.map(function (d) {
          return '<tr><td>' + esc(d.doc_type || '기타') + '</td><td class="gw-doc-name">' + esc(d.file_name) + '</td><td>' + kb(d.size) +
            '</td><td>' + (d.uploaded_by === 'flex-import' ? '<span class="text-muted">플렉스에서 옮김</span>' : esc((d.created_at || '').slice(0, 10))) +
            '</td><td class="gw-doc-act">' +
            '<a class="btn btn-sm btn-outline-secondary" href="/api/hr/documents/' + d.id + '?view=1" target="_blank" rel="noopener" aria-label="' + esc(d.file_name) + ' 보기">보기</a> ' +
            '<a class="btn btn-sm btn-outline-secondary" href="/api/hr/documents/' + d.id + '" aria-label="' + esc(d.file_name) + ' 다운로드">다운로드</a>' +
            (p.can_edit && CFG.mode !== 'me' ? ' <button type="button" class="btn btn-sm btn-outline-danger" data-doc-del="' + d.id + '" aria-label="' + esc(d.file_name) + ' 삭제">삭제</button>' : '') +
            '</td></tr>';
        }));
      } else {
        h += '<div class="gw-log">업로드된 서류가 없습니다.</div>';
      }
      if (p.can_edit && CFG.mode !== 'me') {
        h += '<form class="gw-doc-upload" id="hrDocForm">' +
          '<select class="form-select form-select-sm" name="doc_type" aria-label="서류 종류">' +
          (p.doc_types || []).map(function (t) { return '<option>' + esc(t) + '</option>'; }).join('') + '</select>' +
          '<input type="file" class="form-control form-control-sm" name="file" aria-label="파일" required>' +
          '<button type="submit" class="btn btn-sm btn-outline-primary">업로드</button>' +
          '<span class="gw-form-error" id="hrDocError" role="alert"></span></form>';
      }
    }
    if (CFG.mode === 'me' && p.self_doc_types) {
      if (!docs.length && !p.can_edit) h += '<div class="gw-sec-title">서류</div>';
            h += '<form class="gw-doc-upload" id="hrSelfDocForm">' +
        '<select class="form-select form-select-sm" name="doc_type" aria-label="올릴 서류 종류">' +
        p.self_doc_types.map(function (t) { return '<option>' + esc(t) + '</option>'; }).join('') + '</select>' +
        '<input type="file" class="form-control form-control-sm" name="file" aria-label="올릴 파일" required>' +
        '<button type="submit" class="btn btn-sm btn-outline-primary">서류 올리기</button>' +
        '<span class="gw-form-error" id="hrSelfDocError" role="alert"></span></form>' +
        '<div class="gw-log">올리면 바로 위 서류 목록에 들어가요. 잘못 올렸으면 경영지원에 지워 달라고 해 주세요.</div>';
    }
    return h;
  }

  function bindPrivate(p) {
    cardEl.querySelectorAll('[data-doc-del]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        if (!window.confirm('이 서류를 지울까요? 지우면 되돌릴 수 없어요.')) return;
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
        if (file.size > 8 * 1024 * 1024) { err.textContent = '8MB까지 업로드할 수 있습니다.'; return; }
        var fd = new FormData(form);
        var btn = form.querySelector('button[type=submit]');
        btn.disabled = true;
        err.textContent = '';
        fetch('/api/hr/people/' + p.id + '/documents', { method: 'POST', body: fd, credentials: 'same-origin' })
          .then(function (r) { return r.json().then(function (j) { if (!r.ok || !j.success) throw new Error(j.message || '업로드하지 못했습니다.'); return j; }); })
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

  // 주민등록번호 전체 보기 ↔ 숨기기 — 본인·가족 공용. 한 번 받은 번호는 이 화면에서만 기억
  function bindSsnToggles() {
    cardEl.querySelectorAll('[data-ssn-url]').forEach(function (btn) {
      var val = btn.parentElement.querySelector('.gw-ssn-val');
      if (!val) return;
      var masked = val.textContent, full = null;
      btn.addEventListener('click', function () {
        var what = btn.getAttribute('data-what') || '주민등록번호';
        if (btn.getAttribute('aria-pressed') === 'true') {
          val.textContent = masked;
          btn.textContent = '전체 보기';
          btn.setAttribute('aria-label', what + ' 전체 보기');
          btn.setAttribute('aria-pressed', 'false');
          return;
        }
        function show() {
          val.textContent = full;
          btn.textContent = '숨기기';
          btn.setAttribute('aria-label', what + ' 숨기기');
          btn.setAttribute('aria-pressed', 'true');
        }
        if (full) { show(); return; }
        btn.disabled = true;
        api(btn.getAttribute('data-ssn-url'), { method: 'POST' }).then(function (j) {
          full = j.ssn;
          btn.disabled = false;
          show();
        }).catch(function (e) { btn.disabled = false; window.alert(e.message); });
      });
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
    // 다시 그려도 키보드 포커스가 같은 사람 줄에 남게
    var ae = document.activeElement;
    var focusId = ae && listEl.contains(ae) && ae.getAttribute('data-id');
    // 목록은 보통 단추 목록(listbox·option 역할은 화살표 키 조작까지 약속하므로 뺌) — 고른 사람은 aria-current 로
    listEl.innerHTML = rows.map(function (p) {
      var sub = [p.department, p.job_rank].filter(Boolean).join(' · ');
      return '<li><button type="button" class="gw-list-item' + (p.status === '퇴사' ? ' resigned' : '') +
        '" data-id="' + p.id + '"' + (p.id === state.selectedId ? ' aria-current="true"' : '') + '>' +
        '<span class="gw-avatar" aria-hidden="true">' + initial(p.name) + '</span>' +
        '<span><span class="gw-list-name">' + esc(p.name) + '</span>' +
        (p.status !== '재직' ? ' ' + statusBadge(p.status) : '') +
        (p.status !== '퇴사' ? fieldBadge(p) : '') +
        '<br><span class="gw-list-sub">' + esc(sub || '-') + '</span></span></button></li>';
    }).join('') || '<li class="gw-empty">맞는 구성원이 없습니다.</li>';
    if (focusId) {
      var back = listEl.querySelector('.gw-list-item[data-id="' + focusId + '"]');
      if (back) back.focus({ preventScroll: true });
    }
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
      if (o.value === 'field') o.textContent = '로그인 번호·슬랙 휴가 앱 링크 (현장·미화)';
    });
    var accLabel = formEl.querySelector('label[for="f_access"]');
    if (accLabel) accLabel.textContent = '로그인 방식';
    var hint = formEl.elements.login_email && formEl.elements.login_email.parentNode.querySelector('.form-text');
    if (hint) hint.textContent = '회사 구글 계정으로 들어오는 메일이에요. 현장·미화 직원(로그인 번호·슬랙 링크)은 비워 둬도 돼요.';
  }

  function modal() {
    return window.bootstrap ? window.bootstrap.Modal.getOrCreateInstance(modalEl) : null;
  }

  var dirty = false;

  function openEdit(p) {
    if (!formEl) return;
    editingId = p ? p.id : null;
    dirty = false;
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
    fieldErr(formEl.elements.name, '');
    fieldErr(formEl.elements.ssn, '');
    var m = modal();
    if (m) { returnFocusOnHide(modalEl); m.show(); }
  }

  // 칸 오류는 그 칸 바로 아래 — 창 맨 아래에만 나와 어느 칸인지 찾아야 했음
  function fieldErr(input, msg) {
    var id = input.id + '_err';
    var box = document.getElementById(id);
    if (!box) {
      box = document.createElement('div');
      box.id = id;
      box.className = 'gw-form-error gw-field-err';
      input.parentNode.appendChild(box);
    }
    box.textContent = msg || '';
    if (msg) {
      input.setAttribute('aria-invalid', 'true');
      input.setAttribute('aria-describedby', id);
      input.focus();
    } else {
      input.removeAttribute('aria-invalid');
      input.removeAttribute('aria-describedby');
    }
  }

  if (formEl) {
    ['name', 'ssn'].forEach(function (k) {
      var el = formEl.elements[k];
      if (el) el.addEventListener('input', function () { if (el.getAttribute('aria-invalid')) fieldErr(el, ''); });
    });
    formEl.addEventListener('submit', function (e) {
      e.preventDefault();
      errEl.textContent = '';
      fieldErr(formEl.elements.name, '');
      fieldErr(formEl.elements.ssn, '');
      var body = {};
      FIELDS.forEach(function (k) { if (formEl.elements[k]) body[k] = formEl.elements[k].value; });
      if (!body.name.trim()) { fieldErr(formEl.elements.name, '이름을 입력하세요.'); return; }
      var ssn = formEl.elements.ssn.value.trim();
      if (ssn) {
        if (ssn.replace(/\D/g, '').length !== 13) { fieldErr(formEl.elements.ssn, '주민등록번호는 숫자 13자리여야 합니다.'); return; }
        body.ssn = ssn;
      }
      var saveBtn = document.getElementById('hrEditSave');
      saveBtn.disabled = true;
      saveBtn.textContent = '저장하는 중…';
      var req = editingId
        ? api('/api/hr/people/' + editingId, { method: 'PUT', body: body })
        : api('/api/hr/people', { method: 'POST', body: body });
      req.then(function (j) {
        var m = modal();
        dirty = false;   // 저장했으니 닫기 확인 없이
        if (m) m.hide();
        state.selectedId = j.person.id;
        return loadList(j.person.id);
      }).catch(function (err) {
        errEl.textContent = err.message;
      }).then(function () { saveBtn.disabled = false; saveBtn.textContent = '저장'; });
    });
    formEl.elements.ssn.addEventListener('input', function () { errEl.textContent = ''; });
    // 고친 내용이 있는데 닫으면 확인 — 실수로 닫아 입력한 걸 잃지 않게
    formEl.addEventListener('input', function () { dirty = true; });
    formEl.addEventListener('change', function () { dirty = true; });
    modalEl.addEventListener('hide.bs.modal', function (e) {
      if (dirty && !window.confirm('저장하지 않은 내용이 있어요.\n닫으면 고친 것이 사라져요. 닫을까요?')) { e.preventDefault(); return; }
      dirty = false;
    });
  }

  // ── 시작 ───────────────────────────────────────────
  if (CFG.mode === 'me' && CFG.personId) {
    loadCard(CFG.personId);
    return;
  }

  if (CFG.mode === 'people') {
    listEl.addEventListener('click', function (e) {
      var btn = e.target.closest('.gw-list-item');
      if (!btn) return;
      select(parseInt(btn.getAttribute('data-id'), 10));
      // 휴대폰은 목록 아래에 카드가 나오므로 카드로 내려 줌(고른 게 화면 밖이라 안 바뀐 것처럼 보였다, 11-15 점검)
      if (window.matchMedia('(max-width: 767.98px)').matches) {
        var card = document.getElementById('hrCard');
        if (card) card.scrollIntoView({ block: 'start', behavior: 'smooth' });
      }
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
