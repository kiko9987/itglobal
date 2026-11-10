/* 그룹웨어 설정 — 시범 명단·전체 켜기·인사카드 관리자·휴가 현황 열람자 */
(function () {
  'use strict';

  var data = { pilots: [], open_all: false, hr_admins: [], leave_viewers: [] };

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function api(url, body) {
    return fetch(url, {
      method: body ? 'POST' : 'GET',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: body ? JSON.stringify(body) : undefined
    }).then(function (r) {
      return r.json().catch(function () { return {}; }).then(function (j) {
        if (!r.ok || j.success === false) throw new Error(j.message || ('처리하지 못했습니다 (' + r.status + ')'));
        return j;
      });
    });
  }

  function chips(el, emails, onRemove, lockLast) {
    el.innerHTML = emails.map(function (e) {
      var lock = lockLast && emails.length === 1;
      return '<li class="gw-chip">' + esc(e) + (lock ? '' :
        '<button type="button" data-email="' + esc(e) + '" aria-label="' + esc(e) + ' 빼기">×</button>') + '</li>';
    }).join('') || '<li class="gw-panel-desc">아무도 없음</li>';
    el.querySelectorAll('button[data-email]').forEach(function (b) {
      b.addEventListener('click', function () { onRemove(b.getAttribute('data-email')); });
    });
  }

  function render() {
    chips(document.getElementById('gwPilotList'), data.pilots.map(function (p) { return p.email; }), function (email) {
      setPilot(email, false);
    });
    document.getElementById('gwOpenAll').checked = !!data.open_all;
    chips(document.getElementById('gwHrAdmins'), data.hr_admins, function (email) {
      saveRole('hr_admins', data.hr_admins.filter(function (x) { return x !== email; }));
    }, true);
    chips(document.getElementById('gwLeaveViewers'), data.leave_viewers, function (email) {
      saveRole('leave_viewers', data.leave_viewers.filter(function (x) { return x !== email; }));
    });
  }

  function showErr(id, msg) {
    var el = document.getElementById(id);
    if (el) el.textContent = msg || '';
  }

  function load() {
    api('/api/gw/roles').then(function (j) {
      data = j;
      render();
    }).catch(function (e) { showErr('gwPilotError', e.message); });
  }

  function setPilot(email, on) {
    showErr('gwPilotError', '');
    api('/api/gw/pilots', { email: email, on: on }).then(function (j) {
      data.pilots = j.pilots;
      data.open_all = j.open_all;
      render();
    }).catch(function (e) { showErr('gwPilotError', e.message); });
  }

  function saveRole(key, emails) {
    showErr('gwRoleError', '');
    api('/api/gw/roles', { key: key, emails: emails }).then(function (j) {
      data[key] = j[key];
      render();
    }).catch(function (e) { showErr('gwRoleError', e.message); });
  }

  function valid(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }

  document.getElementById('gwPilotAdd').addEventListener('click', function () {
    var input = document.getElementById('gwPilotInput');
    var email = input.value.trim().toLowerCase();
    if (!valid(email)) { showErr('gwPilotError', '이메일을 정확히 입력하세요.'); input.focus(); return; }
    input.value = '';
    setPilot(email, true);
  });
  document.getElementById('gwPilotInput').addEventListener('input', function () { showErr('gwPilotError', ''); });

  document.querySelectorAll('button[data-add]').forEach(function (btn) {
    btn.addEventListener('click', function () {
      var key = btn.getAttribute('data-add');
      var input = document.getElementById(btn.getAttribute('data-input'));
      var email = input.value.trim().toLowerCase();
      if (!valid(email)) { showErr('gwRoleError', '이메일을 정확히 입력하세요.'); input.focus(); return; }
      input.value = '';
      if (data[key].indexOf(email) === -1) saveRole(key, data[key].concat([email]));
    });
  });

  document.getElementById('gwOpenAll').addEventListener('change', function (e) {
    var on = e.target.checked;
    var msg = on ? '전 직원에게 새 화면을 엽니다(정식 오픈). 계속할까요?' : '정식 오픈을 끄고 시범 명단만 새 화면을 보게 합니다. 계속할까요?';
    if (!window.confirm(msg)) { e.target.checked = !on; return; }
    api('/api/gw/pilots', { open_all: on }).then(function (j) {
      data.open_all = j.open_all;
      render();
    }).catch(function (err) { e.target.checked = !on; showErr('gwPilotError', err.message); });
  });

  // 이메일 자동완성(웹 계정 목록)
  fetch('/api/users', { credentials: 'same-origin' }).then(function (r) { return r.json(); }).then(function (j) {
    var list = (j && (j.users || j.data)) || [];
    document.getElementById('gwUserEmails').innerHTML = list.filter(function (u) { return u && u.email && u.is_active !== false; })
      .map(function (u) { return '<option value="' + esc(u.email) + '">' + esc(u.name || '') + '</option>'; }).join('');
  }).catch(function () { /* 자동완성만 없음 */ });

  load();
})();
