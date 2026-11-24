/* 그룹웨어 설정 — 시범 명단·전체 켜기·인사카드 관리자·휴가 현황 열람자 */
(function () {
  'use strict';

  // 로그인이 끝나 로그인 화면(200 HTML)으로 돌려보내졌거나 401 이면 멈춘다 — 예전엔 그 화면을 {} 로 읽어
  // 처리가 '된' 것처럼 보였음(11-24 조용한 실패 점검 4번, 결재함 gw-approval.js 와 같은 판정)
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  function jsonOrEnd(r) {
    var ct = r.headers.get('content-type') || '';
    if (r.status === 401 || r.redirected || (r.ok && ct.indexOf('application/json') === -1)) {
      return Promise.reject(new Error(SESSION_ENDED));
    }
    return r.json().catch(function () { return {}; });
  }

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
      return jsonOrEnd(r).then(function (j) {
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

  // 빼기는 확인 뒤 — 누르는 즉시 권한이 사라졌음
  function render() {
    chips(document.getElementById('gwPilotList'), data.pilots.map(function (p) { return p.email; }), function (email) {
      if (!window.confirm(email + ' — 시범 명단에서 뺄까요?\n바로 예전 화면으로 돌아갑니다.')) return;
      setPilot(email, false);
    });
    document.getElementById('gwOpenAll').checked = !!data.open_all;
    chips(document.getElementById('gwHrAdmins'), data.hr_admins, function (email) {
      if (!window.confirm(email + ' — 인사카드 관리자에서 뺄까요?\n바로 전 직원 인사카드를 볼 수 없게 됩니다.')) return;
      saveRole('hr_admins', data.hr_admins.filter(function (x) { return x !== email; }));
    }, true);
    chips(document.getElementById('gwLeaveViewers'), data.leave_viewers, function (email) {
      if (!window.confirm(email + ' — 휴가 현황 열람자에서 뺄까요?')) return;
      saveRole('leave_viewers', data.leave_viewers.filter(function (x) { return x !== email; }));
    });
  }

  // 오류는 고친 칸 바로 아래에 — 인사카드 관리자 오류가 아래 '휴가 현황' 칸에 나오던 것
  var ROLE_ERR = { hr_admins: 'gwHrAdminError', leave_viewers: 'gwRoleError' };

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

  // 성공하면 true, 실패하면 false 로 끝남 — [추가]는 성공했을 때만 입력칸을 비움(실패하면 친 메일이 그대로 남게)
  function setPilot(email, on) {
    showErr('gwPilotError', '');
    return api('/api/gw/pilots', { email: email, on: on }).then(function (j) {
      data.pilots = j.pilots;
      data.open_all = j.open_all;
      render();
      return true;
    }).catch(function (e) { showErr('gwPilotError', e.message); return false; });
  }

  function saveRole(key, emails) {
    showErr(ROLE_ERR[key], '');
    return api('/api/gw/roles', { key: key, emails: emails }).then(function (j) {
      data[key] = j[key];
      render();
      return true;
    }).catch(function (e) { showErr(ROLE_ERR[key], e.message); return false; });
  }

  // [추가] 처리 중엔 단추를 잠그고, 성공한 뒤에만 칸을 비움
  function addFlow(btn, input, req) {
    btn.disabled = true;
    req.then(function (ok) {
      btn.disabled = false;
      if (ok) input.value = '';
      else input.focus();
    });
  }

  function valid(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }

  document.getElementById('gwPilotAdd').addEventListener('click', function () {
    var input = document.getElementById('gwPilotInput');
    var email = input.value.trim().toLowerCase();
    if (!valid(email)) { showErr('gwPilotError', '이메일을 정확히 입력하세요.'); input.focus(); return; }
    addFlow(this, input, setPilot(email, true));
  });
  document.getElementById('gwPilotInput').addEventListener('input', function () { showErr('gwPilotError', ''); });

  document.querySelectorAll('button[data-add]').forEach(function (btn) {
    btn.addEventListener('click', function () {
      var key = btn.getAttribute('data-add');
      var input = document.getElementById(btn.getAttribute('data-input'));
      var email = input.value.trim().toLowerCase();
      if (!valid(email)) { showErr(ROLE_ERR[key], '이메일을 정확히 입력하세요.'); input.focus(); return; }
      if (data[key].indexOf(email) !== -1) { showErr(ROLE_ERR[key], '이미 들어 있어요.'); input.focus(); return; }
      addFlow(btn, input, saveRole(key, data[key].concat([email])));
    });
  });

  // 입력칸에서 Enter 로도 추가
  [['gwPilotInput', 'gwPilotAdd'], ['gwHrAdminInput', null], ['gwLeaveViewerInput', null]].forEach(function (pair) {
    var input = document.getElementById(pair[0]);
    if (!input) return;
    input.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter' || e.isComposing) return;
      e.preventDefault();
      var b = pair[1] ? document.getElementById(pair[1]) : document.querySelector('button[data-input="' + pair[0] + '"]');
      if (b) b.click();
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
