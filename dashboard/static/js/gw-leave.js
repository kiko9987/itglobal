/* 그룹웨어 휴가 화면 — 내 휴가(본인 신청·취소, 11-15) / 전사 현황(달력·직원별, 보기만).
   신청·취소는 슬랙 휴가 앱과 같은 장부·같은 알림(서버 leave_web).
   window.GW_LEAVE = { mode: 'me' | 'team' } */
(function () {
  'use strict';

  var CFG = window.GW_LEAVE || {};
  var WD = ['월', '화', '수', '목', '금', '토', '일'];

  // 창을 JS 로 열면 닫힌 뒤 포커스가 여는 단추로 돌아가지 않음(부트스트랩은 data-bs 로 열 때만) → 직접 되돌림
  function returnFocusOnHide(modalEl) {
    var opener = document.activeElement;
    modalEl.addEventListener('hidden.bs.modal', function h() {
      modalEl.removeEventListener('hidden.bs.modal', h);
      if (opener && opener.focus && document.contains(opener)) opener.focus();
    });
  }

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';

  function api(url, opts) {
    opts = opts || {};
    var method = opts.method || 'GET';
    var init = { method: method, credentials: 'same-origin' };
    // 본문 있을 때만 JSON 머리글(빈 본문 + JSON 머리글 = 운영 400). 본문 없는 POST(휴가 취소)는 로그인이 끝나면
    // 로그인 화면으로 돌려보내지므로 아래 redirected·HTML 판정으로 잡는다(예전엔 {} 로 읽혀 취소가 "됐다"고 떴음, 11-16 점검)
    if (opts.body) { init.headers = { 'Content-Type': 'application/json' }; init.body = JSON.stringify(opts.body); }
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      // 로그인 화면으로 돌려보내졌거나(redirected·200 HTML) 401 이면 로그인이 끝난 것
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) {
        var se = new Error(SESSION_ENDED);
        se.fields = {};
        se.sessionEnded = true;
        throw se;
      }
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          j = j || {};
          var fallback = method === 'GET'
            ? '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'
            : '처리하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')';
          var err = new Error(j.message || j.error || fallback);
          err.fields = j.errors || {};
          throw err;
        }
        return j;
      });
    });
  }

  function days(x) {
    if (x == null) return '-';
    var n = Math.round(Number(x) * 10) / 10;
    return (n % 1 === 0 ? n.toFixed(0) : n.toFixed(1)) + '일';
  }

  function todayIso() {
    var t = new Date();
    return t.getFullYear() + '-' + (t.getMonth() < 9 ? '0' : '') + (t.getMonth() + 1) + '-' + (t.getDate() < 10 ? '0' : '') + t.getDate();
  }

  function md(s) {
    if (!s) return '';
    var p = s.split('-');
    var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    return Number(p[1]) + '월 ' + Number(p[2]) + '일 (' + WD[(d.getDay() + 6) % 7] + ')';
  }

  function period(e, withYear) {
    var y = withYear ? e.start.slice(0, 4) + '년 ' : '';
    if (e.half) return y + md(e.start) + ' ' + e.half;
    if (!e.end || e.end === e.start) return y + md(e.start);
    return y + md(e.start) + ' ~ ' + md(e.end);
  }

  function statusBadge(s) {
    var cls = { '승인': 'gw-badge-active', '대기': 'gw-badge-leave', '반려': 'gw-badge-reject', '취소': 'gw-badge-resigned' }[s] || 'gw-badge-resigned';
    return '<span class="gw-badge ' + cls + '">' + esc(s === '대기' ? '승인 대기' : s) + '</span>';
  }

  // 본인 취소 단추 — 서버가 정한 것(cancel: 'all' | 'rest')만. 문구는 슬랙 카드와 같다
  function cancelBtn(e) {
    if (!e.cancel) return '';
    return '<button type="button" class="btn btn-outline-danger btn-sm gw-lv-cancel" data-lv-cancel="' + esc(e.id) + '"' +
      ' data-what="' + esc(period(e) + ' · ' + e.type) + '" data-confirm="' + esc(e.cancel_confirm) + '"' +
      ' data-mode="' + esc(e.cancel) + '">' + esc(e.cancel_label) + '</button>';
  }

  // ── 한 사람 휴가(내 휴가·직원별 상세 공용) ───────────────────────
  function renderSummary(L, opts) {
    opts = opts || {};
    var h = '';
    if (!L.linked) {
      h += '<div class="gw-note mb-2">휴가 장부에 아직 연결되지 않아 남은 연차를 계산하지 못했습니다. 지난 기록만 보입니다.</div>';
    }
    h += '<div class="gw-stats">' +
      '<div class="gw-stat gw-stat-main"><div class="k">남은 연차</div><div class="v">' + (L.balance == null ? '-' : days(L.balance)) + '</div></div>' +
      '<div class="gw-stat"><div class="k">' + L.year + '년 쓴 연차</div><div class="v">' + days(L.annual_used) + '</div></div>' +
      '<div class="gw-stat"><div class="k">' + L.year + '년 병가</div><div class="v">' + days(L.sick_used) + ' <small>/ ' + L.sick_limit + '일</small></div></div>' +
      '<div class="gw-stat"><div class="k">승인 대기</div><div class="v">' + L.pending + '건</div></div>' +
      '</div>';
    // 남은 연차는 승인 대기 중인 것까지 미리 뺀 값(leave.balance) — 신청하자마자 줄어 "왜 벌써?" 묻지 않게(11-15 점검)
    if (L.pending && L.balance != null) {
      h += '<div class="gw-lv-note">남은 연차는 승인 대기 중인 휴가까지 미리 뺀 날수예요.</div>';
    }

    if (L.grants && L.grants.length) {
      h += '<div class="gw-sec-title">언제까지 써야 하나요</div><ul class="gw-plain-list">';
      L.grants.forEach(function (g) {
        h += '<li><strong>' + days(g.days) + '</strong> — ' + esc(g.valid_to.slice(0, 4)) + '년 ' + esc(md(g.valid_to)) + '까지</li>';
      });
      h += '</ul>';
    }

    if (L.upcoming && L.upcoming.length) {
      h += '<div class="gw-sec-title">다가오는 휴가</div>';
      if (opts.mine) {
        // 내 휴가: 한 줄씩 카드 + 오른쪽 취소 단추(휴대폰에서 누르기 쉽게)
        h += '<ul class="gw-lv-upcoming">';
        L.upcoming.forEach(function (e) {
          h += '<li><div class="gw-lv-up-main"><div class="gw-lv-up-when">' + esc(period(e)) + '</div>' +
            '<div class="gw-lv-up-sub">' + esc(e.type) + ' · ' + days(e.days) + ' ' + statusBadge(e.status) + '</div>' +
            // 플렉스에서 옮겨 온 휴가는 여기서 취소할 수 없다 — 대기뿐 아니라 승인된 것도(끝나지 않았으면) 단추 없이 비면
            // 막막하니 할 일을 적어 둠(11-16 점검: 승인된 플렉스 휴가에 안내가 없던 것)
            (!e.cancel && e.source === 'flex' && (e.status === '대기' || e.status === '승인') && (e.end || e.start) >= todayIso()
              ? '<div class="gw-lv-up-why">플렉스에서 신청한 휴가예요. 취소는 경영지원에 말씀해 주세요.</div>' : '') +
            '</div>' + cancelBtn(e) + '</li>';
        });
        h += '</ul>';
      } else {
        h += '<ul class="gw-plain-list">';
        L.upcoming.forEach(function (e) {
          h += '<li>' + esc(period(e)) + ' · ' + esc(e.type) + ' · ' + days(e.days) + ' ' + statusBadge(e.status) + '</li>';
        });
        h += '</ul>';
      }
    }

    if (opts.applyNote) {
      h += '<div class="gw-callout">' + opts.applyNote + '</div>';
    } else if (!opts.hideApplyNote) {
      h += '<div class="gw-callout"><i class="fab fa-slack me-1" aria-hidden="true"></i>휴가 신청·취소는 슬랙 <strong>휴가</strong> 앱에서 해요.</div>';
    }

    var years = [];
    (L.entries || []).forEach(function (e) { var y = e.start.slice(0, 4); if (years.indexOf(y) === -1) years.push(y); });
    // 기록은 처음에 올해만 — 3년치 45줄이 휴대폰에 다 펼쳐지던 것(11-15 점검). 올해 기록이 없으면 전체
    var thisYear = String(L.year || '');
    var pickYear = years.length > 1 && years.indexOf(thisYear) !== -1 ? thisYear : '';
    h += '<div class="gw-sec-head"><div class="gw-sec-title m-0">휴가 기록</div>' +
      (years.length > 1 ? '<select class="form-select form-select-sm gw-year-select" aria-label="연도">' +
        '<option value="">전체</option>' + years.map(function (y) {
          return '<option value="' + y + '"' + (y === pickYear ? ' selected' : '') + '>' + y + '년</option>';
        }).join('') +
        '</select>' : '') + '</div>';
    var cols = opts.mine ? 6 : 5;
    h += '<div class="gw-table-wrap"><table class="gw-table gw-leave-table' + (opts.mine ? ' gw-lv-mine-table' : '') + '"><thead><tr><th>기간</th><th>종류</th><th class="n">일수</th><th>상태</th><th>사유</th>' +
      (opts.mine ? '<th><span class="visually-hidden">취소</span></th>' : '') + '</tr></thead><tbody>';
    if (!L.entries || !L.entries.length) {
      h += '<tr><td colspan="' + cols + '" class="gw-empty">기록이 없습니다.</td></tr>';
    } else {
      L.entries.forEach(function (e) {
        h += '<tr data-year="' + esc(e.start.slice(0, 4)) + '"' + (e.status === '취소' || e.status === '반려' ? ' class="muted"' : '') +
          (pickYear && e.start.slice(0, 4) !== pickYear ? ' hidden' : '') + '>' +
          '<td class="gw-lv-when">' + esc(period(e, true)) + '</td><td>' + esc(e.type) + '</td><td class="n">' + days(e.days) + '</td>' +
          '<td>' + statusBadge(e.status) + '</td><td class="gw-reason">' + esc(e.reason) + '</td>' +
          (opts.mine ? '<td class="gw-lv-act">' + cancelBtn(e) + '</td>' : '') + '</tr>';
      });
    }
    h += '</tbody></table></div>';
    h += '<div class="gw-note">2026년 11월 이전 기록은 예전 휴가 프로그램(플렉스)에서 옮겨 온 것이에요.</div>';
    return h;
  }

  function bindYearFilter(root) {
    var sel = root.querySelector('.gw-year-select');
    if (!sel) return;
    sel.addEventListener('change', function () {
      var y = sel.value;
      root.querySelectorAll('.gw-leave-table tbody tr[data-year]').forEach(function (tr) {
        tr.hidden = !!y && tr.getAttribute('data-year') !== y;
      });
    });
  }

  // ── 내 휴가 ────────────────────────────────────────────────
  if (CFG.mode === 'me') {
    initMine();
    return;
  }

  function initMine() {
    var mine = document.getElementById('lvMine');
    var flash = document.getElementById('lvFlash');
    var applyBtn = document.getElementById('lvApplyBtn');
    var form = null;   // 서버가 준 신청 창 정보

    function showFlash(text, kind) {
      flash.textContent = text;
      flash.className = 'gw-lv-flash ' + (kind === 'err' ? 'err' : 'ok');
      flash.hidden = false;
      try { flash.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } catch (e) { /* 무시 */ }
    }

    function load() {
      return api('/api/leave/me').then(function (j) {
        form = j.form || { can_apply: false };
        applyBtn.hidden = !form.can_apply;
        var note = form.can_apply
          ? '신청하면 승인하는 분께 바로 알림이 가요. 결과는 아래 휴가 기록과 위쪽 <i class="fas fa-bell" aria-hidden="true"></i> 알림에서 볼 수 있어요.'
          : '<i class="fas fa-info-circle me-1" aria-hidden="true"></i>' + esc(form.why || '지금은 휴가를 신청할 수 없어요.');
        mine.innerHTML = renderSummary(j.leave, { mine: true, applyNote: note });
        bindYearFilter(mine);
      }).catch(function (e) {
        mine.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>';
      });
    }

    // ── 신청 창 ──
    var modalEl = document.getElementById('lvApplyModal');
    var f = document.getElementById('lvApplyForm');
    var typeSel = document.getElementById('lvType');
    var subSel = document.getElementById('lvSub');
    var submitBtn = document.getElementById('lvApplySubmit');

    function clearErrors() {
      f.querySelectorAll('.gw-lv-err').forEach(function (el) { el.textContent = ''; });
      f.querySelectorAll('.is-invalid').forEach(function (el) {
        el.classList.remove('is-invalid');
        el.removeAttribute('aria-invalid');
        var d = (el.getAttribute('aria-describedby') || '').split(' ').filter(function (x) { return x && x.indexOf('lvErr_') !== 0; }).join(' ');
        if (d) el.setAttribute('aria-describedby', d); else el.removeAttribute('aria-describedby');
      });
    }

    function showErrors(errs, fallback) {
      var shown = false;
      Object.keys(errs || {}).forEach(function (k) {
        var box = f.querySelector('[data-err="' + k + '"]');
        if (!box) return;
        box.textContent = errs[k];
        var input = f.querySelector('[name="' + k + '"]');
        if (input) {
          input.classList.add('is-invalid');
          // 낭독기가 '이 칸이 틀렸고 이유는 …'으로 읽게 칸과 오류 글을 잇는다
          input.setAttribute('aria-invalid', 'true');
          if (!box.id) box.id = 'lvErr_' + k;
          var d = (input.getAttribute('aria-describedby') || '').split(' ').filter(Boolean);
          if (d.indexOf(box.id) === -1) d.push(box.id);
          input.setAttribute('aria-describedby', d.join(' '));
        }
        if (!shown) {
          if (input) input.focus();
          try { box.scrollIntoView({ block: 'nearest' }); } catch (e) { /* 무시 */ }
        }
        shown = true;
      });
      if (!shown) {
        var all = f.querySelector('[data-err="_"]');
        all.textContent = fallback || '신청하지 못했어요.';
        try { all.scrollIntoView({ block: 'nearest' }); } catch (e) { /* 무시 */ }
      }
    }

    function syncType() {
      var t = typeSel.value;
      var half = (form.half_types || []).indexOf(t) !== -1;
      document.getElementById('lvSubWrap').hidden = t !== form.event;
      document.getElementById('lvEndWrap').hidden = half;
      document.getElementById('lvStartLabel').textContent = half ? '날짜' : '시작일';
      document.getElementById('lvProof').hidden = (form.needs_proof || []).indexOf(t) === -1;
    }

    // 칸 비우기는 처음 열 때·신청이 된 뒤에만 — 실수로 닫았다 다시 열면 적던 것이 그대로 있게(11-16 점검)
    var applyFresh = true;
    var closingAfterApply = false;   // 신청 성공으로 닫을 때는 묻지 않음

    function applyDirty() {
      return !!(f.start.value || f.end.value || f.reason.value.trim() || subSel.value);
    }

    function selectIfThere(sel, v) {
      if (v && Array.prototype.some.call(sel.options, function (o) { return o.value === v; })) sel.value = v;
    }

    function openApply() {
      if (!form || !form.can_apply) return;
      clearErrors();
      closingAfterApply = false;
      var keepType = applyFresh ? '' : typeSel.value;
      var keepSub = applyFresh ? '' : subSel.value;
      if (applyFresh) {
        f.reset();
        document.getElementById('lvEnd').min = '';
        applyFresh = false;
      }
      document.getElementById('lvApplyLeft').innerHTML = '<strong>' + esc(form.name) + '</strong> 님 오늘 기준 남은 연차: <strong>' + days(form.balance) + '</strong>';
      typeSel.innerHTML = form.types.map(function (o) { return '<option value="' + esc(o.value) + '">' + esc(o.label) + '</option>'; }).join('');
      subSel.innerHTML = '<option value="">고르기</option>' + form.event_subs.map(function (o) { return '<option value="' + esc(o.value) + '">' + esc(o.label) + '</option>'; }).join('');
      typeSel.value = form.types.length ? form.types[0].value : '';
      selectIfThere(typeSel, keepType);
      selectIfThere(subSel, keepSub);
      syncType();
      if (window.bootstrap) { returnFocusOnHide(modalEl); window.bootstrap.Modal.getOrCreateInstance(modalEl).show(); }
    }

    // 적던 중에 [닫기]·✕·Esc 를 누르면 한 번 묻는다(바깥 클릭은 data-bs-backdrop="static" 으로 막음)
    modalEl.addEventListener('hide.bs.modal', function (ev) {
      if (closingAfterApply || !applyDirty()) return;
      if (!window.confirm('아직 신청하지 않았어요. 닫을까요?')) ev.preventDefault();
    });
    modalEl.addEventListener('hidden.bs.modal', function () { closingAfterApply = false; });

    applyBtn.addEventListener('click', openApply);
    typeSel.addEventListener('change', function () { clearErrors(); syncType(); });
    // 시작일을 고르면 마지막 날 달력이 그 달에서 열리게
    document.getElementById('lvStart').addEventListener('change', function () {
      var end = document.getElementById('lvEnd');
      end.min = this.value || '';
    });

    f.addEventListener('submit', function (ev) {
      ev.preventDefault();
      clearErrors();
      var t = typeSel.value;
      var half = (form.half_types || []).indexOf(t) !== -1;
      var body = {
        type: t,
        sub: t === form.event ? subSel.value : '',
        start: f.start.value,
        end: half ? '' : f.end.value,
        reason: f.reason.value.trim()
      };
      // 서버와 같은 검사 일부를 미리(칸 비움) — 나머지(겹침·남은 연차·한도)는 서버가
      var errs = {};
      if (t === form.event && !body.sub) errs.sub = '경조사 종류를 골라 주세요.';
      if (!body.start) errs.start = '날짜를 골라 주세요.';
      if (!body.reason) errs.reason = '사유를 적어 주세요.';
      if (Object.keys(errs).length) { showErrors(errs); return; }

      submitBtn.disabled = true;
      submitBtn.textContent = '신청하는 중…';
      api('/api/leave/apply', { method: 'POST', body: body }).then(function (j) {
        applyFresh = true;
        closingAfterApply = true;
        if (window.bootstrap) window.bootstrap.Modal.getOrCreateInstance(modalEl).hide();
        showFlash(j.status === 'approved'
          ? '휴가가 확정됐어요 (' + days(j.days) + ').'
          : '휴가를 신청했어요 (' + days(j.days) + '). 승인되면 아래 기록이 「승인」으로 바뀌어요.', 'ok');
        load();
      }).catch(function (e) {
        showErrors(e.fields, e.message);
      }).then(function () {
        submitBtn.disabled = false;
        submitBtn.textContent = '신청';
      });
    });

    // ── 취소 ──
    var cModalEl = document.getElementById('lvCancelModal');
    var cYes = document.getElementById('lvCancelYes');
    var cErr = document.getElementById('lvCancelErr');
    var pending = null;

    mine.addEventListener('click', function (ev) {
      var b = ev.target.closest('[data-lv-cancel]');
      if (!b) return;
      pending = { id: b.getAttribute('data-lv-cancel'), mode: b.getAttribute('data-mode') };
      document.getElementById('lvCancelWhat').innerHTML = '<strong>' + esc(b.getAttribute('data-what')) + '</strong>';
      document.getElementById('lvCancelText').textContent = b.getAttribute('data-confirm');
      cErr.textContent = '';
      if (window.bootstrap) { returnFocusOnHide(cModalEl); window.bootstrap.Modal.getOrCreateInstance(cModalEl).show(); }
    });

    cYes.addEventListener('click', function () {
      if (!pending) return;
      var p = pending;
      cYes.disabled = true;
      cYes.textContent = '취소하는 중…';
      cErr.textContent = '';
      api('/api/leave/' + encodeURIComponent(p.id) + '/cancel', { method: 'POST' }).then(function () {
        pending = null;
        if (window.bootstrap) window.bootstrap.Modal.getOrCreateInstance(cModalEl).hide();
        showFlash(p.mode === 'rest' ? '오늘부터 남은 휴가를 취소했어요.' : '휴가를 취소했어요.', 'ok');
        load();
      }).catch(function (e) {
        cErr.textContent = e.message;
      }).then(function () { cYes.disabled = false; cYes.textContent = '취소하기'; });
    });

    load();
  }

  if (CFG.mode !== 'team') return;

  // ── 전사 현황: 탭 ───────────────────────────────────────────
  var tabs = { cal: document.getElementById('lvTabCal'), list: document.getElementById('lvTabList') };
  var panes = { cal: document.getElementById('lvCalPane'), list: document.getElementById('lvListPane') };
  var listLoaded = false;

  function showTab(key) {
    Object.keys(tabs).forEach(function (k) {
      tabs[k].setAttribute('aria-selected', String(k === key));
      tabs[k].tabIndex = k === key ? 0 : -1;
      panes[k].hidden = k !== key;
    });
    try { localStorage.setItem('gw.leave.tab', key); } catch (e) { /* 무시 */ }
    if (key === 'list' && !listLoaded) loadTeam();
  }
  tabs.cal.addEventListener('click', function () { showTab('cal'); });
  tabs.list.addEventListener('click', function () { showTab('list'); });
  tabs.cal.parentElement.addEventListener('keydown', function (e) {
    if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].indexOf(e.key) === -1) return;
    e.preventDefault();
    var next = (e.key === 'Home') ? 'cal' : (e.key === 'End') ? 'list'
      : (tabs.cal.getAttribute('aria-selected') === 'true' ? 'list' : 'cal');   // 탭이 둘이라 좌우 모두 바꿈
    showTab(next);
    tabs[next].focus();
  });

  // ── 달력 ───────────────────────────────────────────────────
  var calEl = document.getElementById('lvCal');
  var titleEl = document.getElementById('lvMonthTitle');
  var now = new Date();
  var cur = { y: now.getFullYear(), m: now.getMonth() + 1 };

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function iso(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }

  function renderCal(entries) {
    var first = new Date(cur.y, cur.m - 1, 1);
    var lastDay = new Date(cur.y, cur.m, 0).getDate();
    var lead = (first.getDay() + 6) % 7;     // 월요일 시작
    var todayIso = iso(now.getFullYear(), now.getMonth() + 1, now.getDate());
    var h = '<div class="gw-cal-grid">';
    WD.forEach(function (w, i) { h += '<div class="gw-cal-wd' + (i >= 5 ? ' we' : '') + '">' + w + '</div>'; });
    for (var i = 0; i < lead; i++) h += '<div class="gw-cal-day out"></div>';
    for (var d = 1; d <= lastDay; d++) {
      var di = iso(cur.y, cur.m, d);
      var wd = (lead + d - 1) % 7;
      var todays = entries.filter(function (e) { return e.start <= di && di <= e.end; });
      h += '<div class="gw-cal-day' + (wd >= 5 ? ' we' : '') + (di === todayIso ? ' today' : '') + '">' +
        '<div class="gw-cal-num">' + d + '</div>';
      todays.forEach(function (e) {
        var label = e.half ? e.half + ' 반차' : (e.type === '연차' ? '' : e.type);
        h += '<div class="gw-cal-chip ' + (e.status === '대기' ? 'wait' : 'ok') + '" title="' +
          esc(e.name + ' · ' + e.type + (e.half ? ' ' + e.half : '') + (e.status === '대기' ? ' · 승인 대기' : '')) + '">' +
          esc(e.name) + (label ? ' <small>' + esc(label) + '</small>' : '') +
          '<span class="visually-hidden">' + (e.status === '대기' ? ' 승인 대기' : ' 승인됨') + '</span></div>';
      });
      h += '</div>';
    }
    var tail = (7 - ((lead + lastDay) % 7)) % 7;
    for (var t = 0; t < tail; t++) h += '<div class="gw-cal-day out"></div>';
    h += '</div>';
    calEl.innerHTML = h;
  }

  function loadCal() {
    titleEl.textContent = cur.y + '년 ' + cur.m + '월';
    calEl.innerHTML = '<div class="gw-empty">불러오는 중…</div>';
    api('/api/leave/calendar?month=' + cur.y + '-' + pad(cur.m)).then(function (j) {
      renderCal(j.entries || []);
    }).catch(function (e) {
      calEl.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>';
    });
  }

  document.getElementById('lvPrev').addEventListener('click', function () {
    cur.m -= 1; if (cur.m < 1) { cur.m = 12; cur.y -= 1; } loadCal();
  });
  document.getElementById('lvNext').addEventListener('click', function () {
    cur.m += 1; if (cur.m > 12) { cur.m = 1; cur.y += 1; } loadCal();
  });
  document.getElementById('lvToday').addEventListener('click', function () {
    cur = { y: now.getFullYear(), m: now.getMonth() + 1 }; loadCal();
  });

  // ── 직원별 ─────────────────────────────────────────────────
  var tbody = document.querySelector('#lvTeamTable tbody');

  function loadTeam() {
    api('/api/leave/team').then(function (j) {
      listLoaded = true;
      var rows = j.people || [];
      tbody.innerHTML = rows.map(function (p) {
        var nxt = p.next ? esc(period(p.next)) + (p.next.status === '대기' ? ' <span class="gw-badge gw-badge-leave">대기</span>' : '') : '<span class="text-muted">-</span>';
        return '<tr tabindex="0" data-id="' + p.person_id + '" data-name="' + esc(p.name) + '">' +
          '<td><strong>' + esc(p.name) + '</strong>' + (p.access === 'field' ? ' <span class="gw-badge gw-badge-field">' + esc(p.job_role || '현장') + '</span>' : '') + '</td>' +
          '<td>' + esc(p.department) + '</td><td>' + esc(p.hire_date) + '</td>' +
          '<td class="n">' + (p.linked ? days(p.balance) : '<span class="text-muted" title="휴가 장부에 없음">-</span>') + '</td>' +
          '<td class="n">' + days(p.annual_used) + '</td><td class="n">' + days(p.sick_used) + '</td>' +
          '<td class="n">' + (p.pending ? '<strong>' + p.pending + '</strong>' : '0') + '</td><td>' + nxt + '</td></tr>';
      }).join('') || '<tr><td colspan="8" class="gw-empty">재직자가 없습니다.</td></tr>';
    }).catch(function (e) {
      tbody.innerHTML = '<tr><td colspan="8" class="gw-empty">' + esc(e.message) + '</td></tr>';
    });
  }

  var modalEl = document.getElementById('lvPersonModal');
  var bodyEl = document.getElementById('lvPersonBody');

  function openPerson(id, name) {
    document.getElementById('lvPersonTitle').textContent = name + ' 휴가';
    bodyEl.innerHTML = '<div class="gw-empty">불러오는 중…</div>';
    if (window.bootstrap) { returnFocusOnHide(modalEl); window.bootstrap.Modal.getOrCreateInstance(modalEl).show(); }
    api('/api/leave/person/' + id).then(function (j) {
      bodyEl.innerHTML = renderSummary(j.leave, { hideApplyNote: true });
      bindYearFilter(bodyEl);
    }).catch(function (e) {
      bodyEl.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>';
    });
  }

  tbody.addEventListener('click', function (e) {
    var tr = e.target.closest('tr[data-id]');
    if (tr) openPerson(tr.getAttribute('data-id'), tr.getAttribute('data-name'));
  });
  tbody.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' && e.key !== ' ') return;
    var tr = e.target.closest('tr[data-id]');
    if (tr) { e.preventDefault(); openPerson(tr.getAttribute('data-id'), tr.getAttribute('data-name')); }
  });

  var saved = 'cal';
  try { saved = localStorage.getItem('gw.leave.tab') || 'cal'; } catch (e) { /* 무시 */ }
  loadCal();
  showTab(saved === 'list' ? 'list' : 'cal');
})();
