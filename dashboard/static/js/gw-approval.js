/* 그룹웨어 결재함·플렉스 보관함 (2026-11-10).
   결재함: 내 차례·내가 올린·전체 → 줄을 누르면 오른쪽 창에 상세 + 처리 단추(슬랙 카드와 같은 처리).
   서버에서 온 글자는 모두 esc() 로만 넣는다. */
(function () {
  'use strict';

  var WD = ['일', '월', '화', '수', '목', '금', '토'];
  function $(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  function api(url, opts) {
    opts = opts || {};
    var method = opts.method || 'GET';
    var init = { method: method, credentials: 'same-origin' };
    if (opts.body) {   // 본문 있을 때만 JSON 머리글(빈 본문 + JSON 머리글 = 운영 400)
      init.headers = { 'Content-Type': 'application/json' };
      init.body = JSON.stringify(opts.body);
    }
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      // 로그인 화면으로 돌려보내졌거나(redirected·200 HTML) 401 이면 로그인이 끝난 것.
      // 예전엔 로그인 화면을 {} 로 읽어 처리가 "된" 것처럼 보였음(11-16 점검)
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) {
        var se = new Error(SESSION_ENDED);
        se.stale = false;
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
          err.stale = !!j.stale;
          err.status = r.status;
          throw err;
        }
        return j;
      });
    });
  }
  function won(n) { return (Number(n) || 0).toLocaleString('ko-KR') + '원'; }
  function when(s) {
    if (!s) return '';
    var d = new Date(s.replace(' ', 'T'));
    if (isNaN(d)) return s;
    return (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ') ' + s.slice(11, 16);
  }

  // ── 오른쪽 창 ─────────────────────────────────────
  var drawer = $('apDrawer');
  var drawerOpener = null;   // 창을 연 줄 — 닫으면 그 줄로 포커스를 돌려준다(11-15 점검)
  function drawerTitle() {
    var t = $('apDocTitle');
    if (t && !t.hasAttribute('tabindex')) t.setAttribute('tabindex', '-1');
    return t;
  }
  function openDrawer(title, html) {
    $('apDocTitle').textContent = title;
    $('apDoc').innerHTML = html;
    if (drawer.hidden) {
      var a = document.activeElement;
      drawerOpener = a && a !== document.body && !drawer.contains(a) ? a : null;
    }
    drawer.hidden = false;
    document.body.classList.add('gw-drawer-open');
    // 키보드·화면 읽기 사용자가 창 안에서 시작하게 — 제목으로 포커스
    var t = drawerTitle();
    if (t) t.focus({ preventScroll: true });
  }
  function closeDrawer() {
    drawer.hidden = true;
    document.body.classList.remove('gw-drawer-open');
    if (history.replaceState && /[?&]id=/.test(location.search)) {
      var u = new URL(location.href); u.searchParams.delete('id'); history.replaceState({}, '', u);
    }
    if (drawerOpener && document.contains(drawerOpener)) {
      try { drawerOpener.focus({ preventScroll: true }); } catch (e) { /* 무시 */ }
    }
    drawerOpener = null;
  }
  // 창이 열린 동안 Tab 이 뒤쪽 표로 빠져나가지 않게(aria-modal 인데 포커스가 갇히지 않던 것)
  function trapTab(e) {
    if (e.key !== 'Tab' || drawer.hidden) return;
    var panel = drawer.querySelector('.gw-drawer-panel');
    var items = Array.prototype.filter.call(
      panel.querySelectorAll('a[href], button:not([disabled]), textarea:not([disabled]), input:not([disabled]), select:not([disabled]), [tabindex]:not([tabindex="-1"])'),
      function (el) { return el.offsetParent !== null; });
    if (!items.length) return;
    var first = items[0], last = items[items.length - 1];
    var cur = document.activeElement;
    if (!panel.contains(cur)) { e.preventDefault(); first.focus(); return; }
    if (e.shiftKey && (cur === first || cur === drawerTitle())) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && cur === last) { e.preventDefault(); first.focus(); }
  }
  // 반려·수정 요청 사유를 적던 중에 바깥·✕·Esc 로 닫으면 글이 말없이 사라졌음 → 한 번 묻는다(11-16 점검).
  // 사유 칸이 닫혀 있거나(보냄·[닫기]) 비어 있으면 묻지 않음
  function reasonPending() {
    var box = $('apReason'), t = $('apReasonText');
    return !!(box && !box.hidden && t && t.value.trim());
  }
  function tryCloseDrawer() {
    if (reasonPending() && !window.confirm('적은 사유가 지워져요. 닫을까요?')) return;
    closeDrawer();
  }
  if (drawer) {
    drawer.addEventListener('click', function (e) { if (e.target.closest('[data-close]')) tryCloseDrawer(); });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && !drawer.hidden) tryCloseDrawer();
      trapTab(e);
    });
  }

  // 칸 값 — 줄이 http(s):// 로 시작하면 그 주소를 누를 수 있는 링크로(비품 신청 '상품 링크', 11-16 점검).
  // javascript: 같은 다른 주소는 글자 그대로. 주소·글자 모두 esc() 를 거친다
  function valueHtml(v) {
    return String(v == null ? '' : v).split('\n').map(function (line) {
      var m = /^(\s*)(https?:\/\/\S+)(.*)$/i.exec(line);
      if (m) {
        return esc(m[1]) + '<a href="' + esc(m[2]) + '" target="_blank" rel="noopener">' + esc(m[2]) + '</a>' + esc(m[3]);
      }
      return esc(line);
    }).join('<br>');
  }

  function emptyState(icon, title, desc) {
    return '<div class="gw-empty-state"><span class="ic" aria-hidden="true"><i class="fas ' + icon + '"></i></span>' +
      '<div class="t">' + esc(title) + '</div>' + (desc ? '<div class="d">' + esc(desc) + '</div>' : '') + '</div>';
  }

  function fieldsHtml(fields) {
    if (!fields || !fields.length) return '<div class="gw-log">적힌 칸이 없어요.</div>';
    var h = '<dl class="gw-dl gw-dl-1col">';
    fields.forEach(function (f) {
      if (f.section) h += '</dl><div class="gw-ap-section">' + esc(f.section) + '</div><dl class="gw-dl gw-dl-1col">';
      h += '<div' + (f.wide ? ' class="gw-dl-wide"' : '') + '><dt>' + esc(f.label) + '</dt><dd>' + valueHtml(f.value) + '</dd></div>';
    });
    return h + '</dl>';
  }

  // ── 휴대폰 필터 접기(11-15 부장님 제보: 결재함 화면이 엉망) ──────────────────
  // 양식·상태·날짜 칩이 네 줄을 차지해 목록이 화면 중간부터 시작했다 → 휴대폰에선 검색 + [필터 n] 만,
  // 누르면 펼침. 넓은 화면에선 [필터] 단추를 CSS 로 숨겨 지금 그대로
  document.querySelectorAll('.gw-approval .gw-ap-filters').forEach(function (box) {
    var search = box.querySelector('.gw-chip-search');
    if (!search || box.querySelector('.gw-chip-more')) return;
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'gw-chip-more';
    btn.setAttribute('aria-expanded', 'false');
    btn.innerHTML = '<i class="fas fa-sliders-h me-1" aria-hidden="true"></i>필터<b class="gw-chip-more-n" hidden></b>';
    search.insertAdjacentElement('afterend', btn);
    var nEl = btn.querySelector('b');
    function sync() {
      var n = Array.prototype.filter.call(box.querySelectorAll('select.gw-chip-select, input.gw-chip-date'),
        function (el) { return !!el.value; }).length;
      nEl.textContent = n;
      nEl.hidden = !n;
      btn.classList.toggle('on', n > 0);
    }
    btn.addEventListener('click', function () {
      var open = !box.classList.contains('open');
      box.classList.toggle('open', open);
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
    box.addEventListener('change', sync);
    box.addEventListener('click', function () { setTimeout(sync, 0); });   // [초기화]는 값만 비움(change 없음)
    sync();
  });

  // ── 결재함 ────────────────────────────────────────
  var page = $('apPage');
  if (page) {
    var state = { box: 'todo', loaded: false };
    var params = new URLSearchParams(location.search);
    if (params.get('box')) state.box = params.get('box');

    function filters() {
      return { box: state.box, form: $('apForm').value, status: $('apStatus').value,
               from: $('apFrom').value, to: $('apTo').value, q: $('apQ').value.trim() };
    }
    function qs(f) {
      return Object.keys(f).filter(function (k) { return f[k]; })
        .map(function (k) { return k + '=' + encodeURIComponent(f[k]); }).join('&');
    }
    function setTabs() {
      document.querySelectorAll('#apTabs button').forEach(function (b) {
        var on = b.getAttribute('data-box') === state.box;
        b.classList.toggle('active', on);
        b.setAttribute('aria-selected', on ? 'true' : 'false');
        b.tabIndex = on ? 0 : -1;
      });
      var f = filters();
      var active = f.form || f.status || f.from || f.to || f.q;
      $('apFilters').classList.toggle('has-active-filters', !!active);
      ['apForm', 'apStatus'].forEach(function (id) { $(id).classList.toggle('on', !!$(id).value); });
    }

    var timer = null;
    // 요청 번호 — 탭·필터를 빨리 바꾸면 먼저 보낸 요청의 응답이 늦게 와서 지금 고른 탭을 덮어쓰던 것(11-16 점검).
    // 마지막으로 보낸 요청의 응답만 그린다
    var listSeq = 0;
    function load() {
      setTabs();
      var f = filters();
      $('apExcel').href = '/api/approval/export.xlsx?' + qs(f);
      var seq = ++listSeq;
      api('/api/approval/list?' + qs(f)).then(function (j) {
        if (seq !== listSeq) return;
        state.box = j.box;
        $('apCntTodo').textContent = j.counts.todo;
        $('apCntMine').textContent = j.counts.mine;
        if (j.see_all) { $('apTabAll').hidden = false; $('apCntAll').textContent = j.counts.all; }
        $('apUnlinked').hidden = j.linked;
        if (!j.linked) {
          $('apUnlinked').textContent = j.sid_unavailable
            ? '슬랙 확인이 잠시 안 돼요. 잠시 뒤 새로 고쳐 주세요.'
            : '인사카드에 슬랙 계정이 연결되지 않아 결재함을 쓸 수 없어요. 경영지원에 말씀해 주세요.';
        }
        if (!state.loaded) {
          var sel = $('apForm');
          j.forms.forEach(function (x) { var o = document.createElement('option'); o.value = x.key; o.textContent = x.label; sel.appendChild(o); });
          state.loaded = true;
        }
        setTabs();
        renderRows(j.items);
      }).catch(function (e) {
        if (seq !== listSeq) return;
        $('apRows').innerHTML = '<tr><td colspan="8" class="gw-empty">' + esc(e.message) + '</td></tr>';
      });
    }

    function renderRows(items) {
      $('apCount').textContent = items.length.toLocaleString('ko-KR') + '건';
      if (!items.length) {
        // 빈 화면 — 글 한 줄 대신 무엇을 하면 되는지까지(11-28 점검)
        var es = state.box === 'todo'
          ? emptyState('fa-check', '지금 내 차례인 결재가 없어요', '새 결재가 오면 슬랙과 위쪽 알림 종으로 알려 드려요.')
          : emptyState('fa-filter', '맞는 결재가 없어요', '거르기 조건을 바꾸거나 [초기화]를 눌러 보세요.');
        $('apRows').innerHTML = '<tr><td colspan="8" class="gw-empty-cell">' + es + '</td></tr>';
        return;
      }
      $('apRows').innerHTML = items.map(function (it) {
        // 줄 전체가 누르는 자리 — 화면 읽기엔 '단추'로, 이름은 번호·내용·상태(Enter·Space 로 열림, 아래 keydown)
        var label = '결재 #' + it.id + ' ' + (it.what || it.form_label || '') + ', ' + (it.status_label || '') +
          (it.my_turn ? ', 내 차례' : '') + ' 열기';
        return '<tr data-id="' + it.id + '" tabindex="0" role="button" aria-label="' + esc(label) + '"' + (it.my_turn ? ' class="my-turn"' : '') + '>' +
          '<td class="muted">#' + it.id + '</td>' +
          '<td>' + esc(it.form_label) + '</td>' +
          '<td class="main">' + (it.urgent ? '<span class="gw-badge gw-badge-urgent">긴급</span> ' : '') + esc(it.what) + '</td>' +
          '<td>' + esc(it.project_code) + '</td>' +
          '<td class="num">' + (it.amount ? won(it.amount) : '') + '</td>' +
          '<td>' + esc(it.requester) + '</td>' +
          '<td><span class="gw-ap-st st-' + esc(it.status) + '">' + esc(it.status_label) + '</span>' + (it.my_turn ? ' <span class="gw-ap-turn">내 차례</span>' : '') + '</td>' +
          '<td class="muted">' + esc(when(it.created_at)) + '</td></tr>';
      }).join('');
    }

    $('apRows').addEventListener('click', function (e) {
      var tr = e.target.closest('tr[data-id]');
      if (tr) openDoc(Number(tr.getAttribute('data-id')));
    });
    $('apRows').addEventListener('keydown', function (e) {
      var tr = e.target.closest('tr[data-id]');
      if (tr && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); openDoc(Number(tr.getAttribute('data-id'))); }
    });
    $('apTabs').addEventListener('keydown', function (e) {
      if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].indexOf(e.key) === -1) return;
      var tabs = Array.prototype.filter.call($('apTabs').querySelectorAll('button[data-box]'), function (b) { return !b.hidden; });
      var i = tabs.indexOf(document.activeElement);
      if (i < 0) return;
      e.preventDefault();
      var n = e.key === 'Home' ? 0 : e.key === 'End' ? tabs.length - 1 : (i + (e.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
      tabs[n].focus();
      tabs[n].click();
    });
    $('apTabs').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-box]');
      if (!b) return;
      state.box = b.getAttribute('data-box');
      load();
    });
    ['apForm', 'apStatus', 'apFrom', 'apTo'].forEach(function (id) { $(id).addEventListener('change', load); });
    $('apQ').addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(load, 250); });
    $('apReset').addEventListener('click', function () {
      ['apForm', 'apStatus', 'apFrom', 'apTo', 'apQ'].forEach(function (id) { $(id).value = ''; });
      load();
    });

    function openDoc(id) {
      openDrawer('결재 #' + id, '<div class="gw-empty">불러오는 중…</div>');
      if (history.replaceState) { var u = new URL(location.href); u.searchParams.set('id', id); history.replaceState({}, '', u); }
      api('/api/approval/' + id).then(function (j) { renderDoc(j.doc); })
        .catch(function (e) { $('apDoc').innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>'; });
    }

    function renderDoc(d) {
      $('apDocTitle').textContent = d.title;
      var h = '<div class="gw-ap-doc-top">' +
        '<span class="gw-ap-st st-' + esc(d.status) + '">' + esc(d.status_label) + '</span>' +
        (d.urgent ? ' <span class="gw-badge gw-badge-urgent">긴급</span>' : '') +
        '<span class="gw-ap-doc-meta">' + esc(d.form_label) + ' · ' + esc(d.requester) + ' · ' + esc(when(d.created_at)) + '</span></div>';
      if (d.amount) h += '<div class="gw-ap-amount">' + won(d.amount) + '</div>';
      // 슬랙 카드와 같은 ⚠️ 줄 — 결재함으로 처리해도 계좌 바뀜·중복을 놓치지 않게(5차 점검)
      var warns = d.warnings || [];
      var hasActions = d.actions && d.actions.length;
      if (warns.length) {
        h += '<div class="gw-ap-warn" role="alert">' +
          (hasActions ? '<div class="gw-ap-warn-head">아래 내용을 확인하고 처리해 주세요</div>' : '') +
          warns.map(function (w) { return '<div>' + esc(w) + '</div>'; }).join('') + '</div>';
      }
      // 처리 단추는 내용(계좌·예금주) 아래 — 판단 근거를 보기 전에 [승인]을 누르지 않게(11-15 점검).
      // 내 차례인데 위에서 단추가 안 보이면 헷갈리니 한 줄로 알려 줌
      if (hasActions) {
        h += '<div class="gw-ap-turn-hint"><i class="fas fa-arrow-down me-1" aria-hidden="true"></i>내용을 확인한 뒤 아래에서 처리해 주세요.</div>';
      }

      h += '<div class="gw-sec-title">내용</div>' + fieldsHtml(d.fields);

      if (hasActions) {
        h += '<div class="gw-ap-actions' + (warns.length ? ' warned' : '') + '" id="apActions">' + d.actions.map(function (a) {
          var warnedPrimary = a.primary && warns.length;
          return '<button type="button" class="btn ' + (a.primary ? 'btn-primary' : (a.key === 'reject' ? 'btn-outline-danger' : 'btn-outline-secondary')) +
            '" data-act="' + a.key + '" data-reason="' + (a.needs_reason ? '1' : '') + '">' +
            (warnedPrimary ? '<i class="fas fa-exclamation-triangle me-1" aria-hidden="true"></i>' : '') + esc(a.label) +
            (warnedPrimary ? '<span class="visually-hidden"> (경고 있음)</span>' : '') + '</button>';
        }).join('') + '</div>' +
        '<div class="gw-ap-reason" id="apReason" hidden><textarea class="form-control" rows="3" maxlength="300" id="apReasonText" aria-label="사유"></textarea>' +
        // 사유 칸 닫기 — '취소'는 결재 [취소] 처리와 헷갈려서 [닫기](11-16 점검)
        '<div class="gw-ap-reason-btns"><button type="button" class="btn btn-sm btn-outline-secondary" id="apReasonCancel">닫기</button>' +
        '<button type="button" class="btn btn-sm btn-primary" id="apReasonOk">보내기</button></div></div>' +
        '<div class="gw-form-error" id="apActError" role="alert"></div>';
      }

      h += '<div class="gw-sec-title">진행</div>';
      if (d.steps.length) {
        // 단계 줄(점과 선) — 반려 빨강·보완 주황·취소 회색·그 밖 초록, 끝나지 않았으면 맨 아래 '다음' 칸(11-28 점검)
        var STEP_OPEN = { pending: 1, final: 1, approved: 1, revise: 1 };
        var stepKind = function (s) {
          var r = String(s.result || s.step || '');
          return /반려/.test(r) ? 'bad' : /보완|수정/.test(r) ? 'warn' : /취소/.test(r) ? 'mute' : 'ok';
        };
        h += '<ol class="gw-ap-steps gw-tl">' + d.steps.map(function (s) {
          return '<li class="gw-tl-' + stepKind(s) + '"><span class="who">' + esc(s.actor || '-') + '</span> ' + esc(s.step) +
            (s.result && s.result !== s.step ? ' · ' + esc(s.result) : '') +
            (s.reason ? '<div class="why">' + esc(s.reason) + '</div>' : '') +
            '<span class="at">' + esc(when(s.at)) + '</span></li>';
        }).join('') + (STEP_OPEN[d.status] ? '<li class="gw-tl-next"><span class="who">다음</span> ' + esc(d.status_label) + '</li>' : '') + '</ol>';
      } else {
        h += '<div class="gw-log">아직 처리한 사람이 없어요.</div>';
      }
      if (d.slack_link) {
        h += '<a class="gw-ap-slack" href="' + esc(d.slack_link) + '" target="_blank" rel="noopener"><i class="fab fa-slack me-1" aria-hidden="true"></i>슬랙 카드 열기 — 첨부 파일·댓글은 카드 스레드에 있어요</a>';
      }
      var hadFocusInside = drawer.contains(document.activeElement);
      $('apDoc').innerHTML = h;
      bindActions(d);
      // 처리 뒤 다시 그리면 누른 단추가 사라져 포커스를 잃음 → 제목으로
      if (hadFocusInside && !drawer.contains(document.activeElement)) {
        var t = drawerTitle();
        if (t) t.focus({ preventScroll: true });
      }
    }

    function bindActions(d) {
      var box = $('apActions');
      if (!box) return;
      var pending = null;
      box.addEventListener('click', function (e) {
        var b = e.target.closest('button[data-act]');
        if (!b) return;
        var act = b.getAttribute('data-act');
        $('apActError').textContent = '';
        if (b.getAttribute('data-reason')) {
          pending = act;
          $('apReason').hidden = false;
          var t = $('apReasonText');
          t.placeholder = act === 'reject' ? '반려 사유를 적어 주세요.' : '무엇을 고쳐야 하는지 적어 주세요.';
          t.setAttribute('aria-label', act === 'reject' ? '반려 사유' : '수정 요청 내용');
          t.removeAttribute('aria-invalid');
          // 단추 이름에 무엇을 보내는지 — '보내기'만으로는 반려인지 수정 요청인지 안 드러났음
          $('apReasonOk').textContent = act === 'reject' ? '반려 보내기' : '수정 요청 보내기';
          t.focus();
          return;
        }
        var label = b.textContent.replace(' (경고 있음)', '');
        var warnText = (d.warnings || []).filter(function (w) { return w.indexOf('⚠️') !== -1; });
        var ask = d.title + '\n' + (warnText.length ? '\n' + warnText.join('\n') + '\n\n' : '') +
          '「' + label + '」 처리할까요? 슬랙 카드도 같이 바뀌어요.';
        if (!window.confirm(ask)) return;
        send(act, '');
      });
      $('apReasonCancel').addEventListener('click', function () { $('apReason').hidden = true; pending = null; });
      $('apReasonOk').addEventListener('click', function () {
        var r = $('apReasonText').value.trim();
        if (!r) {
          $('apActError').textContent = '내용을 적어 주세요.';
          $('apReasonText').setAttribute('aria-invalid', 'true');
          $('apReasonText').focus();
          return;
        }
        send(pending, r);
      });
      // 처리 중엔 사유 칸의 [보내기]·[닫기]도 막음 — 처리 단추만 막아 두 번 누르면 두 번째에 '이미 처리된 결재' 오류가 떴음
      function lock(on) {
        box.querySelectorAll('button').forEach(function (x) { x.disabled = on; });
        ['apReasonOk', 'apReasonCancel'].forEach(function (id) { var x = $(id); if (x) x.disabled = on; });
      }
      function send(act, reason) {
        lock(true);
        // 처리 중 글자 — 단추만 흐려져 눌렸는지 알기 어려웠음
        var busy = reason ? $('apReasonOk') : box.querySelector('button[data-act="' + act + '"]');
        var busyHtml = busy ? busy.innerHTML : '';
        if (busy) busy.textContent = '처리하는 중…';
        function restore() { if (busy && document.contains(busy)) busy.innerHTML = busyHtml; }
        // rev = 이 화면을 그릴 때의 수정 회차 — 그 사이 요청자가 고쳤으면 서버가 거절한다(11-14)
        api('/api/approval/' + d.id + '/act', { method: 'POST', body: { action: act, reason: reason, rev: d.rev } })
          .then(function (j) { renderDoc(j.doc); load(); })
          .catch(function (e) {
            // 내용이 바뀌었거나(stale) 다른 사람이 먼저 처리해 409 '이미 처리된 결재예요'가 오면 —
            // 옛 단추가 그대로 남아 또 누르게 되던 것(11-16 점검) → 지금 상태로 다시 그리고 목록도 새로
            var conflict = e.stale || (e.status === 409 && /^(이미 처리된 결재|지금 단계에서는)/.test(e.message));
            if (conflict) {
              api('/api/approval/' + d.id).then(function (j) {
                renderDoc(j.doc);
                var el = $('apActError');
                if (el) {
                  el.textContent = e.message;
                } else {   // 처리 단추가 없어진 화면 — 안내를 본문 맨 위에
                  var note = document.createElement('div');
                  note.className = 'gw-ap-warn';
                  note.setAttribute('role', 'alert');
                  note.textContent = e.message;
                  $('apDoc').insertBefore(note, $('apDoc').firstChild);
                }
                load();
              }).catch(function () {
                lock(false);
                restore();
                $('apActError').textContent = e.message;
              });
              return;
            }
            lock(false);
            restore();
            $('apActError').textContent = e.message;
          });
      }
    }

    if (params.get('id')) openDoc(Number(params.get('id')));
    load();
  }

  // ── 플렉스 보관함 ──────────────────────────────────
  var arPage = $('arPage');
  if (arPage) {
    var ar = { page: 1, loaded: false };
    var arTimer = null;
    var arSeq = 0;   // 마지막 요청의 응답만 그림(검색어 입력·쪽 넘김이 겹칠 때)
    // focusDir = 쪽 넘김 단추로 불렀을 때 그 방향('prev'|'next') — 다시 그린 뒤 같은 단추로 포커스(11-16 점검)
    function arLoad(focusDir) {
      var f = { form: $('arForm').value, from: $('arFrom').value, to: $('arTo').value, q: $('arQ').value.trim(), page: ar.page };
      var q = Object.keys(f).filter(function (k) { return f[k]; }).map(function (k) { return k + '=' + encodeURIComponent(f[k]); }).join('&');
      arPage.querySelector('.gw-ap-filters').classList.toggle('has-active-filters', !!(f.form || f.from || f.to || f.q));
      $('arForm').classList.toggle('on', !!f.form);
      var seq = ++arSeq;
      api('/api/approval/archive?' + q).then(function (j) {
        if (seq !== arSeq) return;
        if (!ar.loaded) {
          j.forms.forEach(function (x) { var o = document.createElement('option'); o.value = x; o.textContent = x; $('arForm').appendChild(o); });
          ar.loaded = true;
        }
        $('arCount').textContent = j.total.toLocaleString('ko-KR') + '건';
        $('arRows').innerHTML = j.items.length ? j.items.map(function (it) {
          var label = '문서 ' + (it.doc_no || '') + ' ' + (it.title || '') + ', ' + (it.status || '') + ' 열기';
          return '<tr data-doc="' + esc(it.doc_no) + '" tabindex="0" role="button" aria-label="' + esc(label) + '"><td class="muted">' + esc(it.doc_no) + '</td><td>' + esc(it.form) + '</td>' +
            '<td class="main">' + esc(it.title) + (it.attachments ? ' <i class="fas fa-paperclip muted" role="img" aria-label="플렉스 첨부 ' + it.attachments + '개" title="플렉스 첨부 ' + it.attachments + '개(파일은 옮기지 않았어요)"></i>' : '') + '</td>' +
            '<td>' + esc(it.project_code || '') + '</td><td>' + esc(it.partner || '') + '</td>' +
            '<td class="num">' + (it.amount ? won(it.amount) : '') + '</td><td>' + esc(it.writer) + '</td>' +
            '<td><span class="gw-ap-st flex-' + esc(it.status) + '">' + esc(it.status) + '</span></td>' +
            '<td class="muted">' + esc((it.created_at || '').slice(0, 10)) + '</td></tr>';
        }).join('') : '<tr><td colspan="9" class="gw-empty">맞는 문서가 없어요.</td></tr>';
        var pages = Math.max(1, Math.ceil(j.total / j.size));
        $('arPager').innerHTML = pages > 1
          ? '<button type="button" class="btn btn-sm btn-outline-secondary" data-p="' + (j.page - 1) + '" data-dir="prev"' + (j.page <= 1 ? ' disabled' : '') + '>이전</button>' +
            '<span>' + j.page + ' / ' + pages + '</span>' +
            '<button type="button" class="btn btn-sm btn-outline-secondary" data-p="' + (j.page + 1) + '" data-dir="next"' + (j.page >= pages ? ' disabled' : '') + '>다음</button>'
          : '';
        // 쪽을 넘기면 단추가 새로 그려져 포커스가 사라졌음 → 같은 방향 단추로, 끝 쪽이라 막혔으면 반대 단추로
        if (focusDir) {
          var same = $('arPager').querySelector('button[data-dir="' + focusDir + '"]:not([disabled])');
          var other = $('arPager').querySelector('button[data-dir]:not([disabled])');
          var target = same || other;
          if (target) target.focus({ preventScroll: true });
        }
      }).catch(function (e) {
        if (seq !== arSeq) return;
        $('arRows').innerHTML = '<tr><td colspan="9" class="gw-empty">' + esc(e.message) + '</td></tr>';
      });
    }
    $('arPager').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-p]');
      if (b && !b.disabled) { ar.page = Number(b.getAttribute('data-p')); arLoad(b.getAttribute('data-dir')); }
    });
    ['arForm', 'arFrom', 'arTo'].forEach(function (id) { $(id).addEventListener('change', function () { ar.page = 1; arLoad(); }); });
    $('arQ').addEventListener('input', function () { clearTimeout(arTimer); arTimer = setTimeout(function () { ar.page = 1; arLoad(); }, 250); });
    $('arReset').addEventListener('click', function () {
      ['arForm', 'arFrom', 'arTo', 'arQ'].forEach(function (id) { $(id).value = ''; });
      ar.page = 1; arLoad();
    });
    // 플렉스에서 옮긴 금액 칸은 '5900000'처럼 숫자만 — 금액 칸이면 쉼표를 넣어 보여 줌(값은 그대로, 11-28 점검)
    var MONEY_LABEL = /(금액|금\(|금$|대금|비용|합계|단가|공급가|부가세|VAT|원\))/;
    var NOT_MONEY = /(번호|코드|연락처|전화|계좌|사업자|일자|날짜)/;
    function archiveMoney(fields) {
      return (fields || []).map(function (f) {
        var v = String(f.value == null ? '' : f.value).trim();
        if (f.label && MONEY_LABEL.test(f.label) && !NOT_MONEY.test(f.label) && /^-?\d{4,}$/.test(v)) {
          var o = {}; Object.keys(f).forEach(function (k) { o[k] = f[k]; });
          o.value = Number(v).toLocaleString('ko-KR');
          return o;
        }
        return f;
      });
    }
    function openArchive(no) {
      openDrawer('문서 ' + no, '<div class="gw-empty">불러오는 중…</div>');
      api('/api/approval/archive/' + encodeURIComponent(no)).then(function (j) {
        var d = j.doc;
        $('apDocTitle').textContent = d.title;
        var h = '<div class="gw-ap-doc-top"><span class="gw-ap-st flex-' + esc(d.status) + '">' + esc(d.status) + '</span>' +
          '<span class="gw-ap-doc-meta">' + esc(d.form) + ' · ' + esc(d.writer) + ' · ' + esc((d.created_at || '').slice(0, 16)) + '</span></div>';
        h += '<div class="gw-sec-title">내용</div>' + fieldsHtml(archiveMoney(d.fields));
        if (d.memo) h += '<div class="gw-sec-title">메모</div><div class="gw-ap-memo">' + esc(d.memo).replace(/\n/g, '<br>') + '</div>';
        if (d.attachments) h += '<div class="gw-log mt-2">플렉스 첨부 파일 ' + d.attachments + '개 — 파일은 옮기지 않았어요(목록·내용만 보관).</div>';
        $('apDoc').innerHTML = h;
      }).catch(function (e) { $('apDoc').innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>'; });
    }
    $('arRows').addEventListener('click', function (e) {
      var tr = e.target.closest('tr[data-doc]');
      if (tr) openArchive(tr.getAttribute('data-doc'));
    });
    $('arRows').addEventListener('keydown', function (e) {
      var tr = e.target.closest('tr[data-doc]');
      if (tr && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); openArchive(tr.getAttribute('data-doc')); }
    });
    arLoad();
  }
})();
