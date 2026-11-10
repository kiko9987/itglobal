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
  function api(url, opts) {
    opts = opts || {};
    var init = { method: opts.method || 'GET', credentials: 'same-origin' };
    if (opts.body) {   // 본문 있을 때만 JSON 머리글(빈 본문 + JSON 머리글 = 운영 400)
      init.headers = { 'Content-Type': 'application/json' };
      init.body = JSON.stringify(opts.body);
    }
    return fetch(url, init).then(function (r) {
      return r.json().catch(function () { return {}; }).then(function (j) {
        if (!r.ok || j.success === false) throw new Error(j.message || ('처리하지 못했습니다 (' + r.status + ')'));
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
  function openDrawer(title, html) {
    $('apDocTitle').textContent = title;
    $('apDoc').innerHTML = html;
    drawer.hidden = false;
    document.body.classList.add('gw-drawer-open');
  }
  function closeDrawer() {
    drawer.hidden = true;
    document.body.classList.remove('gw-drawer-open');
    if (history.replaceState && /[?&]id=/.test(location.search)) {
      var u = new URL(location.href); u.searchParams.delete('id'); history.replaceState({}, '', u);
    }
  }
  if (drawer) {
    drawer.addEventListener('click', function (e) { if (e.target.closest('[data-close]')) closeDrawer(); });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !drawer.hidden) closeDrawer(); });
  }

  function fieldsHtml(fields) {
    if (!fields || !fields.length) return '<div class="gw-log">적힌 칸이 없어요.</div>';
    var h = '<dl class="gw-dl gw-dl-1col">';
    fields.forEach(function (f) {
      if (f.section) h += '</dl><div class="gw-ap-section">' + esc(f.section) + '</div><dl class="gw-dl gw-dl-1col">';
      h += '<div' + (f.wide ? ' class="gw-dl-wide"' : '') + '><dt>' + esc(f.label) + '</dt><dd>' + esc(f.value).replace(/\n/g, '<br>') + '</dd></div>';
    });
    return h + '</dl>';
  }

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
      });
      var f = filters();
      var active = f.form || f.status || f.from || f.to || f.q;
      $('apFilters').classList.toggle('has-active-filters', !!active);
      ['apForm', 'apStatus'].forEach(function (id) { $(id).classList.toggle('on', !!$(id).value); });
    }

    var timer = null;
    function load() {
      setTabs();
      var f = filters();
      $('apExcel').href = '/api/approval/export.xlsx?' + qs(f);
      api('/api/approval/list?' + qs(f)).then(function (j) {
        state.box = j.box;
        $('apCntTodo').textContent = j.counts.todo;
        $('apCntMine').textContent = j.counts.mine;
        if (j.see_all) { $('apTabAll').hidden = false; $('apCntAll').textContent = j.counts.all; }
        $('apUnlinked').hidden = j.linked;
        if (!state.loaded) {
          var sel = $('apForm');
          j.forms.forEach(function (x) { var o = document.createElement('option'); o.value = x.key; o.textContent = x.label; sel.appendChild(o); });
          state.loaded = true;
        }
        setTabs();
        renderRows(j.items);
      }).catch(function (e) {
        $('apRows').innerHTML = '<tr><td colspan="8" class="gw-empty">' + esc(e.message) + '</td></tr>';
      });
    }

    function renderRows(items) {
      $('apCount').textContent = items.length.toLocaleString('ko-KR') + '건';
      if (!items.length) {
        var msg = state.box === 'todo' ? '지금 내 차례인 결재가 없어요.' : '맞는 결재가 없어요.';
        $('apRows').innerHTML = '<tr><td colspan="8" class="gw-empty">' + msg + '</td></tr>';
        return;
      }
      $('apRows').innerHTML = items.map(function (it) {
        return '<tr data-id="' + it.id + '" tabindex="0"' + (it.my_turn ? ' class="my-turn"' : '') + '>' +
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
      if (tr && e.key === 'Enter') openDoc(Number(tr.getAttribute('data-id')));
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

      if (d.actions && d.actions.length) {
        h += '<div class="gw-ap-actions" id="apActions">' + d.actions.map(function (a) {
          return '<button type="button" class="btn ' + (a.primary ? 'btn-primary' : (a.key === 'reject' ? 'btn-outline-danger' : 'btn-outline-secondary')) +
            '" data-act="' + a.key + '" data-reason="' + (a.needs_reason ? '1' : '') + '">' + esc(a.label) + '</button>';
        }).join('') + '</div>' +
        '<div class="gw-ap-reason" id="apReason" hidden><textarea class="form-control" rows="3" maxlength="300" id="apReasonText"></textarea>' +
        '<div class="gw-ap-reason-btns"><button type="button" class="btn btn-sm btn-outline-secondary" id="apReasonCancel">취소</button>' +
        '<button type="button" class="btn btn-sm btn-primary" id="apReasonOk">보내기</button></div></div>' +
        '<div class="gw-form-error" id="apActError" role="alert"></div>';
      }

      h += '<div class="gw-sec-title">내용</div>' + fieldsHtml(d.fields);
      h += '<div class="gw-sec-title">진행</div>';
      if (d.steps.length) {
        h += '<ol class="gw-ap-steps">' + d.steps.map(function (s) {
          return '<li><span class="who">' + esc(s.actor || '-') + '</span> ' + esc(s.step) +
            (s.result && s.result !== s.step ? ' · ' + esc(s.result) : '') +
            (s.reason ? '<div class="why">' + esc(s.reason) + '</div>' : '') +
            '<span class="at">' + esc(when(s.at)) + '</span></li>';
        }).join('') + '</ol>';
      } else {
        h += '<div class="gw-log">아직 처리한 사람이 없어요.</div>';
      }
      if (d.slack_link) {
        h += '<a class="gw-ap-slack" href="' + esc(d.slack_link) + '" target="_blank" rel="noopener"><i class="fab fa-slack me-1" aria-hidden="true"></i>슬랙 카드 열기 — 첨부 파일·댓글은 카드 스레드에 있어요</a>';
      }
      $('apDoc').innerHTML = h;
      bindActions(d);
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
          t.focus();
          return;
        }
        var label = b.textContent;
        if (!window.confirm(d.title + '\n「' + label + '」 처리할까요? 슬랙 카드도 같이 바뀌어요.')) return;
        send(act, '');
      });
      $('apReasonCancel').addEventListener('click', function () { $('apReason').hidden = true; pending = null; });
      $('apReasonOk').addEventListener('click', function () {
        var r = $('apReasonText').value.trim();
        if (!r) { $('apActError').textContent = '내용을 적어 주세요.'; return; }
        send(pending, r);
      });
      function send(act, reason) {
        box.querySelectorAll('button').forEach(function (x) { x.disabled = true; });
        api('/api/approval/' + d.id + '/act', { method: 'POST', body: { action: act, reason: reason } })
          .then(function (j) { renderDoc(j.doc); load(); })
          .catch(function (e) {
            box.querySelectorAll('button').forEach(function (x) { x.disabled = false; });
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
    function arLoad() {
      var f = { form: $('arForm').value, from: $('arFrom').value, to: $('arTo').value, q: $('arQ').value.trim(), page: ar.page };
      var q = Object.keys(f).filter(function (k) { return f[k]; }).map(function (k) { return k + '=' + encodeURIComponent(f[k]); }).join('&');
      arPage.querySelector('.gw-ap-filters').classList.toggle('has-active-filters', !!(f.form || f.from || f.to || f.q));
      $('arForm').classList.toggle('on', !!f.form);
      api('/api/approval/archive?' + q).then(function (j) {
        if (!ar.loaded) {
          j.forms.forEach(function (x) { var o = document.createElement('option'); o.value = x; o.textContent = x; $('arForm').appendChild(o); });
          ar.loaded = true;
        }
        $('arCount').textContent = j.total.toLocaleString('ko-KR') + '건';
        $('arRows').innerHTML = j.items.length ? j.items.map(function (it) {
          return '<tr data-doc="' + esc(it.doc_no) + '" tabindex="0"><td class="muted">' + esc(it.doc_no) + '</td><td>' + esc(it.form) + '</td>' +
            '<td class="main">' + esc(it.title) + (it.attachments ? ' <i class="fas fa-paperclip muted" title="플렉스 첨부 ' + it.attachments + '개(파일은 옮기지 않았어요)"></i>' : '') + '</td>' +
            '<td>' + esc(it.project_code || '') + '</td><td>' + esc(it.partner || '') + '</td>' +
            '<td class="num">' + (it.amount ? won(it.amount) : '') + '</td><td>' + esc(it.writer) + '</td>' +
            '<td><span class="gw-ap-st flex-' + esc(it.status) + '">' + esc(it.status) + '</span></td>' +
            '<td class="muted">' + esc((it.created_at || '').slice(0, 10)) + '</td></tr>';
        }).join('') : '<tr><td colspan="9" class="gw-empty">맞는 문서가 없어요.</td></tr>';
        var pages = Math.max(1, Math.ceil(j.total / j.size));
        $('arPager').innerHTML = pages > 1
          ? '<button type="button" class="btn btn-sm btn-outline-secondary" data-p="' + (j.page - 1) + '"' + (j.page <= 1 ? ' disabled' : '') + '>이전</button>' +
            '<span>' + j.page + ' / ' + pages + '</span>' +
            '<button type="button" class="btn btn-sm btn-outline-secondary" data-p="' + (j.page + 1) + '"' + (j.page >= pages ? ' disabled' : '') + '>다음</button>'
          : '';
      }).catch(function (e) {
        $('arRows').innerHTML = '<tr><td colspan="9" class="gw-empty">' + esc(e.message) + '</td></tr>';
      });
    }
    $('arPager').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-p]');
      if (b && !b.disabled) { ar.page = Number(b.getAttribute('data-p')); arLoad(); }
    });
    ['arForm', 'arFrom', 'arTo'].forEach(function (id) { $(id).addEventListener('change', function () { ar.page = 1; arLoad(); }); });
    $('arQ').addEventListener('input', function () { clearTimeout(arTimer); arTimer = setTimeout(function () { ar.page = 1; arLoad(); }, 250); });
    $('arReset').addEventListener('click', function () {
      ['arForm', 'arFrom', 'arTo', 'arQ'].forEach(function (id) { $(id).value = ''; });
      ar.page = 1; arLoad();
    });
    function openArchive(no) {
      openDrawer('문서 ' + no, '<div class="gw-empty">불러오는 중…</div>');
      api('/api/approval/archive/' + encodeURIComponent(no)).then(function (j) {
        var d = j.doc;
        $('apDocTitle').textContent = d.title;
        var h = '<div class="gw-ap-doc-top"><span class="gw-ap-st flex-' + esc(d.status) + '">' + esc(d.status) + '</span>' +
          '<span class="gw-ap-doc-meta">' + esc(d.form) + ' · ' + esc(d.writer) + ' · ' + esc((d.created_at || '').slice(0, 16)) + '</span></div>';
        h += '<div class="gw-sec-title">내용</div>' + fieldsHtml(d.fields);
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
      if (tr && e.key === 'Enter') openArchive(tr.getAttribute('data-doc'));
    });
    arLoad();
  }
})();
