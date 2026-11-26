/* 그룹웨어 재무 > 이카운트 올리기(11-26). 넣을 매출전표·새 프로젝트를 이카운트 웹자료올리기 칸 순서 그대로 보여 주고,
   [표 복사]로 붙여넣기용 글(탭으로 칸 나눔)을 클립보드에 담는다. [넣음 표시]는 장부에 적어 목록에서 뺀다.
   규칙은 서버(services/ecount_sales.py) — 이카운트에는 아무것도 보내지 않는다. */
(function () {
  'use strict';

  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  var page = document.getElementById('fePage');
  if (!page || !document.getElementById('feSales')) return;
  var $ = function (id) { return document.getElementById(id); };
  var S = { sales: [], projects: [], tab: 'sales', pick: { sales: {}, projects: {} }, edit: {} };
  var MEMO_COL = 8;   // 매출전표 칸 중 적요 — 화면에서 고친 글로 복사(11-26 부장님 '수정 사유를 적을 수 없다')
  // 재시작 전 옛 템플릿이어도 '확인할 점' 칸·[제외] 버튼이 있게(11-26 추가)
  (function () {
    var head = document.querySelector('#fePaneSales thead tr');
    if (head && head.children.length === 9) head.insertAdjacentHTML('beforeend', '<th>확인할 점</th>');
    if (!$('feExclude')) {
      $('feDone').insertAdjacentHTML('afterend', '<button type="button" class="btn btn-sm btn-outline-secondary" id="feExclude" ' +
        'title="이카운트에 넣지 않기로 한 줄을 목록에서 빼요"><i class="fas fa-ban me-1" aria-hidden="true"></i>제외</button>');
    }
  })();

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function won(n) { return n == null || n === '' ? '' : Number(n).toLocaleString('ko-KR'); }
  function md(d) { return d ? Number(d.slice(5, 7)) + '/' + Number(d.slice(8, 10)) : '오늘'; }
  function msg(text, kind) { var m = $('feMsg'); m.className = 'gw-vh-msg ' + (kind || 'ok'); m.textContent = text || ''; }

  function api(url, body) {
    var init = { method: body ? 'POST' : 'GET', credentials: 'same-origin' };
    if (body) { init.headers = { 'Content-Type': 'application/json' }; init.body = JSON.stringify(body); }
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error(SESSION_ENDED);
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          throw new Error((j && j.message) || (body
            ? '처리하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'
            : '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'));
        }
        return j;
      });
    });
  }

  // ── 불러오기 ──
  function load(refresh) {
    $('feRefresh').disabled = true;
    return api('/api/finance/ecount' + (refresh ? '?refresh=1' : '')).then(function (j) {
      S.sales = j.sales || [];
      S.excludeOk = !!j.exclude_ok;
      S.projects = j.projects || [];
      // 목록이 바뀌면 없어진 줄은 고른 것에서 뺌
      ['sales', 'projects'].forEach(function (t) {
        var keys = {};
        S[t].forEach(function (r) { keys[r.key] = 1; });
        Object.keys(S.pick[t]).forEach(function (k) { if (!keys[k]) delete S.pick[t][k]; });
      });
      render();
    }).catch(function (e) {
      $('feSales').innerHTML = '<tr><td colspan="10" class="gw-empty">' + esc(e.message) + '</td></tr>';
      $('feProjects').innerHTML = '<tr><td colspan="3" class="gw-empty">' + esc(e.message) + '</td></tr>';
      msg(e.message, 'err');
    }).then(function () { $('feRefresh').disabled = false; });
  }

  function render() {
    var supply = S.sales.reduce(function (a, r) { return a + Number(r.supply || 0); }, 0);
    var warns = S.sales.filter(function (r) { return r.warn.length; }).length;
    var fresh = S.sales.filter(function (r) { return r.kind === '공사 확정'; }).length;
    var changed = {};   // 부가세 변경은 두 줄(11 감액 + 1Z) — 공사 수는 묶음으로 셈
    S.sales.forEach(function (r) { if (r.kind !== '공사 확정') changed[r.group || r.key] = 1; });
    $('feStats').innerHTML =
      '<div class="gw-stat gw-stat-main"><div class="k">넣을 매출전표</div><div class="v">' + S.sales.length + '<small> 건</small></div>' +
        '<div class="gw-fb-stat-sub">공급가 합 ' + won(supply) + '원</div></div>' +
      '<div class="gw-stat"><div class="k">새로 확정된 공사</div><div class="v">' + fresh + '<small> 건</small></div></div>' +
      '<div class="gw-stat"><div class="k">금액·부가세 바뀐 공사</div><div class="v">' + Object.keys(changed).length + '<small> 건</small></div></div>' +
      '<div class="gw-stat' + (warns ? ' gw-fb-stat-warn' : '') + '"><div class="k">확인 필요</div><div class="v">' + warns + '<small> 건</small></div></div>';
    $('feCntSales').textContent = S.sales.length;
    $('feCntProjects').textContent = S.projects.length;

    $('feSales').innerHTML = S.sales.length ? S.sales.map(function (r) {
      var on = !!S.pick.sales[r.key];
      var warn = r.warn.map(function (w) { return '<span class="gw-badge gw-badge-wait">' + esc(w) + '</span>'; }).join(' ');
      var kindCls = r.kind === '공사 확정' ? 'gw-badge-active' : (r.kind === '공사 취소' ? 'gw-badge-reject' : 'gw-badge-wait');
      return '<tr data-key="' + esc(r.key) + '"' + (r.warn.length ? ' class="gw-fe-warnrow"' : '') + '>' +
        '<td class="gw-fe-chk"><input type="checkbox" class="form-check-input" data-pick="sales" value="' + esc(r.key) + '"' + (on ? ' checked' : '') +
          ' aria-label="' + esc(r.code + ' 고르기') + '"></td>' +
        '<td class="text-nowrap">' + esc(md(r.date)) + '</td>' +
        '<td><span class="gw-badge ' + kindCls + '">' + esc(r.kind) + '</span>' +
          (r.lines > 1 ? '<div class="gw-fe-line">' + r.lines + '줄 중 ' + r.line + '</div>' : '') + '</td>' +
        '<td>' + esc(r.vat_type === '1Z' ? '1Z 기타매출' : '11 세금계산서') + '</td>' +
        '<td>' + esc(r.partner || '-') + '</td>' +
        '<td class="num">' + won(r.supply) + '</td>' +
        '<td class="num">' + (r.vat ? won(r.vat) : '<span class="text-muted">-</span>') + '</td>' +
        '<td class="gw-fe-memo">' + memoCell(r) + '</td>' +
        '<td class="text-nowrap"><strong>' + esc(r.code) + '</strong></td>' +
        '<td class="gw-fe-check">' + (warn || (r.trace ? '' : '<span class="text-muted">-</span>')) +
          (r.trace ? '<div class="gw-fe-trace">' + esc(r.trace) + '</div>' : '') + '</td></tr>';
    }).join('') : '<tr><td colspan="10" class="gw-empty">이카운트에 넣을 매출전표가 없어요.</td></tr>';

    $('feProjects').innerHTML = S.projects.length ? S.projects.map(function (r) {
      var on = !!S.pick.projects[r.key];
      return '<tr data-key="' + esc(r.key) + '">' +
        '<td class="gw-fe-chk"><input type="checkbox" class="form-check-input" data-pick="projects" value="' + esc(r.key) + '"' + (on ? ' checked' : '') +
          ' aria-label="' + esc(r.code + ' 고르기') + '"></td>' +
        '<td class="text-nowrap"><strong>' + esc(r.code) + '</strong></td>' +
        '<td class="gw-fe-memo">' + esc(r.name) + (r.warn.length ? '<div class="gw-fe-warn">' + r.warn.map(esc).join('<br>') + '</div>' : '') + '</td></tr>';
    }).join('') : '<tr><td colspan="3" class="gw-empty">새로 등록할 프로젝트가 없어요.</td></tr>';
    syncPick();
  }

  // 적요 — 사유를 적는 줄(수정 사유·취소 사유)만 사유 부분을 고칠 수 있게. 현장 주소(확정·재개)는 그대로(11-26 부장님)
  var REASON_RE = /^(수정 사유|취소 사유) ?: ?/;
  function reasonParts(r) {
    var m = REASON_RE.exec(r.memo || '');
    return m ? { head: m[1] + ' : ', text: r.memo.slice(m[0].length) } : null;
  }
  function memoOf(r) {
    var p = reasonParts(r);
    if (!p || S.edit[r.key] == null) return r.memo;
    return (p.head + S.edit[r.key].trim()).trim();
  }
  function memoCell(r) {
    var p = reasonParts(r);
    if (!p) return esc(r.memo);
    var edited = S.edit[r.key] != null;
    return '<div class="gw-fe-reason"><span class="gw-fe-reason-h">' + esc(p.head.replace(/ : $/, '')) + '</span>' +
      '<input type="text" class="form-control form-control-sm gw-fe-memo-in" maxlength="80" data-memo="' + esc(r.key) + '"' +
      ' value="' + esc(edited ? S.edit[r.key] : p.text) + '" placeholder="사유를 적어 주세요"' +
      ' aria-label="' + esc(r.code + ' ' + p.head.replace(/ : $/, '')) + '"' + (edited ? ' data-edited="1"' : '') + '></div>';
  }
  function cellsOf(r, t) {
    if (t !== 'sales') return r.cells;
    var c = r.cells.slice();
    c[MEMO_COL] = memoOf(r);
    return c;
  }
  // 적요 고치기 — 복사할 때 고친 글이 들어감(넣음 표시·새로 고침 전까지 기억)
  $('feSales').addEventListener('input', function (ev) {
    var k = ev.target.getAttribute && ev.target.getAttribute('data-memo');
    if (k == null) return;
    S.edit[k] = ev.target.value;
    ev.target.setAttribute('data-edited', '1');
  });

  function picked(t) { return S[t].filter(function (r) { return S.pick[t][r.key]; }); }
  function syncPick() {
    var t = S.tab, n = picked(t).length, all = S[t].length;
    $('fePick').textContent = n ? n + '줄 골랐어요' : (all ? '고르지 않으면 표 전체(' + all + '줄)를 복사해요' : '');
    $('feDone').hidden = t !== 'sales';
    $('feExclude').hidden = t !== 'sales' || !S.excludeOk;
    $('feDone').disabled = $('feExclude').disabled = !picked('sales').length;
    var box = $(t === 'sales' ? 'feAllSales' : 'feAllProjects');
    box.checked = !!all && n === all;
    box.indeterminate = n > 0 && n < all;
    $('feCopy').disabled = !all;
  }

  // ── 탭 ──
  function setTab(tab, focus) {
    S.tab = tab;
    $('feTabs').querySelectorAll('[role=tab]').forEach(function (b) {
      var on = b.getAttribute('data-tab') === tab;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
      b.tabIndex = on ? 0 : -1;
      if (on && focus) b.focus();
    });
    $('fePaneSales').hidden = tab !== 'sales';
    $('fePaneProjects').hidden = tab !== 'projects';
    syncPick();
  }
  $('feTabs').addEventListener('click', function (ev) {
    var b = ev.target.closest('[role=tab]');
    if (b) setTab(b.getAttribute('data-tab'));
  });
  $('feTabs').addEventListener('keydown', function (ev) {
    if (ev.key !== 'ArrowLeft' && ev.key !== 'ArrowRight' && ev.key !== 'Home' && ev.key !== 'End') return;
    ev.preventDefault();
    setTab(ev.key === 'Home' ? 'sales' : ev.key === 'End' ? 'projects' : (S.tab === 'sales' ? 'projects' : 'sales'), true);
  });

  // ── 고르기 ──
  page.addEventListener('change', function (ev) {
    var el = ev.target;
    if (el.id === 'feAllSales' || el.id === 'feAllProjects') {
      var t = el.id === 'feAllSales' ? 'sales' : 'projects';
      S.pick[t] = {};
      if (el.checked) S[t].forEach(function (r) { S.pick[t][r.key] = 1; });
      page.querySelectorAll('input[data-pick="' + t + '"]').forEach(function (c) { c.checked = el.checked; });
      syncPick();
      return;
    }
    var t2 = el.getAttribute && el.getAttribute('data-pick');
    if (!t2) return;
    // 두 줄 전표(부가세 변경)는 같이 골라짐 — 한 줄만 넣으면 이카운트 금액이 어긋남
    var row = S[t2].filter(function (r) { return r.key === el.value; })[0];
    var grp = row && row.group;
    S[t2].forEach(function (r) {
      if (r.key !== el.value && !(grp && r.group === grp)) return;
      if (el.checked) S.pick[t2][r.key] = 1; else delete S.pick[t2][r.key];
      var c = page.querySelector('input[data-pick="' + t2 + '"][value="' + CSS.escape(r.key) + '"]');
      if (c) c.checked = el.checked;
    });
    syncPick();
  });
  // 줄 아무 데나 눌러도 고르기(체크 칸·글자 고르기 빼고)
  ['feSales', 'feProjects'].forEach(function (id) {
    $(id).addEventListener('click', function (ev) {
      if (ev.target.closest('input') || String(window.getSelection && window.getSelection()) !== '') return;
      var tr = ev.target.closest('tr[data-key]');
      var c = tr && tr.querySelector('input[data-pick]');
      if (c) { c.checked = !c.checked; c.dispatchEvent(new Event('change', { bubbles: true })); }
    });
  });

  // ── 표 복사(이카운트 칸 순서 그대로, 탭으로 칸 나눔 — 엑셀에서 복사한 것과 같은 모양) ──
  function clean(v) { return String(v == null ? '' : v).replace(/[\t\r\n]+/g, ' ').trim(); }
  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text);
    return new Promise(function (ok, fail) {
      var ta = document.createElement('textarea');
      ta.value = text; ta.setAttribute('readonly', ''); ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy') ? ok() : fail(new Error('복사하지 못했어요.')); } catch (e) { fail(e); }
      document.body.removeChild(ta);
    });
  }
  $('feCopy').addEventListener('click', function () {
    var t = S.tab, rows = picked(t);
    if (!rows.length) rows = S[t];
    if (!rows.length) return;
    var text = rows.map(function (r) { return cellsOf(r, t).map(clean).join('\t'); }).join('\r\n');
    copyText(text).then(function () {
      msg((t === 'sales' ? '매출전표 ' : '프로젝트 ') + rows.length + '줄을 복사했어요. 이카운트 웹자료올리기 표의 첫 칸을 누르고 붙여넣어 주세요.', 'ok');
    }).catch(function () { msg('복사하지 못했어요. 브라우저에서 클립보드 권한을 허용해 주세요.', 'err'); });
  });

  // ── 넣음 표시 · 제외(이카운트에 넣지 않기로 한 줄 — 장부엔 '제외'로 남음) ──
  function finish(btn, exclude) {
    var rows = picked('sales');
    if (!rows.length) return;
    var ask = exclude
      ? '고른 매출전표 ' + rows.length + '건을 이카운트에 넣지 않고 목록에서 제외할까요? 장부에는 "제외"로 남아요.'
      : '고른 매출전표 ' + rows.length + '건을 이카운트에 넣었다고 표시할까요? 목록에서 빠져요.';
    if (!window.confirm(ask)) return;
    btn.disabled = true;
    api('/api/finance/ecount/done', { keys: rows.map(function (r) { return r.key; }), exclude: exclude }).then(function (j) {
      S.pick.sales = {}; S.pick.projects = {};
      rows.forEach(function (r) { delete S.edit[r.key]; });
      msg(j.done + '건을 ' + (exclude ? '제외했어요.' : '넣음으로 표시했어요.') +
        (j.skipped ? ' ' + j.skipped + '건은 그사이 목록이 바뀌어 건너뛰었어요.' : ''), 'ok');
      return load(false);
    }).catch(function (e) { msg(e.message, 'err'); btn.disabled = false; });
  }
  $('feDone').addEventListener('click', function () { finish(this, false); });
  $('feExclude').addEventListener('click', function () { finish(this, true); });

  $('feRefresh').addEventListener('click', function () { msg(''); load(true); });
  load(false);
})();
