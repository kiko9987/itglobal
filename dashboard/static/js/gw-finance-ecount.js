/* 그룹웨어 재무 > 이카운트 올리기(11-26). 넣을 매출전표·새 프로젝트를 이카운트 웹자료올리기 칸 순서 그대로 보여 주고,
   [표 복사]로 붙여넣기용 글(탭으로 칸 나눔)을 클립보드에 담는다. [넣음 표시]는 장부에 적어 목록에서 뺀다.
   규칙은 서버(services/ecount_sales.py) — 이카운트에는 아무것도 보내지 않는다. */
(function () {
  'use strict';

  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  var page = document.getElementById('fePage');
  if (!page || !document.getElementById('feSales')) return;
  var $ = function (id) { return document.getElementById(id); };
  var S = { sales: [], projects: [], tab: 'sales', pick: { sales: {}, projects: {} }, edit: {}, first: true, later: 0,
            af: { side: 'sales', cat: '', st: 'open', q: '', more: 0 } };
  var MEMO_COL = 8;   // 매출전표 칸 중 적요 — 화면에서 고친 글로 복사(11-26 부장님 '수정 사유를 적을 수 없다')
  // 재시작 전 옛 템플릿이어도 '확인할 점' 칸·[제외] 버튼이 있게(11-26 추가)
  (function () {
    // 메뉴 이름 '이카운트 업로드'(11-26 부장님) — 재시작 전 옛 템플릿 제목도 바꿈
    var h1 = document.querySelector('.gw-page-title');
    if (h1 && h1.textContent.trim() === '이카운트 올리기') h1.textContent = '이카운트 업로드';
    document.title = document.title.replace('이카운트 올리기', '이카운트 업로드');
    // 탭 순서 ① 프로젝트 등록 → ② 매출전표 I — 새 프로젝트를 먼저 저장해야 매출전표가 들어감(11-26 부장님)
    var tabs = $('feTabs'), pb = tabs && tabs.querySelector('[data-tab="projects"]'), sb = tabs && tabs.querySelector('[data-tab="sales"]');
    if (pb && sb && sb.compareDocumentPosition(pb) & Node.DOCUMENT_POSITION_FOLLOWING) tabs.insertBefore(pb, sb);
    // 탭 줄은 숫자 상자 바로 아래(입출금과 같은 자리, 11-28 부장님). ③ 에서도 숫자 상자를 대조 숫자로 채워 탭 줄이 튀지 않음.
    // 재시작 전 옛 템플릿(탭이 맨 위)도 여기서 옮김
    var stats = $('feStats');
    if (tabs && stats && stats.nextElementSibling !== tabs) stats.parentNode.insertBefore(tabs, stats.nextSibling);
    if (pb && pb.firstChild && pb.firstChild.nodeType === 3 && pb.firstChild.nodeValue.indexOf('①') < 0) pb.firstChild.nodeValue = '① ' + pb.firstChild.nodeValue;
    if (sb && sb.firstChild && sb.firstChild.nodeType === 3 && sb.firstChild.nodeValue.indexOf('②') < 0) sb.firstChild.nodeValue = '② ' + sb.firstChild.nodeValue;
    // ③ 이카운트 대조 탭(11-27) — 재시작 전 옛 템플릿이면 탭·칸을 만들어 둠
    if (tabs && !tabs.querySelector('[data-tab="audit"]')) {
      tabs.insertAdjacentHTML('beforeend', '<button type="button" role="tab" data-tab="audit" aria-selected="false" tabindex="-1" ' +
        'aria-controls="fePaneAudit">③ 이카운트 대조 <b id="feCntAudit">-</b></button>');
    }
    if (!$('fePaneAudit') && $('fePaneProjects')) {
      $('fePaneProjects').insertAdjacentHTML('afterend', '<div id="fePaneAudit" role="tabpanel" aria-label="이카운트 대조" hidden>' +
        '<div id="feAudit"><p class="gw-empty">불러오는 중…</p></div></div>');
    }
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
      // 바로 넣을 줄을 위로, 확인이 필요한 줄(확인할 점·원래 전표 고치기)은 아래로 — 표 안 차례는 그대로(11-26 부장님)
      S.sales = (j.sales || []).map(function (r, i) { r._i = i; return r; }).sort(function (a, b) {
        return (ready(a) ? 0 : 1) - (ready(b) ? 0 : 1) || a._i - b._i;
      });
      S.excludeOk = !!j.exclude_ok;
      S.later = j.later || 0;   // 오늘 확정·변경분 — SB 는 전날 것을 다음 날 넣음(11-28), 내일 목록에 나옴
      S.projects = j.projects || [];
      // 처음 열 때 새 프로젝트가 있으면 ① 프로젝트 등록부터, 없으면 ② 매출전표
      if (S.first) { S.first = false; setTab(S.projects.length ? 'projects' : 'sales'); }
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
    S.salesStats =
      '<div class="gw-stat gw-stat-main"><div class="k">넣을 매출전표</div><div class="v">' + S.sales.length + '<small> 건</small></div>' +
        '<div class="gw-fb-stat-sub">공급가 합 ' + won(supply) + '원</div></div>' +
      '<div class="gw-stat"><div class="k">새로 확정된 공사</div><div class="v">' + fresh + '<small> 건</small></div></div>' +
      '<div class="gw-stat"><div class="k">금액·부가세 바뀐 공사</div><div class="v">' + Object.keys(changed).length + '<small> 건</small></div></div>' +
      '<div class="gw-stat' + (warns ? ' gw-fb-stat-warn' : '') + '"><div class="k">확인 필요</div><div class="v">' + warns + '<small> 건</small></div></div>';
    paintStats();
    $('feCntSales').textContent = S.sales.length;
    $('feCntProjects').textContent = S.projects.length;

    var divided = false;
    $('feSales').innerHTML = S.sales.length ? S.sales.map(function (r) {
      var on = !!S.pick.sales[r.key];
      var divider = '';
      if (ready(r) && r === S.sales[0]) {
        var nReady = S.sales.filter(ready).length;
        divider = '<tr class="gw-fe-divider gw-fe-divider-ok"><td colspan="10"><i class="fas fa-circle-check me-1" aria-hidden="true"></i>' +
          '바로 붙여넣어도 되는 줄 ' + nReady + '건 — 아무것도 고르지 않고 [표 복사]를 누르면 이 줄들만 복사돼요.</td></tr>';
      }
      if (!ready(r) && !divided) {
        divided = true;
        divider = '<tr class="gw-fe-divider"><td colspan="10"><i class="fas fa-triangle-exclamation me-1" aria-hidden="true"></i>' +
          '확인·수정이 필요한 줄 ' + S.sales.filter(function (x) { return !ready(x); }).length + '건 — 확인할 점을 보고, 원래 전표를 고칠 줄은 먼저 고친 뒤 골라서 복사해 주세요. 고르지 않으면 복사되지 않아요.</td></tr>';
      }
      var warn = r.warn.map(function (w) { return '<span class="gw-badge gw-badge-wait">' + esc(w) + '</span>'; }).join(' ');
      var kindCls = r.kind === '공사 확정' ? 'gw-badge-active' : (r.kind === '공사 취소' ? 'gw-badge-reject' : 'gw-badge-wait');
      return divider + '<tr data-key="' + esc(r.key) + '"' + (r.warn.length ? ' class="gw-fe-warnrow"' : '') + '>' +
        '<td class="gw-fe-chk"><input type="checkbox" class="form-check-input" data-pick="sales" value="' + esc(r.key) + '"' + (on ? ' checked' : '') +
          ' aria-label="' + esc(r.code + ' 고르기') + '"></td>' +
        '<td class="text-nowrap">' + esc(md(r.date)) + '</td>' +
        '<td><span class="gw-badge ' + kindCls + '">' + esc(r.kind) + '</span>' +
          (r.lines > 1 ? '<div class="gw-fe-line">' + r.lines + '줄 중 ' + r.line + '</div>' : '') + '</td>' +
        '<td>' + (r.cells ? esc(r.vat_type === '1Z' ? '1Z 기타매출' : '11 세금계산서') : '<span class="gw-badge gw-fe-fix">원래 전표 고치기</span>') + '</td>' +
        '<td>' + esc(r.partner || '-') + '</td>' +
        '<td class="num">' + (r.cells ? won(r.supply) : '<span class="text-muted">-</span>') + '</td>' +
        '<td class="num">' + (r.cells && r.vat ? won(r.vat) : '<span class="text-muted">-</span>') + '</td>' +
        '<td class="gw-fe-memo">' + memoCell(r) + '</td>' +
        '<td class="text-nowrap"><strong>' + esc(r.code) + '</strong>' + (ready(r) ? '' : copyBtn('code', r.code, '코드 복사')) + '</td>' +
        '<td class="gw-fe-check">' + (warn || (r.trace ? '' : '<span class="text-muted">-</span>')) +
          (r.trace ? '<div class="gw-fe-trace">' + esc(r.trace) + '</div>' : '') + '</td></tr>';
    }).join('') : '<tr><td colspan="10" class="gw-empty">이카운트에 넣을 매출전표가 없어요.</td></tr>';
    if (S.later) {
      $('feSales').insertAdjacentHTML('beforeend', '<tr class="gw-fe-divider gw-fe-later"><td colspan="10">' +
        '<i class="fas fa-clock me-1" aria-hidden="true"></i>오늘 확정되거나 금액이 바뀐 ' + S.later + '건은 내일 목록에 나와요.</td></tr>');
    }

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
  // 원래 전표 고치기(부가세 있음↔없음, SB 방식 11-26) — 고칠 내용을 보여 주고, 붙여넣을 줄이 없으면 입력칸도 없음
  // 복사 아이콘 — PM 프로젝트 코드 복사와 같은 모양(흐린 아이콘, 누르면 잠깐 체크). 줄 고르기는 안 건드림
  var COPY_SVG = '<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true">' +
    '<rect x="5.5" y="5.5" width="8" height="8" rx="1.5"/><path d="M10.5 3.5v-.5a1.5 1.5 0 0 0-1.5-1.5H4A1.5 1.5 0 0 0 2.5 3v5A1.5 1.5 0 0 0 4 9.5h.5"/></svg>';
  var CHECK_SVG = '<svg width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M3 8.5l3.2 3L13 4.5"/></svg>';
  function copyBtn(kind, value, label, text) {
    return '<button type="button" class="gw-fe-copy" data-copy-kind="' + kind + '" data-copy="' + esc(value) + '" title="' + esc(label) +
      '" aria-label="' + esc(label) + '">' + COPY_SVG + (text ? '<span>' + esc(text) + '</span>' : '') + '</button>';
  }
  // 적요 — 바로 넣을 줄은 사유를 고칠 수 있게, 확인·수정 줄은 글로만(고칠 일 없음, 11-26 부장님) + 고칠 내용 안내
  function memoCell(r) {
    var line = r.cells && !ready(r) ? ' ' + copyBtn('line', r.key, '이 줄만 복사(이카운트 웹자료올리기에 붙여넣기)', '이 줄 복사') : '';
    var guide = r.guide ? '<div class="gw-fe-guide"><i class="fas fa-pen-to-square me-1" aria-hidden="true"></i>' + esc(r.guide) + line + '</div>' : '';
    var memo = !r.cells ? '' : ready(r) ? memoInput(r) : '<div class="gw-fe-memo-text">' + esc(r.memo) + '</div>';
    return memo + guide + (r.cells && !ready(r) && !r.guide ? line : '');
  }
  function memoInput(r) {
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

  function ready(r) { return !!r.cells && !(r.warn && r.warn.length) && !r.guide; }
  function picked(t) { return S[t].filter(function (r) { return S.pick[t][r.key]; }); }
  function copyable(rows) { return rows.filter(function (r) { return !!r.cells; }); }   // 고치기만 하는 줄은 복사 안 함
  // 고르지 않고 복사할 때 들어가는 줄 — 매출전표는 확인할 점·고치기 안내가 없는 줄만(11-26 부장님 '전체 복사하면 안 되잖아')
  function defaults(t) { return t === 'sales' ? S.sales.filter(ready) : S[t]; }
  function syncPick() {
    // 대조 탭은 표 복사·넣음 표시와 상관없음 — 도구 줄·안내·통계를 숨김
    var audit = S.tab === 'audit';
    ['.gw-fe-tools', '.gw-fe-steps', '.gw-vh-foot'].forEach(function (q) { var el = document.querySelector(q); if (el) el.hidden = audit; });
    if (audit) return;
    var t = S.tab, n = picked(t).length, all = S[t].length, can = defaults(t).length;
    $('fePick').textContent = n ? n + '줄 골랐어요'
      : t === 'sales' ? (can ? '고르지 않으면 바로 넣을 줄(' + can + '줄)만 복사해요. 확인이 필요한 줄은 골라서 복사해요' : '바로 넣을 줄이 없어요. 확인이 필요한 줄은 골라서 복사해요')
      : (can ? '고르지 않으면 표 전체(' + can + '줄)를 복사해요' : '');
    $('feDone').hidden = t !== 'sales';
    $('feExclude').hidden = t !== 'sales' || !S.excludeOk;
    $('feDone').disabled = $('feExclude').disabled = !picked('sales').length;
    var box = $(t === 'sales' ? 'feAllSales' : 'feAllProjects');
    box.checked = !!all && n === all;
    box.indeterminate = n > 0 && n < all;
    $('feCopy').disabled = !copyable(S[t]).length;
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
    if ($('fePaneAudit')) $('fePaneAudit').hidden = tab !== 'audit';
    if (tab === 'audit' && !S.auditLoaded) loadAudit();
    paintStats();
    syncPick();
  }
  $('feTabs').addEventListener('click', function (ev) {
    var b = ev.target.closest('[role=tab]');
    if (b) setTab(b.getAttribute('data-tab'));
  });
  $('feTabs').addEventListener('keydown', function (ev) {
    if (ev.key !== 'ArrowLeft' && ev.key !== 'ArrowRight' && ev.key !== 'Home' && ev.key !== 'End') return;
    ev.preventDefault();
    var order = ['projects', 'sales', 'audit'], i = order.indexOf(S.tab);
    setTab(ev.key === 'Home' ? order[0] : ev.key === 'End' ? order[order.length - 1]
      : order[(i + (ev.key === 'ArrowRight' ? 1 : order.length - 1)) % order.length], true);
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
      if (ev.target.closest('input, button') || String(window.getSelection && window.getSelection()) !== '') return;
      var tr = ev.target.closest('tr[data-key]');
      var c = tr && tr.querySelector('input[data-pick]');
      if (c) { c.checked = !c.checked; c.dispatchEvent(new Event('change', { bubbles: true })); }
    });
  });

  // ── 표 복사(이카운트 칸 순서 그대로, 탭으로 칸 나눔 — 엑셀에서 복사한 것과 같은 모양) ──
  function clean(v) { return String(v == null ? '' : v).replace(/[\t\r\n]+/g, ' ').trim(); }
  // 클립보드 쓰기가 막히면(권한·창 포커스 — 11-26 내장 브라우저에서 'Document is not focused') 예전 방식으로 한 번 더
  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text).catch(function () { return legacyCopy(text); });
    }
    return legacyCopy(text);
  }
  function legacyCopy(text) {
    return new Promise(function (ok, fail) {
      var ta = document.createElement('textarea');
      ta.value = text; ta.setAttribute('readonly', ''); ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy') ? ok() : fail(new Error('복사하지 못했어요.')); } catch (e) { fail(e); }
      document.body.removeChild(ta);
    });
  }
  $('feCopy').addEventListener('click', function () {
    var t = S.tab, rows = copyable(picked(t));
    if (!picked(t).length) rows = defaults(t);
    if (!rows.length) { msg('고른 줄은 이카운트에서 원래 전표를 고치는 줄이라 복사할 게 없어요.', 'err'); return; }
    var text = rows.map(function (r) { return cellsOf(r, t).map(clean).join('\t'); }).join('\r\n');
    copyText(text).then(function () {
      msg(t === 'sales'
        ? '매출전표 ' + rows.length + '줄을 복사했어요. 이카운트 매출전표 I › 웹자료올리기 표의 첫 칸을 누르고 붙여넣어 주세요.'
        : '프로젝트 ' + rows.length + '줄을 복사했어요. 이카운트 프로젝트등록 › 웹자료올리기에 붙여넣고 저장한 뒤 ② 매출전표 I 탭으로 넘어가 주세요.', 'ok');
    }).catch(function () { msg('복사하지 못했어요. 브라우저에서 클립보드 권한을 허용해 주세요.', 'err'); });
  });

  // 복사 아이콘 — 코드 또는 그 줄 하나(탭으로 칸 나눔)
  $('feSales').addEventListener('click', function (ev) {
    var b = ev.target.closest('.gw-fe-copy');
    if (!b) return;
    var kind = b.getAttribute('data-copy-kind'), v = b.getAttribute('data-copy'), text = v;
    if (kind === 'line') {
      var r = S.sales.filter(function (x) { return x.key === v; })[0];
      if (!r || !r.cells) return;
      text = cellsOf(r, 'sales').map(clean).join('\t');
    }
    var icon = b.innerHTML;
    copyText(text).then(function () {
      b.innerHTML = CHECK_SVG + (kind === 'line' ? '<span>복사됨</span>' : '');
      b.classList.add('done');
      msg(kind === 'line' ? '이 줄을 복사했어요. 원래 전표를 고친 뒤 이카운트 매출전표 I › 웹자료올리기에 붙여넣어 주세요.'
        : v + ' 코드를 복사했어요. 이카운트 회계거래조회 프로젝트 칸에 붙여넣으면 원래 전표를 찾을 수 있어요.', 'ok');
      setTimeout(function () { b.innerHTML = icon; b.classList.remove('done'); }, 1500);
    }).catch(function () { msg('복사하지 못했어요. 브라우저에서 클립보드 권한을 허용해 주세요.', 'err'); });
  });

  // ── ③ 이카운트 대조(11-27 부장님) — 이카운트 매출과 공사 현황이 어긋난 공사를 줄마다 처리 ──
  //   규칙은 services/ecount_audit.py. 이카운트를 읽어 맞춘 결과(대조 날짜)를 보여 주고, [이카운트 고침]·[시트가 틀림]·[그대로 둠]을 남긴다.
  var ACTS = [['ecount_fixed', '이카운트 고침', 'btn-outline-primary'], ['sheet_wrong', '시트가 틀림', 'btn-outline-danger'],
              ['keep', '그대로 둠', 'btn-outline-secondary']];
  function amt(s, v, buy) {
    if (s == null && v == null) return '<span class="text-muted">-</span>';
    if (buy) return s == null ? '<span class="text-muted">-</span>' : won(s);
    return won(s || 0) + '<div class="gw-fa-vat">부가세 ' + (v ? won(v) : '없음') + '</div>';
  }
  // 근거 칸(11-27 부장님 '근거가 눈에 딱 안 들어와') — 한 문단이던 것을 나눈다:
  //   ① 차이 한 줄(무엇이 얼마나) ② 안내 문장 ③ 이카운트 전표(날짜·금액 한 줄씩, 빼는 전표는 빨강) ④ 계산서 단계 꼬리표 ⑤ 덧붙임 꼬리표.
  //   대조 결과는 글로 저장돼 있어(ecount_audit_items.note·evidence) 화면에서 나눠 읽는다.
  var FA_TAGS = [['SH 정산 공사예요.', 'SH 정산'], ['이카운트 업로드 ② 매출전표 목록에도 있어요.', '② 매출전표에도 있음']];
  var FA_INV = { '발행': 'ok', '미발행': 'warn', 'N입금': 'cash', '카드': 'card' };
  var FA_INV_LABEL = { 'N입금': '현금(계산서 없음)' };
  function faDiff(r) {
    if (r.eco_supply == null || r.sheet_supply == null) return '';
    if (r.side === 'buy' || r.side === 'cost') {
      var bd = (r.eco_supply || 0) - (r.sheet_supply || 0);
      return Math.abs(bd) > 10 ? '<div class="gw-fa-diff">이카운트 매입이 <b>' + won(Math.abs(bd)) + '원 ' + (bd > 0 ? '많아요' : '적어요') + '</b></div>' : '';
    }
    var es = r.eco_supply || 0, ss = r.sheet_supply || 0, ev = r.eco_vat || 0, sv = r.sheet_vat || 0;
    var d = es - ss, out = [];
    if (Math.abs((es + ev) - (ss + sv)) <= 2 && Math.abs(d) >= 1) {
      out.push('<b>합계는 같아요</b> — 공급가·부가세 나눔만 달라요');          // 나눔만 다름: 공급가 차이로 보이면 오해
    } else {
      if (Math.abs(d) >= 1) out.push('이카운트 공급가가 <b>' + won(Math.abs(d)) + '원 ' + (d > 0 ? '많아요' : '적어요') + '</b>');
      if (!!ev !== !!sv) out.push('부가세가 <b>' + (ev ? '이카운트에만' : '공사 현황에만') + ' 있어요</b>');
      else if (ev && Math.abs(ev - sv) >= 1 && Math.abs(d) < 1) out.push('부가세가 <b>' + won(Math.abs(ev - sv)) + '원 달라요</b>');
    }
    return out.length ? '<div class="gw-fa-diff">' + out.join(' · ') + '</div>' : '';
  }
  function faWhy(r) {
    var note = String(r.note || ''), tags = [], vouchers = [];
    FA_TAGS.forEach(function (t) { if (note.indexOf(t[0]) !== -1) { tags.push(t[1]); note = note.replace(t[0], ''); } });
    // '이카운트 공급가 전표: 26.05.11 6,000,000 · …' / 매입은 '이카운트 매입 전표: 26.11.20 80,000 지급운반비 · …'(11-28)
    var m = note.match(/이카운트 (?:공급가|매입) 전표:\s*([^]*?)(?=\s(?:SH 정산 공사예요\.|이카운트 업로드 ②)|$)/);
    var vmore = '';
    if (m) {
      note = note.slice(0, m.index) + note.slice(m.index + m[0].length);
      m[1].split(' · ').forEach(function (v) {
        var p = v.trim().match(/^(\d{2}\.\d{2}\.\d{2})\s+([−-]?[\d,]+)(?:\s+(.+))?$/);
        if (p) vouchers.push({ day: p[1], amt: p[2], who: p[3] || '', neg: /^[−-]/.test(p[2]) });
        else if (/^외 \d+장$/.test(v.trim())) vmore = v.trim();
      });
    }
    note = note.trim();
    var inv = [];
    var e = String(r.evidence || '').replace(/^계산서:\s*/, '');
    if (e) e.split(' · ').forEach(function (x) {
      var q = x.trim().match(/^(\S+)\s+(.+)$/);
      if (q && q[2] !== '-') inv.push([q[1], q[2]]);
    });
    var html = faDiff(r) + (note ? '<div class="gw-fa-note">' + esc(note) + '</div>' : '');
    if (vouchers.length) {
      html += '<div class="gw-fa-sub">이카운트 전표 ' + vouchers.length + '장</div><ul class="gw-fa-vlist' + (vouchers.some(function (v) { return v.who; }) ? ' gw-fa-vlist3' : '') + '">' + vouchers.map(function (v) {
        return '<li' + (v.neg ? ' class="neg"' : '') + '><span>' + esc(v.day) + '</span><b>' + esc(v.amt.replace('-', '−')) + '</b>' +
          (v.who ? '<i>' + esc(v.who) + '</i>' : '') + '</li>';
      }).join('') + '</ul>' + (vmore ? '<div class="gw-fa-sub">' + esc(vmore) + '</div>' : '');
    }
    if (/^원가:/.test(r.evidence || '')) {          // 매입 — 공사 현황 원가 칸 나눔
      html += '<div class="gw-fa-inv"><span class="gw-fa-sub">공사 현황 원가</span>' + r.evidence.replace(/^원가:\s*/, '').split(' · ')
        .map(function (x) { return '<span class="gw-fa-chip">' + esc(x) + '</span>'; }).join('') + '</div>';
    } else if (r.evidence) {
      html += '<div class="gw-fa-inv"><span class="gw-fa-sub">계산서</span>' + (inv.length ? inv.map(function (x) {
        return '<span class="gw-fa-chip ' + (FA_INV[x[1]] || '') + '">' + esc(x[0]) + ' ' + esc(FA_INV_LABEL[x[1]] || x[1]) + '</span>';
      }).join('') : '<span class="text-muted">기록 없음</span>') + '</div>';
    }
    if (tags.length) html += '<div class="gw-fa-tags">' + tags.map(function (t) { return '<span class="gw-fa-tag">' + esc(t) + '</span>'; }).join('') + '</div>';
    // 처리 기록(11-27) — 되돌린 앞 처리·메모, 다시 대조한 결과가 날짜순으로 쌓임
    var hist = String(r.history || '').split('\n').filter(function (x) { return x.trim(); });
    if (hist.length) {
      var older = hist.slice(0, -1);
      html += '<div class="gw-fa-sub">기록</div><ul class="gw-fa-hist"><li>' + esc(hist[hist.length - 1]) + '</li></ul>' +
        (older.length ? '<details class="gw-fa-older"><summary>이전 기록 ' + older.length + '건</summary><ul class="gw-fa-hist">' +
          older.map(function (x) { return '<li>' + esc(x) + '</li>'; }).join('') + '</ul></details>' : '');
    }
    return html;
  }
  function auditRow(r) {
    var done = r.status !== 'open';
    var codeCell = '<strong>' + esc(r.code) + '</strong>' + copyBtn('code', r.code, '코드 복사') +
      (r.ecount_code && r.ecount_code !== r.code ? '<div class="gw-fa-eco">이카운트 <span>' + esc(r.ecount_code) + '</span>' +
        copyBtn('code', r.ecount_code, '이카운트 코드 복사') + '</div>' : '');
    var why = faWhy(r);
    var act = done
      ? '<span class="gw-badge ' + (r.status === 'sheet_wrong' ? 'gw-badge-reject' : 'gw-badge-active') + '">' + esc(r.status_label) + '</span>' +
        '<div class="gw-fa-by">' + esc(r.resolved_by || '') + (r.memo ? ' · ' + esc(r.memo) : '') + '</div>' +
        (r.status === 'resolved' ? '' :     // '맞음'은 밤 자동 대조가 붙인 것 — 다시 다르면 자동으로 처리 전이 됨
          '<button type="button" class="btn btn-link btn-sm p-0 gw-fa-undo" data-aid="' + r.id + '" data-act="open">되돌리기</button>')
      : r.side === 'cost'     // 시트 원가 — SB 가 누를 일 없음(11-28), 시트 원가를 고치면 다음 대조에서 자동으로 맞음
      ? '<div class="gw-fa-by">공사 현황 원가를 채우거나 고치면 다음 대조에서 자동으로 맞음이 돼요.</div>'
      : '<input type="text" class="form-control form-control-sm gw-fa-memo" maxlength="200" placeholder="메모(선택)" aria-label="' + esc(r.code + ' 메모') + '">' +
        '<div class="gw-fa-btns">' + ACTS.map(function (a) {
          return '<button type="button" class="btn btn-sm ' + a[2] + '" data-aid="' + r.id + '" data-act="' + a[0] + '">' + a[1] + '</button>';
        }).join('') + '</div>';
    return '<tr data-aid="' + r.id + '"' + (done ? ' class="gw-fa-done"' : '') + '>' +
      '<td class="text-nowrap">' + codeCell + '</td>' +
      '<td class="gw-fe-memo">' + esc(r.site || '') + '</td>' +
      '<td class="num">' + amt(r.eco_supply, r.eco_vat, r.side !== 'sales') + '</td>' +
      '<td class="num">' + amt(r.sheet_supply, r.sheet_vat, r.side !== 'sales') + '</td>' +
      '<td class="gw-fe-memo gw-fa-why">' + (why || '<span class="text-muted">-</span>') + '</td>' +
      '<td class="gw-fa-act">' + act + '</td></tr>';
  }
  var AUDIT_PAGE = 150;     // 매입은 지난 차이가 많아(11-28 첫 대조 약 1,100건) 나눠 그림
  function auditHead(A) {
    var R = A.run;
    if (A.batch !== 'auto' || !R) {
      var b = Number(A.batch.slice(5, 7)) + '/' + Number(A.batch.slice(8, 10));
      // 건수는 위 숫자 상자에 있어 뺌, 버튼 설명은 접기(11-28 부장님, 토스 원칙)
      return '<div class="gw-callout gw-fa-head"><b>' + esc(b) + ' 대조</b> — 이카운트 매출전표(2025년~)를 공사별로 합산해 공사 현황과 맞춰 본 결과예요.</div>' +
        auditHelp('이카운트를 고쳤으면 [이카운트 고침], 공사 현황(시트)이 틀렸으면 [시트가 틀림], 차이가 맞으면 [그대로 둠]을 눌러 주세요.');
    }
    var warn = R.status === 'failed'
      ? '<div class="gw-fa-runwarn"><i class="fas fa-triangle-exclamation me-1" aria-hidden="true"></i>' + esc(R.at) +
        ' 대조는 이카운트를 읽지 못했어요. 아래는 ' + esc(R.ok_at || '이전') + ' 대조 결과예요.</div>' : '';
    return '<div class="gw-callout gw-fa-head">' + warn + '<b>' + esc(R.ok_at || R.at) + ' 자동 대조</b> · 평일 밤마다 이카운트를 읽어 공사 현황과 맞춰 봐요.</div>' +
      auditHelp('같아지면 다음 대조 때 자동으로 <b>맞음</b>이 돼요.', '[이카운트 고침]을 눌러도 다음 대조에서 아직 다르면 <b>처리 전</b>으로 돌아와요.');
  }
  // ③ 처리 방법 — 접힌 도움말. 다시 그려도 열어 둔 상태 유지
  function auditHelp() {
    var li = Array.prototype.map.call(arguments, function (t) { return '<li>' + t + '</li>'; }).join('');
    return '<details class="gw-help sm gw-fa-help"' + (S.auditHelpOpen ? ' open' : '') + '><summary><i class="fas fa-circle-question" aria-hidden="true"></i> 처리 방법' +
      '<i class="fas fa-chevron-down gw-help-fold" aria-hidden="true"></i></summary><div class="gw-help-body"><ul>' + li + '</ul></div></details>';
  }
  document.addEventListener('toggle', function (e) {
    if (e.target && e.target.classList && e.target.classList.contains('gw-fa-help')) S.auditHelpOpen = e.target.open;
  }, true);
  function auditFilters(A) {
    var F = S.af, sides = A.sides && A.sides.length ? A.sides : [{ key: 'sales', label: '매출', open: A.counts.open, total: A.counts.total }];
    var cats = A.groups.filter(function (g) { return g.side === F.side; });
    if (F.cat && !cats.some(function (g) { return g.key === F.cat; })) F.cat = '';
    return '<div class="gw-ap-filters gw-fa-filters">' +
      '<select class="gw-chip-select on" data-af="side" aria-label="매출·매입">' + sides.map(function (x) {
        return '<option value="' + x.key + '"' + (x.key === F.side ? ' selected' : '') + '>' + esc(x.label) + ' · 처리 전 ' + won(x.open) + '</option>';
      }).join('') + '</select>' +
      '<select class="gw-chip-select' + (F.cat ? ' on' : '') + '" data-af="cat" aria-label="묶음"><option value="">묶음 전체</option>' +
        cats.map(function (g) { return '<option value="' + g.key + '"' + (g.key === F.cat ? ' selected' : '') + '>' + esc(g.label) + ' ' + won(g.items.length) + '</option>'; }).join('') +
      '</select>' +
      '<select class="gw-chip-select' + (F.st !== 'open' ? ' on' : '') + '" data-af="st" aria-label="처리 상태">' +
        '<option value="open"' + (F.st === 'open' ? ' selected' : '') + '>처리 전만</option>' +
        '<option value=""' + (F.st === '' ? ' selected' : '') + '>처리한 것도</option></select>' +
      '<input type="search" class="gw-chip-search" data-af="q" placeholder="공사 코드·현장" aria-label="찾기" maxlength="40" autocomplete="off" value="' + esc(F.q) + '">' +
      '<span class="gw-ap-count" id="feAuditCount" aria-live="polite"></span></div>';
  }
  function auditMatch(r) {
    var F = S.af, q = F.q.trim().toUpperCase();
    if (F.st === 'open' && r.status !== 'open') return false;
    return !q || String(r.code).toUpperCase().indexOf(q) !== -1 || String(r.ecount_code || '').toUpperCase().indexOf(q) !== -1 ||
      String(r.site || '').toUpperCase().indexOf(q) !== -1;
  }
  function renderAuditBody() {
    var A = S.audit, F = S.af, budget = AUDIT_PAGE * (F.more + 1), shown = 0, total = 0, html = '';
    A.groups.forEach(function (g) {
      if (g.side !== F.side || (F.cat && g.key !== F.cat)) return;
      var rows = g.items.filter(auditMatch);
      total += rows.length;
      if (!rows.length) return;
      var take = rows.slice(0, Math.max(0, budget - shown));
      shown += take.length;
      if (!take.length) return;
      var buy = g.side === 'buy' || g.side === 'cost';
      html += '<section class="gw-fa-group"><h3 class="gw-fa-title">' + esc(g.label) +
          ' <span>' + (g.open ? '처리 전 ' + won(g.open) + ' / ' : '') + won(g.items.length) + '건</span></h3>' +
        '<p class="gw-fa-desc">' + esc(g.desc) + '</p>' +
        '<div class="gw-panel gw-ap-tablewrap"><table class="gw-table gw-fe-table gw-fa-table">' +
        '<colgroup><col class="c-code"><col class="c-site"><col class="c-amt"><col class="c-amt"><col><col class="c-act"></colgroup><thead><tr>' +
        '<th>프로젝트코드</th><th>현장</th><th class="num">' + (buy ? '이카운트 매입' : '이카운트') + '</th>' +
        '<th class="num">' + (buy ? '공사 현황 원가' : '공사 현황') + '</th><th>근거</th><th>처리</th>' +
        '</tr></thead><tbody>' + take.map(auditRow).join('') + '</tbody></table></div></section>';
    });
    if (shown < total) {
      html += '<div class="gw-fa-more"><button type="button" class="btn btn-sm btn-outline-secondary" data-af-more="1">' +
        '더 보기 (남은 ' + won(total - shown) + '건)</button></div>';
    }
    $('feAuditBody').innerHTML = html || '<p class="gw-empty">' + (F.st === 'open' ? '처리할 것이 없어요.' : '맞는 줄이 없어요.') + '</p>';
    $('feAuditCount').textContent = won(total) + '건';
  }
  // 숫자 상자 — ①·② 는 매출전표 숫자, ③ 은 대조 숫자. ③ 에서도 상자를 채워 두어 탭 줄 자리가 그대로(11-28 부장님)
  function paintStats() {
    var el = $('feStats');
    if (!el) return;
    el.hidden = false;   // 예전엔 ③ 에서 숨겼음
    el.innerHTML = S.tab === 'audit' ? auditStats() : (S.salesStats || '');
  }
  function auditStats() {
    var A = S.audit, box = function (k, v, main, sub) {
      return '<div class="gw-stat' + (main ? ' gw-stat-main' : '') + '"><div class="k">' + k + '</div><div class="v">' + v +
        (v === '-' ? '' : '<small> 건</small>') + '</div>' + (sub ? '<div class="gw-fb-stat-sub">' + sub + '</div>' : '') + '</div>';
    };
    if (!A || !A.batch) return box('처리할 것', '-', true) + box('매출 처리 전', '-') + box('매입 처리 전', '-') + box('처리함', '-');
    var side = function (key) { return (A.sides || []).filter(function (x) { return x.key === key; })[0]; };
    var sales = side('sales'), buy = side('buy');
    var todo = sales || buy ? (sales ? sales.open : 0) + (buy ? buy.open : 0) : A.counts.open;
    var when = A.run ? (A.run.ok_at || A.run.at) : '';
    return box('처리할 것', won(todo), true, when ? esc(when) + ' 대조' : '') +
      box('매출 처리 전', won(sales ? sales.open : A.counts.open)) +
      box('매입 처리 전', buy ? won(buy.open) : '-') +
      box('처리함', won(A.counts.done));
  }
  function renderAudit() {
    var A = S.audit || { groups: [], counts: { open: 0, done: 0, total: 0 } };
    // 재시작 전 옛 서버 응답엔 side 가 없음(11-28) — 매출로 봄. 줄에도 붙여 매입 칸 모양이 섞이지 않게
    (A.groups || []).forEach(function (g) {
      g.side = g.side || 'sales';
      g.items.forEach(function (r) { r.side = r.side || g.side; });
    });
    // 탭 옆 숫자 = SB 할 일(매출·매입 처리 전). 시트 원가 확인은 보기만이라 뺌(11-28 부장님)
    var work = (A.sides || []).filter(function (x) { return x.key === 'sales' || x.key === 'buy'; });
    if ($('feCntAudit')) $('feCntAudit').textContent = work.length ? work.reduce(function (n, x) { return n + x.open; }, 0) : A.counts.open;
    var box = $('feAudit');
    if (!box) return;
    if (!A.batch) { box.innerHTML = '<p class="gw-empty">아직 대조한 결과가 없어요.</p>'; return; }
    box.innerHTML = auditHead(A) + auditFilters(A) + '<div id="feAuditBody"></div>';
    renderAuditBody();
    if (S.tab === 'audit') paintStats();
  }
  function loadAudit() {
    return api('/api/finance/ecount/audit').then(function (j) {
      S.audit = j; S.auditLoaded = true; renderAudit();
    }).catch(function (e) {
      if ($('feAudit')) $('feAudit').innerHTML = '<p class="gw-empty">' + esc(e.message) + '</p>';
    });
  }
  if ($('feAudit')) {
    $('feAudit').addEventListener('change', function (ev) {
      var k = ev.target.getAttribute && ev.target.getAttribute('data-af');
      if (!k || k === 'q') return;
      S.af[k] = ev.target.value; S.af.more = 0;
      if (k === 'side') S.af.cat = '';
      renderAudit();
    });
    $('feAudit').addEventListener('input', function (ev) {
      if (!ev.target.getAttribute || ev.target.getAttribute('data-af') !== 'q') return;
      S.af.q = ev.target.value; S.af.more = 0;
      renderAuditBody();
    });
    $('feAudit').addEventListener('click', function (ev) {
      if (ev.target.closest('[data-af-more]')) { S.af.more++; renderAuditBody(); return; }
      var cp = ev.target.closest('.gw-fe-copy');
      if (cp) {
        var v = cp.getAttribute('data-copy'), icon = cp.innerHTML;
        copyText(v).then(function () {
          cp.innerHTML = CHECK_SVG; cp.classList.add('done'); msg(v + ' 코드를 복사했어요.', 'ok');
          setTimeout(function () { cp.innerHTML = icon; cp.classList.remove('done'); }, 1500);
        }).catch(function () { msg('복사하지 못했어요. 브라우저에서 클립보드 권한을 허용해 주세요.', 'err'); });
        return;
      }
      var b = ev.target.closest('button[data-act]');
      if (!b) return;
      var tr = b.closest('tr'), memoEl = tr && tr.querySelector('.gw-fa-memo');
      b.disabled = true;
      api('/api/finance/ecount/audit/' + b.getAttribute('data-aid'),
          { status: b.getAttribute('data-act'), memo: memoEl ? memoEl.value : '' }).then(function (j) {
        var code = tr && tr.querySelector('strong') ? tr.querySelector('strong').textContent : '';
        msg(code + (j.status === 'open' ? ' — 처리 전으로 되돌렸어요.' : " — '" + j.status_label + "'으로 남겼어요."), 'ok');
        return loadAudit();
      }).catch(function (e) { msg(e.message, 'err'); b.disabled = false; });
    });
  }

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
  loadAudit();   // 탭 옆 숫자(처리 전 건수)
})();
