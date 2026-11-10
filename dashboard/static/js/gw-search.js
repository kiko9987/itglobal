/* 그룹웨어 빠른 찾기 — 위쪽 검색창·/ 키·Ctrl+K 로 열고, 메뉴(화면에서)·프로젝트·리드(/api/gw/search)를 한 번에.
   당근 비즈니스의 「/ 키로 메뉴 바로 찾기」를 본떴다(2026-11-10). 새 틀(body.gw-shell)에서만 읽힌다.
   서버에서 온 글자(이름·주소)는 textContent 로만 넣는다(innerHTML 금지). */
(function () {
  'use strict';

  var body = document.body;
  var root = document.getElementById('gwFind');
  if (!body || !root || !body.classList.contains('gw-shell')) return;

  var input = document.getElementById('gwFindInput');
  var list = document.getElementById('gwFindList');
  var trigger = document.getElementById('gwFindBtn');
  var scopeAll = root.getAttribute('data-scope') === 'all';

  var DEBOUNCE = 180;
  var timer = null;
  var ctrl = null;
  var items = [];      // 지금 보이는 고를 수 있는 줄 [{href, el}]
  var active = -1;
  var lastFocus = null;
  var menus = collectMenus();

  // ── 메뉴: 왼쪽 줄에 이미 그려진 것(권한대로 걸러져 있음)에서 모은다 ──
  function collectMenus() {
    var out = [];
    document.querySelectorAll('#gwRail .gw-nav-group').forEach(function (g) {
      var head = g.querySelector('.gw-nav-area');
      if (!head) return;
      var area = (head.querySelector('.gw-nav-label') || head).textContent.trim();
      var iconEl = head.querySelector('.gw-nav-icon i');
      var icon = iconEl ? iconEl.className : 'fas fa-compass';
      var subs = g.querySelectorAll('.gw-nav-subitem');
      if (subs.length) {
        subs.forEach(function (a) {
          var label = a.textContent.trim();
          out.push({ title: label, path: area + ' › ' + label, href: a.getAttribute('href'), icon: icon });
        });
      } else if (head.tagName === 'A') {
        out.push({ title: area, path: area, href: head.getAttribute('href'), icon: icon });
      }
    });
    return out;
  }

  function compact(s) { return String(s || '').replace(/\s+/g, '').toLowerCase(); }
  function digits(s) { return String(s || '').replace(/\D/g, ''); }

  function usableForServer(q) {
    var c = compact(q);
    return c.length >= 2 || digits(q).length >= 4;
  }

  // ── 그리기 ──
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  function addGroup(title, extra) {
    var h = el('div', 'gw-find-group');
    h.appendChild(el('span', null, title));
    if (extra) h.appendChild(extra);
    list.appendChild(h);
  }

  function addRow(opt) {
    var a = el('a', 'gw-find-item' + (opt.dim ? ' dim' : ''));
    a.href = opt.href;
    a.setAttribute('role', 'option');
    a.id = 'gwFindOpt' + items.length;
    var ic = el('span', 'gw-find-ic');
    ic.appendChild(el('i', opt.icon));
    a.appendChild(ic);
    var main = el('span', 'gw-find-main');
    var line = el('span', 'gw-find-title');
    if (opt.code) line.appendChild(el('span', 'gw-find-code', opt.code));
    line.appendChild(el('span', 'gw-find-name', opt.title || '(이름 없음)'));
    if (opt.badge) line.appendChild(el('span', 'gw-find-badge', opt.badge));
    main.appendChild(line);
    if (opt.sub) main.appendChild(el('span', 'gw-find-sub', opt.sub));
    a.appendChild(main);
    if (opt.meta) a.appendChild(el('span', 'gw-find-meta', opt.meta));
    var idx = items.length;
    a.addEventListener('mousemove', function () { if (active !== idx) setActive(idx); });
    list.appendChild(a);
    items.push({ href: opt.href, el: a });
  }

  function addNote(text) {
    list.appendChild(el('div', 'gw-find-note', text));
  }

  function moreLink(text, href) {
    var a = el('a', 'gw-find-more', text);
    a.href = href;
    return a;
  }

  function reset() {
    list.textContent = '';
    items = [];
    active = -1;
  }

  function setActive(i) {
    if (!items.length) { active = -1; input.removeAttribute('aria-activedescendant'); return; }
    active = (i + items.length) % items.length;
    items.forEach(function (it, k) { it.el.classList.toggle('active', k === active); it.el.setAttribute('aria-selected', k === active ? 'true' : 'false'); });
    var cur = items[active].el;
    input.setAttribute('aria-activedescendant', cur.id);
    cur.scrollIntoView({ block: 'nearest' });
  }

  function matchMenus(q) {
    var c = compact(q);
    if (!c) return menus;
    return menus.filter(function (m) { return compact(m.path).indexOf(c) !== -1; });
  }

  function renderMenus(q, max) {
    var ms = matchMenus(q).slice(0, max);
    if (!ms.length) return 0;
    addGroup('메뉴');
    ms.forEach(function (m) { addRow({ href: m.href, icon: m.icon, title: m.title, sub: m.path !== m.title ? m.path : '' }); });
    return ms.length;
  }

  function render(q, data, state) {
    reset();
    var shown = renderMenus(q, q ? 4 : 20);
    if (scopeAll && q && usableForServer(q)) {
      if (state === 'loading') {
        addNote('찾는 중…');
      } else if (state === 'error') {
        addNote('프로젝트·리드를 찾지 못했어요. 잠시 후 다시 해 주세요.');
      } else if (data) {
        var enc = encodeURIComponent(q.trim());
        if (data.projects.length) {
          addGroup('프로젝트', data.project_total > data.projects.length
            ? moreLink(data.project_total.toLocaleString('ko-KR') + '건 모두 보기', '/projects?search=' + enc + '&fresh=1') : null);
          data.projects.forEach(function (p) {
            addRow({ href: p.href, icon: 'fas fa-folder-open', code: p.code, title: p.title, sub: p.sub, meta: p.meta,
                     badge: p.cancelled ? '취소' : '', dim: p.cancelled });
          });
          shown += data.projects.length;
        }
        if (data.leads.length) {
          addGroup('리드', data.lead_total > data.leads.length
            ? moreLink(data.lead_total.toLocaleString('ko-KR') + '건 모두 보기', '/leads?search=' + enc) : null);
          data.leads.forEach(function (l) {
            addRow({ href: l.href, icon: 'fas fa-user-tie', code: l.code, title: l.title, sub: l.sub, meta: l.meta });
          });
          shown += data.leads.length;
        }
        if (!shown) addNote('‘' + q.trim() + '’에 맞는 결과가 없어요.');
      }
    } else if (q && !shown) {
      addNote(scopeAll ? '두 글자 이상 입력해 주세요. 전화번호는 숫자 4자리부터 찾아요.' : '맞는 메뉴가 없어요.');
    }
    if (items.length) setActive(0);
  }

  function search() {
    var q = input.value;
    clearTimeout(timer);
    if (ctrl) { ctrl.abort(); ctrl = null; }
    if (!scopeAll || !usableForServer(q)) { render(q, null, null); return; }
    render(q, null, 'loading');
    timer = setTimeout(function () {
      ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
      fetch('/api/gw/search?q=' + encodeURIComponent(q.trim()), {
        credentials: 'same-origin', signal: ctrl ? ctrl.signal : undefined, headers: { 'Accept': 'application/json' }
      }).then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.json();
      }).then(function (j) {
        if (input.value !== q) return;            // 그새 바뀐 검색어면 버림
        render(q, j, null);
      }).catch(function (e) {
        if (e && e.name === 'AbortError') return;
        if (input.value === q) render(q, null, 'error');
      });
    }, DEBOUNCE);
  }

  // ── 열기·닫기 ──
  function open(prefill) {
    if (!root.hidden) { input.focus(); return; }
    lastFocus = document.activeElement;
    root.hidden = false;
    body.classList.add('gw-find-open');
    input.value = prefill || '';
    search();
    setTimeout(function () { input.focus(); input.select(); }, 0);
  }

  function close() {
    if (root.hidden) return;
    root.hidden = true;
    body.classList.remove('gw-find-open');
    clearTimeout(timer);
    if (ctrl) { ctrl.abort(); ctrl = null; }
    if (lastFocus && lastFocus.focus) { try { lastFocus.focus(); } catch (e) { /* noop */ } }
  }

  function go(i, newTab) {
    var it = items[i];
    if (!it) return;
    if (newTab) { window.open(it.href, '_blank', 'noopener'); return; }
    close();
    window.location.href = it.href;
  }

  input.addEventListener('input', search);
  input.addEventListener('keydown', function (e) {
    if (e.key === 'ArrowDown') { e.preventDefault(); setActive(active + 1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(active - 1); }
    else if (e.key === 'Enter') { e.preventDefault(); go(active, e.ctrlKey || e.metaKey); }
    else if (e.key === 'Escape') { e.preventDefault(); close(); }
  });
  root.addEventListener('click', function (e) {
    if (e.target.closest('[data-close]')) close();
  });
  list.addEventListener('click', function (e) {
    var a = e.target.closest('a');
    if (!a || e.ctrlKey || e.metaKey || e.shiftKey || e.button !== 0) return;   // 새 탭 열기는 브라우저에 맡김
    close();
  });
  if (trigger) trigger.addEventListener('click', function () { open(''); });

  function typing(t) {
    if (!t) return false;
    var tag = (t.tagName || '').toLowerCase();
    return tag === 'input' || tag === 'textarea' || tag === 'select' || t.isContentEditable;
  }

  document.addEventListener('keydown', function (e) {
    if (e.defaultPrevented) return;
    var k = e.key;
    if ((e.ctrlKey || e.metaKey) && !e.altKey && (k === 'k' || k === 'K')) {
      e.preventDefault();
      if (root.hidden) open(''); else close();
      return;
    }
    if (k === '/' && !e.ctrlKey && !e.metaKey && !e.altKey && root.hidden && !typing(e.target)
        && !document.querySelector('.modal.show')) {
      e.preventDefault();
      open('');
    }
  });
})();
