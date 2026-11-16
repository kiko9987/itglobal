/* 그룹웨어 알림 🔔 (2026-11-10) — 안 읽은 게 있으면 빨간 점, 누르면 최근 30일 알림 창.
   창을 열면 '읽음'(빨간 점 꺼짐). 화면을 띄워 둔 동안 1분마다 다시 확인(창이 안 보일 땐 쉼).
   서버 글자는 textContent 로만 넣는다. */
(function () {
  'use strict';

  var btn = document.getElementById('gwNotifyBtn');
  var panel = document.getElementById('gwNotifyPanel');
  if (!btn || !panel || !document.body.classList.contains('gw-shell')) return;
  var dot = document.getElementById('gwNotifyDot');
  var list = document.getElementById('gwNotifyList');
  var tabs = document.getElementById('gwNotifyTabs');
  var sr = document.getElementById('gwNotifySr');
  var data = null;
  var failed = false;   // 마지막 불러오기 실패 — 실패도 '알림 없음'으로 보이던 것
  var cat = '';
  var POLL = 60000;

  function api(url, method) {
    return fetch(url, { method: method || 'GET', credentials: 'same-origin' }).then(function (r) {
      return r.json().catch(function () { return {}; }).then(function (j) {
        if (!r.ok || j.success === false) throw new Error(j.message || ('HTTP ' + r.status));
        return j;
      });
    });
  }

  function ago(s) {
    var d = new Date((s || '').replace(' ', 'T'));
    if (isNaN(d)) return '';
    var m = Math.floor((Date.now() - d.getTime()) / 60000);
    if (m < 1) return '방금';
    if (m < 60) return m + '분 전';
    var h = Math.floor(m / 60);
    if (h < 24) return h + '시간 전';
    var day = Math.floor(h / 24);
    if (day < 7) return day + '일 전';
    return (d.getMonth() + 1) + '/' + d.getDate();
  }

  function setDot(n) {
    dot.hidden = !n;
    sr.textContent = n ? '알림 — 안 읽은 ' + n + '개' : '알림';
  }

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  function render() {
    list.textContent = '';
    if (!data) {
      list.appendChild(el('div', 'gw-notify-empty', failed ? '알림을 불러오지 못했어요. 잠시 뒤 다시 열어 주세요.' : '불러오는 중…'));
      return;
    }
    // 창을 열 때 새로 받기에 실패했으면 예전 목록이라는 걸 알림(11-16 점검 — 실패해도 지난 목록이 최신처럼 보였다)
    if (failed) {
      var stale = el('div', 'gw-notify-stale', '새 알림을 받아 오지 못했어요.');
      stale.setAttribute('role', 'status');
      list.appendChild(stale);
    }
    var items = (data.items || []).filter(function (i) { return !cat || i.cat === cat; });
    if (!items.length) {
      list.appendChild(el('div', 'gw-notify-empty', '최근 30일 동안 알림이 없어요.'));
      return;
    }
    items.forEach(function (i) {
      var a = el(i.href ? 'a' : 'div', 'gw-notify-item' + (i.unread ? ' unread' : ''));
      if (i.href) {
        a.href = i.href;
        if (/^https?:/.test(i.href)) { a.target = '_blank'; a.rel = 'noopener'; }
      }
      var ic = el('span', 'gw-notify-ic', i.icon || '🔔');
      ic.setAttribute('aria-hidden', 'true');
      a.appendChild(ic);
      var main = el('span', 'gw-notify-main');
      var title = el('span', 'gw-notify-title', i.title);
      // 안 읽은 줄은 주황 점·굵은 제목으로만 보여서 낭독기엔 같은 줄로 들렸다 → 앞에 '안 읽음'
      if (i.unread) title.insertBefore(el('span', 'visually-hidden', '안 읽음. '), title.firstChild);
      main.appendChild(title);
      if (i.body) main.appendChild(el('span', 'gw-notify-body', i.body));
      var meta = el('span', 'gw-notify-meta');
      meta.appendChild(el('span', 'gw-notify-cat', i.cat_label));
      meta.appendChild(el('span', null, ago(i.at)));
      main.appendChild(meta);
      a.appendChild(main);
      list.appendChild(a);
    });
  }

  function load() {
    return api('/api/gw/notify').then(function (j) {
      data = j;
      failed = false;
      if (panel.hidden) setDot(j.unread);   // 창이 열려 있는 동안엔 읽음 처리 중이라 점을 다시 켜지 않음
    }).catch(function () { failed = true; /* 화면은 그대로, 창을 열면 '불러오지 못했어요' */ });
  }

  function markSeen() {
    if (!data || !data.unread) { setDot(0); return; }
    api('/api/gw/notify/seen', 'POST').then(function () { data.unread = 0; setDot(0); }).catch(function () {});
  }

  function open() {
    panel.hidden = false;
    btn.setAttribute('aria-expanded', 'true');
    render();
    // 최신으로 다시 받고 읽음 처리 — 이번엔 안 읽은 줄을 표시한 채로 두고, 빨간 점만 끈다.
    // 다시 받기에 실패했으면 읽음 처리하지 않음(못 본 새 알림까지 읽음이 되던 것, 11-16 점검)
    load().then(function () {
      if (panel.hidden) return;
      render();
      if (!failed) markSeen();
    });
  }

  function close() {
    panel.hidden = true;
    btn.setAttribute('aria-expanded', 'false');
    if (data && !failed) data.items.forEach(function (i) { i.unread = false; });   // 받기 실패로 읽음 처리 안 했으면 그대로
  }

  btn.addEventListener('click', function (e) {
    e.stopPropagation();
    if (panel.hidden) open(); else close();
  });
  tabs.addEventListener('click', function (e) {
    var b = e.target.closest('button[data-cat]');
    if (!b) return;
    cat = b.getAttribute('data-cat');
    tabs.querySelectorAll('button').forEach(function (x) {
      x.classList.toggle('active', x === b);
      x.setAttribute('aria-pressed', x === b ? 'true' : 'false');   // 고른 분류를 낭독기에도
    });
    render();
  });
  document.addEventListener('click', function (e) {
    if (!panel.hidden && !panel.contains(e.target) && !btn.contains(e.target)) close();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && !panel.hidden) { close(); btn.focus(); }
  });

  load();
  setInterval(function () { if (!document.hidden) load(); }, POLL);
  document.addEventListener('visibilitychange', function () { if (!document.hidden) load(); });
})();
