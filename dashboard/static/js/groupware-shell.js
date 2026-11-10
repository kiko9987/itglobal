/* 그룹웨어 새 틀 — 왼쪽 줄 접기·펼치기, 가장자리 마우스로 잠깐 펼치기, 예전 화면으로 돌아가기.
   새 틀(body.gw-shell)에서만 읽힌다. */
(function () {
  'use strict';

  var body = document.body;
  if (!body || !body.classList.contains('gw-shell')) return;

  var KEY = 'gw.rail.collapsed';
  var rail = document.getElementById('gwRail');
  var hotzone = document.getElementById('gwRailHotzone');
  var toggleBtn = document.getElementById('gwRailToggle');
  var topMenuBtn = document.getElementById('gwTopbarMenu');
  var classicBtn = document.getElementById('gwClassicBtn');

  var OPEN_DELAY = 220;   // 가장자리에 잠깐 머물러야 열림(지나가다 열리는 것 방지)
  var CLOSE_DELAY = 280;
  var openTimer = null;
  var closeTimer = null;

  function isNarrow() {
    return window.matchMedia('(max-width: 767.98px)').matches;
  }

  function save(collapsed) {
    try { localStorage.setItem(KEY, collapsed ? '1' : '0'); } catch (e) { /* 저장 못 해도 동작은 함 */ }
  }

  function setCollapsed(collapsed) {
    body.classList.toggle('gw-collapsed', collapsed);
    body.classList.remove('gw-peek');
    save(collapsed);
    var label = collapsed ? '메뉴 펼치기' : '메뉴 접기';
    if (toggleBtn) { toggleBtn.title = label; toggleBtn.setAttribute('aria-label', label); }
    // 표 폭이 바뀌므로 DataTables 등이 다시 계산하게
    setTimeout(function () { window.dispatchEvent(new Event('resize')); }, 200);
  }

  function peek(on) {
    body.classList.toggle('gw-peek', on);
    if (!on) clearTimers();
  }

  function clearTimers() {
    clearTimeout(openTimer);
    clearTimeout(closeTimer);
    closeTimer = null;
  }

  // ☰ (왼쪽 줄 안): 넓은 화면=접기/펼치기, 좁은 화면=닫기
  if (toggleBtn) {
    toggleBtn.addEventListener('click', function () {
      if (isNarrow()) { peek(false); return; }
      if (body.classList.contains('gw-peek')) { setCollapsed(false); return; }  // 잠깐 펼친 상태에서 누르면 고정
      setCollapsed(!body.classList.contains('gw-collapsed'));
    });
  }

  // ☰ (위쪽 줄, 접혔을 때만 보임): 넓은 화면=다시 고정, 좁은 화면=잠깐 펼치기
  if (topMenuBtn) {
    topMenuBtn.addEventListener('click', function () {
      if (isNarrow()) { peek(!body.classList.contains('gw-peek')); return; }
      setCollapsed(false);
    });
  }

  // 가장자리 마우스 → 잠깐 펼침
  if (hotzone) {
    hotzone.addEventListener('mouseenter', function () {
      clearTimers();
      openTimer = setTimeout(function () { peek(true); }, OPEN_DELAY);
    });
    hotzone.addEventListener('mouseleave', function () { clearTimeout(openTimer); });
  }

  // 잠깐 펼친 동안: 마우스가 메뉴 오른쪽으로 충분히 벗어나면 닫는다.
  // (mouseleave 만 쓰면 가장자리에서 표 쪽으로 빠르게 옮길 때 메뉴에 들어간 적이 없어 안 닫힘)
  document.addEventListener('mousemove', function (e) {
    if (!body.classList.contains('gw-peek') || isNarrow()) return;
    var edge = (rail ? rail.getBoundingClientRect().right : 64) + 24;
    if (e.clientX > edge) {
      if (!closeTimer) {
        closeTimer = setTimeout(function () { closeTimer = null; peek(false); }, CLOSE_DELAY);
      }
    } else if (closeTimer) {
      clearTimeout(closeTimer);
      closeTimer = null;
    }
  });

  // 바깥 클릭·Esc 로 잠깐 펼친 메뉴 닫기
  document.addEventListener('click', function (e) {
    if (!body.classList.contains('gw-peek')) return;
    if (rail && rail.contains(e.target)) return;
    if (topMenuBtn && topMenuBtn.contains(e.target)) return;
    peek(false);
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && body.classList.contains('gw-peek')) peek(false);
  });

  // 예전 화면으로(이번 로그인 동안) — 사용자 메뉴에서 다시 새 화면으로 돌아올 수 있음
  if (classicBtn) {
    classicBtn.addEventListener('click', function () {
      classicBtn.disabled = true;
      fetch('/api/gw/layout', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({ 'new': false })
      }).then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        window.location.reload();
      }).catch(function () {
        classicBtn.disabled = false;
        alert('예전 화면으로 바꾸지 못했습니다. 잠시 후 다시 눌러 주세요.');
      });
    });
  }
})();
