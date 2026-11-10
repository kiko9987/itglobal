/* 그룹웨어 새 틀 — 왼쪽 줄(당근 비즈니스처럼 마우스를 대면 펼침), 묶음 접고 펴기,
   ☰ 로 숨기기·가장자리 마우스로 잠깐 꺼내기, 예전 화면으로 돌아가기.
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

  var PEEK_DELAY = 220;    // 숨긴 상태에서 가장자리에 잠깐 머물러야 꺼냄(지나가다 열리는 것 방지)
  var CLOSE_DELAY = 150;   // 메뉴에서 마우스가 잠깐 삐져나가도 바로 닫히지 않게
  var peekTimer = null;
  var closeTimer = null;

  function isNarrow() {
    return window.matchMedia('(max-width: 767.98px)').matches;
  }

  function has(c) { return body.classList.contains(c); }

  function save(collapsed) {
    try { localStorage.setItem(KEY, collapsed ? '1' : '0'); } catch (e) { /* 저장 못 해도 동작은 함 */ }
  }

  function expand(on) {
    clearTimeout(closeTimer);
    closeTimer = null;
    body.classList.toggle('gw-expand', on);
  }

  // 펼침을 닫는다. 숨긴 상태에서 꺼낸 것(peek)이면 다시 숨긴다.
  function closeAll() {
    clearTimeout(peekTimer);
    clearTimeout(closeTimer);
    closeTimer = null;
    body.classList.remove('gw-expand', 'gw-peek');
  }

  function setCollapsed(collapsed) {
    body.classList.toggle('gw-collapsed', collapsed);
    closeAll();
    save(collapsed);
    // 표 폭이 바뀌므로 DataTables 등이 다시 계산하게
    setTimeout(function () { window.dispatchEvent(new Event('resize')); }, 220);
  }

  function scheduleClose() {
    if (closeTimer) return;
    closeTimer = setTimeout(function () {
      closeTimer = null;
      if (rail && rail.matches(':hover')) return;                       // 그새 다시 들어옴
      var a = document.activeElement;
      if (rail && a && rail.contains(a) && a.matches(':focus-visible')) return;  // 키보드로 고르는 중
      closeAll();
    }, CLOSE_DELAY);
  }

  if (rail) {
    // 마우스를 대면 바로 펼침(당근 비즈니스 실측: 지연 없음, 0.3초 동안 넓어짐)
    rail.addEventListener('mouseenter', function () {
      if (isNarrow()) return;
      expand(true);
    });
    rail.addEventListener('mouseleave', function () {
      if (isNarrow()) return;
      scheduleClose();
    });

    // 키보드(Tab)로 들어오면 펼치고, 밖으로 나가면 닫음
    rail.addEventListener('focusin', function () {
      if (!isNarrow()) expand(true);
    });
    rail.addEventListener('focusout', function (e) {
      if (isNarrow()) return;
      if (e.relatedTarget && rail.contains(e.relatedTarget)) return;
      if (rail.matches(':hover')) return;
      closeAll();
    });

    // 묶음(영업·휴가·인사 …) 접고 펴기
    rail.addEventListener('click', function (e) {
      var btn = e.target.closest('.gw-nav-toggle');
      if (!btn) return;
      var group = btn.closest('.gw-nav-group');
      // 펼치지 않은 줄(터치 등)에서 누르면 먼저 펼치고 그 묶음을 연다
      if (!has('gw-expand')) {
        expand(true);
        group.classList.add('open');
      } else {
        group.classList.toggle('open');
      }
      btn.setAttribute('aria-expanded', group.classList.contains('open') ? 'true' : 'false');
    });
  }

  // « (펼친 줄 맨 위): 왼쪽 줄 숨기기 / 좁은 화면이면 닫기
  if (toggleBtn) {
    toggleBtn.addEventListener('click', function (e) {
      e.stopPropagation();
      if (isNarrow()) { closeAll(); return; }
      setCollapsed(true);
    });
  }

  // ☰ (위쪽 줄): 넓은 화면=다시 보이기, 좁은 화면=꺼내기/닫기
  if (topMenuBtn) {
    topMenuBtn.addEventListener('click', function () {
      if (isNarrow()) {
        if (has('gw-peek')) { closeAll(); } else { body.classList.add('gw-peek', 'gw-expand'); }
        return;
      }
      setCollapsed(false);
    });
  }

  // 숨긴 상태: 왼쪽 끝에 잠깐 머물면 펼친 모양으로 꺼냄
  if (hotzone) {
    hotzone.addEventListener('mouseenter', function () {
      clearTimeout(peekTimer);
      peekTimer = setTimeout(function () { body.classList.add('gw-peek', 'gw-expand'); }, PEEK_DELAY);
    });
    hotzone.addEventListener('mouseleave', function () { clearTimeout(peekTimer); });
  }

  // 펼친 동안 마우스가 메뉴 오른쪽으로 충분히 벗어나면 닫는다.
  // (가장자리에서 표 쪽으로 빠르게 옮기면 메뉴에 들어간 적이 없어 mouseleave 가 안 옴)
  document.addEventListener('mousemove', function (e) {
    if (!has('gw-expand') || isNarrow()) return;
    var edge = (rail ? rail.getBoundingClientRect().right : 208) + 24;
    if (e.clientX > edge) scheduleClose();
  });

  // 바깥 클릭·Esc 로 닫기
  document.addEventListener('click', function (e) {
    if (!has('gw-expand') && !has('gw-peek')) return;
    if (rail && rail.contains(e.target)) return;
    if (topMenuBtn && topMenuBtn.contains(e.target)) return;
    closeAll();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && (has('gw-expand') || has('gw-peek'))) {
      closeAll();
      if (rail && rail.contains(document.activeElement)) document.activeElement.blur();
    }
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
