/* 그룹웨어 새 틀 — 왼쪽 줄(당근 비즈니스처럼 마우스를 대면 펼침), 묶음 접고 펴기,
   ☰ 로 숨기기·가장자리 마우스로 잠깐 꺼내기, 예전 화면으로 돌아가기.
   새 틀(body.gw-shell)에서만 읽힌다. */
(function () {
  'use strict';

  var body = document.body;
  if (!body || !body.classList.contains('gw-shell')) return;

  // 나 안 테마 먼저 보기(11-28) — 주소 끝 ?theme=b 로 켜고 ?theme=off 로 끔. 이 브라우저에만 기억
  (function () {
    var TKEY = 'gw.theme';
    try {
      var q = new URLSearchParams(location.search).get('theme');
      if (q === 'b') localStorage.setItem(TKEY, 'b');
      else if (q === 'off') localStorage.removeItem(TKEY);
      if (localStorage.getItem(TKEY) === 'b') document.documentElement.classList.add('gw-tb');
    } catch (e) { /* 저장소를 못 쓰면 지금 모양 그대로 */ }
  })();

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

  // 좁은 화면: 닫힌 왼쪽 줄은 화면 밖에 있어도 Tab 이 들어가던 것 → 닫혀 있는 동안 inert(11-15 점검).
  // ☰ 에 열림 상태도 알려 줌
  function syncNarrowRail() {
    var narrowClosed = isNarrow() && !has('gw-peek');
    if (rail) {
      if (narrowClosed) rail.setAttribute('inert', ''); else rail.removeAttribute('inert');
    }
    if (topMenuBtn) topMenuBtn.setAttribute('aria-expanded', isNarrow() ? String(has('gw-peek')) : String(!has('gw-collapsed')));
  }
  if (window.MutationObserver) {
    new MutationObserver(syncNarrowRail).observe(body, { attributes: true, attributeFilter: ['class'] });
  }
  try { window.matchMedia('(max-width: 767.98px)').addEventListener('change', syncNarrowRail); } catch (e) { /* 옛 브라우저 */ }
  syncNarrowRail();

  // 하위 메뉴가 넘치면 지금 탭이 보이게 가로로 옮김(휴대폰에서 '연말정산' 등이 찾기 단추 뒤로 가려지던 것)
  var subnav = document.querySelector('.gw-subnav');
  var activeSub = subnav && subnav.querySelector('.gw-subnav-item.active');
  if (activeSub && subnav.scrollWidth > subnav.clientWidth) {
    var left = activeSub.getBoundingClientRect().left - subnav.getBoundingClientRect().left + subnav.scrollLeft;
    subnav.scrollLeft = Math.max(0, left - (subnav.clientWidth - activeSub.offsetWidth) / 2);
  }

  // 그룹웨어 화면의 창(휴가 신청·취소·앱 설치 등) 포커스(11-15 점검):
  // 열면 첫 칸(칸이 없으면 아래 첫 단추 = '그대로 두기'·'확인' 같은 안전한 쪽), 닫으면 연 단추로 돌아감.
  // 예전 PM 화면 창은 건드리지 않게 그룹웨어 화면(.gw-page)·설치 창만
  function gwModal(m) {
    return m && m.classList && m.classList.contains('modal') &&
      (m.id === 'gwInstallModal' || !!document.querySelector('.gw-page'));
  }
  document.addEventListener('show.bs.modal', function (e) {
    var m = e.target;
    if (!gwModal(m)) return;
    var a = document.activeElement;
    // 사용자 메뉴 안 항목([앱으로 설치])에서 열었으면 메뉴가 닫히므로 이름 단추로 돌아가게
    var dd = a && a.closest && a.closest('.dropdown');
    if (dd && a.closest('.dropdown-menu')) a = dd.querySelector('[data-bs-toggle="dropdown"]') || a;
    m._gwOpener = a && a !== document.body && !m.contains(a) ? a : null;
  });
  document.addEventListener('shown.bs.modal', function (e) {
    var m = e.target;
    if (!gwModal(m)) return;
    // 화면 코드가 이미 창 안 어딘가(제목 등)에 초점을 줬으면 그대로 둔다 — 덮어쓰면 스케줄 근무 '쉬는 사람' 창에서
    // 첫 체크박스에 초점 테두리가 생겨 첫 사람이 고른 것처럼 보였다(11-25 부장님)
    var cur = document.activeElement;
    if (cur && cur !== m && m.contains(cur)) return;
    var first = m.querySelector('.modal-body select:not([disabled]), .modal-body input:not([disabled]):not([type=hidden]), .modal-body textarea:not([disabled])');
    if (first && first.offsetParent === null) first = null;
    if (!first) first = m.querySelector('.modal-footer button:not([disabled])');
    if (first) first.focus();
  });
  document.addEventListener('hidden.bs.modal', function (e) {
    var m = e.target;
    if (!gwModal(m)) return;
    var o = m._gwOpener;
    m._gwOpener = null;
    if (o && document.contains(o) && o.offsetParent !== null) {
      try { o.focus({ preventScroll: true }); } catch (x) { /* 무시 */ }
    } else {
      var h = document.querySelector('.gw-page-title');   // 연 단추가 다시 그려져 없어졌으면 화면 제목으로
      if (h) { if (!h.hasAttribute('tabindex')) h.setAttribute('tabindex', '-1'); h.focus({ preventScroll: true }); }
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

/* 날짜·시각 칸(11-26 부장님 '그룹웨어 안에 캘린더 다 점검') — 마우스로 칸 어디를 눌러도 달력이 열리고,
   칸 안 글자(연·월·일)가 파랗게 선택되지 않게. 키보드로 고치는 건 그대로, 휴대폰(터치)은 원래 달력 그대로.
   모든 그룹웨어 화면에 한 번에 — 새 화면에 날짜 칸을 넣어도 따로 할 일 없음 */
(function () {
  'use strict';
  if (!document.body || !document.body.classList.contains('gw-shell')) return;
  var TYPES = { date: 1, month: 1, week: 1, 'datetime-local': 1, time: 1 };
  document.addEventListener('mousedown', function (e) {
    var el = e.target;
    if (e.button !== 0 || !el || el.tagName !== 'INPUT' || !TYPES[el.type]) return;
    if (el.disabled || el.readOnly || typeof el.showPicker !== 'function') return;
    e.preventDefault();
    try { el.showPicker(); } catch (err) { el.focus(); }   // 못 열면 예전처럼 칸에 들어감
  }, true);
})();
