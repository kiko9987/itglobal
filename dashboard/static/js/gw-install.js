/* 사용자 메뉴 [앱으로 설치] (2026-11-15, 다우오피스처럼).
   - 크롬·엣지(안드로이드·PC): 브라우저가 준 설치 신호(beforeinstallprompt, 머리에서 window.__gwInstallEvt 로 받아 둠)가
     있으면 눌렀을 때 바로 설치 창.
   - 아이폰(사파리는 사이트가 설치 창을 띄울 수 없음)·신호가 없을 때: 기기에 맞는 안내 창.
   - 이미 홈 화면 앱으로 열려 있으면 메뉴를 숨긴다.
   서비스 워커는 쓰지 않는다(sw.js 머리말 — 2026-07-08 끈 이유). */
(function () {
  'use strict';

  var item = document.getElementById('gwInstallItem');
  var btn = document.getElementById('gwInstallBtn');
  if (!item || !btn) return;

  function standalone() {
    try {
      return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
    } catch (e) { return false; }
  }

  if (standalone()) { item.hidden = true; return; }

  function os() {
    var ua = navigator.userAgent || '';
    // 아이패드 사파리는 맥처럼 보이므로 터치 지원으로 가린다
    if (/iPhone|iPad|iPod/.test(ua) || (/Macintosh/.test(ua) && navigator.maxTouchPoints > 1)) return 'ios';
    if (/Android/.test(ua)) return 'android';
    return 'pc';
  }

  function showGuide() {
    var modal = document.getElementById('gwInstallModal');
    if (!modal) return;
    var which = os();
    modal.querySelectorAll('[data-os]').forEach(function (el) { el.hidden = el.getAttribute('data-os') !== which; });
    if (window.bootstrap) window.bootstrap.Modal.getOrCreateInstance(modal).show();
  }

  btn.addEventListener('click', function (ev) {
    ev.preventDefault();
    var evt = window.__gwInstallEvt;
    if (evt && typeof evt.prompt === 'function') {
      window.__gwInstallEvt = null;   // 설치 창은 한 번만 띄울 수 있음
      try {
        evt.prompt();
        if (evt.userChoice) {
          evt.userChoice.then(function (c) { if (c && c.outcome === 'accepted') item.hidden = true; });
        }
        return;
      } catch (e) { /* 아래 안내로 */ }
    }
    showGuide();
  });

  window.addEventListener('appinstalled', function () { item.hidden = true; });
})();
