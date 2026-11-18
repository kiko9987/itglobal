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

  // 아이폰 안내를 브라우저별로(11-18 부장님 "안내가 부족"). 템플릿을 고치면 재시작이 필요해 여기서 채운다.
  //   - 사파리: iOS 26 은 공유가 ⋯ 메뉴 안, 예전 iOS 는 아래쪽 가운데 공유 버튼 → 둘 다 적음
  //   - 아이폰 크롬(CriOS): 주소창 오른쪽 공유 버튼
  //   - 카카오톡·슬랙 등 다른 앱 안 화면: 바로 추가가 안 됨 → 사파리로 먼저
  //   아이폰은 홈 화면 앱과 사파리가 로그인을 따로 기억해서, 아이콘으로 처음 열면 한 번 더 로그인해야 한다.
  function iosKind() {
    var ua = navigator.userAgent || '';
    if (/KAKAOTALK|NAVER\(inapp|Line\/|FBAN|FBAV|Instagram|Slack/i.test(ua)) return 'inapp';
    if (/CriOS/.test(ua)) return 'chrome';
    return 'safari';
  }

  var SHARE = '<i class="fas fa-arrow-up-from-bracket" aria-hidden="true"></i>';
  var LOGIN_NOTE = '<p class="gw-install-note">아이폰은 홈 화면 아이콘과 사파리가 로그인을 따로 기억해요. ' +
    '아이콘으로 처음 열 때 한 번 더 로그인해 주세요. 번호로 로그인하는 분은 경영지원에서 새 번호를 받아 넣으면 돼요.</p>';

  function iosHtml(kind) {
    if (kind === 'inapp') {
      return '<p class="mb-2"><strong>지금은 카카오톡·슬랙 같은 다른 앱 안에서 열린 화면이라 바로 추가할 수 없어요.</strong></p>' +
        '<ol>' +
        '<li>화면 아래나 위의 <strong>⋯</strong> 또는 나침반 아이콘을 눌러 <strong>Safari로 열기</strong>를 고르세요.</li>' +
        '<li>사파리에서 이름 메뉴의 <strong>앱으로 설치</strong>를 다시 누르면 이어서 안내가 나와요.</li>' +
        '</ol>';
    }
    if (kind === 'chrome') {
      return '<ol>' +
        '<li>주소창 오른쪽의 공유 버튼 ' + SHARE + '을 누르세요.</li>' +
        '<li>목록에서 <strong>홈 화면에 추가</strong>를 누르세요. 안 보이면 목록을 위로 밀어 보세요.</li>' +
        '<li>이름이 <strong>아이티글로벌</strong>인지 보고 오른쪽 위 <strong>추가</strong>를 누르세요.</li>' +
        '<li>홈 화면에 생긴 <strong>아이티글로벌</strong> 아이콘으로 여세요.</li>' +
        '</ol>' + LOGIN_NOTE;
    }
    return '<ol>' +
      '<li>사파리 아래쪽 주소창 옆의 <strong>⋯</strong>을 누르고 <strong>공유</strong> ' + SHARE + '를 누르세요. ' +
      '화면 아래 가운데에 공유 버튼 ' + SHARE + '이 바로 보이면 그걸 누르면 돼요.</li>' +
      '<li>나온 목록을 위로 밀어 <strong>홈 화면에 추가</strong> <i class="fas fa-square-plus" aria-hidden="true"></i>를 누르세요.</li>' +
      '<li><strong>웹 앱으로 열기</strong>가 보이면 켜 둔 채로, 오른쪽 위 <strong>추가</strong>를 누르세요.</li>' +
      '<li>홈 화면에 생긴 <strong>아이티글로벌</strong> 아이콘으로 여세요. 주소창 없이 앱처럼 열려요.</li>' +
      '</ol>' + LOGIN_NOTE +
      '<p class="gw-install-note">슬랙에서 링크를 눌러 열었다면 오른쪽 아래 나침반 아이콘으로 사파리에서 먼저 열어 주세요.</p>';
  }

  function showGuide() {
    var modal = document.getElementById('gwInstallModal');
    if (!modal) return;
    var which = os();
    if (which === 'ios') {
      var box = modal.querySelector('[data-os="ios"]');
      if (box) box.innerHTML = iosHtml(iosKind());   // 고정 문구만(사용자 값 없음)
    }
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
