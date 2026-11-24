/* 그룹웨어 실시간 위치 (2026-11-24) — U+커넥트 '실시간 차량 위치'와 같은 구성:
   왼쪽 차량 목록(전체·운행 중·정지) + 오른쪽 카카오 지도(차마다 이름 꼬리표), 30초마다 새로. 위 주소 칸 = 가까운 차 순서.
   서버에서 온 글자는 모두 esc() 로만 넣는다. 지도 키가 없거나 못 불러오면 목록만으로도 쓸 수 있게. */
(function () {
  'use strict';

  function $(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  function api(url) {
    return fetch(url, { credentials: 'same-origin' }).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error(SESSION_ENDED);
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) throw new Error((j && j.message) || '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')');
        return j;
      });
    });
  }
  function msg(text, kind) {
    var el = $('vnMsg');
    el.className = 'gw-vh-msg' + (text ? ' ' + (kind || 'info') : '');
    el.textContent = text || '';
    if (typeof S !== 'undefined') S.loadMsg = false;
  }
  // 불러오기 실패 문구 — 다음에 잘 불러오면 지운다(전엔 한 번 뜬 실패 문구가 계속 남았음, 11-24 조용한 실패 점검).
  // 가까운 차 찾기 문구처럼 다른 데서 띄운 문구는 건드리지 않는다.
  function loadMsg(text) {
    if (text) { msg(text, 'err'); S.loadMsg = true; }
    else if (S.loadMsg) msg('');
  }
  function ago(m) {
    if (m == null) return '';
    if (m < 1) return '방금';
    if (m < 60) return m + '분 전';
    if (m < 1440) return Math.floor(m / 60) + '시간 전';
    return Math.floor(m / 1440) + '일 전';
  }
  // 주소는 시·구·동까지만 짧게(지번 끝자리는 목록에서 빼고 지도 꼬리표에서도 뺌)
  function shortAddr(a) {
    a = String(a || '').replace(/^(서울특별시|서울)\s+/, '').replace(/^경기(도)?\s+/, '').replace(/^인천(광역시)?\s+/, '인천 ');
    return a.replace(/\s+\(.*\)$/, '');
  }

  var page = $('vnPage');
  var S = { cars: [], filter: 'all', near: null, loadMsg: false, map: null, kakao: null, overlays: {}, target: null, timer: null, key: page.getAttribute('data-kakao') || '' };

  function isStop(c) { return c.lat != null && !c.moving; }     // 위치를 모르는 차는 정지로 세지 않는다(11-24)
  function counts() {
    $('vnCntAll').textContent = S.cars.length;
    $('vnCntMoving').textContent = S.cars.filter(function (c) { return c.moving; }).length;
    $('vnCntStop').textContent = S.cars.filter(isStop).length;
  }
  function shown() {
    var list = S.near ? S.near : S.cars;
    return list.filter(function (c) { return S.filter === 'all' || (S.filter === 'moving' ? c.moving : isStop(c)); });
  }
  function row(c) {
    var state = c.lat == null ? '<span class="gw-vn-state none">위치 없음</span>'
      : (c.moving ? '<span class="gw-vn-state moving">운행 중</span>' : '<span class="gw-vn-state stop">정지</span>');
    var who = (c.riders || []).length ? esc(c.riders.join(', ')) : '<span class="text-muted">오늘 배정 없음</span>';
    return '<li><button type="button" class="gw-vn-car" data-plate="' + esc(c.plate) + '"' + (c.lat == null ? ' disabled' : '') + '>' +
      '<span class="gw-vn-top"><strong>' + esc(c.label) + '</strong> <span class="gw-vn-plate">' + esc(c.plate) + '</span>' +
      (c.km != null ? '<span class="gw-vn-km">' + esc(c.km) + ' km</span>' : '') + '</span>' +
      '<span class="gw-vn-mid">' + state + ' <span class="gw-vn-who">' + who + '</span></span>' +
      // 'N분 전'은 위치를 받은 시각 기준 — 신호만 오고 위치가 비면 '새 위치 안 들어옴'(11-24 조용한 실패 점검)
      '<span class="gw-vn-addr">' + esc(shortAddr(c.addr) || '-') + (c.mins != null ? ' · ' + esc(ago(c.mins)) : '') +
        (c.pos_stale ? ' · 새 위치 안 들어옴' : '') + '</span>' +
      '</button></li>';
  }
  function renderList() {
    counts();
    var list = shown();
    $('vnCars').innerHTML = list.length ? list.map(row).join('')
      : '<li class="gw-empty">' + (S.cars.length ? '이 조건에 맞는 차가 없어요.' : 'U+ 연동 승인 뒤 차 위치가 여기에 보여요.') + '</li>';
  }

  // ── 지도 ──
  function mapMsg(text) { $('vnMap').innerHTML = '<p class="gw-vn-mapmsg">' + esc(text) + '</p>'; }
  function initMap() {
    if (!S.key) { mapMsg('지도를 준비하고 있어요. 차 위치는 왼쪽 목록에서 볼 수 있어요.'); return Promise.resolve(); }
    return window.GWKMap.load(S.key).then(function (kakao) {
      S.kakao = kakao;
      $('vnMap').innerHTML = '';
      S.map = new kakao.maps.Map($('vnMap'), { center: new kakao.maps.LatLng(37.62, 126.72), level: 9 });
      S.map.addControl(new kakao.maps.ZoomControl(), kakao.maps.ControlPosition.RIGHT);
    }).catch(function () { mapMsg('지도를 불러오지 못했어요. 차 위치는 왼쪽 목록에서 볼 수 있어요.'); });
  }
  function drawMarkers(fit) {
    if (!S.map) return;
    var kakao = S.kakao, seen = {}, bounds = new kakao.maps.LatLngBounds(), any = false;
    S.cars.forEach(function (c) {
      if (c.lat == null) return;
      var pos = new kakao.maps.LatLng(c.lat, c.lon);
      seen[c.plate] = 1;
      var el = document.createElement('button');
      el.type = 'button';
      el.className = 'gw-vn-pin' + (c.moving ? ' moving' : '') + (c.crew ? ' crew' : '');
      el.setAttribute('data-plate', c.plate);
      el.textContent = c.label;
      var o = S.overlays[c.plate];
      if (o) { o.setPosition(pos); o.setContent(el); }
      else S.overlays[c.plate] = new kakao.maps.CustomOverlay({ map: S.map, position: pos, content: el, yAnchor: 1.3, clickable: true });
      bounds.extend(pos);
      any = true;
    });
    Object.keys(S.overlays).forEach(function (p) { if (!seen[p]) { S.overlays[p].setMap(null); delete S.overlays[p]; } });
    if (S.target) bounds.extend(S.target.getPosition());
    if (fit && any) S.map.setBounds(bounds, 40, 40, 40, 40);
  }
  function focusCar(plate) {
    var c = S.cars.filter(function (x) { return x.plate === plate; })[0];
    if (!c || c.lat == null || !S.map) return;
    S.map.setLevel(5);
    S.map.panTo(new S.kakao.maps.LatLng(c.lat, c.lon));
  }

  // ── 불러오기 ──
  var first = true;
  function load(quiet) {
    if (!quiet) $('vnWhen').textContent = '불러오는 중…';
    return api('/api/vehicles/now').then(function (j) {
      S.cars = j.cars || [];
      if (S.near) {           // 가까운 차 순서를 보고 있으면 거리만 남기고 위치는 새 값으로
        var km = {};
        S.near.forEach(function (c) { km[c.plate] = c.km; });
        S.near = S.cars.filter(function (c) { return km[c.plate] != null; })
          .map(function (c) { return Object.assign({}, c, { km: km[c.plate] }); })
          .sort(function (a, b) { return a.km - b.km; });
      }
      // U+ 조회가 실패하면 옛 위치를 '지금 기준'으로 보이지 않게 — 마지막으로 받은 시각을 머리에(11-24 조용한 실패 점검)
      var when = $('vnWhen');
      var failed = j.enabled && j.error;
      // .env 의 U+ 값을 잘못 적은 것은 '승인 전'과 구별(11-24 조용한 실패 점검)
      var confBad = !j.enabled && !!j.config_error;
      when.textContent = confBad ? 'U+ 설정 오류 — 관리자 확인 필요'
        : !j.enabled ? 'U+ 연동 승인 전'
        : failed ? 'U+ 연결 실패 — ' + (j.last_ok ? j.last_ok + ' 위치예요' : '받은 위치가 없어요')
        : j.checked_at + ' 기준 · 30초마다';
      when.classList.toggle('warn', !!failed || confBad);
      loadMsg(failed ? j.error + ' 목록과 지도는 마지막으로 받은 위치예요. 잠시 뒤 저절로 다시 받아요.'
        : (j.enabled && j.config_error ? 'U+ 설정 일부가 잘못 적혀 있어요 — 관리자 확인 필요' : ''));
      renderList();
      drawMarkers(first);
      first = false;
    }).catch(function (e) { $('vnWhen').textContent = ''; $('vnWhen').classList.remove('warn'); loadMsg(e.message); });
  }
  function schedule() {
    clearInterval(S.timer);
    S.timer = setInterval(function () { if (document.visibilityState === 'visible') load(true); }, 30000);
  }

  // ── 가까운 차 ──
  $('vnNear').addEventListener('submit', function (e) {
    e.preventDefault();
    var addr = $('vnAddr').value.trim();
    if (addr.length < 2) { msg('주소를 적어 주세요.', 'err'); $('vnAddr').focus(); return; }
    var btn = this.querySelector('button[type=submit]');
    btn.disabled = true;
    msg('');
    api('/api/vehicles/now/near?addr=' + encodeURIComponent(addr)).then(function (j) {
      S.near = j.cars || [];
      S.filter = 'all';
      setFilter('all');
      $('vnNearClear').hidden = false;
      msg(S.near.length ? '가까운 차부터 보여 줘요. 가장 가까운 차: ' + S.near[0].label + ', ' + S.near[0].km + ' km' : '위치가 있는 차가 없어요.', 'ok');
      if (S.map) {
        var pos = new S.kakao.maps.LatLng(j.point[0], j.point[1]);
        if (S.target) S.target.setMap(null);
        var d = document.createElement('div');
        d.className = 'gw-kmap-pin end';
        d.textContent = '방문지';
        S.target = new S.kakao.maps.CustomOverlay({ map: S.map, position: pos, content: d, yAnchor: 1.4 });
        drawMarkers(true);
      }
    }).catch(function (err) { msg(err.message, 'err'); }).then(function () { btn.disabled = false; });
  });
  $('vnNearClear').addEventListener('click', function () {
    S.near = null;
    this.hidden = true;
    $('vnAddr').value = '';
    if (S.target) { S.target.setMap(null); S.target = null; }
    msg('');
    renderList();
    drawMarkers(true);
    $('vnAddr').focus();
  });

  // ── 거르기(전체·운행 중·정지) — 결재함 탭과 같은 키보드 ──
  function setFilter(f) {
    S.filter = f;
    Array.prototype.forEach.call($('vnFilter').querySelectorAll('button[data-f]'), function (b) {
      var on = b.getAttribute('data-f') === f;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
      b.tabIndex = on ? 0 : -1;
    });
    renderList();
  }
  $('vnFilter').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-f]');
    if (b) setFilter(b.getAttribute('data-f'));
  });
  $('vnFilter').addEventListener('keydown', function (e) {
    if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].indexOf(e.key) === -1) return;
    var tabs = Array.prototype.slice.call($('vnFilter').querySelectorAll('button[data-f]'));
    var i = tabs.indexOf(document.activeElement);
    if (i < 0) return;
    e.preventDefault();
    var n = e.key === 'Home' ? 0 : e.key === 'End' ? tabs.length - 1 : (i + (e.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
    tabs[n].focus();
    tabs[n].click();
  });
  $('vnCars').addEventListener('click', function (e) {
    var b = e.target.closest('[data-plate]');
    if (b) focusCar(b.getAttribute('data-plate'));
  });
  $('vnMap').addEventListener('click', function (e) {     // 지도 꼬리표를 누르면 목록의 그 차로
    var b = e.target.closest('.gw-vn-pin[data-plate]');
    if (!b) return;
    var li = $('vnCars').querySelector('[data-plate="' + b.getAttribute('data-plate').replace(/"/g, '') + '"]');
    if (li) { li.scrollIntoView({ block: 'nearest' }); li.focus(); }
  });
  $('vnReload').addEventListener('click', function () { load(); });
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible') load(true); });

  // 목록은 지도와 따로 동시에 — 지도가 늦거나 실패해도 목록은 뜬다. 지도가 준비되면 그때 꼬리표를 그린다
  //   (전엔 지도 불러오기가 안 끝나면 목록도 영영 안 떴음, 11-24 조용한 실패 점검)
  load().then(schedule);
  initMap().then(function () { if (S.map) drawMarkers(true); });
})();
