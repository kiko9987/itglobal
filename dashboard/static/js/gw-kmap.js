/* 카카오 지도 불러오기 (2026-11-24) — 실시간 위치·운행 경로가 같이 쓴다.
   키(KAKAO_JS_KEY)는 서버가 화면에 넘겨 준다. 키가 없거나 못 불러오면 reject — 화면은 목록만 보여 준다.
   카카오 개발자 › 내 애플리케이션 › 플랫폼 › Web 에 그룹웨어 주소(https://gw.itg-aircon.com 등)를 넣어야 지도가 뜬다. */
(function () {
  'use strict';
  var loading = null;
  function load(key) {
    if (window.kakao && window.kakao.maps && window.kakao.maps.LatLng) return Promise.resolve(window.kakao);
    if (!key) return Promise.reject(new Error('no_key'));
    if (loading) return loading;
    loading = new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = 'https://dapi.kakao.com/v2/maps/sdk.js?autoload=false&appkey=' + encodeURIComponent(key);
      s.async = true;
      s.onload = function () {
        if (!window.kakao || !window.kakao.maps) { loading = null; reject(new Error('load_failed')); return; }
        window.kakao.maps.load(function () { resolve(window.kakao); });
      };
      s.onerror = function () { loading = null; reject(new Error('load_failed')); };
      document.head.appendChild(s);
    });
    return loading;
  }
  // 경로 선을 그린 지도 하나 — el 안에 새로 만든다. pts = [[위도, 경도], ...]
  function route(el, key, pts) {
    return load(key).then(function (kakao) {
      el.innerHTML = '';
      var path = pts.map(function (p) { return new kakao.maps.LatLng(p[0], p[1]); });
      var map = new kakao.maps.Map(el, { center: path[0], level: 6 });
      new kakao.maps.Polyline({ map: map, path: path, strokeWeight: 4, strokeColor: '#2563eb', strokeOpacity: 0.85 });
      var ends = [[path[0], '출발'], [path[path.length - 1], '도착']];
      ends.forEach(function (e) {
        var d = document.createElement('div');
        d.className = 'gw-kmap-pin' + (e[1] === '도착' ? ' end' : '');
        d.textContent = e[1];
        new kakao.maps.CustomOverlay({ map: map, position: e[0], content: d, yAnchor: 1.4 });
      });
      var b = new kakao.maps.LatLngBounds();
      path.forEach(function (p) { b.extend(p); });
      map.setBounds(b, 30, 30, 30, 30);
      return map;
    });
  }
  window.GWKMap = { load: load, route: route };
})();
