/* 매출 통계(11-28 전체 점검 3차) — 옛 PM 통계(src/js/pages/stats-page.js)와 같은 계산을 그룹웨어 틀로.
 *  - 기준: 공사 확정 날짜, 총액1(없으면 '총액 1'), 담당자별은 직원 명단 비영업(non_sales) 제외
 *  - 바뀐 점: 올해·작년을 한 그래프에 겹침, 아직 안 온 달은 그리지 않음(전엔 12월이 0으로 떨어져 보였음),
 *    요약 칸은 '작년 같은 날짜까지'와 비교, 담당자별은 한 색 가로 막대
 */
(function () {
  'use strict';

  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  var CUR = new Date().getFullYear();
  var PREV = CUR - 1;
  // 그래프 색 — 홈 매출 그래프와 같은 색. 테마 색 값(--s-chart*, 1안 주황)이 있으면 그걸, 없으면(옛 남색 ?theme=b) 예전 파랑.
  // 테마 클래스가 붙은 뒤 읽도록 그릴 때마다 다시 읽음(palette())
  var C_CUR = '#2f6fde';      // 올해
  var C_PREV = '#cbd5e1';     // 작년 — 옅은 회색
  var C_G = '#2f6fde';        // 글로벌
  var C_R = '#334155';        // 글로벌그룹
  function palette() {
    var cs = getComputedStyle(document.body);
    var v = function (n, fb) { return cs.getPropertyValue(n).trim() || fb; };
    C_CUR = C_G = v('--s-chart', '#2f6fde');
    C_PREV = v('--s-chart-prev', '#cbd5e1');
    C_R = v('--s-chart-2', '#334155');
  }
  var charts = {};

  function $(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function getJson(url) {
    return fetch(url, { credentials: 'same-origin', headers: { Accept: 'application/json', 'X-Requested-With': 'XMLHttpRequest' } })
      .then(function (r) {
        var ct = r.headers.get('content-type') || '';
        if (r.status === 401 || r.redirected || (r.ok && ct.indexOf('application/json') === -1)) throw new Error(SESSION_ENDED);
        return r.json().catch(function () { return {}; }).then(function (j) {
          if (!r.ok || j.success === false) throw new Error(j.message || j.error || ('불러오지 못했어요 (' + r.status + ')'));
          return j;
        });
      });
  }

  // '2026-11-27', '2026. 11. 27', '2026/11/27 …' 모두
  function ymd(v) {
    var m = /(\d{4})\D+(\d{1,2})\D+(\d{1,2})/.exec(String(v || ''));
    return m ? { y: +m[1], m: +m[2] - 1, d: +m[3] } : null;
  }
  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var n = parseFloat(String(v == null ? '' : v).replace(/[^\d.-]/g, ''));
    return isFinite(n) ? n : 0;
  }
  function eok(v) {   // 억 원 한 자리
    return (v / 1e8).toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + '억';
  }
  function pct(cur, prev) {
    if (!prev) return null;
    return Math.round((cur - prev) / prev * 1000) / 10;
  }
  function delta(cur, prev, unit) {
    var p = pct(cur, prev);
    if (p === null) return '<span class="gw-st-d">작년 같은 기간 자료 없음</span>';
    var cls = p > 0 ? 'up' : p < 0 ? 'down' : 'flat';
    var arrow = p > 0 ? '▲' : p < 0 ? '▼' : '–';
    return '<span class="gw-st-d ' + cls + '">' + arrow + ' ' + Math.abs(p).toLocaleString('ko-KR') + '%</span>' +
      ' <span class="gw-st-vs">작년 같은 기간 ' + esc(unit(prev)) + '</span>';
  }

  function compute(projects, nonSales) {
    var today = new Date();
    var cm = today.getMonth(), cd = today.getDate();
    var S = {
      month: { cur: Array(12).fill(0), prev: Array(12).fill(0) },
      co: { cur: { g: Array(12).fill(0), r: Array(12).fill(0) }, prevTotal: { g: 0, r: 0 }, curTotal: { g: 0, r: 0 } },
      mgr: {},
      rev: { cur: 0, prev: 0, prevSame: 0 }, cnt: { cur: 0, prev: 0, prevSame: 0 },
      out: { cur: 0, prev: 0 }
    };
    var skip = {};
    (nonSales || []).forEach(function (n) { skip[String(n).trim()] = 1; });
    projects.forEach(function (p) {
      var d = ymd(p['공사 확정']);
      if (!d || (d.y !== CUR && d.y !== PREV)) return;
      var amt = num(p['총액1'] != null && p['총액1'] !== '' ? p['총액1'] : p['총액 1']);
      var key = d.y === CUR ? 'cur' : 'prev';
      S.month[key][d.m] += amt;
      S.rev[key] += amt;
      S.cnt[key] += 1;
      S.out[key] += Math.max(0, num(p['미수금']));
      // 작년 같은 날짜까지(오늘이 11/28 이면 작년 1/1~11/28)
      if (key === 'prev' && (d.m < cm || (d.m === cm && d.d <= cd))) { S.rev.prevSame += amt; S.cnt.prevSame += 1; }
      var co = String(p['사업자'] || '').trim();
      var ck = co === '글로벌' ? 'g' : co === '글로벌그룹' ? 'r' : '';
      if (ck) {
        if (key === 'cur') { S.co.cur[ck][d.m] += amt; S.co.curTotal[ck] += amt; } else { S.co.prevTotal[ck] += amt; }
      }
      var who = String(p['담당자'] || '').trim();
      if (who && !skip[who]) {
        var m = S.mgr[who] || (S.mgr[who] = { cur: 0, prev: 0 });
        m[key] += amt;
      }
    });
    return S;
  }

  function renderStats(S) {
    var months = new Date().getMonth() + 1;
    var avgCur = S.rev.cur / months, avgPrev = S.rev.prev / 12;
    $('stStats').innerHTML =
      '<div class="gw-stat gw-st-main"><div class="k">' + CUR + '년 매출 · 1월~오늘</div>' +
        '<div class="v">' + esc(eok(S.rev.cur)) + '</div><div class="s">' + delta(S.rev.cur, S.rev.prevSame, eok) + '</div></div>' +
      '<div class="gw-stat"><div class="k">' + CUR + '년 확정 공사</div>' +
        '<div class="v">' + S.cnt.cur.toLocaleString('ko-KR') + ' <small>건</small></div>' +
        '<div class="s">' + delta(S.cnt.cur, S.cnt.prevSame, function (v) { return v.toLocaleString('ko-KR') + '건'; }) + '</div></div>' +
      '<div class="gw-stat"><div class="k">월 평균 매출 · ' + months + '개월</div>' +
        '<div class="v">' + esc(eok(avgCur)) + '</div><div class="s"><span class="gw-st-vs">' + PREV + '년 월 평균 ' + esc(eok(avgPrev)) + '</span></div></div>' +
      '<div class="gw-stat"><div class="k">미수금 · ' + CUR + '년 확정분</div>' +
        '<div class="v">' + esc(eok(S.out.cur)) + '</div><div class="s"><span class="gw-st-vs">' + PREV + '년 확정분 ' + esc(eok(S.out.prev)) + '</span></div></div>';
  }

  function axisEok() {
    return { beginAtZero: true, grid: { color: '#eef2f7' }, ticks: { callback: function (v) { return v + '억'; }, color: '#64748b' } };
  }
  function tipEok(ctx) {
    var v = ctx.parsed.y != null && ctx.chart.options.indexAxis !== 'y' ? ctx.parsed.y : ctx.parsed.x;
    return ' ' + ctx.dataset.label + ' ' + (v == null ? '-' : v.toLocaleString('ko-KR', { maximumFractionDigits: 2 }) + '억');
  }
  function draw(id, cfg) {
    if (charts[id]) charts[id].destroy();
    var c = $(id);
    if (!c || typeof Chart === 'undefined') return;
    cfg.options = Object.assign({ responsive: true, maintainAspectRatio: false, animation: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: tipEok } } } }, cfg.options || {});
    charts[id] = new Chart(c.getContext('2d'), cfg);
  }

  function renderCharts(S) {
    palette();
    var M = ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'];
    var cm = new Date().getMonth();
    var e = function (v) { return Math.round(v / 1e6) / 100; };   // 억, 소수 둘째 자리
    // 아직 안 온 달은 null — 그리지 않음(전엔 0으로 떨어져 매출이 급감한 것처럼 보였음)
    var cur = S.month.cur.map(function (v, i) { return i <= cm ? e(v) : null; });
    draw('stMonth', {
      type: 'bar',
      data: { labels: M, datasets: [
        { label: PREV + '년', data: S.month.prev.map(e), backgroundColor: C_PREV, borderRadius: 4, order: 2 },
        { label: CUR + '년', data: cur, backgroundColor: C_CUR, borderRadius: 4, order: 1 }
      ] },
      options: { scales: { y: axisEok(), x: { grid: { display: false }, ticks: { color: '#64748b' } } } }
    });
    $('stMonth').setAttribute('aria-label', '월별 매출 — ' + M.map(function (m, i) {
      return m + ' ' + CUR + '년 ' + (cur[i] == null ? '-' : cur[i] + '억') + ', ' + PREV + '년 ' + e(S.month.prev[i]) + '억';
    }).join('; '));

    draw('stCompany', {
      type: 'bar',
      data: { labels: ['글로벌', '글로벌그룹'], datasets: [
        // 올해를 앞(왼쪽)에 — 월별·담당자별 그래프와 같은 순서(작년이 앞이던 것, 11-28 부장님)
        { label: CUR + '년', data: [e(S.co.curTotal.g), e(S.co.curTotal.r)], backgroundColor: C_CUR, borderRadius: 4 },
        { label: PREV + '년', data: [e(S.co.prevTotal.g), e(S.co.prevTotal.r)], backgroundColor: C_PREV, borderRadius: 4 }
      ] },
      options: { plugins: { legend: { display: true, position: 'bottom', labels: { boxWidth: 12, color: '#475569' } }, tooltip: { callbacks: { label: tipEok } } },
        scales: { y: axisEok(), x: { grid: { display: false }, ticks: { color: '#334155' } } } }
    });
    $('stCompany').setAttribute('aria-label', '사업자별 매출 — 글로벌 ' + CUR + '년 ' + e(S.co.curTotal.g) + '억(' + PREV + '년 ' + e(S.co.prevTotal.g) +
      '억), 글로벌그룹 ' + CUR + '년 ' + e(S.co.curTotal.r) + '억(' + PREV + '년 ' + e(S.co.prevTotal.r) + '억)');

    var mgr = Object.keys(S.mgr).map(function (k) { return { n: k, cur: S.mgr[k].cur, prev: S.mgr[k].prev }; })
      .filter(function (x) { return x.cur > 0 || x.prev > 0; })
      .sort(function (a, b) { return b.cur - a.cur || b.prev - a.prev; });
    var box = $('stManager').parentNode;
    box.style.height = Math.max(240, mgr.length * 34 + 40) + 'px';
    draw('stManager', {
      type: 'bar',
      data: { labels: mgr.map(function (x) { return x.n; }), datasets: [
        { label: CUR + '년', data: mgr.map(function (x) { return e(x.cur); }), backgroundColor: C_CUR, borderRadius: 3, barPercentage: 0.9, categoryPercentage: 0.75 },
        { label: PREV + '년', data: mgr.map(function (x) { return e(x.prev); }), backgroundColor: C_PREV, borderRadius: 3, barPercentage: 0.9, categoryPercentage: 0.75 }
      ] },
      options: { indexAxis: 'y', scales: { x: axisEok(), y: { grid: { display: false }, ticks: { color: '#334155' } } } }
    });
    $('stManager').setAttribute('aria-label', '담당자별 매출 — ' + mgr.map(function (x) {
      return x.n + ' ' + CUR + '년 ' + e(x.cur) + '억, ' + PREV + '년 ' + e(x.prev) + '억';
    }).join('; '));

    draw('stCoMonth', {
      type: 'bar',
      data: { labels: M.slice(0, cm + 1), datasets: [
        { label: '글로벌', data: S.co.cur.g.slice(0, cm + 1).map(e), backgroundColor: C_G, borderRadius: 2, stack: 'a' },
        { label: '글로벌그룹', data: S.co.cur.r.slice(0, cm + 1).map(e), backgroundColor: C_R, borderRadius: 2, stack: 'a' }
      ] },
      options: { scales: { y: Object.assign(axisEok(), { stacked: true }), x: { stacked: true, grid: { display: false }, ticks: { color: '#64748b' } } } }
    });
    $('stCoMonth').setAttribute('aria-label', CUR + '년 사업자별 월 매출 — ' + M.slice(0, cm + 1).map(function (m, i) {
      return m + ' 글로벌 ' + e(S.co.cur.g[i]) + '억, 글로벌그룹 ' + e(S.co.cur.r[i]) + '억';
    }).join('; '));
  }

  function load() {
    var btn = $('stRefresh');
    if (btn) btn.disabled = true;
    Promise.all([
      getJson('/api/projects/list'),
      getJson('/api/staff-directory').catch(function () { return {}; })
    ]).then(function (res) {
      var j = res[0];
      var projects = Array.isArray(j) ? j : (Array.isArray(j.data) ? j.data : []);
      var S = compute(projects, (res[1] && res[1].non_sales) || []);
      renderStats(S);
      if (typeof Chart === 'undefined') {
        document.querySelectorAll('.gw-st-chart').forEach(function (el) { el.innerHTML = '<div class="gw-empty">그래프 도구를 불러오지 못했어요. 새로 고쳐 주세요.</div>'; });
        return;
      }
      renderCharts(S);
    }).catch(function (e) {
      $('stStats').innerHTML = '<div class="gw-stat"><div class="k">매출 통계</div><div class="v gw-st-err">' + esc(e.message || '불러오지 못했어요.') + '</div></div>';
    }).then(function () { if (btn) btn.disabled = false; });
  }

  document.querySelectorAll('[data-y="cur"]').forEach(function (el) { el.textContent = CUR + '년'; });
  document.querySelectorAll('[data-y="prev"]').forEach(function (el) { el.textContent = PREV + '년'; });
  if ($('stRefresh')) $('stRefresh').addEventListener('click', load);
  load();
})();
