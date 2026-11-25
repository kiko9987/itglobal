/* 그룹웨어 차량 배정 (2026-11-20) — 다음 날 차마다 탈 사람. 고치기는 배정 담당자(SJ)·인사 관리자만.
   카톡 배정을 대신한다(카톡 글 붙여넣기는 넣지 않음 — 부장님 11-20).
   코스 중심(11-23 부장님 1안): 방문 일정과 같은 코스 카드에서 차를 고른다. 코스 전체 한 차면 카드 머리에서 한 번,
   갈라 타면 사람 칩마다(오전·오후 따로도). 왼쪽은 차 목록(누가 탔는지) — 방문 일정처럼 줄을 눌러 체크한 뒤 코스를 누름 — 차를 끌어 코스나 사람에 놓아도 된다.
   저장은 예전 그대로 차마다 사람(오전 people · 오후 pm). 서버에서 온 글자는 모두 esc() 로만 넣는다. */
(function () {
  'use strict';

  function $(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  function api(url, init) {
    init = init || {};
    init.credentials = 'same-origin';
    var method = init.method || 'GET';
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error(SESSION_ENDED);
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          var err = new Error((j && j.message) || (method === 'GET'
            ? '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'
            : '저장하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'));
          err.status = r.status;       // 409 = 다른 분이 먼저 고침(11-24 조용한 실패 점검 17번)
          throw err;
        }
        return j;
      });
    });
  }
  function post(url, body) {
    return api(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  }
  // near = true 면 아래 단추 줄 바로 위에(저장 결과 — 휴대폰에서 위쪽 안내는 화면 밖, 11-21 점검)
  function msg(text, kind, near) {
    var bar = $('vaBarMsg');
    var useBar = near && bar && !$('vaBar').hidden;
    [$('vaMsg'), bar].forEach(function (el) {
      if (!el) return;
      var on = useBar ? el === bar : el.id === 'vaMsg';
      el.className = 'gw-vh-msg' + (el === bar ? ' gw-va-barmsg' : '') + (on && text ? ' ' + (kind || 'info') : '');
      el.textContent = on ? (text || '') : '';
    });
  }

  var TOUCH = !!(window.matchMedia && window.matchMedia('(pointer: coarse)').matches);
  var S = { day: '', cars: [], staff: [], names: {}, canEdit: false, confirmed: false, dirty: false,
            visits: {}, places: {}, teams: [], seq: 0, focus: null, placeTries: 0, placeTimer: null,
            off: {}, half: {}, online: {}, work: {}, schedule: [], courses: null, plan: null, dayLabel: '',
            splitOpen: {}, splitCourse: {}, memoOpen: {}, pick: '', pickHalf: '', duty: {}, ownPick: {}, extra: [], carWhen: '', draftNote: '',
            base: null, partial: [] };
  // base = 불러올 때 받은 그날 배정 판 — 저장할 때 돌려보내 그 사이 다른 분이 고쳤으면 서버가 409(11-24 조용한 실패 점검 17번)
  // partial = 서버가 못 읽은 칸('leads'·'off' 등) — 화면 위 경고 줄(11-24 조용한 실패 점검 18번)
  // ownPick = 여러 명 코스에서 [본인 차량으로 이동]을 켜 '누가 자차로 가는지' 고르는 중인 코스
  // pickHalf = 이미 코스에 든 차를 체크했을 때 새 코스가 쓸 반나절('am'|'pm') · duty = 온라인 당번 {이니셜: '오전'|'오후'}
  // pick = 체크한 차(방문 일정처럼 줄을 눌러 체크 → 코스를 누르거나 위 줄에서 코스 고르기, 11-24 부장님) · splitCourse = [갈라 타기]를 연 코스
  // courses = 그룹웨어에서 확정한 방문 일정의 코스 — 슬랙으로 확정한 날은 null(그땐 일정확정 팀을 카드로)
  // splitOpen = {이니셜: 1} 오전·오후 따로 고르는 중 · extra = '그 밖'에 더한 사람(코스에 없는 사람 — 미팅 등)
  // off = 종일 휴무(차에 못 넣음), half = 반차·확정 전 토요 휴무(꼬리표만), online = 온라인 당번, work = 사내 업무

  function name(code) { return S.names[code] || code; }
  function carOf(plate) { return S.cars.filter(function (c) { return c.plate === plate; })[0]; }
  function isOwn(c) { return c.usage === 'own'; }
  // 시공팀 차(11-23 부장님) — 포터 2대·전기트럭은 평소 시공팀이 끌고 나가고, 영업은 현장 지원·물품 운반 때만.
  //   차 종류 이름(포터·트럭)으로 알아본다(차량 관리 용도는 그대로 '공용'). 영업 차 대수에서 빼고 목록 맨 아래 따로
  function isCrew(c) { return c.usage === 'shared' && /포터|트럭/.test((c.model || '') + ' ' + c.label); }
  function carsInOrder() {      // 영업 차 → 시공팀 차. 자차 칸은 목록에서 뺐다(11-24 부장님 — 차 없는 코스는 메모 한 줄로)
    return S.cars.filter(function (c) { return !isCrew(c) && !isOwn(c); }).concat(S.cars.filter(isCrew));
  }
  function ownCar() { return S.cars.filter(isOwn)[0] || null; }           // 자차 칸(11-23 — 차가 모자라면 자기 차로)
  function isSplit(c) { return !!c.pm; }
  function pmOf(c) { return c.pm ? c.pm : c.people; }
  function ridersOf(c) {
    var out = c.people.slice();
    pmOf(c).forEach(function (p) { if (out.indexOf(p) === -1) out.push(p); });
    return out;
  }
  function whereAt(code, half) {
    for (var i = 0; i < S.cars.length; i++) {
      if ((half === 'pm' ? pmOf(S.cars[i]) : S.cars[i].people).indexOf(code) !== -1) return S.cars[i];
    }
    return null;
  }
  function whereIs(code) { return whereAt(code, 'am') || whereAt(code, 'pm'); }
  function sameList(a, b) { return a.length === b.length && a.every(function (x, i) { return x === b[i]; }); }
  function tidy(c) { if (c.pm && sameList(c.pm, c.people)) c.pm = null; }
  function uniq(list) { return list.filter(function (x, i) { return list.indexOf(x) === i; }); }
  // 한 사람을 그 반나절의 차에서 뺀다. 종일 탄 차에서 반나절만 빼면 그 차는 오전·오후로 나뉜다
  function removeHalf(code, half) {
    S.cars.forEach(function (c) {
      if (half === 'am' && c.people.indexOf(code) !== -1) {
        if (!isSplit(c)) c.pm = c.people.slice();
        c.people.splice(c.people.indexOf(code), 1);
      } else if (half === 'pm' && pmOf(c).indexOf(code) !== -1) {
        if (!isSplit(c)) c.pm = c.people.slice();
        c.pm.splice(c.pm.indexOf(code), 1);
      }
      tidy(c);
    });
  }
  // 한 사람의 차를 정한다. half = 'all'(종일) / 'am' / 'pm', plate '' = 뺀다. 휴무(종일)는 못 넣음
  function setCar(code, plate, half) {
    if (S.off[code]) return;
    if (half === 'all') { removeHalf(code, 'am'); removeHalf(code, 'pm'); } else removeHalf(code, half);
    if (plate) {
      var c = carOf(plate);
      if (half === 'all') { c.people.push(code); if (c.pm) c.pm.push(code); }
      else {
        if (!c.pm) c.pm = c.people.slice();
        (half === 'am' ? c.people : c.pm).push(code);
      }
      tidy(c);
    }
    setDirty(true);
  }
  function addDays(day, n) {
    var d = new Date(day + 'T00:00:00');
    d.setDate(d.getDate() + n);
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function setDirty(v) {
    S.dirty = v;
    var b = $('vaSave');
    if (b) b.disabled = !v;                              // 바뀐 게 없으면 [저장만]은 꺼 둔다
    var c = $('vaConfirm');
    if (c && S.confirmed) c.disabled = !v;               // 확정 뒤에는 고친 게 있을 때만
  }
  function focusAfter(sel) { S.focus = sel; }
  function restoreFocus() {
    var sel = S.focus;
    S.focus = null;
    if (!sel) return;
    var el = document.querySelector(sel);
    if (el) { try { el.focus({ preventScroll: true }); } catch (e) { el.focus(); } }
  }
  function cssq(v) { return String(v).replace(/["\\]/g, '\\$&'); }
  function ownCars() {                // 고정 차 주인 → 그 차(YG 카니발·MW 스파크)
    var own = {};
    S.cars.forEach(function (c) { if (c.usage === 'fixed' && c.fixed_user && S.names[c.fixed_user]) own[c.fixed_user] = c; });
    return own;
  }
  // 본인 차(개인 소유 고정 차 — 스파크 1783 = MW, 인센티브 구간이라 늘 자기 차로 다님, 11-25 부장님)
  function isPersonal(c) { return !!c && c.usage === 'fixed' && !!c.fixed_user && /^개인/.test(c.corp || ''); }
  // 방문 일정 코스에 든 본인 차 주인이 아직 차가 없으면 그 차에 종일로 넣는다(저장해야 반영). ['박민우 님 → 스파크 1783', …]
  function ownToCar() {
    if (!S.courses) return [];
    var inG = groupOf(), out = [];
    S.cars.forEach(function (c) {
      var m = c.fixed_user;
      if (!isPersonal(c) || !inG[m] || !S.names[m] || S.off[m] || !noCar(m)) return;
      setCar(m, c.plate, 'all');
      out.push(name(m) + ' 님 → ' + c.label);
    });
    return out;
  }
  // 고정 차 주인을 '배정 없어도 자기 차로 간다'고 짐작하던 것 없앰(11-24 부장님 — 빈 데서 차를 넣는 화면).
  //   고정 차도 다른 차처럼 코스에 넣어야 들어간다. (운행일지 운전자 추정의 '고정 사용자'는 그대로)
  function byOwnFixed() { return false; }
  function noCar(code) { return !S.off[code] && !whereIs(code) && !byOwnFixed(code); }
  function riders() {
    var seen = {};
    S.cars.forEach(function (c) { ridersOf(c).forEach(function (p) { seen[p] = 1; }); });
    return Object.keys(seen).length;
  }
  function whereText(code) {       // '레이 6854' / '레이 6854 오전 · 포터 4904 오후' / '레이 6854 오전만'
    var am = whereAt(code, 'am'), pm = whereAt(code, 'pm');
    if (!am && !pm) return byOwnFixed(code) ? ownCars()[code].label : '';
    if (am && pm && am === pm) return am.label;
    if (am && pm) return am.label + ' 오전 · ' + pm.label + ' 오후';
    return am ? am.label + ' 오전만' : pm.label + ' 오후만';
  }
  function splitPerson(code) { return whereAt(code, 'am') !== whereAt(code, 'pm') || !!S.splitOpen[code]; }

  // ── 코스(카드 한 장 = 함께 가는 사람 묶음) ──
  function courseMode() { return !!S.courses; }
  function groups() {
    if (S.courses) {
      return S.courses.courses.filter(function (c) { return c.kind !== 'work'; }).map(function (c, i) {
        return { key: 'c' + i, label: c.label, members: c.members.filter(function (m) { return S.names[m]; }),
                 memo: c.memo || '', sites: c.sites || [] };
      });
    }
    var n = 0;
    return S.teams.map(function (t, i) {
      var label = t.label || (t.members.length > 1 ? '함께 ' + (++n) : '혼자');
      return { key: 't' + i, label: label, members: t.members.filter(function (m) { return S.names[m]; }),
               memo: t.memo || '', sites: [], visits: t.visits };
    }).filter(function (g) { return g.members.length; });
  }
  function groupOf() {             // 이니셜 → 든 코스 이름들(두 코스면 둘)
    var out = {};
    groups().forEach(function (g) { g.members.forEach(function (m) { (out[m] = out[m] || []).push(g.label); }); });
    return out;
  }
  function others(inG) {           // 코스 밖인데 차에 넣었거나 '그 밖'에 더한 사람
    return S.staff.map(function (p) { return p.code; }).filter(function (c) {
      return !inG[c] && (whereIs(c) || S.extra.indexOf(c) !== -1);
    });
  }
  function noCarList() {          // 방문 있는데(또는 코스에 있는데) 차 없음 — '이름(코스 N)'
    var inG = groupOf();
    var codes = uniq(Object.keys(inG).concat(Object.keys(S.visits).filter(function (c) { return S.names[c]; })));
    return codes.filter(noCar).map(function (c) { return name(c) + (inG[c] ? '(' + inG[c][0] + ')' : ''); });
  }
  function staffNoCar() { return noCarList(); }

  // ── 차 고르기 칸 ──
  // 고르는 칸의 차 이름 옆 = 지금 그 차에 탄 사람(고르는 사람 빼고) — 비었는지 바로 보이게
  function carOptions(sel, except) {
    function opt(c) {
      var who = isOwn(c) ? [] : ridersOf(c).filter(function (p) { return except.indexOf(p) === -1; }).map(name);
      // 고른 차는 이름만(같이 탄 사람은 왼쪽 차 목록에), 나머지는 지금 탄 사람 또는 '빈 차'
      var text = isOwn(c) || sel === c.plate ? c.label
        : c.label + (who.length ? ' · ' + who.join(', ') : isCrew(c) ? '' : ' · 빈 차');   // 시공팀 차는 '빈 차'가 아님
      return '<option value="' + esc(c.plate) + '"' + (sel === c.plate ? ' selected' : '') + '>' + esc(text) + '</option>';
    }
    var shared = S.cars.filter(function (c) { return c.usage === 'shared' && !isCrew(c); });
    var fixed = S.cars.filter(function (c) { return c.usage === 'fixed'; });
    var crew = S.cars.filter(isCrew);
    return (shared.length ? '<optgroup label="회사 차">' + shared.map(opt).join('') + '</optgroup>' : '') +
      (fixed.length ? '<optgroup label="고정 차(빌려 쓰기)">' + fixed.map(opt).join('') + '</optgroup>' : '') +
      (crew.length ? '<optgroup label="시공팀 차량(현장 지원·물품 운반 때)">' + crew.map(opt).join('') + '</optgroup>' : '');
  }
  function personSel(code, half) {
    var at = whereAt(code, half === 'pm' ? 'pm' : 'am');
    var own = null;
    var none = '차 없음';
    var hl = half === 'am' ? ' 오전' : half === 'pm' ? ' 오후' : '';
    return '<select class="form-select form-select-sm gw-vc-sel' + (!at && !own ? ' none' : '') + '" data-pcar="' + esc(code) +
      '" data-half="' + half + '" aria-label="' + esc(name(code) + hl) + ' 차">' +
      '<option value=""' + (at ? '' : ' selected') + '>' + esc(none) + '</option>' + carOptions(at ? at.plate : '', [code]) +
      '</select>';
  }
  // 코스의 차 — 모두 같은 차(종일)면 그 번호판, 아무도 없으면 '', 섞였으면 '__mixed', 모두 자기 고정 차면 '__own'
  function onOwn(m) {               // 종일 본인 차량(자차 칸)
    var am = whereAt(m, 'am');
    return !!am && isOwn(am) && whereAt(m, 'pm') === am;
  }
  function withoutOwn(ms) {        // 몇 사람만 본인 차량이면 그 사람들은 빼고 코스 차를 본다(모두면 그대로 — 코스 전체 본인 차량)
    var rest = ms.filter(function (m) { return !onOwn(m); });
    return rest.length ? rest : ms;
  }
  // 현장 꼬리표 — '공사확정'은 안 보인다(11-24 부장님). 서버에서도 뺐고, 재시작 전 응답에 남아 있어도 거른다
  function marksOf(ms) { return (ms || []).filter(function (m) { return m !== '공사확정'; }); }
  function courseCar(g) {
    var ms = g.members.filter(function (m) { return !S.off[m]; });
    if (!ms.length) return '__none';
    var inG = groupOf();
    var one = ms.filter(function (m) { return (inG[m] || []).length < 2; });   // 두 코스에 든 사람은 다른 코스 차일 수 있어 빼고 본다
    if (one.length) ms = one;
    ms = withoutOwn(ms);
    var vals = uniq(ms.map(function (m) {
      if (splitPerson(m)) return '__mixed';
      var at = whereAt(m, 'am');
      return at ? at.plate : (byOwnFixed(m) ? '__own' : '');
    }));
    return vals.length === 1 ? vals[0] : '__mixed';
  }
  // ── 차 한 대를 두 코스가 오전·오후로 나눠 쓰기(11-24 부장님 — 온라인 당번이 오전·오후로 갈리면 생김.
  //    예: 8021 오전 민재(오후 당번) / 오후 윤구·태훈(태훈 오전 당번)). 저장은 차의 people(오전)·pm(오후) 그대로
  function courseMembersForCar(g) {   // courseCar 와 같은 사람들(휴무·두 코스에 든 사람 빼고)
    var ms = g.members.filter(function (m) { return !S.off[m]; });
    var inG = groupOf();
    var one = ms.filter(function (m) { return (inG[m] || []).length < 2; });
    return withoutOwn(one.length ? one : ms);
  }
  function courseHalf(g) {            // 코스 사람 모두 같은 차의 같은 반나절만 → {car, half} / 아니면 null
    var ms = courseMembersForCar(g);
    if (!ms.length) return null;
    var out = null;
    for (var i = 0; i < ms.length; i++) {
      var am = whereAt(ms[i], 'am'), pm = whereAt(ms[i], 'pm');
      if ((am && pm) || (!am && !pm)) return null;
      var k = { car: am || pm, half: am ? 'am' : 'pm' };
      if (out && (out.car !== k.car || out.half !== k.half)) return null;
      out = k;
    }
    return out;
  }
  function halfName(h) { return h === 'am' ? '오전' : '오후'; }
  // 체크할 수 있는 '이미 코스에 든 차' — 비어 있는 반나절('am'|'pm'), 종일 한 코스면 'both', 못 나누면 ''
  function splitFree(c) {
    if (isOwn(c) || !ridersOf(c).length) return '';
    if (!isSplit(c)) return 'both';
    if (!c.people.length) return 'am';
    if (!c.pm.length) return 'pm';
    return '';
  }
  function halfGuess(c) {             // 새 코스가 쓸 반나절 — 지금 타는 사람의 온라인 당번으로(오후 당번이면 그 사람은 오전에 차를 씀)
    var f = splitFree(c);
    if (f === 'am' || f === 'pm') return f;
    var rs = ridersOf(c);
    for (var i = 0; i < rs.length; i++) {
      if (S.duty[rs[i]] === '오후') return 'pm';
      if (S.duty[rs[i]] === '오전') return 'am';
    }
    return 'pm';
  }
  function unsplitEmpty(keep) {       // 나눠 쓰던 짝이 빠지면 남은 코스가 종일 쓰는 것으로 되돌림(keep = 건드리지 않을 차)
    S.cars.forEach(function (c) {
      if (!c.pm || c === keep) return;
      if (c.people.length && !c.pm.length) c.pm = null;
      else if (!c.people.length && c.pm.length) { c.people = c.pm; c.pm = null; }
    });
  }
  function putSplit(key, plate, half) {   // 이미 다른 코스가 쓰는 차를 이 코스가 반나절(half) 쓴다 — 다른 코스는 반대 반나절만
    var c = carOf(plate), g = groupByKey(key);
    if (!c || !g) return;
    var mine = courseMembersForCar(g);
    (half === 'am' ? c.people : pmOf(c)).slice().forEach(function (m) {
      if (mine.indexOf(m) === -1) removeHalf(m, half);
    });
    mine.forEach(function (m) {
      delete S.splitOpen[m];
      setCar(m, '', 'all');
      setCar(m, plate, half);
    });
    unsplitEmpty(c);
  }
  // 사람마다 차 고르기(갈라 타기·오전·오후 따로)는 없앴다(11-24 부장님 — 드문 경우라 코스 메모 한 줄로).
  //   예전에 사람마다 다르게 저장한 날은 차 칸에 실제 차 이름만 나열해 보여 주고, 차를 다시 넣으면 코스 모두 같은 차로
  // 코스 머리의 차 칸(11-23 부장님: 코스·사람 묶음은 끝났고 차를 코스에 넣는 일) — 왼쪽 차를 누르면 여기로
  function courseSlot(g) {
    var cur = courseCar(g);
    if (cur === '__none') return '';
    var x = function (label) {
      return S.canEdit ? '<button type="button" class="gw-va-x" data-clearc="' + esc(g.key) + '" aria-label="' + esc(g.label) + '에서 ' +
        esc(label) + ' 빼기"><i class="fas fa-xmark" aria-hidden="true"></i></button>' : '';
    };
    // 갈라 탄 코스 — '사람마다 다름' 대신 실제 차를 나열하고, 차 없는 사람은 주황으로 따로(11-23 부장님)
    //   예) '쏘렌토 8021 · 자차' / '쏘렌토 4041' + '차 없음 1명' / '쏘렌토 8021 · 레이 5358(오후)'
    var hf = courseHalf(g);
    if (hf) {
      return '<span class="gw-vc-slot on"><i class="fas fa-car-side" aria-hidden="true"></i> ' + esc(hf.car.label) +
        ' <span class="gw-vc-half">' + halfName(hf.half) + '</span>' + x(hf.car.label + ' ' + halfName(hf.half)) + '</span>';
    }
    if (cur === '__mixed') {
      var ms = g.members.filter(function (m) { return !S.off[m]; });
      var names = [];
      var add = function (t) { if (t && names.indexOf(t) === -1) names.push(t); };
      var lab = function (c) { return c.label; };
      ms.forEach(function (m) {
        var am = whereAt(m, 'am'), pm = whereAt(m, 'pm');
        if (am) add(lab(am));
        if (pm && pm !== am) add(lab(pm) + '(오후)');
        if (!am && !pm && byOwnFixed(m)) add(ownCars()[m].label);
      });
      var miss = ms.filter(noCar).length;
      return (names.length ? '<span class="gw-vc-slot on"><i class="fas fa-car-side" aria-hidden="true"></i> ' + esc(names.join(' · ')) + '</span>' : '') +
        (miss ? '<span class="gw-vc-slot empty"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> 차 없음 ' + miss + '명</span>' : '');
    }
    if (cur === '__own') {
      var one = g.members.filter(function (m) { return !S.off[m]; });
      var own = one.length === 1 ? ownCars()[one[0]] : null;
      return '<span class="gw-vc-slot own"><i class="fas fa-car-side" aria-hidden="true"></i> ' + esc(own ? own.label : '각자 고정 차') + '</span>';
    }
    if (cur) {     // 회사 차든 [본인 차량으로 이동] 체크든 같은 모양(11-24 부장님)
      var c = carOf(cur);
      return '<span class="gw-vc-slot on"><i class="fas fa-car-side" aria-hidden="true"></i> ' + esc(c.label) + x(c.label) + '</span>';
    }
    return '<span class="gw-vc-slot empty"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> 차 없음</span>';
  }
  function tagsOf(code) {
    return (S.half[code] ? '<small class="gw-va-leave">' + esc(S.half[code]) + '</small>' : '') +
      (S.online[code] ? '<small class="gw-va-leave">' + esc(S.online[code] + (S.visits[code] ? '·방문 후' : '')) + '</small>' : '');
  }
  function pchip(code, detail) {
    if (S.off[code]) {      // 종일 휴무 — 차에 못 넣음
      return '<span class="gw-va-chip gw-vp-chip gw-vc-p off"><span class="gw-va-cname">' + esc(name(code)) + '</span>' +
        '<small class="gw-va-leave gw-vp-offtag">' + esc(S.off[code]) + '</small></span>';
    }
    var w = detail && noCar(code);
    var ctl;
    if (!detail) {
      ctl = '';
    } else if (!S.canEdit) {
      ctl = '<small class="gw-vc-ptext">' + esc(whereText(code) || '차 없음') + '</small>';
    } else if (splitPerson(code)) {
      ctl = '<span class="gw-vc-half"><small>오전</small>' + personSel(code, 'am') + '</span>' +
        '<span class="gw-vc-half"><small>오후</small>' + personSel(code, 'pm') + '</span>' +
        '<button type="button" class="gw-va-tool gw-vc-join" data-joinp="' + esc(code) + '">하루 같게</button>';
    } else {
      ctl = personSel(code, 'all');
    }
    return '<span class="gw-va-chip gw-vp-chip gw-vc-p' + (w ? ' w' : '') + (detail && splitPerson(code) && S.canEdit ? ' split' : '') +
      '"' + (detail ? ' data-cperson="' + esc(code) + '"' : '') + '>' + (w ? '<i class="fas fa-triangle-exclamation" aria-hidden="true"></i>' : '') +
      '<span class="gw-va-cname">' + esc(name(code)) + '</span>' + tagsOf(code) + ctl +
      (w ? '<span class="visually-hidden"> — 차 없음</span>' : '') + '</span>';
  }
  // 코스가 가는 지역 — 현장을 넣은 순서일 뿐 이동 순서가 아니라 화살표 없이, 같은 지역은 한 번만 쓰고 곳 수를 붙임
  //   예) '서울 동대문구 2 · 서울 용산구 · 서울 마포구' (11-23 부장님)
  function siteRoute(sites) {
    var out = [], n = {};
    sites.forEach(function (s) {
      var p = s.place;
      if (!p) return;
      if (!n[p]) out.push(p);
      n[p] = (n[p] || 0) + 1;
    });
    return out.map(function (p) { return n[p] > 1 ? p + ' ' + n[p] : p; }).join(' · ');
  }
  function placeRoute(members) {   // 슬랙으로 확정한 날 — 사람들이 갈 곳(도로명 기준 시·구, 11-24)을 코스 줄과 같은 모양으로
    var seen = {}, sites = [];
    members.forEach(function (m) {
      (S.places[m] || []).forEach(function (p) {
        if (!seen[p.addr]) { seen[p.addr] = 1; sites.push({ place: p.label }); }   // 같은 현장(주소)은 한 번만
      });
    });
    return siteRoute(sites);
  }
  // 코스 현장 한 줄 — 모양은 방문 일정과 같게(번호 동그라미·지역 파란 꼬리표·담당 꼬리표), 글 구성은 차량 배정 그대로
  //   (첫 줄 꼬리표·전화·날짜·번호 / 둘째 줄 주소 / 셋째 줄 문의, 11-24 부장님)
  function whoMark(s) {           // 방문 일정 who() 와 같은 모양 — 본인 방문 필수는 진한 채움, 담당만은 테두리
    if (!s.contact) return s.must ? '<span class="gw-va-smark">본인 방문 필수</span>' : '';
    return '<span class="gw-va-smark ' + (s.must ? 'must' : 'who') + '"><i class="fas fa-user" aria-hidden="true"></i> ' +
      esc(s.contact) + (s.must ? ' · 본인 방문 필수' : '') + '</span>';
  }
  function courseSite(s, k) {
    if (s.missing) {
      return '<li class="gw-vp-row gw-vc-srow"><span class="gw-vp-no" aria-hidden="true">' + (k + 1) + '</span><span class="gw-vp-rmain">' +
        '<span class="gw-vp-vtop"><span class="gw-va-sno">' + esc(s.lead_no) + '</span></span>' +
        '<span class="gw-vp-vtext">그날 현장 목록에 없어요(날짜가 바뀌었을 수 있어요)</span></span></li>';
    }
    var tags = (s.place ? '<span class="gw-vp-place">' + esc(s.place) + '</span>' : '') +
      (s.platform ? '<span class="gw-va-smark plat">' + esc(s.platform) + '</span>' : '') + whoMark(s) +
      marksOf(s.marks).map(function (m) { return '<span class="gw-va-smark">' + esc(m) + '</span>'; }).join('');
    var phone = s.phone ? '<a class="gw-vc-sphone" href="tel:' + esc(s.phone.replace(/[^\d+]/g, '')) + '">' + esc(s.phone) + '</a>' : '';
    return '<li class="gw-vp-row gw-vc-srow"><span class="gw-vp-no" aria-hidden="true">' + (k + 1) + '</span><span class="gw-vp-rmain">' +
      '<span class="gw-vp-vtop">' + tags + phone + '<span class="gw-va-sno">' + esc((s.date ? s.date + ' · ' : '') + s.lead_no) + '</span></span>' +
      (s.addr ? '<span class="gw-vp-vaddr">' + esc(s.addr) + '</span>' : '') +
      (s.text ? '<span class="gw-vp-vtext">' + esc(s.text) + '</span>' : '') + '</span></li>';
  }
  function courseCard(g, inG) {
    var route = g.sites.length ? siteRoute(g.sites) : placeRoute(g.members);
    var count = g.sites.length ? '현장 ' + g.sites.length : (g.visits ? '방문 ' + g.visits : '');
    // 두 코스에 든 사람(오전 코스·오후 코스) — 차는 먼저 넣은 코스 것, 다르면 메모로
    var hint = g.members.filter(function (m) { return (inG[m] || []).length > 1 && !S.off[m]; }).map(function (m) {
      var other = inG[m].filter(function (l) { return l !== g.label; });
      return '<p class="gw-vc-hint">' + esc(name(m)) + ' 님은 ' + esc(other.join(', ')) + '에도 있어요.' +
        (S.canEdit ? ' 차가 다르면 메모에 적어 주세요.' : '') + '</p>';
    }).join('');
    var head = '<div class="gw-vp-chead"><strong>' + esc(g.label) + '</strong>' + courseSlot(g) +
      '<div class="gw-va-people">' + (g.members.map(function (m) { return pchip(m, false); }).join('') ||
        '<span class="gw-va-none">사람 없음</span>') + '</div>' + memoBtn(g) +
      (count ? '<span class="gw-vp-ccount">' + esc(count) + '</span>' : '') + ownCheck(g) + '</div>';
    var memo = g.memo ? '<p class="gw-vp-memotext"><i class="fas fa-note-sticky" aria-hidden="true"></i> ' + esc(g.memo) + '</p>' : '';
    var sites = g.sites.length
      ? '<details class="gw-vc-sites"><summary><i class="fas fa-chevron-right gw-va-cfold" aria-hidden="true"></i>' +
        '<span>' + esc(route || '현장') + '</span></summary><ol class="gw-vp-rows gw-vc-srows">' + g.sites.map(courseSite).join('') + '</ol></details>'
      : (route ? '<p class="gw-vc-route"><i class="fas fa-location-dot" aria-hidden="true"></i> ' + esc(route) + '</p>' : '');
    var pk = S.canEdit && S.pick ? carOf(S.pick) : null;     // 차를 체크했으면 코스 카드를 눌러 그 차를 넣는다
    return '<section class="gw-va-card gw-vp-course gw-vc-course' + (pk ? ' pickable' : '') +
      '" data-ccourse="' + esc(g.key) + '"' + (pk ? ' tabindex="0" role="button" aria-label="' + esc(g.label + '에 ' + pk.label + ' 넣기') + '"' : '') +
      '>' + head + ownPickRow(g) + carMemo(g) + hint + memo + sites + '</section>';
  }
  // [본인 차량으로 이동] 체크(11-24 부장님) — 코스 전체가 자기 차로 가면 체크만. 체크하면 차 칸이 '본인 차량'(안에선 자차 칸)
  function coursePeople(g) { return g.members.filter(function (m) { return !S.off[m]; }); }
  function ownCheck(g) {
    var cur = courseCar(g), oc = ownCar();
    if (!S.canEdit || !oc || cur === '__none') return '';
    var many = coursePeople(g).length > 1;
    var on = cur === oc.plate || (many && (!!S.ownPick[g.key] || coursePeople(g).some(onOwn)));
    return '<label class="gw-vc-owncheck"><input type="checkbox" class="form-check-input" data-ownc="' + esc(g.key) + '"' +
      (on ? ' checked' : '') + '> 본인 차량으로 이동</label>';
  }
  // 여러 명 코스 — 자차로 가는 사람만 고른다(11-24 부장님: 체크하면 코스 차가 통째로 본인 차량이 되던 것).
  //   고른 사람만 본인 차량, 나머지는 코스 차 그대로. 모두 고르면 코스 전체 본인 차량(YG+TH 처럼 한 사람 차로 같이 감)
  function ownPickRow(g) {
    var ps = coursePeople(g);
    if (ps.length < 2 || !ownCar()) return '';
    var mine = ps.filter(onOwn);
    if (!S.canEdit) {
      return mine.length && mine.length < ps.length
        ? '<p class="gw-vc-ownpick"><i class="fas fa-car-side" aria-hidden="true"></i> 본인 차량: ' + esc(mine.map(name).join(', ')) + '</p>' : '';
    }
    if (!S.ownPick[g.key] && !mine.length) return '';
    return '<div class="gw-vc-ownpick" role="group" aria-label="' + esc(g.label) + ' 본인 차량으로 가는 사람"><span>자차로 가는 사람</span>' +
      ps.map(function (m) {
        return '<label class="gw-vc-owncheck"><input type="checkbox" class="form-check-input" data-ownp="' + esc(m) + '" data-ckey="' +
          esc(g.key) + '"' + (onOwn(m) ? ' checked' : '') + '> ' + esc(name(m)) + '</label>';
      }).join('') + '</div>';
  }
  function courseBack(g, m) {      // 본인 차량에서 뺀 사람 — 코스 차로(반나절 코스면 그 반나절), 코스 차가 없으면 차 없음
    var rest = { key: g.key, label: g.label, members: g.members.filter(function (x) { return x !== m; }) };   // 그 사람 빼고 본 코스 차
    var hf = courseHalf(rest);
    if (hf) { setCar(m, hf.car.plate, hf.half); return; }
    var cur = courseCar(rest);
    setCar(m, cur && cur.indexOf('__') !== 0 && !isOwn(carOf(cur)) ? cur : '', 'all');
  }
  // 코스 메모 한 줄(11-24 부장님) — 차가 들어간 코스에 아무 메모나(그 차의 메모로 저장, 배정 DM 에 같이 감)
  function memoCar(g) {           // 메모를 붙일 차 — 차 한 대인 코스(회사 차·본인 차량, 반나절 쓰는 차 포함)
    var hf = courseHalf(g);
    if (hf) return hf.car;            // 나눠 쓰는 두 코스는 메모 한 줄을 같이 씀(차의 메모)
    var cur = courseCar(g);
    return cur && cur.indexOf('__') !== 0 ? carOf(cur) : null;
  }
  function memoText(g) { return (memoCar(g) || {}).note || ''; }
  function memoBtn(g) {           // 평소엔 [메모] 단추만(칸이 늘 보이면 복잡, 11-24) — 눌러야 한 줄 칸이 열린다
    var c = memoCar(g);
    if (!S.canEdit || !c || memoText(g) || S.memoOpen[g.key]) return '';
    return '<button type="button" class="gw-va-tool gw-vc-memobtn" data-memoopen="' + esc(g.key) + '" aria-label="' + esc(g.label) +
      ' 메모 쓰기"><i class="fas fa-pen" aria-hidden="true"></i> 메모</button>';
  }
  function carMemo(g) {
    var c = memoCar(g), text = memoText(g);
    if (!c) return '';
    if (!S.canEdit || !(text || S.memoOpen[g.key])) return text ? '<p class="gw-vc-cmemo"><i class="fas fa-note-sticky" aria-hidden="true"></i> ' + esc(text) + '</p>' : '';
    return '<div class="gw-vc-cmemo"><i class="fas fa-note-sticky" aria-hidden="true"></i>' +
      '<input type="text" class="form-control form-control-sm" maxlength="60" data-cmemo="' + esc(c.plate) + '" data-ckey="' + esc(g.key) +
      '" value="' + esc(text) + '" placeholder="메모 한 줄' +
      '" aria-label="' + esc(g.label) + ' 메모" autocomplete="off">' +
      // 메모 지우고 닫기(11-24 부장님 — [메모]를 눌렀다가 되돌릴 길이 없었음). 빈 칸에서 Esc 도 닫힘
      '<button type="button" class="gw-va-x gw-vc-memox" data-memoclose="' + esc(g.key) + '" aria-label="' + esc(g.label) +
      ' 메모 지우고 닫기" title="메모 지우고 닫기"><i class="fas fa-xmark" aria-hidden="true"></i></button></div>';
  }
  // 메모 닫기 — 글을 지우고 [메모] 단추로 되돌린다(차 없는 코스는 다시 '차 없음')
  function closeMemo(key) {
    var g = groupByKey(key), c = g && memoCar(g);
    delete S.memoOpen[key];
    if (c && memoText(g)) {
      c.note = '';
      setDirty(true);
    }
    focusAfter('[data-memoopen="' + cssq(key) + '"]');
    render();
  }
  function otherCard(inG) {        // 그 밖 — 코스에 없는데 차를 타는 사람(미팅·자재 심부름 등)
    var list = others(inG);
    var cand = S.staff.map(function (p) { return p.code; }).filter(function (c) {
      return !inG[c] && list.indexOf(c) === -1 && !S.off[c];
    });
    if (!list.length && !S.canEdit) return '';
    var add = S.canEdit && cand.length
      ? '<select class="form-select form-select-sm gw-vc-sel gw-vc-add" data-add="1" aria-label="그 밖에 사람 더하기">' +
        '<option value="">+ 사람 더하기</option>' + cand.map(function (c) {
          var why = S.work[c] ? ' (사내 업무)' : S.online[c] ? ' (' + S.online[c] + ')' : '';
          return '<option value="' + esc(c) + '">' + esc(name(c) + why) + '</option>';
        }).join('') + '</select>' : '';
    return '<section class="gw-va-card gw-vp-course gw-vc-course gw-vc-other" data-ccourse="other" aria-label="그 밖">' +
      '<div class="gw-vp-chead"><strong>그 밖</strong><div class="gw-va-people">' +
      (list.map(function (m) { return pchip(m, true); }).join('') || '<span class="gw-va-none">코스에 없는데 차를 탈 사람(미팅 등)</span>') + '</div>' + add + '</div></section>';
  }
  function leftCard() {            // 코스에 없는 현장(방문 일정에서 코스를 안 정한 현장)
    var left = S.courses ? S.courses.left : [];
    if (!left.length) return '';
    return '<section class="gw-va-card gw-vp-course gw-vc-course gw-vc-left"><details class="gw-vc-sites"><summary>' +
      '<i class="fas fa-chevron-right gw-va-cfold" aria-hidden="true"></i><strong>코스에 없는 현장 ' + left.length + '건</strong>' +
      '<span>' + esc(siteRoute(left)) + '</span></summary><ol class="gw-vp-rows gw-vc-srows">' + left.map(courseSite).join('') + '</ol></details></section>';
  }
  function nocarStrip(inG) {       // 차가 필요 없는 사람 — 사내 업무·온라인 당번(방문 없음)·휴무
    var codes = S.staff.map(function (p) { return p.code; });
    var work = codes.filter(function (c) { return S.work[c] && !S.visits[c] && !inG[c] && !whereIs(c); });
    var duty = codes.filter(function (c) { return S.online[c] && !S.visits[c] && !inG[c] && !whereIs(c) && !S.work[c]; });
    var off = codes.filter(function (c) { return S.off[c] && !inG[c]; });
    var bits = [];
    if (work.length) bits.push('<strong>사내 업무</strong> ' + esc(work.map(function (c) { return name(c) + '(' + S.work[c] + ')'; }).join(', ')));
    if (duty.length) bits.push('<strong>온라인 당번</strong> ' + esc(duty.map(name).join(', ')));
    if (off.length) bits.push('<strong>휴무</strong> ' + esc(off.map(function (c) { return name(c) + '(' + S.off[c] + ')'; }).join(', ')));
    return bits.length ? '<p class="gw-vc-strip"><span>차 필요 없음</span> ' + bits.join('<span class="gw-va-dot">·</span>') + '</p>' : '';
  }

  // ── 왼쪽 차 목록 ──
  function carRow(c, inG) {
    var rs = ridersOf(c);
    var labels = uniq(rs.map(function (m) { return (inG[m] || ['그 밖'])[0]; })).filter(function (l) { return l !== '혼자'; });
    var who;
    if (isSplit(c)) who = '오전 ' + (c.people.map(name).join('·') || '없음') + ' / 오후 ' + (c.pm.map(name).join('·') || '없음');
    else who = rs.map(name).join('·');
    var line = rs.length
      ? '<span class="gw-vc-cwho"><span class="gw-vc-clabel">' + esc(labels.join('·')) + '</span> ' + esc(who) + '</span>'
      : '';      // 빈 차는 '코스 없음' 묶음에 있으니 줄 글 없음('평소 ○○'·'빈 차' 글은 뺐다, 11-24 부장님)
    // 차 줄의 메모 칸은 없앴다(11-24 부장님 — 예외는 코스 메모 한 줄로)
    var on = S.canEdit && S.pick === c.plate;
    // 코스에 넣은 차는 흐리게, 체크·끌기 막음(11-24 부장님 — 차는 코스 두 곳에 들어갈 일이 없음). 바꾸려면 코스의 차 꼬리표 ×로 뺀 뒤 다시
    //   단 오전·오후 나눠 쓰기(11-24)는 됨 — 체크하면 위 줄에서 반나절을 고르고 다른 코스에 넣는다
    var free = splitFree(c);
    var check = !S.canEdit ? '' : rs.length && !free
      ? '<input type="checkbox" class="form-check-input gw-vp-check" disabled aria-label="' + esc(c.label) + ' — 오전·오후 모두 코스에 넣음">'
      : '<input type="checkbox" class="form-check-input gw-vp-check" data-carcheck="' + esc(c.plate) + '"' +
        (on ? ' checked' : '') + ' aria-label="' + esc(c.label) + (rs.length ? ' 오전·오후 나눠 쓰기' : ' 고르기') + '">';
    var tip = !S.canEdit || !rs.length ? '' : free
      ? '이미 코스에 넣은 차예요. 다른 코스와 오전·오후로 나눠 쓰려면 체크하세요. 바꾸려면 코스의 차 꼬리표 ×로 빼 주세요.'
      : '오전·오후 모두 코스에 넣은 차예요. 바꾸려면 코스의 차 꼬리표 ×로 빼 주세요.';
    return '<div class="gw-vc-car' + (rs.length ? ' used' : '') + (isOwn(c) ? ' own' : '') + (on ? ' on' : '') + '" data-car="' + esc(c.plate) + '"' +
      (tip ? ' title="' + esc(tip) + '"' : '') + '>' +
      '<div class="gw-vc-carhead">' + check +
      '<strong>' + (isOwn(c) ? '<i class="fas fa-user" aria-hidden="true"></i> ' : '') + esc(c.label) + '</strong>' +
      (isOwn(c) ? '' : '<small>' + esc(c.plate) + '</small>') + '</div>' + line + '</div>';
  }
  // 차 목록 위 한 줄 — 안내 글 ↔ '레이 5358 → [코스에 넣기]' 가 같은 자리(방문 일정 사람 칸과 같게)
  function carBar(gs) {
    if (!S.canEdit) return '누가 어느 차를 타는지 참고만 해요.';
    var c = S.pick ? carOf(S.pick) : null;
    if (!c) return '<span class="gw-vp-how">차를 체크한 뒤 코스를 ' + (TOUCH ? '고르세요.' : '고르거나 드래그하세요.') + '</span>'   // 11-24 부장님 문구;
    if (ridersOf(c).length) {         // 나눠 쓰기 — '쏘렌토 8021 [오후 ▾] → [코스에 넣기 ▾]'
      var free = splitFree(c), h = S.pickHalf || halfGuess(c);
      var using = gs.filter(function (g) { return g.members.some(function (m) { return ridersOf(c).indexOf(m) !== -1; }); });
      var halfSel = free === 'both'
        ? '<select class="form-select form-select-sm gw-vc-halfsel" data-carhalf="1" aria-label="' + esc(c.label) + ' 새 코스가 쓸 때">' +
          ['am', 'pm'].map(function (k) { return '<option value="' + k + '"' + (k === h ? ' selected' : '') + '>' + halfName(k) + '</option>'; }).join('') + '</select>'
        : '<span class="gw-vc-half">' + halfName(free) + '</span>';
      var byDuty = free === 'both' && !S.pickHalf && ridersOf(c).some(function (m) { return S.duty[m]; });
      return '<strong>' + esc(c.label) + '</strong> ' + halfSel + ' → <select class="form-select form-select-sm gw-vp-liftsel" data-carto="1" aria-label="' +
        esc(c.label) + ' 나눠 쓸 코스"><option value="">코스에 넣기</option>' + gs.filter(function (g) { return using.indexOf(g) === -1; }).map(function (g) {
          return '<option value="' + esc(g.key) + '">' + esc(g.label) + '</option>';
        }).join('') + '</select><button type="button" class="btn btn-sm btn-link" data-carunchk="1">풀기</button>' +
        '<span class="gw-vc-splitnote">' + esc((using.map(function (g) { return g.label; }).join('·') || '지금 코스') + ' → ' +
        halfName(h === 'am' ? 'pm' : 'am') + ', 새 코스 → ' + halfName(h) + (byDuty ? ' (온라인 당번 기준)' : '')) + '</span>';
    }
    return '<strong>' + esc(c.label) + '</strong> → <select class="form-select form-select-sm gw-vp-liftsel" data-carto="1" aria-label="' +
      esc(c.label) + ' 넣을 코스"><option value="">코스에 넣기</option>' + gs.map(function (g) {
        var cur = courseCar(g);
        var now = cur === '' ? ' (차 없음)' : cur === c.plate ? ' (지금 이 차)' : '';
        return '<option value="' + esc(g.key) + '">' + esc(g.label + now) + '</option>';
      }).join('') + '</select><button type="button" class="btn btn-sm btn-link" data-carunchk="1">풀기</button>';
  }

  // ── 그리기 ──
  function board() {               // 서버 재시작 전 옛 화면이면 판을 새 모양으로 바꿔 끼운다
    var b = $('vaBoard');
    if (b && !$('vcCourses')) {
      b.className = 'gw-va-board gw-vc-board';
      b.innerHTML = '<aside class="gw-va-staff gw-vc-cars" aria-label="법인차량 목록"><div class="gw-va-staff-head"><strong>법인차량</strong>' +
        '<span id="vcCarCount"></span></div><p class="gw-va-staff-how" id="vcHow"></p><div id="vcCarList"></div></aside>' +
        '<section class="gw-vc-main"><div class="gw-va-schedbox" id="vaSchedBox" hidden><div class="gw-va-pre" id="vaPre"></div>' +
        '<details class="gw-va-sched" id="vaSched"></details></div>' +
        '<div id="vcStrip"></div><div class="gw-vp-courses" id="vcCourses"><p class="gw-empty">불러오는 중…</p></div></section>';
    }
    if ($('vaBar') && !$('vcWarn')) {
      var w = document.createElement('span');
      w.id = 'vcWarn';
      w.className = 'gw-vc-warn';
      $('vaBar').insertBefore(w, $('vaBar').firstChild);
    }
  }
  function render() {
    board();
    var inG = groupOf();
    var gs = groups();
    $('vcCourses').innerHTML = gs.map(function (g) { return courseCard(g, inG); }).join('') + otherCard(inG) + leftCard() ||
      '<p class="gw-va-none gw-va-noused">' + (visitsMissing() ? '방문 일정을 불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.' : '이날 방문 일정이 없어요.') + '</p>';
    $('vcStrip').innerHTML = nocarStrip(inG);
    var sales = S.cars.filter(function (c) { return !isOwn(c) && !isCrew(c); });     // 영업 차만 센다(시공팀 차 빼고)
    var used = sales.filter(function (c) { return ridersOf(c).length; }).length;
    var crewUsed = S.cars.filter(function (c) { return isCrew(c) && ridersOf(c).length; }).length;
    var hd = document.querySelector('.gw-vc-cars .gw-va-staff-head strong');   // 재시작 전 옛 템플릿('차')도 같은 이름으로
    if (hd && hd.textContent !== '법인차량') hd.textContent = '법인차량';
    $('vcCarCount').textContent = '배정 ' + used + ' / ' + sales.length + '대' + (crewUsed ? ' + 시공팀 차량 ' + crewUsed : '');   // 11-24 부장님 이름
    if (S.pick && !carOf(S.pick)) S.pick = '';
    $('vcHow').className = 'gw-va-staff-how gw-vc-bar' + (S.pick ? ' on' : '') +
      (S.pick && ridersOf(carOf(S.pick)).length ? ' split' : '');      // 나눠 쓰기 — 세 줄로(차 / 반나절·코스 / 안내)
    $('vcHow').innerHTML = carBar(gs);
    // 방문 일정 사람 칸처럼 '코스 없음 / 코스 있음'으로 나눠 보인다(11-24 부장님 — 넣은 차가 아래로 빠져 남은 차만 위에)
    var crewRows = S.cars.filter(isCrew);
    var sales = carsInOrder().filter(function (c) { return !isCrew(c); });
    var gi = {};
    gs.forEach(function (g, i) { g.members.forEach(function (m) { if (!(m in gi)) gi[m] = i; }); });
    var firstCourse = function (c) {      // 코스 있음 묶음은 코스 번호 순(코스 1 차가 맨 위)
      return Math.min.apply(null, ridersOf(c).map(function (m) { return m in gi ? gi[m] : 999; }).concat([999]));
    };
    var grp = function (title, list) {
      return list.length ? '<div class="gw-va-group gw-vc-cgroup"><div class="gw-va-gtitle">' + esc(title) + ' <span>' + list.length + '</span></div>' +
        '<div class="gw-vc-clist">' + list.map(function (c) { return carRow(c, inG); }).join('') + '</div></div>' : '';
    };
    $('vcCarList').innerHTML = grp('코스 없음', sales.filter(function (c) { return !ridersOf(c).length; })) +
      grp('코스 있음', sales.filter(function (c) { return ridersOf(c).length; }).sort(function (a, b) { return firstCourse(a) - firstCourse(b); })) +
      (crewRows.length ? '<div class="gw-vc-sub">시공팀 차량 <small>현장 지원·물품 운반 때만</small></div>' +
        '<div class="gw-vc-clist">' + crewRows.map(function (c) { return carRow(c, inG); }).join('') + '</div>' : '');
    // 아래 바 경고 — 코스 단위로 짧게('차 없는 코스: 코스 1, 코스 3'), 코스 밖 사람만 이름으로
    var missG = gs.filter(function (g) { return g.members.some(noCar); }).map(function (g) {
      return g.label === '혼자' ? g.members.map(name).join('·') : g.label;     // 코스 이름이 '혼자'인 팀(확정 전)은 사람 이름으로 — '혼자, 혼자, 혼자'로 나오던 것
    });
    var missP = Object.keys(S.visits).filter(function (c) { return S.names[c] && !inG[c] && noCar(c); }).map(name);
    var warn = $('vcWarn');
    if (warn) {
      var wt = (missG.length ? '차 없는 코스: ' + missG.join(', ') : '') + (missP.length ? (missG.length ? ' · ' : '') + '차 없음: ' + missP.join(', ') : '');
      warn.innerHTML = wt && S.canEdit ? '<i class="fas fa-triangle-exclamation" aria-hidden="true"></i> ' + esc(wt) : '';
      warn.hidden = !(wt && S.canEdit);
    }
    $('vaBar').hidden = !S.canEdit;
    $('vaBoard').classList.toggle('editing', S.canEdit);
    $('vaConfirm').textContent = S.dmOff ? (S.confirmed ? '저장(DM 꺼짐)' : '확정(DM 꺼짐)')
      : (S.confirmed ? '저장하고 바뀐 사람에게 알리기' : '확정하고 알리기');
    $('vaSave').textContent = S.confirmed ? '알리지 않고 저장' : '저장만';
    $('vaSave').hidden = false;
    setDirty(S.dirty);
    statusLine();
    restoreFocus();
  }
  // 슬랙으로 확정한 날(코스 없음) — 위에 그날 방문 일정(캔버스1 묶음)을 접어서 둔다
  // 확정 전 날만 보이는 칸 — 위에 '확정 전' 안내 한 줄, 방문 목록은 기본 접힘(11-24 부장님: 목록이 눈에 안 들어옴.
  //   차는 방문 일정 확정 뒤에 넣는 흐름이라 할 일은 목록 읽기보다 코스 짜기). 직접 펼친 것만 기억(새 키 — 예전 '펼침' 기억은 버림)
  var SCHED_KEY = 'gw.va.schedOpen2';
  function schedOpenPref() {
    try { return localStorage.getItem(SCHED_KEY) === '1'; } catch (e) { return false; }
  }
  function schedItem(it) {
    var marks = (it.platform ? '<span class="gw-va-smark plat">' + esc(it.platform) + '</span>' : '') +
      marksOf(it.marks).map(function (m) { return '<span class="gw-va-smark">' + esc(m) + '</span>'; }).join('');
    var who = it.who.length
      ? '<span class="gw-va-swho">' + esc(it.who.join(', ')) + '</span>'
      : '<span class="gw-va-swho none">배분 전</span>';
    var phone = it.phone
      ? '<a href="tel:' + esc(it.phone.replace(/[^\d+]/g, '')) + '">' + esc(it.phone) + '</a>' : '';
    return '<li class="gw-va-sitem">' +
      '<div class="gw-va-sline">' + marks + '<strong>' + esc(it.date) + '</strong>' + (phone ? '<span>' + phone + '</span>' : '') +
        (it.lead_no ? '<span class="gw-va-sno">' + esc(it.lead_no) + '</span>' : '') + who + '</div>' +
      (it.addr ? '<div class="gw-va-saddr">' + esc(it.addr) + '</div>' : '') +
      (it.text ? '<div class="gw-va-stext">' + esc(it.text) + '</div>' : '') + '</li>';
  }
  function schedPanel(j) {
    var el = $('vaSched'), box = $('vaSchedBox'), pre = $('vaPre');
    if (!el) return;
    if (!box) {          // 재시작 전 옛 템플릿(details 만) — 그 자리를 감싼다
      box = document.createElement('div');
      box.className = 'gw-va-schedbox'; box.id = 'vaSchedBox'; box.hidden = true;
      el.parentNode.insertBefore(box, el);
      pre = document.createElement('div');
      pre.className = 'gw-va-pre'; pre.id = 'vaPre'; pre.hidden = true;
      box.appendChild(pre); box.appendChild(el);
      el.hidden = false;
    }
    if (S.courses || !('schedule' in j)) { box.hidden = true; return; }   // 코스로 볼 때는 현장이 코스 카드 안에
    var n = S.schedule.reduce(function (a, g) { return a + g.items.length; }, 0);
    var todo = S.schedule.reduce(function (a, g) { return a + g.items.filter(function (it) { return !it.who.length; }).length; }, 0);
    var sum = '방문 ' + n + '건' + (todo ? ' · 배분 전 ' + todo + '건' : '');
    var body = S.schedule.map(function (g) {
      return '<div class="gw-va-sgroup"><h3>' + esc(g.title) + ' <span>' + g.items.length + '</span></h3>' +
        '<ul>' + g.items.map(schedItem).join('') + '</ul></div>';
    }).join('') || '<p class="gw-va-sempty">이날 방문 일정이 없어요.</p>';
    // 방문 일정 탭이 보이는 사람(두 화면 다 여는 사람)에게만 [방문 일정으로] — 같은 날짜로
    var canPlan = !!document.querySelector('[data-vtab][href^="/visits/plan"]');
    pre.innerHTML = n ? '<i class="fas fa-circle-info" aria-hidden="true"></i><span>방문 일정이 아직 확정 전이에요. ' +
      '코스를 먼저 짜면 현장이 코스 카드에 들어와요.</span>' +
      (canPlan ? '<a class="btn btn-sm btn-outline-secondary gw-va-pretogo" href="/visits/plan?day=' + encodeURIComponent(S.day) +
        '">방문 일정으로 <i class="fas fa-arrow-right" aria-hidden="true"></i></a>' : '') : '';
    pre.hidden = !n;
    var wasOpen = el.dataset.ready ? el.open : schedOpenPref();
    el.innerHTML = '<summary><i class="fas fa-calendar-day" aria-hidden="true"></i><span class="gw-va-stitle">' +
      esc(j.day_label || '') + ' 방문 일정</span><span class="gw-va-ssum">' + esc(sum) + '</span>' +
      '<i class="fas fa-chevron-down gw-va-fold" aria-hidden="true"></i></summary><div class="gw-va-sbody">' + body + '</div>';
    el.open = wasOpen;
    box.hidden = false;
    el.dataset.ready = '1';
  }
  function courseMembers() {
    var seen = {};
    S.teams.forEach(function (t) {
      t.members.forEach(function (m) { seen[m] = 1; });
      Object.keys(t.also || {}).forEach(function (m) { seen[m] = 1; });
    });
    return seen;
  }
  // 차량 배정을 저장한 뒤 방문 일정이 다시 확정됐고, 지금 배정과 코스가 어긋나면(차 없는 코스 사람·코스에서 빠진 사람) 경고 글
  function staleText() {
    var p = S.plan;
    if (!S.canEdit || !p || !p.confirmed_at || !p.cars_saved_at || !(p.confirmed_at > p.cars_saved_at)) return '';
    var inCourse = courseMembers();
    var none = staffNoCar();
    var out = [];
    S.cars.forEach(function (c) {
      if (isOwn(c)) return;
      ridersOf(c).forEach(function (m) { if (!inCourse[m] && !S.visits[m] && out.indexOf(name(m)) === -1) out.push(name(m)); });
    });
    if (!none.length && !out.length) return '';
    return '차량 배정 뒤에 방문 일정이 다시 확정됐어요.' + (none.length ? ' 차 없음: ' + none.join(', ') + '.' : '') +
      (out.length ? ' 코스에서 빠짐: ' + out.join(', ') + '.' : '');
  }
  function status(j) {
    S.carWhen = j.confirmed_at ? j.confirmed_at.slice(5, 16).replace('-', '/') + (j.confirmed_name ? ' · ' + j.confirmed_name : '') : '';
    S.draftNote = '';      // 초안 없앰(11-23 부장님 — 타는 차·인원이 매일 달라서). 빈 날은 비어서 열린다
  }
  // 서버가 못 읽은 칸(11-24 조용한 실패 점검 18번) — 빈 목록을 '없음'으로 믿지 않게 화면 위에 한 줄. 배정은 막지 않는다
  var PARTIAL_NAMES = { leads: '리드 시트', visits: '방문 일정', teams: '방문 일정 코스', schedule: '방문 일정 목록',
                        off: '휴가', online: '온라인 당번', plan: '방문 일정 확정' };
  function visitsMissing() {
    return S.partial.some(function (k) { return k === 'leads' || k === 'visits' || k === 'teams' || k === 'plan'; });
  }
  function partialText() {
    var keys = S.partial.filter(function (k) { return PARTIAL_NAMES[k]; });
    if (!keys.length) return '';
    var what = [];
    if (keys.some(function (k) { return k !== 'off' && k !== 'online'; })) what.push('방문');
    if (keys.indexOf('off') !== -1) what.push('휴무');
    if (keys.indexOf('online') !== -1) what.push('당번');
    return keys.map(function (k) { return PARTIAL_NAMES[k]; }).join('·') + ' 정보를 못 불러왔어요 — ' + what.join('·') +
      ' 표시가 빠졌을 수 있어요.' + (keys.indexOf('off') !== -1 ? ' 쉬는 사람이 차에 들어가지 않았는지 확인해 주세요.' : '') +
      ' 잠시 뒤 새로 고쳐 보세요.';
  }
  function statusLine() {        // 날짜 · 확정 상태 · 방문 일정 · 초안 안내 · (경고) — 한 줄
    var el = $('vaStatus');
    if (!el) return;
    var pl = $('vaPlanLine');
    if (pl) { pl.hidden = true; pl.innerHTML = ''; }
    var head = '<strong>' + esc(S.dayLabel) + '</strong> ' + (S.confirmed
      ? '<span class="gw-badge gw-badge-ok">확정</span>' + (S.carWhen ? ' <span class="gw-va-when">' + esc(S.carWhen) + '</span>' : '')
      : '<span class="gw-badge gw-badge-wait">확정 전</span>');
    var bits = [];
    var p = S.plan;
    if (p && p.confirmed_at) {
      var t = '방문 일정 확정 ' + p.confirmed_at.slice(11, 16) + (p.confirmed_name ? ' · ' + p.confirmed_name : '');
      bits.push('<i class="fas fa-route" aria-hidden="true"></i> ' +
        (p.can_open ? '<a href="/visits/plan?day=' + encodeURIComponent(S.day) + '">' + esc(t) + '</a>' : esc(t)));
    } else if (S.courses && S.courses.source === 'sheet') {
      bits.push('<i class="fas fa-route" aria-hidden="true"></i> 슬랙으로 정한 날 — 시트 담당자대로 코스');
    }
    // 방문 일정 DM 이 차량 배정 확정을 기다리는 중(11-25 부장님) — 확정하면 한 번에, 급하면 일정만 먼저
    if (p && p.dm_mode === 'held') {
      bits.push('<span class="gw-badge gw-badge-wait">방문 일정 DM 대기</span> 확정하면 방문 일정과 차를 한 번에 보내요' +
        (S.canEdit ? ' <button type="button" class="btn btn-sm btn-outline-secondary gw-va-sendnow" data-act="sendnow">일정만 먼저 보내기</button>' : ''));
    } else if (p && p.dm_mode === 'plain' && p.dm_at) {
      bits.push('방문 일정 DM 은 먼저 보냄 ' + esc(p.dm_at.slice(11, 16)) + (S.confirmed ? '' : ' — 차는 확정할 때 따로 가요'));
    }
    if (S.draftNote && S.canEdit && !S.confirmed) bits.push(esc(S.draftNote));
    var warn = staleText();
    var part = partialText();
    el.innerHTML = head + (bits.length ? '<span class="gw-va-sinfo">' + bits.join('<span class="gw-va-dot">·</span>') + '</span>' : '') +
      (part ? '<span class="gw-va-stale gw-va-partial" role="alert"><i class="fas fa-triangle-exclamation" aria-hidden="true"></i> ' + esc(part) + '</span>' : '') +
      (warn ? '<span class="gw-va-stale" role="status">' + esc(warn) + '</span>' : '');
  }

  // ── 불러오기·저장 ──
  function load(day) {
    msg('');
    var my = ++S.seq;     // ◀ ▶ 를 빨리 누르면 늦게 온 앞 날짜 응답이 덮지 않게
    return api('/api/vehicles/assign' + (day ? '?day=' + encodeURIComponent(day) : '')).then(function (j) {
      if (my !== S.seq) return;
      board();
      S.day = j.day; S.cars = j.cars; S.staff = j.staff; S.canEdit = !!j.can_edit; S.dmOff = !!j.dm_off;
      S.base = j.base == null ? null : j.base;          // 재시작 전 옛 서버면 없음 → 예전처럼 비교 없이 저장
      S.partial = Array.isArray(j.partial) ? j.partial : [];
      S.cars.forEach(function (c) {
        c.pm = Array.isArray(c.pm) ? c.pm : null;
        if (isOwn(c)) c.label = '본인 차량';     // 화면에서 부르는 이름(11-24 부장님 — 꼬리표는 '본인 차량'만)
      });
      S.confirmed = !!j.confirmed_at;
      S.names = {}; S.visits = {}; S.places = {}; S.off = {}; S.half = {}; S.online = {}; S.work = {}; S.duty = {}; S.pickHalf = '';
      S.splitOpen = {}; S.splitCourse = {}; S.memoOpen = {}; S.ownPick = {}; S.extra = []; S.pick = '';
      j.staff.forEach(function (p) { S.names[p.code] = p.name; });
      // 확정 전 토요 휴무는 바뀔 수 있어 반차처럼 꼬리표만(방문 일정과 같게 — 코스에 넣을 수 있음, 11-23)
      (j.off || []).forEach(function (o) { if (o.soft) S.half[o.code] = o.why; else S.off[o.code] = o.why; });
      (j.half || []).forEach(function (o) { S.half[o.code] = o.why; });
      (j.online || []).forEach(function (o) {
        S.online[o.code] = '온라인 당번' + (o.shift ? '(' + o.shift + ')' : '');
        if (o.shift) S.duty[o.code] = o.shift;
      });
      (j.work || []).forEach(function (o) { S.work[o.code] = o.task; });
      S.schedule = j.schedule || [];
      S.courses = j.courses || null;
      S.plan = j.plan || null;
      S.dayLabel = j.day_label || '';
      (j.visitors || []).forEach(function (v) { S.visits[v.code] = v.visits; S.places[v.code] = v.places || []; });
      S.teams = j.teams || [];
      schedPanel(j);
      S.dirty = false;
      $('vaDay').value = j.day;
      try { history.replaceState(null, '', '/vehicles/assign?day=' + j.day); } catch (e) { /* 무시 */ }
      status(j);
      var own = S.canEdit ? ownToCar() : [];
      var moved = S.canEdit ? followCourses() : [];
      render();
      var auto = [];
      if (own.length) auto.push('본인 차로 넣었어요 — ' + own.join(', ') + '.');
      if (moved.length) auto.push('방문 일정 코스대로 차에 넣었어요 — ' + moved.join(', ') + '.');
      if (auto.length) msg(auto.join(' ') + ' 저장해 주세요.', 'info', true);
      S.placeTries = 0;
      placesLater(j.resolving);
    }).catch(function (e) { if (my === S.seq) msg(e.message, 'err'); });
  }
  // 갈 곳의 동을 서버가 뒤에서 찾는 중이면 잠시 뒤 갈 곳만 다시 받는다(배정 고친 것은 그대로)
  function placesLater(resolving) {
    clearTimeout(S.placeTimer);
    if (!resolving || S.placeTries >= 3) return;
    var my = S.seq, day = S.day;
    S.placeTimer = setTimeout(function () {
      S.placeTries++;
      api('/api/vehicles/assign?day=' + encodeURIComponent(day)).then(function (j) {
        if (my !== S.seq || j.day !== S.day) return;
        S.places = {};
        (j.visitors || []).forEach(function (v) { S.places[v.code] = v.places || []; });
        var act = document.activeElement;
        // 고르는 중이거나 메모를 쓰는 중이면 다시 그리지 않는다 — 다음에 그릴 때 들어감
        var inBoard = act && act.closest && act.closest('.gw-va-board');
        if (!D && !(inBoard && act.matches('input, select'))) render();
        placesLater(j.resolving);
      }).catch(function () { /* 다음에 불러올 때 */ });
    }, 15000);
  }
  function rowsForSave() {
    return S.cars.map(function (c) {   // 코스 메모(11-24) — 자차 칸은 탄 사람이 없으면 메모도 비움
      return { plate: c.plate, people: c.people, pm: c.pm || null, note: isOwn(c) && !ridersOf(c).length ? '' : (c.note || '') };
    });
  }
  // notify: 슬랙으로 알릴지. 확정 전 [저장만]·확정 뒤 [알리지 않고 저장]은 false
  function save(notify) {
    var label = $('vaStatus').querySelector('strong');
    var dayText = label ? label.textContent : S.day;
    var miss = staffNoCar();
    var missText = miss.length ? '\n\n아직 차 없는 사람: ' + miss.join(', ') : '';
    if (notify && S.dmOff) {          // 차량 DM 꺼 둠(11-23 부장님) — 확정 기록만 남고 슬랙은 안 나간다
      if (!window.confirm(dayText + ' 배정을 ' + (S.confirmed ? '저장할까요?' : '확정할까요?') + missText +
          '\n\n지금은 차량 DM 을 꺼 두어서 슬랙으로는 알리지 않아요.')) return;
    } else if (notify && !S.confirmed && S.plan && S.plan.dm_mode === 'held') {
      if (!window.confirm(dayText + ' 배정을 확정하고, 방문 일정과 탈 차를 사람마다 한 통으로 슬랙에 보낼까요?' +
          '\n방문 없이 차만 타는 사람에게는 차량 안내만 가요.' + missText + '\n\n보낸 알림은 되돌릴 수 없어요.')) return;
    } else if (notify && !S.confirmed) {
      if (!window.confirm(dayText + ' 배정을 확정하고, 차에 넣은 ' + riders() + '명에게 슬랙으로 알릴까요?' + missText +
          '\n\n보낸 알림은 되돌릴 수 없어요.')) return;
    } else if (notify) {
      if (!window.confirm('고친 내용을 저장하고, 차나 같이 타는 사람이 바뀐 사람에게만 슬랙으로 알릴까요?')) return;
    }
    var btns = [$('vaSave'), $('vaConfirm')];
    btns.forEach(function (b) { b.disabled = true; });
    msg(notify ? '저장하고 알리는 중…' : '저장하는 중…', 'info', true);
    post('/api/vehicles/assign', { day: S.day, rows: rowsForSave(), confirm: !!notify && !S.confirmed, notify: !!notify,
                                   base: S.base })
      .then(function (j) {
        S.dirty = false;
        if (j.base != null) S.base = j.base;     // 같은 사람이 연달아 저장해도 막히지 않게(다시 불러오기가 실패해도)
        var t = j.changed.length ? '저장했어요.' : '바뀐 내용이 없어요.';
        var missed = j.failed || [];
        var pd = j.plan_dm;      // 방문 일정 + 차 한 통(11-25)
        var pdBad = pd && (!pd.ok || pd.dm_failed || pd.bot_missing);
        if (pd && pd.ok && pd.dm) t += ' 방문 일정과 차를 ' + pd.dm + '명에게 한 번에 보냈어요.';
        if (pd && pd.ok && pd.dm_failed) t += ' 방문 일정 DM 을 ' + pd.dm_failed + '명이 받지 못했어요. 직접 알려 주세요.';
        if (pd && pd.bot_missing) t += ' 방문봇 연결이 없어 방문 일정 DM 을 못 보냈어요. 직접 알려 주세요.';
        if (pd && !pd.ok) t += ' 방문 일정 DM 은 보내지 못했어요(' + (pd.message || '') + ') — 차량 안내만 보냈어요. 위 [일정만 먼저 보내기]로 다시 보내 주세요.';
        if (j.notified) t += ' ' + (pd && pd.ok ? '차량 안내만 ' : '') + j.notified + '명에게 슬랙으로 알렸어요.';
        else if (notify && j.dm_off) t += ' 차량 DM 은 꺼 두어서 슬랙으로는 알리지 않았어요.';
        else if (notify && j.changed.length && !missed.length && !(pd && pd.dm)) t += ' 슬랙으로 알릴 사람이 없었어요.';
        // 못 받은 사람은 이름으로(11-24) — 이미 확정이라 다시 눌러도 안 가니 직접 알려야 한다
        if (missed.length) t += ' ' + missed.join(', ') + ' 님은 슬랙 알림을 받지 못했어요. 직접 알려 주세요.';
        return load(S.day).then(function () {
          msg(t, missed.length || pdBad ? 'err' : 'ok', true);
          if ((j.retry || []).length) askResend(j.retry);   // 보내다 실패한 사람은 다시 보내 볼 수 있게(11-24 추가분)
        });
      }).catch(function (e) {
        if (e.status === 409) { staleReload(e.message); return; }
        msg(e.message, 'err', true);
      })
      .then(function () { btns.forEach(function (b) { b.disabled = false; }); setDirty(S.dirty); });
  }
  // 다른 분이 먼저 고친 날(409, 11-24 조용한 실패 점검 17번) — 새로 불러오면 내가 고친 것은 사라지니 묻고 불러온다.
  //   취소하면 화면은 그대로(고친 내용을 보고 옮겨 적을 수 있게). 저장은 새로 불러와야 된다.
  function staleReload(text) {
    msg(text, 'err', true);
    if (!window.confirm(text + '\n\n새로 불러오면 지금 고친 내용은 사라져요. 새로 불러올까요?')) return;
    S.dirty = false;
    load(S.day).then(function () { msg('새로 불러왔어요. 다른 분이 고친 내용을 보고 다시 해 주세요.', 'info', true); });
  }
  // 슬랙 계정은 있는데 보내다 실패한 사람(retry) — 한 번 묻고 그 사람에게만 다시 보낸다
  // (슬랙 계정이 없는 사람은 다시 보내도 안 가므로 빼고, 그 사람들은 위 안내대로 직접)
  function askResend(codes) {
    if (!window.confirm('슬랙을 못 받은 ' + codes.length + '명에게 차량 배정을 지금 다시 보내 볼까요?\n\n취소하면 직접 알려 주세요.')) return;
    msg('다시 보내는 중…', 'info', true);
    post('/api/vehicles/assign/resend', { day: S.day, codes: codes })
      .then(function (j) {
        var again = j.failed || [];
        if (again.length) msg(again.join(', ') + ' 님에게는 또 못 보냈어요. 직접 알려 주세요.', 'err', true);
        else msg(j.notified + '명에게 다시 보냈어요.', 'ok', true);
      }).catch(function (e) { msg(e.message, 'err', true); });
  }

  // ── 고르기 ──
  function groupByKey(key) { return groups().filter(function (g) { return g.key === key; })[0]; }
  // 체크한 차를 코스에 넣고 체크를 푼다
  function putPicked(key) {
    var c = carOf(S.pick), g = groupByKey(key);
    if (!c || !g) return;
    if (ridersOf(c).length) {                 // 나눠 쓰기(11-24)
      if (g.members.some(function (m) { return ridersOf(c).indexOf(m) !== -1; })) { msg('이미 ' + g.label + '에 넣은 차예요.', 'info'); return; }
      var h = S.pickHalf || halfGuess(c);
      putSplit(key, c.plate, h);
      S.pick = ''; S.pickHalf = '';
      msg(g.label + '에 넣었어요: ' + c.label + ' ' + halfName(h), 'ok');
      focusAfter('[data-ccourse="' + cssq(key) + '"]');
      render();
      return;
    }
    setCourse(key, c.plate, true);
    delete S.splitCourse[key];
    S.pick = '';
    msg(g.label + '에 ' + c.label + ' 넣었어요.', 'ok');
    focusAfter('[data-carcheck="' + cssq(c.plate) + '"]');
    render();
  }
  // 확정한 방문 일정 기준(11-24 부장님) — 코스에 차가 이미 한 대 들어가 있는데 차 없는 사람이 있으면(배정 뒤 코스에 더해진 사람 등)
  //   그 코스 차에 넣는다. 차를 고르는 게 아니라 코스 묶음을 따르는 것. 저장해야 반영. 돌려주는 값: ['김민욱 님 → 쏘렌토 4041', …]
  function followCourses() {
    if (!S.courses) return [];
    var inG = groupOf(), out = [];
    groups().forEach(function (g) {
      var ms = g.members.filter(function (m) { return !S.off[m] && (inG[m] || []).length < 2; });
      var placed = ms.filter(function (m) { return whereIs(m); });
      var hf = placed.length ? courseHalf({ members: placed }) : null;
      if (hf) {                                                   // 나눠 쓰는 코스 — 새로 든 사람도 같은 반나절
        ms.filter(noCar).forEach(function (m) {
          setCar(m, hf.car.plate, hf.half);
          out.push(name(m) + ' 님 → ' + g.label + ' ' + hf.car.label + ' ' + halfName(hf.half));
        });
        return;
      }
      var cars = uniq(ms.map(whereIs).filter(Boolean));
      // 차가 없거나 여러 대·'본인 차량'·개인 차(스파크 — 주인만 자동으로 넣음, 11-25)인 코스는 그대로
      if (cars.length !== 1 || isOwn(cars[0]) || isPersonal(cars[0])) return;
      ms.filter(noCar).forEach(function (m) {
        setCar(m, cars[0].plate, 'all');
        out.push(name(m) + ' 님 → ' + g.label + ' ' + cars[0].label);
      });
    });
    return out;
  }
  function setCourse(key, plate, keepOwn) {   // 코스 사람 모두(휴무 빼고) 종일 같은 차 / '' = 모두 뺀다
    var g = groupByKey(key);                    // keepOwn = 몇 사람만 자차로 고른 코스면 그 사람은 그대로(11-24)
    if (!g) return;
    var prev = courseCar(g), inG = groupOf();
    var some = keepOwn && coursePeople(g).some(onOwn) && !coursePeople(g).every(onOwn);
    g.members.forEach(function (m) {
      if (S.off[m]) return;
      if (some && onOwn(m)) return;
      var at = whereIs(m);     // 두 코스에 든 사람 — 다른 코스 차를 이미 탔으면 그대로(다르면 메모로, 11-24)
      if ((inG[m] || []).length > 1 && at && at.plate !== prev) return;
      delete S.splitOpen[m];
      setCar(m, plate, 'all');
    });
    unsplitEmpty();
  }
  document.addEventListener('change', function (e) {
    var t = e.target;
    if (!t.dataset) return;
    if (t.dataset.cmemo) {                    // 코스 메모 — 다시 그리지 않는다(쓰는 중 칸이 흔들리지 않게)
      var mc = carOf(t.dataset.cmemo), v = t.value.trim();
      mc.note = v;
      setDirty(true);
      return;
    }
    if (t.dataset.ownc) {                     // [본인 차량으로 이동] — 한 명 코스: 켜면 본인 차량·끄면 차 없음 / 여러 명: 켜면 사람 고르기
      var oc = ownCar(), og = groupByKey(t.dataset.ownc);
      if (og && coursePeople(og).length > 1) {
        if (t.checked) S.ownPick[og.key] = 1;
        else {                                // 끄면 자차로 고른 사람 모두 코스 차로
          delete S.ownPick[og.key];
          coursePeople(og).filter(onOwn).forEach(function (m) { setCar(m, '', 'all'); courseBack(og, m); });
        }
      } else {
        setCourse(t.dataset.ownc, t.checked ? oc.plate : '');
      }
      if (!ridersOf(oc).length) oc.note = '';
      focusAfter('[data-ownc="' + cssq(t.dataset.ownc) + '"]');
      render();
      return;
    }
    if (t.dataset.ownp) {                     // 여러 명 코스에서 자차로 가는 사람 하나씩
      var pg = groupByKey(t.dataset.ckey), po = ownCar();
      if (pg && po) {
        if (t.checked) { setCar(t.dataset.ownp, po.plate, 'all'); unsplitEmpty(); }
        else { setCar(t.dataset.ownp, '', 'all'); courseBack(pg, t.dataset.ownp); }
        if (!ridersOf(po).length) po.note = '';
        S.ownPick[pg.key] = 1;
      }
      focusAfter('[data-ownp="' + cssq(t.dataset.ownp) + '"][data-ckey="' + cssq(t.dataset.ckey) + '"]');
      render();
      return;
    }
    if (t.dataset.pcar) {
      var code = t.dataset.pcar, half = t.dataset.half || 'all';
      if (t.value === '__split') {             // 오전·오후 따로 — 지금 차를 양쪽에 둔 채로 칸만 둘로
        S.splitOpen[code] = 1;
        focusAfter('[data-pcar="' + cssq(code) + '"][data-half="am"]');
      } else {
        setCar(code, t.value, half);
        focusAfter('[data-pcar="' + cssq(code) + '"][data-half="' + (splitPerson(code) ? half : 'all') + '"]');
      }
      render();
    } else if (t.dataset.add) {
      if (!t.value) return;
      S.extra.push(t.value);
      focusAfter('[data-pcar="' + cssq(t.value) + '"]');
      render();
    } else if (t.dataset.carhalf) {
      S.pickHalf = t.value;
      focusAfter('[data-carhalf]');
      render();
    } else if (t.dataset.carto) {
      if (t.value) putPicked(t.value);
    } else if (t.dataset.carcheck) {          // 체크 칸 자체(키보드 Space) — 한 번에 한 대만
      S.pick = t.checked ? t.dataset.carcheck : ''; S.pickHalf = '';
      focusAfter('[data-carcheck="' + cssq(t.dataset.carcheck) + '"]');
      render();
    } else if (t.dataset.fixed) {
      var fc = carOf(t.dataset.fixed);
      post('/api/vehicles/fleet/' + encodeURIComponent(fc.plate), { fixed_user: t.value }).then(function (j) {
        fc.fixed_user = j.car.fixed_user; fc.fixed_name = name(j.car.fixed_user);
        msg(fc.label + '의 평소 사용자를 ' + (fc.fixed_user ? fc.fixed_name : '없음') + '(으)로 바꿨어요.', 'ok');
        render();
      }).catch(function (e2) { msg(e2.message, 'err'); t.value = fc.fixed_user || ''; });
    }
  });
  document.addEventListener('keydown', function (e) {
    var t = e.target;
    if (e.key !== 'Escape' || !t.dataset || !t.dataset.ckey) return;
    e.stopPropagation();
    if (!t.value.trim()) { delete S.memoOpen[t.dataset.ckey]; focusAfter('[data-memoopen="' + cssq(t.dataset.ckey) + '"]'); render(); }
    else t.blur();
  }, true);
  document.addEventListener('input', function (e) {
    var t = e.target;
    if (t.dataset && t.dataset.cmemo) { carOf(t.dataset.cmemo).note = t.value.trim(); setDirty(true); }   // 코스에 넣기는 change(칸을 떠날 때)
  });
  document.addEventListener('click', function (e) {
    if (S.suppressClick) { S.suppressClick = false; return; }
    if (!S.canEdit) return;
    var mx = e.target.closest('[data-memoclose]');
    if (mx) { closeMemo(mx.dataset.memoclose); return; }
    var mo = e.target.closest('[data-memoopen]');
    if (mo) {
      S.memoOpen[mo.dataset.memoopen] = 1;
      focusAfter('[data-ckey="' + cssq(mo.dataset.memoopen) + '"]');
      render();
      return;
    }
    if (e.target.closest('[data-carunchk]')) { S.pick = ''; focusAfter('#vcCarList'); render(); return; }
    // 차 줄은 아무 곳이나 눌러도 체크(방문 일정 사람 줄과 같게) — 다른 차를 누르면 그 차로 바뀜
    var crow = e.target.closest('[data-car]');
    if (crow && !e.target.closest('input, select, a, button')) {
      var rc = carOf(crow.dataset.car);
      if (ridersOf(rc).length && !splitFree(rc)) return;     // 오전·오후 모두 찬 차는 체크 안 됨(나눠 쓸 수 있으면 됨, 11-24)
      S.pick = S.pick === crow.dataset.car ? '' : crow.dataset.car;
      S.pickHalf = '';
      focusAfter('[data-carcheck="' + cssq(crow.dataset.car) + '"]');
      render();
      return;
    }
    var cl = e.target.closest('[data-clearc]');
    if (cl) {
      setCourse(cl.dataset.clearc, '', true);
      var oc = ownCar();
      if (oc && !ridersOf(oc).length) oc.note = '';
      focusAfter('[data-ccourse="' + cssq(cl.dataset.clearc) + '"]');
      render();
      return;
    }
    var card = e.target.closest('[data-ccourse]');
    if (S.pick && card && card.dataset.ccourse !== 'other' && !e.target.closest('select, input, button, a, summary, label')) {
      putPicked(card.dataset.ccourse);
      return;
    }
    var jp = e.target.closest('[data-joinp]');
    if (jp) {                                 // 하루 같게 — 오전 차를 종일로(오전이 없으면 오후 차)
      var code = jp.dataset.joinp;
      var at = whereAt(code, 'am') || whereAt(code, 'pm');
      delete S.splitOpen[code];
      setCar(code, at ? at.plate : '', 'all');
      focusAfter('[data-pcar="' + cssq(code) + '"][data-half="all"]');
      render();
    }
  });

  // 차 끌기(마우스만) — 왼쪽 차를 코스 카드(코스 모두)나 사람 칩(그 사람만)에 놓는다
  var D = null;
  document.addEventListener('pointerdown', function (e) {
    if (!S.canEdit || e.pointerType !== 'mouse' || e.button !== 0) return;
    var row = e.target.closest('[data-car]');
    if (!row || e.target.closest('input, select, a, button')) return;
    var dc = carOf(row.dataset.car);
    if (ridersOf(dc).length && !splitFree(dc)) return;          // 오전·오후 모두 찬 차는 끌기 안 됨(나눠 쓸 수 있으면 됨 — 다른 코스에 놓으면 나눠 쓰기, 11-24)
    D = { plate: row.dataset.car, x: e.clientX, y: e.clientY, ghost: null, over: null };
  });
  document.addEventListener('pointermove', function (e) {
    if (!D) return;
    if (!D.ghost) {
      if (Math.abs(e.clientX - D.x) + Math.abs(e.clientY - D.y) < 6) return;
      D.ghost = document.createElement('div');
      D.ghost.className = 'gw-va-ghost';
      D.ghost.textContent = carOf(D.plate).label;
      document.body.appendChild(D.ghost);
      document.body.classList.add('gw-va-dragging');
    }
    D.ghost.style.left = e.clientX + 'px';
    D.ghost.style.top = e.clientY + 'px';
    var t = document.elementFromPoint(e.clientX, e.clientY);
    var over = t && (t.closest('[data-cperson]') || t.closest('[data-ccourse]'));
    if (over && over.dataset.ccourse === 'other') over = null;   // 그 밖 카드는 사람 칩에만
    if (over && ridersOf(carOf(D.plate)).length && !over.dataset.ccourse) over = null;   // 나눠 쓰기는 코스 카드에만
    if (over !== D.over) {
      if (D.over) D.over.classList.remove('over');
      if (over) over.classList.add('over');
      D.over = over;
    }
  });
  function dropGhost() {
    if (!D) return;
    if (D.ghost) D.ghost.remove();
    if (D.over) D.over.classList.remove('over');
    document.body.classList.remove('gw-va-dragging');
    D = null;
  }
  document.addEventListener('keydown', function (e) {     // 차를 체크한 채 코스 카드에서 Enter·Space = 그 코스에 넣기 · Esc = 체크 풀기
    if (!S.canEdit || !S.pick) return;
    if (e.key === 'Escape') { S.pick = ''; render(); return; }
    var card = e.target.closest && e.target.closest('[data-ccourse]');
    if (!card || card !== e.target || (e.key !== 'Enter' && e.key !== ' ')) return;
    e.preventDefault();
    putPicked(card.dataset.ccourse);
  });
  document.addEventListener('pointercancel', dropGhost);
  window.addEventListener('blur', dropGhost);
  document.addEventListener('pointerup', function () {
    if (!D) return;
    var d = D;
    D = null;
    if (!d.ghost) return;
    d.ghost.remove();
    document.body.classList.remove('gw-va-dragging');
    if (!d.over) return;
    d.over.classList.remove('over');
    S.suppressClick = true;
    setTimeout(function () { S.suppressClick = false; }, 0);
    if (d.over.dataset.cperson) {
      delete S.splitOpen[d.over.dataset.cperson];
      setCar(d.over.dataset.cperson, d.plate, 'all');
    } else if (ridersOf(carOf(d.plate)).length) {      // 이미 코스에 든 차 — 체크해서 넣을 때와 같게 오전·오후 나눠 쓰기
      if (S.pick !== d.plate) S.pickHalf = '';         // 다른 차를 체크해 둔 반나절은 쓰지 않음
      S.pick = d.plate;
      putPicked(d.over.dataset.ccourse);
      if (S.pick === d.plate) { S.pick = ''; S.pickHalf = ''; render(); }   // 이미 그 코스 차였으면 체크만 풀기
      return;
    } else {
      setCourse(d.over.dataset.ccourse, d.plate, true);
    }
    S.pick = '';
    msg('');
    render();
  });

  // [일정만 먼저 보내기](11-25 부장님) — 차량 배정이 늦어질 때 기다리던 방문 일정 DM 만 지금. 차는 확정할 때 따로
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-act="sendnow"]');
    if (!b) return;
    if (!window.confirm(S.dayLabel + ' 방문 일정 DM 을 지금 먼저 보낼까요?\n\n차는 차량 배정을 확정할 때 차량 안내로 따로 가요.')) return;
    b.disabled = true;
    msg('방문 일정 DM 을 보내는 중…', 'info');
    post('/api/vehicles/plan/send-now', { day: S.day }).then(function (j) {
      var bad = j.dm_failed || j.bot_missing;
      return load(S.day).then(function () {
        msg('방문 일정 DM 을 ' + (j.dm || 0) + '명에게 먼저 보냈어요.' +
          (bad ? ' ' + (j.bot_missing ? '방문봇 연결이 없어 못 보냈어요.' : j.dm_failed + '명은 받지 못했어요.') + ' 직접 알려 주세요.' : ''),
          bad ? 'err' : 'ok');
      });
    }).catch(function (er) { b.disabled = false; msg(er.message, 'err'); });
  });

  // 슬랙으로 확정한 날의 방문 일정 칸 — 직접 접고 펼친 것만 기억
  document.addEventListener('click', function (e) {
    var s = e.target.closest && e.target.closest('#vaSched > summary');
    if (!s) return;
    setTimeout(function () {
      try { localStorage.setItem(SCHED_KEY, s.parentNode.open ? '1' : '0'); } catch (er) { /* 무시 */ }
    }, 0);
  });
  // 초기화(11-22 부장님) — 지금 화면의 배정을 모두 비운다. 저장해야 반영되고, 저장 전엔 다른 날로 갔다 오면 되돌아간다
  if ($('vaReset')) {
    $('vaReset').addEventListener('click', function () {
      if (!riders()) { msg('비울 배정이 없어요.', 'info', true); return; }
      if (!window.confirm('지금 화면의 차량 배정을 모두 비울까요?\n저장하기 전까지는 반영되지 않아요.')) return;
      S.cars.forEach(function (c) { c.people = []; c.pm = null; c.note = ''; });
      S.splitOpen = {}; S.pick = '';
      setDirty(true);
      render();
      msg('모두 비웠어요. 다시 넣은 뒤 저장해 주세요.', 'info', true);
    });
  }
  $('vaSave').addEventListener('click', function () { save(false); });
  $('vaConfirm').addEventListener('click', function () { save(true); });

  // 날짜는 달력으로만 고른다(11-17 결정 — 글자를 치면 '0001-11-23' 같은 날을 바로 불러왔음, 11-21 점검)
  (function pickerOnly(inp) {
    if (!inp) return;
    inp.addEventListener('click', function () { try { inp.showPicker(); } catch (e) { /* 무시 */ } });
    inp.addEventListener('keydown', function (e) {
      if (e.key === 'Tab' || e.key === 'Escape') return;
      e.preventDefault();
      if (e.key === 'Enter' || e.key === ' ') { try { inp.showPicker(); } catch (er) { /* 무시 */ } }
    });
    inp.addEventListener('paste', function (e) { e.preventDefault(); });
  })($('vaDay'));
  function go(day) {
    if (S.dirty && S.canEdit && !window.confirm('저장하지 않은 배정이 있어요. 다른 날로 갈까요?')) { $('vaDay').value = S.day; return; }
    load(day);
  }
  $('vaPrev').addEventListener('click', function () { go(addDays(S.day, -1)); });
  $('vaNext').addEventListener('click', function () { go(addDays(S.day, 1)); });
  $('vaDay').addEventListener('change', function () { if (this.value) go(this.value); });
  window.addEventListener('beforeunload', function (e) {
    if (S.dirty && S.canEdit) { e.preventDefault(); e.returnValue = ''; }
  });

  var q = new URLSearchParams(location.search).get('day');
  load(q || '');
})();
