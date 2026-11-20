/* 그룹웨어 입찰 공고 (2026-11-18) — 나라장터 공고 목록 + 검토 표시.
   서버에서 온 글자는 모두 esc() 로만 넣는다. 링크는 나라장터(g2b.go.kr) 주소만 그대로 쓴다. */
(function () {
  'use strict';

  var WD = ['일', '월', '화', '수', '목', '금', '토'];
  var MARKS = [['', '표시 없음'], ['review', '검토 중'], ['join', '참여'], ['skip', '안 함']];
  function $(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function safeUrl(u) {   // 나라장터·K-apt 주소만 링크로
    return /^https:\/\/((www|nuri)\.)?(g2b|k-apt)\.go\.kr\//.test(String(u || '')) ? u : '';
  }
  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  function api(url, opts) {
    opts = opts || {};
    var method = opts.method || 'GET';
    var init = { method: method, credentials: 'same-origin' };
    if (opts.body) {
      init.headers = { 'Content-Type': 'application/json' };
      init.body = JSON.stringify(opts.body);
    }
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error(SESSION_ENDED);
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          throw new Error((j && j.message) || (method === 'GET'
            ? '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'
            : '저장하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'));
        }
        return j;
      });
    });
  }
  function won(n) {
    n = Number(n) || 0;
    if (!n) return '-';
    if (n < 1000) return '비공개';   // 발주처가 1원 등으로 넣은 금액 비공개 공고(11-19 동원대학)
    if (n >= 100000000) {
      var eok = Math.floor(n / 100000000), man = Math.floor((n % 100000000) / 10000);
      return eok + '억' + (man ? ' ' + man.toLocaleString('ko-KR') + '만' : '') + ' 원';
    }
    return Math.floor(n / 10000).toLocaleString('ko-KR') + '만 원';
  }
  function when(s) {
    if (!s) return '-';
    var d = new Date(s.replace(' ', 'T'));
    if (isNaN(d)) return s;
    return (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ') ' + s.slice(11, 16);
  }
  function soon(s) {   // 48시간 안에 마감
    var d = new Date(String(s || '').replace(' ', 'T'));
    return !isNaN(d) && d - new Date() < 48 * 3600 * 1000 && d > new Date();
  }


  var state = { tab: 'eligible', q: '', closed: false, seq: 0, items: {}, open: '', anim: false,
    rows: [], area: '', src: '', showSkip: false, folded: {}, enabled: true };

  // ── 마감일 묶음(11-20 부장님: 쭉 나열되면 보기 힘듦) ──
  function dayStart(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function groupOf(it) {
    var d = new Date(String(it.close_at || '').replace(' ', 'T'));
    if (isNaN(d)) return { key: 'z', label: '마감 일시 없음' };
    if (d < new Date()) return { key: '0', label: '마감 지남' };
    var today = dayStart(new Date());
    var days = Math.round((dayStart(d) - today) / 86400000);
    if (days === 0) return { key: '1', label: '오늘 마감' };
    if (days === 1) return { key: '2', label: '내일 마감' };
    if (days <= 7) {   // 일주일 안은 날짜별(금요일에 '다음 주' 하나로 묶으면 24건이 몰림)
      var ymd = d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
      return { key: '3' + ymd, label: (d.getMonth() + 1) + '/' + d.getDate() + '(' + WD[d.getDay()] + ') 마감' };
    }
    return { key: '5', label: '그 이후' };
  }
  function render() {
    var all = state.rows;
    var skipped = all.filter(function (it) { return it.mark === 'skip'; }).length;
    var rows = all.filter(function (it) {
      return (!state.area || it.area === state.area) && (!state.src || (it.source || '나라장터') === state.src) &&
        (state.showSkip || it.mark !== 'skip');
    });
    $('bdSkipCnt').textContent = skipped ? '(' + skipped + ')' : '';
    $('bdCount').textContent = rows.length === all.length ? rows.length + '건' : rows.length + '건 / 전체 ' + all.length + '건';
    if (!rows.length) {
      $('bdRows').innerHTML = '<tr><td colspan="6" class="gw-empty">' + (!state.enabled ? '나라장터 연결 키가 서버에 없어요.'
        : all.length ? '고른 조건에 맞는 공고가 없어요.'
        : state.tab === 'eligible' ? '지금 들어갈 수 있는 마감 전 공고가 없어요.' : '입찰 불가능한 마감 전 공고가 없어요.') + '</td></tr>';
      return;
    }
    var groups = [], byKey = {};
    rows.forEach(function (it) {
      var g = groupOf(it);
      if (!byKey[g.key]) { byKey[g.key] = { key: g.key, label: g.label, items: [] }; groups.push(byKey[g.key]); }
      byKey[g.key].items.push(it);
    });
    groups.sort(function (a, b) { return a.key < b.key ? -1 : 1; });
    $('bdRows').innerHTML = groups.map(function (g) {
      var folded = !!state.folded[g.key];
      return '<tr class="gw-bd-grp' + (g.key === '1' || g.key === '2' ? ' urgent' : '') + '"><td colspan="6">' +
        '<button type="button" class="gw-bd-grp-btn" data-grp="' + g.key + '" aria-expanded="' + (folded ? 'false' : 'true') + '">' +
        '<i class="fas fa-chevron-' + (folded ? 'right' : 'down') + ' me-2" aria-hidden="true"></i>' + esc(g.label) +
        ' <b>' + g.items.length + '건</b></button></td></tr>' +
        (folded ? '' : g.items.map(row).join(''));
    }).join('');
    if (state.open) {   // 표시를 바꿔 다시 그려도 펼친 줄은 그대로(애니메이션 없이)
      var tr = $('bdRows').querySelector('tr.gw-bd-row[data-no="' + CSS.escape(state.open) + '"]');
      if (tr) openDetail(tr, false);
    }
  }

  function row(it) {
    var sel = '<select class="gw-chip-select gw-bd-mark" data-no="' + esc(it.bid_no) + '" aria-label="' + esc(it.title) + ' 표시">' +
      MARKS.map(function (m) {
        return '<option value="' + m[0] + '"' + (m[0] === it.mark ? ' selected' : '') + '>' + m[1] + '</option>';
      }).join('') + '</select>' +
      (it.mark ? '<div class="gw-bd-by">' + esc(it.mark_by) + ' · ' + esc(when(it.mark_at)) + '</div>' : '');
    var close = '<span class="' + (it.closed ? 'gw-bd-closed-t' : (soon(it.close_at) ? 'gw-bd-soon' : '')) + '">' +
      esc(when(it.close_at)) + '</span>' + (it.closed ? '<div class="gw-bd-sub">마감</div>' : '');
    var title = '<div class="gw-bd-title">' + (it.is_new ? '<span class="gw-badge gw-badge-new">New</span> ' : '') +
      esc(it.title) + '</div><div class="gw-bd-sub">' +
      '<span class="gw-bd-src ' + ({ '누리장터': 'nuri', '공동주택': 'apt' }[it.source] || 'g2b') + '">' + esc(it.source || '나라장터') + '</span>' +
      esc(it.demand_org || it.org) +
      (it.contact ? ' · 담당 ' + esc(it.contact) : '') + '</div>';
    var bAt = (it.detail || {}).briefing_at;   // 현장설명회 — 놓치면 못 들어가는 공고가 있어 줄에서 바로 보이게
    if (bAt) {
      title += '<div class="gw-bd-brief ' + (past(bAt) ? 'gone' : (soon(bAt) ? 'soon' : '')) + '">' +
        '<i class="fas fa-person-chalkboard me-1" aria-hidden="true"></i>현장설명회 ' + esc(when(bAt)) +
        (past(bAt) ? ' (지남)' : '') + '</div>';
    }
    var nuri = it.source === '누리장터';   // 자격·지역이 공고서에만 → '전국·면허 제한 없음'으로 보이면 오해
    var apt = it.source === '공동주택';
    var cond = '<div>' + (apt ? esc(it.site_region || '지역 미상') + ' 아파트' : nuri ? '지역·면허: 공고서 참조'
      : it.nationwide ? '<span class="gw-badge gw-badge-nation">전국</span> 지역 제한 없음'
      : esc((it.regions || []).join(', '))) + '</div><div class="gw-bd-sub">' +
      (apt ? '자격: K-apt 공고 참조' : nuri ? '누리장터 민간 공고' : esc((it.licenses || []).join(' 또는 ') || '면허 제한 없음')) + '</div>' +
      (it.category !== 'eligible' && it.reason ? '<div class="gw-bd-miss">' + esc(it.reason) + '</div>' : '');
    var open = state.open === it.bid_no;
    return '<tr class="gw-bd-row' + (it.mark === 'skip' ? ' gw-bd-skip' : '') + (open ? ' gw-bd-open' : '') +
      '" data-no="' + esc(it.bid_no) + '" tabindex="0" aria-expanded="' + (open ? 'true' : 'false') +
      '" title="눌러서 자세히 보기">' +
      '<td class="gw-bd-close">' + close + '</td>' +
      '<td>' + title + '</td>' +
      '<td class="gw-bd-nw">' + esc(it.award_method || it.contract_method) + '</td>' +
      '<td class="num gw-bd-nw">' + esc(won(it.price)) + '</td>' +
      '<td>' + cond + '</td>' +
      '<td>' + sel + '</td></tr>';
  }

  // ── 펼친 상세 — 나라장터 공고 상세처럼 표로, 맨 위엔 우리가 판단할 것만 먼저 ──
  function full(s) { return s ? esc(when(s)) : ''; }
  function past(s) { var d = new Date(String(s || '').replace(' ', 'T')); return !isNaN(d) && d < new Date(); }
  function chk(ok, label, text) {   // ok: true=됨, false=안 됨, null=주의
    var cls = ok === true ? 'ok' : ok === false ? 'no' : 'warn';
    var ic = ok === true ? 'fa-circle-check' : ok === false ? 'fa-circle-xmark' : 'fa-circle-exclamation';
    return '<li class="gw-bd-chk ' + cls + '"><i class="fas ' + ic + '" aria-hidden="true"></i>' +
      '<span class="gw-bd-chk-l">' + esc(label) + '</span><span class="gw-bd-chk-t">' + text + '</span></li>';
  }
  function blank(html) {   // 글자만 남겼을 때 비었거나 '-' 뿐이면 빈 칸
    var t = String(html == null ? '' : html).replace(/<[^>]*>/g, '').replace(/\s+/g, '');
    return !t || t === '-';
  }
  // 나라장터처럼 [항목|값|항목|값]. 빈 칸은 숨긴다(11-20 — 누리장터·공동주택은 없는 항목이 많아 '-'투성이였음):
  //   둘 다 비면 줄째 빼고, 한쪽만 비면 남은 쪽이 끝까지
  function tr2(l1, v1, l2, v2) {
    var e1 = blank(v1), e2 = !l2 || blank(v2);
    if (e1 && e2) return '';
    if (e1) { l1 = l2; v1 = v2; l2 = ''; }
    else if (e2) l2 = '';
    return '<tr><th scope="row">' + esc(l1) + '</th><td' + (l2 ? '' : ' colspan="3"') + '>' + v1 + '</td>' +
      (l2 ? '<th scope="row">' + esc(l2) + '</th><td>' + v2 + '</td>' : '') + '</tr>';
  }
  function sect(title, rows) {
    if (!rows) return '';   // 묶음 전체가 비면 제목도 뺀다
    return '<h3 class="gw-bd-sec">' + esc(title) + '</h3><table class="gw-bd-kv">' +
      '<colgroup><col class="k"><col><col class="k"><col></colgroup>' + rows + '</table>';
  }
  function detail(it) {
    var d = it.detail || {};
    var has = Object.keys(d).length > 0;
    var lic = (it.licenses || []).join(' 또는 ');
    var held = (state.held || []).join(', ');
    // 1) 확인할 것
    var rg = esc((it.regions || []).join(', '));
    var checks = it.region_ok ? chk(true, '참가 지역', rg + ' — 김포 업체 참가 가능')
      : it.nationwide ? chk(true, '참가 지역', '<b>지역 제한 없음(전국 공고)</b> — 어느 지역 업체나 들어와서 경쟁이 셀 수 있어요')
      : chk(false, '참가 지역', '<b>' + rg + '</b> — 김포 업체는 참가할 수 없어요');
    checks += it.license_ok
      ? chk(true, '면허', esc(lic || '제한 없음') + (lic ? ' — 보유' : ''))
      : chk(false, '면허', '<b>' + esc(it.missing || '우리 면허로 안 됨') + (it.missing ? ' 없음' : '') + '</b> — ' + esc(lic) +
          (lic.indexOf('+') > -1 ? ' 모두 있어야 해요' : ''));
    if (it.source === '누리장터' || it.source === '공동주택') checks = '';   // 지역·면허는 아래 '참가 자격' 한 줄로
    checks += it.hard_reason ? chk(false, '참가 조건', esc(it.hard_reason))
      : it.source === '공동주택'   // K-apt 는 실적·PQ 항목이 없어 '없음'이라 하면 틀림 — 입찰·낙찰 방식만
        ? chk(true, '입찰 방식', esc([it.contract_method, it.award_method].filter(Boolean).join(' · ') || '-'))
        : chk(true, '참가 조건', '실적제한·사전심사(PQ) 없음 · ' + esc(it.award_method || it.contract_method));
    if (d.briefing_at) {
      checks += chk(past(d.briefing_at) ? false : null, '현장설명회',
        '<b>' + full(d.briefing_at) + '</b>' + (d.briefing_place ? ' · ' + esc(d.briefing_place) : '') +
        (past(d.briefing_at) ? ' — 지났어요.' : '') + ' <span class="gw-bd-sub">참석이 필수인지 공고문에서 확인하세요</span>');
    }
    if (it.source === '누리장터') {   // 민간 입찰 — 자격·지역이 공고서에만 있어 자동 판정 못 함(11-20)
      checks += chk(null, '참가 자격', '<b>공고서 참조</b> <span class="gw-bd-sub">누리장터 민간 공고는 면허·지역 조건이 첨부 공고서에만 있어요. ' +
        '[누리장터에서 보기]를 누른 뒤 공고번호 ' + esc(it.bid_no) + '로 찾아보세요</span>');
    }
    if (it.source === '공동주택') {   // 아파트 입찰(11-20) — 면허·실적 조건은 K-apt 공고 화면·첨부에
      checks += chk(null, '참가 자격', '<b>K-apt 공고 참조</b>' +
        (d.apt_perf_cert === '예' ? ' · <b>실적증명서 제출</b>' : '') + (d.apt_credit_cert === '예' ? ' · 신용평가등급확인서 제출' : '') +
        ' <span class="gw-bd-sub">입찰 시스템: ' + esc(d.apt_auth || '-') + ' — 응찰은 그 사이트에서 해요</span>');
    }
    if (it.bid_kind === '물품' && it.source !== '공동주택') {   // 물품은 나라장터 물품 분류 등록이 있어야 투찰 — 등록 여부는 우리가 확인
      checks += chk(null, '물품 등록', '<b>' + esc((d.products || []).join(', ') || '공고문의 물품 분류') + '</b>' +
        ' <span class="gw-bd-sub">나라장터에 이 물품으로 등록돼 있어야 입찰할 수 있어요</span>');
    }
    if (d.qual_at) {
      checks += chk(past(d.qual_at) ? false : (soon(d.qual_at) ? null : true), '참가 등록 마감',
        '<b>' + full(d.qual_at) + '</b>' + (past(d.qual_at) ? ' — 지났어요' : '') + ' <span class="gw-bd-sub">(나라장터 입찰참가자격 등록)</span>');
    }
    checks += chk(it.closed ? false : (soon(it.close_at) ? null : true), '입찰 마감',
      '<b>' + full(it.close_at) + '</b>' + (it.closed ? ' — 지났어요' : ''));
    // 2) 표
    var tel = d.contact_tel ? '<a href="tel:' + esc(d.contact_tel.replace(/[^0-9+]/g, '')) + '">' + esc(d.contact_tel) + '</a>' : '';
    var mail = d.contact_email && /^[^\s@<>"']+@[^\s@<>"']+$/.test(d.contact_email)
      ? '<a href="mailto:' + esc(d.contact_email) + '">' + esc(d.contact_email) + '</a>' : '';
    var files = (it.files || []).map(function (f) {
      var u = safeUrl(f.url);
      return u ? '<li><a href="' + esc(u) + '" target="_blank" rel="noopener">' + esc(f.name) + '</a></li>' : '';
    }).join('');
    // 누리장터 공고는 공고별 주소가 없어 누리장터 첫 화면으로(공고번호로 검색) — 11-20
    var du = safeUrl(it.detail_url) || (it.source === '누리장터' ? 'https://nuri.g2b.go.kr/' : '');
    var body =
      sect('일정',
        tr2('입찰 시작', full(d.begin_at), '입찰 마감', '<b>' + full(it.close_at) + '</b>') +
        tr2('참가 등록 마감', full(d.qual_at), '개찰', full(it.open_at)) +
        tr2('현장설명회', full(d.briefing_at), '설명회 장소', esc(d.briefing_place)) +
        tr2('공고 게시', full(it.posted_at), '개찰 장소', esc(d.open_place))) +
      sect('금액·낙찰',
        tr2(it.source === '누리장터' ? '기준금액' : '추정가격', '<b>' + esc(won(it.price)) + '</b>',
            '배정예산', d.budget ? esc(won(d.budget)) : '') +
        tr2('부가세 포함', esc(d.vat_note)) +
        tr2('부가세', d.vat ? esc(won(d.vat)) : '', '낙찰하한율', esc(d.lower_rate)) +
        tr2('계약 방법', esc([it.contract_method, d.bid_method].filter(Boolean).join(' · ')),
            '예정가격', esc([d.price_method, d.price_count].filter(Boolean).join(' · '))) +
        tr2('낙찰 방법', esc(d.award_full || it.award_method)) +
        (it.bid_kind === '물품'
          ? tr2('구매 물품', esc((d.products || []).join(', '))) +
            tr2('납품 기한', d.deliver_days ? '계약일부터 ' + esc(d.deliver_days) + '일' : '', '납품 조건', esc(d.deliver_cond)) +
            tr2('제조업체만', esc(d.maker_only), '물품 분류 제한', esc(d.product_limit))
          : '')) +
      sect('참가 조건',
        (it.source === '나라장터' || !it.source   // 면허·지역 제한 칸은 나라장터만 있음 — 다른 곳에 '제한 없음'은 틀린 말
          ? tr2('참가가능지역', esc((it.regions || []).join(', '))) +
            tr2('면허', esc(lic || '제한 없음') + (held ? ' <span class="gw-bd-sub">(우리 면허: ' + esc(held) + ')</span>' : ''))
          : '') +
        tr2('참가 자격', esc(d.nuri_qual), '지역 제한', esc(d.nuri_region)) +                       // 누리장터
        tr2('실적증명서', esc(d.apt_perf_cert), '신용평가 확인서', esc(d.apt_credit_cert)) +          // 공동주택
        tr2('응찰 시스템', esc(d.apt_auth), '지급 조건', esc(d.apt_pay)) +
        tr2('공고 내용', esc(d.apt_content)) +
        tr2('주공종', esc(d.main_license), '공사현장', esc(d.site_region || it.site_region)) +
        tr2('공동수급', esc(d.joint), '지역 공동도급', esc(d.joint_region))) +
      sect('공고·문의',
        tr2('공고번호', esc(it.bid_no + '-' + it.bid_ord) + (d.re_notice === '예' ? ' (재공고)' : ''), '참조번호', esc(d.ref_no)) +
        tr2('발주기관', esc(it.org), '수요기관', esc(it.demand_org)) +
        tr2('담당자', esc(d.contact_name || it.contact), '연락처', [tel, mail].filter(Boolean).join(' · '))) +
      '<h3 class="gw-bd-sec">첨부파일</h3>' +
      (files ? '<ul class="gw-bd-files">' + files + '</ul>'
        : '<p class="gw-bd-sub">' + (it.source === '공동주택' ? '첨부파일은 K-apt 공고 화면에 있어요.' : '첨부파일이 없어요.') + '</p>');
    var foot = '<div class="gw-bd-acc-foot">처음 본 때 ' + esc(when(it.first_seen)) +
      (it.mark ? ' · ' + esc(it.mark_label) + ' 표시: ' + esc(it.mark_by) + ' ' + esc(when(it.mark_at)) : '') +
      (has ? '' : ' · 일정·낙찰하한율 같은 세부 항목은 다음 모으기 때 채워져요.') + '</div>';
    return '<tr class="gw-bd-detail"><td colspan="6"><div class="gw-bd-acc' + (state.anim ? ' gw-bd-slide' : '') + '">' +
      '<div class="gw-bd-acc-top"><ul class="gw-bd-chks" aria-label="확인할 것">' + checks + '</ul>' +
      (du ? '<a class="btn btn-sm btn-primary gw-bd-acc-go" href="' + esc(du) + '" target="_blank" rel="noopener">' +
        '<i class="fas fa-up-right-from-square me-1" aria-hidden="true"></i>' + ({ '공동주택': 'K-apt에서 보기', '누리장터': '누리장터에서 보기' }[it.source] || '나라장터에서 보기') + '</a>' : '') + '</div>' +
      body + foot + '</div></td></tr>';
  }
  function closeDetail() {
    var rows = $('bdRows');
    Array.prototype.forEach.call(rows.querySelectorAll('tr.gw-bd-detail'), function (x) { x.remove(); });
    Array.prototype.forEach.call(rows.querySelectorAll('tr.gw-bd-open'), function (x) {
      x.classList.remove('gw-bd-open');
      x.setAttribute('aria-expanded', 'false');
    });
  }
  function openDetail(tr, animate) {
    var it = state.items[tr.getAttribute('data-no')];
    if (!it) return;
    closeDetail();
    state.open = it.bid_no;
    state.anim = animate;
    tr.classList.add('gw-bd-open');
    tr.setAttribute('aria-expanded', 'true');
    tr.insertAdjacentHTML('afterend', detail(it));
  }

  function load() {
    var my = ++state.seq;
    var qs = '?tab=' + encodeURIComponent(state.tab) + (state.closed ? '&closed=1' : '') +
      (state.q ? '&q=' + encodeURIComponent(state.q) : '');
    api('/api/bids' + qs).then(function (j) {
      if (my !== state.seq) return;
      $('bdCntEligible').textContent = j.counts.eligible;
      $('bdCntBlocked').textContent = j.counts.blocked;
      $('bdRule').textContent = '나라장터·누리장터·공동주택(K-apt)의 냉난방 관련 공사 공고와 물품(구매·설치) 공고를 모았어요. 입찰 가능 기준: 참가가능지역 제한 없음(전국) 또는 경기도·김포시 포함, ' +
        '지금 면허(' + (j.held || []).join(', ') + ')로 참가 가능, 실적제한·사전심사(PQ)·지명경쟁·4억 이상 적격심사 없음.' +
        (j.planned && j.planned.length ? ' 준비 중인 면허: ' + j.planned.join(', ') + '.' : '');
      state.held = j.held || [];
      state.enabled = !!j.enabled;
      state.rows = j.items || [];
      state.items = {};
      state.rows.forEach(function (it) { state.items[it.bid_no] = it; });
      if (state.open && !state.items[state.open]) state.open = '';
      render();
      $('bdFoot').textContent = j.last_collect ? '마지막으로 모은 때: ' + j.last_collect + ' (하루 세 번 07:20·12:20·16:20)' : '';
    }).catch(function (e) {
      if (my !== state.seq) return;
      $('bdRows').innerHTML = '<tr><td colspan="6" class="gw-empty">' + esc(e.message) + '</td></tr>';
    });
  }

  $('bdTabs').addEventListener('click', function (ev) {
    var b = ev.target.closest('button[data-tab]');
    if (!b) return;
    state.tab = b.getAttribute('data-tab');
    Array.prototype.forEach.call(this.querySelectorAll('button'), function (x) {
      var on = x === b;
      x.classList.toggle('active', on);
      x.setAttribute('aria-selected', on ? 'true' : 'false');
      x.tabIndex = on ? 0 : -1;
    });
    load();
  });
  var qt;
  $('bdQ').addEventListener('input', function () {
    clearTimeout(qt);
    var v = this.value.trim();
    qt = setTimeout(function () { state.q = v; load(); }, 250);
  });
  $('bdClosed').addEventListener('change', function () { state.closed = this.checked; load(); });
  $('bdArea').addEventListener('change', function () { state.area = this.value; render(); });
  (function () {   // 플랫폼 고르기(11-20) — 템플릿을 안 바꿔 재시작 없이 붙도록 여기서 만든다
    var sel = document.createElement('select');
    sel.className = 'gw-chip-select';
    sel.id = 'bdSrc';
    sel.setAttribute('aria-label', '플랫폼');
    [['', '플랫폼 전체'], ['나라장터', '나라장터'], ['누리장터', '누리장터'], ['공동주택', '공동주택']].forEach(function (o) {
      var op = document.createElement('option');
      op.value = o[0];
      op.textContent = o[1];
      sel.appendChild(op);
    });
    $('bdArea').insertAdjacentElement('afterend', sel);
    sel.addEventListener('change', function () { state.src = this.value; render(); });
  })();
  $('bdSkip').addEventListener('change', function () { state.showSkip = this.checked; render(); });
  $('bdRows').addEventListener('click', function (ev) {
    var gb = ev.target.closest('button.gw-bd-grp-btn');   // 묶음 머리줄 = 접기·펴기
    if (gb) {
      var k = gb.getAttribute('data-grp');
      state.folded[k] = !state.folded[k];
      render();
      var again = $('bdRows').querySelector('button.gw-bd-grp-btn[data-grp="' + k + '"]');
      if (again) again.focus();
      return;
    }
    if (ev.target.closest('a, button, select, input, label, .gw-bd-detail')) return;   // 링크·표시 칸은 펼치기 제외
    var tr = ev.target.closest('tr.gw-bd-row');
    if (!tr) return;
    if (state.open === tr.getAttribute('data-no')) {
      state.open = '';
      closeDetail();
    } else {
      openDetail(tr, true);
    }
  });
  $('bdRows').addEventListener('keydown', function (ev) {   // 줄에 커서가 있을 때 Enter·Space 로 펼치기
    if ((ev.key !== 'Enter' && ev.key !== ' ') || !ev.target.classList || !ev.target.classList.contains('gw-bd-row')) return;
    ev.preventDefault();
    ev.target.click();
  });
  document.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Escape' || !state.open) return;
    var tr = $('bdRows').querySelector('tr.gw-bd-open');
    state.open = '';
    closeDetail();
    if (tr) tr.focus();
  });
  $('bdRows').addEventListener('change', function (ev) {
    var s = ev.target.closest('select.gw-bd-mark');
    if (!s) return;
    s.disabled = true;
    api('/api/bids/' + encodeURIComponent(s.getAttribute('data-no')) + '/mark', { method: 'POST', body: { mark: s.value } })
      .then(function () { load(); })
      .catch(function (e) { s.disabled = false; alert(e.message); load(); });
  });

  load();
})();
