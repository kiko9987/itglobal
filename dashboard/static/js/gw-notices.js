/* 그룹웨어 공지사항·개발 노트(11-25). 화면 하나에서 목록(list)·글(view)·쓰기(new/edit)를 그린다.
   권한·대상은 서버(services/notices.py)가 다시 확인한다 — 여기서 숨기는 단추는 편의일 뿐. */
(function () {
  'use strict';

  var SESSION_ENDED = '로그인이 끝났어요. 새로 고친 뒤 다시 해 주세요.';
  var WD = ['일', '월', '화', '수', '목', '금', '토'];
  var FILE_MAX = 15 * 1024 * 1024;

  var page = document.getElementById('ntPage');
  if (!page) return;
  var root = document.getElementById('ntRoot');
  var flashBox = document.getElementById('ntFlash');
  var KIND = page.getAttribute('data-kind');
  var LABEL = page.getAttribute('data-label');
  var MODE = page.getAttribute('data-mode');
  var NID = Number(page.getAttribute('data-id') || 0);
  var CAN_EDIT = page.getAttribute('data-can-edit') === '1';
  var BASE = KIND === 'devnote' ? '/devnotes' : '/notices';

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function api(url, opts) {
    opts = opts || {};
    var method = opts.method || 'GET';
    var init = { method: method, credentials: 'same-origin' };
    if (opts.form) init.body = opts.form;
    else if (opts.body) { init.headers = { 'Content-Type': 'application/json' }; init.body = JSON.stringify(opts.body); }
    return fetch(url, init).then(function (r) {
      var isJson = (r.headers.get('content-type') || '').indexOf('application/json') !== -1;
      if (r.status === 401 || r.redirected || (r.ok && !isJson)) throw new Error(SESSION_ENDED);
      return (isJson ? r.json() : Promise.reject()).catch(function () { return null; }).then(function (j) {
        if (!j || !r.ok || j.success === false) {
          throw new Error((j && j.message) || (method === 'GET'
            ? '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요. (오류 ' + r.status + ')'
            : '처리하지 못했어요. 잠시 뒤 다시 해 주세요. (오류 ' + r.status + ')'));
        }
        return j;
      });
    });
  }

  function flash(msg, kind) {
    flashBox.className = 'gw-lv-flash ' + (kind || 'ok');
    flashBox.textContent = msg;
    flashBox.hidden = !msg;
    if (msg) flashBox.focus();
  }
  // 다른 화면으로 옮긴 뒤 보여 줄 안내(저장했어요 등)
  function flashNext(msg) { try { sessionStorage.setItem('ntFlash', msg); } catch (e) { /* 비공개 창 */ } }
  function flashPending() {
    var m = '';
    try { m = sessionStorage.getItem('ntFlash') || ''; sessionStorage.removeItem('ntFlash'); } catch (e) { /* */ }
    if (m) flash(m, 'ok');
  }

  function when(s, withTime) {
    if (!s) return '';
    var d = new Date(s.replace(' ', 'T'));
    var now = new Date();
    var md = (d.getMonth() + 1) + '월 ' + d.getDate() + '일 (' + WD[d.getDay()] + ')';
    var y = d.getFullYear() !== now.getFullYear() ? d.getFullYear() + '년 ' : '';
    var hm = ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
    return y + md + (withTime ? ' ' + hm : '');
  }
  function shortDate(s) {
    if (!s) return '';
    var d = new Date(s.replace(' ', 'T'));
    var now = new Date();
    if (d.toDateString() === now.toDateString()) return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
    var md = (d.getMonth() + 1) + '.' + d.getDate();
    return d.getFullYear() === now.getFullYear() ? md : String(d.getFullYear()).slice(2) + '.' + md;
  }
  function size(n) {
    if (n >= 1024 * 1024) return (Math.round(n / 1024 / 102.4) / 10) + 'MB';
    return Math.max(1, Math.round(n / 1024)) + 'KB';
  }
  function fileIcon(name) {
    var x = (name.split('.').pop() || '').toLowerCase();
    if (/^(jpg|jpeg|png|gif|webp|heic)$/.test(x)) return 'fa-file-image';
    if (x === 'pdf') return 'fa-file-pdf';
    if (/^(xls|xlsx)$/.test(x)) return 'fa-file-excel';
    if (/^(doc|docx|hwp|hwpx|txt)$/.test(x)) return 'fa-file-lines';
    if (/^(ppt|pptx)$/.test(x)) return 'fa-file-powerpoint';
    if (x === 'zip') return 'fa-file-zipper';
    return 'fa-file';
  }
  function canView(name) { return /\.(pdf|jpe?g|png|gif|webp)$/i.test(name); }

  // 본문 — 글자 그대로(줄바꿈 유지) + 주소는 링크
  function bodyHtml(t) {
    return esc(t).replace(/(https?:\/\/[^\s<]+[^\s<.,)\]])/g, function (u) {
      return '<a href="' + u + '" target="_blank" rel="noopener">' + u + '</a>';
    });
  }

  function setTitle(t) { document.title = t + ' - 아이티글로벌그룹'; }

  // ── 목록 ─────────────────────────────────────────────────────────
  var listState = { cat: 0, q: '' };

  function renderList() {
    var h = '<div class="gw-page-head gw-nt-head"><h1 class="gw-page-title" tabindex="-1" id="ntH1">' + esc(LABEL) +
      ' <span class="gw-nt-unread" id="ntUnread" hidden></span></h1>';
    if (CAN_EDIT) {
      h += '<div class="gw-nt-head-acts">' +
        (KIND === 'notice' ? '<button type="button" class="btn btn-outline-secondary" id="ntCatBtn" aria-expanded="false" aria-controls="ntCats">카테고리 관리</button>' : '') +
        '<a class="btn btn-primary" href="' + BASE + '/new"><i class="fas fa-pen" aria-hidden="true"></i> 글쓰기</a></div>';
    }
    h += '</div>';
    if (KIND === 'devnote') {
      h += '<p class="gw-note gw-nt-lead">그룹웨어·관리 사이트·슬랙 봇이 바뀐 내용을 남겨요.</p>';
    }
    h += '<div class="gw-nt-cats-panel" id="ntCats" hidden></div>';
    h += '<div class="gw-nt-bar">' +
      (KIND === 'notice' ? '<div class="gw-nt-chips" id="ntChips" role="group" aria-label="카테고리"></div>' : '<span></span>') +
      '<label class="gw-nt-search"><span class="visually-hidden">찾기</span><i class="fas fa-magnifying-glass" aria-hidden="true"></i>' +
      '<input type="search" class="form-control" id="ntQ" placeholder="제목·내용 찾기" autocomplete="off" maxlength="50"></label></div>';
    h += '<ul class="gw-nt-list" id="ntList" aria-live="polite"><li class="gw-empty">불러오는 중…</li></ul>';
    root.innerHTML = h;
    setTitle(LABEL);
    var qT = null;
    document.getElementById('ntQ').addEventListener('input', function (e) {
      clearTimeout(qT);
      qT = setTimeout(function () { listState.q = e.target.value.trim(); loadList(); }, 250);
    });
    var cb = document.getElementById('ntCatBtn');
    if (cb) cb.addEventListener('click', function () { toggleCats(cb); });
    loadList();
  }

  function loadList() {
    var url = '/api/notices?kind=' + KIND + (listState.cat ? '&category=' + listState.cat : '') +
      (listState.q ? '&q=' + encodeURIComponent(listState.q) : '');
    api(url).then(function (j) {
      renderChips(j.categories || []);
      var u = document.getElementById('ntUnread');
      if (u) { u.hidden = !j.unread || listState.cat || listState.q; u.textContent = '안 읽음 ' + j.unread; }
      var list = document.getElementById('ntList');
      if (!j.items.length) {
        list.innerHTML = '<li class="gw-empty">' + (listState.q ? '"' + esc(listState.q) + '"에 맞는 글이 없어요.'
          : (KIND === 'devnote' ? '아직 개발 노트가 없어요.' : '아직 공지가 없어요.')) + '</li>';
        return;
      }
      list.innerHTML = j.items.map(function (n) {
        var cls = 'gw-nt-row' + (n.pinned ? ' pinned' : '') + (n.read ? '' : ' unread');
        return '<li class="' + cls + '"><a class="gw-nt-link" href="' + BASE + '/' + n.id + '">' +
          '<span class="gw-nt-mark" aria-hidden="true">' + (n.pinned ? '<i class="fas fa-thumbtack"></i>' : '') + '</span>' +
          '<span class="gw-nt-main">' +
            '<span class="gw-nt-title">' +
              (n.category && KIND === 'notice' ? '<span class="gw-nt-tag">' + esc(n.category) + '</span>' : '') +
              '<span class="gw-nt-t">' + esc(n.title) + '</span>' +
              (n.read ? '' : '<span class="gw-nt-new">새 글</span>') +
              (n.files.length ? '<i class="fas fa-paperclip gw-nt-clip" aria-label="첨부 ' + n.files.length + '개"></i>' : '') +
            '</span>' +
            (n.preview ? '<span class="gw-nt-prev">' + esc(n.preview) + '</span>' : '') +
          '</span>' +
          '<span class="gw-nt-meta"><span class="gw-nt-author">' + esc(n.author) + '</span>' +
          '<time datetime="' + esc(n.created_at) + '">' + esc(shortDate(n.created_at)) + '</time></span>' +
          '</a></li>';
      }).join('');
    }).catch(function (e) {
      document.getElementById('ntList').innerHTML = '<li class="gw-empty">' + esc(e.message) + '</li>';
    });
  }

  var chipsDrawn = false;
  function renderChips(cats) {
    var box = document.getElementById('ntChips');
    if (!box || chipsDrawn) return;
    chipsDrawn = true;
    if (cats.length < 2) { box.hidden = true; return; }
    box.innerHTML = [{ id: 0, name: '전체' }].concat(cats).map(function (c) {
      return '<button type="button" class="gw-nt-chip" data-cat="' + c.id + '" aria-pressed="' +
        (c.id === listState.cat ? 'true' : 'false') + '">' + esc(c.name) + '</button>';
    }).join('');
    box.addEventListener('click', function (e) {
      var b = e.target.closest('[data-cat]');
      if (!b) return;
      listState.cat = Number(b.getAttribute('data-cat'));
      box.querySelectorAll('[data-cat]').forEach(function (x) { x.setAttribute('aria-pressed', x === b ? 'true' : 'false'); });
      loadList();
    });
  }

  // ── 카테고리 관리(관리자) ───────────────────────────────────────────
  var catData = null;
  function toggleCats(btn) {
    var p = document.getElementById('ntCats');
    var open = p.hidden;
    p.hidden = !open;
    btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    if (open) loadCats();
  }
  function loadCats() {
    var p = document.getElementById('ntCats');
    p.innerHTML = '<div class="gw-empty">불러오는 중…</div>';
    api('/api/notices/categories').then(function (j) { catData = j; drawCats(); })
      .catch(function (e) { p.innerHTML = '<div class="gw-empty">' + esc(e.message) + '</div>'; });
  }
  function drawCats(editId) {
    var p = document.getElementById('ntCats');
    var h = '<div class="gw-panel gw-nt-cats"><div class="gw-nt-cats-head"><h2 class="gw-panel-title">카테고리 관리</h2>' +
      '<p class="gw-panel-desc">카테고리마다 볼 수 있는 사람이 정해져요. 슬랙 채널을 넣으면 글을 올릴 때 그 채널로 알릴 수 있어요.</p></div>' +
      '<table class="gw-nt-cat-table"><thead><tr><th scope="col">이름</th><th scope="col">보는 사람</th><th scope="col">슬랙 채널</th>' +
      '<th scope="col"><span class="visually-hidden">고치기</span></th></tr></thead><tbody>';
    catData.categories.forEach(function (c) {
      if (c.id === editId) { h += '<tr><td colspan="4">' + catForm(c) + '</td></tr>'; return; }
      h += '<tr><td>' + esc(c.name) + '</td><td>' + esc(c.audience_label) + '</td><td>' +
        (c.slack_channel ? '<code>' + esc(c.slack_channel) + '</code>' : '<span class="gw-nt-dim">없음</span>') + '</td>' +
        '<td class="gw-nt-cat-act"><button type="button" class="btn btn-sm btn-link" data-cedit="' + c.id + '">고치기</button></td></tr>';
    });
    h += '</tbody></table>';
    h += editId === -1 ? catForm(null) : '<button type="button" class="btn btn-sm btn-outline-secondary gw-nt-cat-add" data-cedit="-1">' +
      '<i class="fas fa-plus" aria-hidden="true"></i> 카테고리 추가</button>';
    h += '</div>';
    p.innerHTML = h;
    p.querySelectorAll('[data-cedit]').forEach(function (b) {
      b.addEventListener('click', function () { drawCats(Number(b.getAttribute('data-cedit'))); });
    });
    var f = p.querySelector('form');
    if (f) wireCatForm(f, editId);
  }
  function catForm(c) {
    var aud = (c && c.audience) || 'all';
    var opts = Object.keys(catData.audiences).map(function (k) {
      return '<option value="' + k + '"' + (k === aud ? ' selected' : '') + '>' + esc(catData.audiences[k]) + '</option>';
    }).join('');
    var depts = catData.depts.map(function (d, i) {
      var on = c && c.depts && c.depts.indexOf(d) !== -1;
      return '<label class="form-check form-check-inline"><input class="form-check-input" type="checkbox" name="dept" value="' +
        esc(d) + '" id="ntDept' + i + '"' + (on ? ' checked' : '') + '><span class="form-check-label">' + esc(d) + '</span></label>';
    }).join('');
    return '<form class="gw-nt-cat-form" novalidate>' +
      '<div class="gw-nt-cat-fields">' +
        '<div><label class="form-label" for="ntCName">이름</label><input class="form-control" id="ntCName" name="name" maxlength="20" required value="' + esc(c ? c.name : '') + '"></div>' +
        '<div><label class="form-label" for="ntCAud">보는 사람</label><select class="form-select" id="ntCAud" name="audience">' + opts + '</select></div>' +
        '<div><label class="form-label" for="ntCCh">슬랙 채널 ID <span class="gw-nt-dim">선택</span></label><input class="form-control" id="ntCCh" name="slack_channel" maxlength="12" autocomplete="off" spellcheck="false" placeholder="C0123ABCD" value="' + esc(c ? (c.slack_channel || '') : '') + '"></div>' +
      '</div>' +
      '<fieldset class="gw-nt-depts"' + (aud === 'dept' ? '' : ' hidden') + '><legend class="form-label">보여 줄 부서</legend>' + depts + '</fieldset>' +
      '<div class="gw-lv-err" role="alert"></div>' +
      '<div class="gw-nt-cat-btns">' +
        (c ? '<button type="button" class="btn btn-sm btn-link text-danger me-auto" data-cdel="' + c.id + '">지우기</button>' : '') +
        '<button type="button" class="btn btn-sm btn-light" data-ccancel>닫기</button>' +
        '<button type="submit" class="btn btn-sm btn-primary">저장</button></div></form>';
  }
  function wireCatForm(f, editId) {
    var err = f.querySelector('.gw-lv-err');
    f.elements.audience.addEventListener('change', function () {
      f.querySelector('.gw-nt-depts').hidden = f.elements.audience.value !== 'dept';
    });
    f.querySelector('[data-ccancel]').addEventListener('click', function () { drawCats(); });
    var del = f.querySelector('[data-cdel]');
    if (del) del.addEventListener('click', function () {
      if (!window.confirm('이 카테고리를 지울까요? 카테고리에 공지가 있으면 지울 수 없어요.')) return;
      api('/api/notices/categories/' + editId + '/delete', { method: 'POST', body: {} })
        .then(function () { chipsDrawn = false; loadCats(); loadList(); })
        .catch(function (e) { err.textContent = e.message; });
    });
    f.elements.name.focus();
    f.addEventListener('submit', function (e) {
      e.preventDefault();
      var depts = Array.prototype.map.call(f.querySelectorAll('input[name=dept]:checked'), function (x) { return x.value; });
      var body = { name: f.elements.name.value, audience: f.elements.audience.value, depts: depts,
                   slack_channel: f.elements.slack_channel.value };
      var sb = f.querySelector('button[type=submit]');
      sb.disabled = true; err.textContent = '';
      api('/api/notices/categories' + (editId > 0 ? '/' + editId : ''), { method: 'POST', body: body })
        .then(function () { chipsDrawn = false; loadCats(); loadList(); })
        .catch(function (x) { err.textContent = x.message; sb.disabled = false; });
    });
  }

  // ── 글 ───────────────────────────────────────────────────────────
  function renderView() {
    api('/api/notices/' + NID).then(function (j) {
      var n = j.notice;
      setTitle(n.title);
      var h = '<a class="gw-nt-back" href="' + BASE + '"><i class="fas fa-arrow-left" aria-hidden="true"></i> ' + esc(LABEL) + '</a>' +
        '<article class="gw-panel gw-nt-article" aria-labelledby="ntH1">' +
        '<header class="gw-nt-ahead">' +
          (n.category ? '<span class="gw-nt-tag">' + esc(n.category) + '</span>' : '') +
          (n.pinned ? '<span class="gw-nt-pin"><i class="fas fa-thumbtack" aria-hidden="true"></i> 맨 위 고정</span>' : '') +
          '<h1 class="gw-nt-h1" id="ntH1" tabindex="-1">' + esc(n.title) + '</h1>' +
          '<p class="gw-nt-ameta">' + esc(n.author) + ' · <time datetime="' + esc(n.created_at) + '">' + esc(when(n.created_at, true)) + '</time>' +
            (n.updated_at ? ' · 고침 ' + esc(when(n.updated_at, true)) : '') +
          '</p>' +
        '</header>' +
        '<div class="gw-nt-body">' + (n.body ? bodyHtml(n.body) : '<span class="gw-nt-dim">본문 없음</span>') + '</div>';
      if (n.files.length) {
        h += '<div class="gw-nt-files"><h2 class="gw-nt-files-h">첨부 ' + n.files.length + '개</h2><ul>' +
          n.files.map(function (f) {
            return '<li><i class="fas ' + fileIcon(f.name) + '" aria-hidden="true"></i>' +
              '<a href="/api/notices/files/' + f.id + '" download>' + esc(f.name) + '</a><span class="gw-nt-size">' + size(f.size) + '</span>' +
              (canView(f.name) ? '<a class="gw-nt-fview" href="/api/notices/files/' + f.id + '?view=1" target="_blank" rel="noopener">열어 보기</a>' : '') +
              '</li>';
          }).join('') + '</ul></div>';
      }
      if (CAN_EDIT) {
        h += '<footer class="gw-nt-afoot">' +
          '<a class="btn btn-outline-secondary btn-sm" href="' + BASE + '/' + n.id + '/edit"><i class="fas fa-pen" aria-hidden="true"></i> 고치기</a>' +
          (n.slack_sent ? '<span class="gw-nt-sent"><i class="fab fa-slack" aria-hidden="true"></i> 슬랙으로 알렸어요</span>'
            : '<button type="button" class="btn btn-outline-secondary btn-sm" id="ntNotify"><i class="fab fa-slack" aria-hidden="true"></i> 슬랙으로 알리기</button>') +
          '<button type="button" class="btn btn-link btn-sm text-danger ms-auto" id="ntDel">지우기</button></footer>';
      }
      h += '</article>';
      if (n.reads) h += readsHtml(n.reads);
      root.innerHTML = h;
      document.getElementById('ntH1').focus({ preventScroll: true });
      flashPending();
      wireView(n);
    }).catch(function (e) {
      root.innerHTML = '<a class="gw-nt-back" href="' + BASE + '"><i class="fas fa-arrow-left" aria-hidden="true"></i> ' + esc(LABEL) + '</a>' +
        '<div class="gw-empty gw-empty-box">' + esc(e.message) + '</div>';
    });
  }

  function readsHtml(r) {
    var pct = r.total ? Math.round(r.read.length / r.total * 100) : 0;
    var h = '<section class="gw-panel gw-nt-reads" aria-labelledby="ntReadsH">' +
      '<div class="gw-nt-reads-head"><h2 class="gw-panel-title" id="ntReadsH">읽음 확인</h2>' +
      '<span class="gw-nt-reads-n"><b>' + r.read.length + '</b> / ' + r.total + '명 읽음</span></div>' +
      '<div class="gw-nt-bar-track" role="img" aria-label="대상 ' + r.total + '명 중 ' + r.read.length + '명 읽음"><span style="width:' + pct + '%"></span></div>';
    if (r.unread.length) {
      h += '<h3 class="gw-nt-reads-sub">안 읽은 사람 ' + r.unread.length + '명</h3><ul class="gw-nt-people">' +
        r.unread.map(function (p) { return '<li>' + esc(p.name) + (p.dept ? '<span>' + esc(p.dept) + '</span>' : '') + '</li>'; }).join('') + '</ul>';
    } else if (r.total) {
      h += '<p class="gw-nt-allread"><i class="fas fa-circle-check" aria-hidden="true"></i> 대상 모두 읽었어요.</p>';
    }
    if (r.read.length) {
      h += '<details class="gw-nt-readlist"><summary>읽은 사람 ' + r.read.length + '명</summary><ul class="gw-nt-people">' +
        r.read.map(function (p) { return '<li>' + esc(p.name) + '<span>' + esc(shortDate(p.read_at)) + '</span></li>'; }).join('') + '</ul></details>';
    }
    return h + '</section>';
  }

  function wireView(n) {
    var nb = document.getElementById('ntNotify');
    if (nb) nb.addEventListener('click', function () {
      if (!window.confirm('"' + n.title + '" 공지를 슬랙으로 알릴까요?')) return;
      nb.disabled = true;
      api('/api/notices/' + n.id + '/notify', { method: 'POST', body: {} })
        .then(function () { flashNext('슬랙으로 알렸어요.'); renderView(); })
        .catch(function (e) { flash(e.message, 'err'); nb.disabled = false; });
    });
    var db = document.getElementById('ntDel');
    if (db) db.addEventListener('click', function () {
      if (!window.confirm('"' + n.title + '" 글을 지울까요? 지우면 목록에서 사라져요.')) return;
      db.disabled = true;
      api('/api/notices/' + n.id + '/delete', { method: 'POST', body: {} })
        .then(function () { flashNext('글을 지웠어요.'); location.href = BASE; })
        .catch(function (e) { flash(e.message, 'err'); db.disabled = false; });
    });
  }

  // ── 쓰기·고치기(관리자) ─────────────────────────────────────────────
  function renderEdit() {
    var edit = MODE === 'edit';
    var jobs = [KIND === 'notice' ? api('/api/notices?kind=notice&limit=1') : Promise.resolve({ categories: [] })];
    if (edit) jobs.push(api('/api/notices/' + NID + '?peek=1'));
    Promise.all(jobs).then(function (res) {
      var cats = res[0].categories || [];
      var n = edit ? res[1].notice : { title: '', body: '', pinned: false, category_id: null, files: [], slack_sent: false };
      setTitle((edit ? '고치기 · ' : '새 글 · ') + LABEL);
      var catSel = '';
      if (KIND === 'notice') {
        catSel = '<div class="gw-nt-f-cat"><label class="form-label" for="ntCat">카테고리</label><select class="form-select" id="ntCat" name="category_id" required>' +
          (edit ? '' : '<option value="">고르세요</option>') +
          cats.map(function (c) {
            return '<option value="' + c.id + '"' + (c.id === n.category_id ? ' selected' : '') + ' data-aud="' + esc(c.audience_label) + '">' + esc(c.name) + '</option>';
          }).join('') + '</select><p class="gw-nt-hint" id="ntCatAud"></p></div>';
      }
      var h = '<a class="gw-nt-back" href="' + (edit ? BASE + '/' + NID : BASE) + '"><i class="fas fa-arrow-left" aria-hidden="true"></i> ' +
        (edit ? '글로 돌아가기' : esc(LABEL)) + '</a>' +
        '<form class="gw-panel gw-nt-form" id="ntForm" novalidate aria-labelledby="ntH1">' +
        '<h1 class="gw-nt-form-h" id="ntH1" tabindex="-1">' + (edit ? '글 고치기' : (KIND === 'devnote' ? '개발 노트 쓰기' : '공지 쓰기')) + '</h1>' +
        catSel +
        '<div><label class="form-label" for="ntTitle">제목</label><input class="form-control" id="ntTitle" name="title" maxlength="120" required autocomplete="off" value="' + esc(n.title) + '"></div>' +
        '<div><label class="form-label" for="ntBody">본문</label><textarea class="form-control gw-nt-ta" id="ntBody" name="body" rows="12">' + esc(n.body || '') + '</textarea>' +
          '<p class="gw-nt-hint">줄바꿈은 그대로 보여요. 주소(https://…)는 누르면 열리는 링크가 돼요.</p></div>' +
        '<fieldset class="gw-nt-f-files"><legend class="form-label">첨부 <span class="gw-nt-dim">한 파일 15MB, 10개까지</span></legend>' +
          '<ul class="gw-nt-flist" id="ntFiles">' + n.files.map(function (f) {
            return '<li data-fid="' + f.id + '"><i class="fas ' + fileIcon(f.name) + '" aria-hidden="true"></i><span>' + esc(f.name) + '</span>' +
              '<span class="gw-nt-size">' + size(f.size) + '</span><button type="button" class="btn btn-sm btn-link text-danger" data-frm="' + f.id + '">빼기</button></li>';
          }).join('') + '</ul>' +
          '<label class="btn btn-sm btn-outline-secondary gw-nt-pick"><i class="fas fa-paperclip" aria-hidden="true"></i> 파일 고르기' +
          '<input type="file" id="ntPick" multiple class="visually-hidden"></label></fieldset>' +
        '<div class="gw-nt-f-opts">' +
          '<label class="form-check"><input class="form-check-input" type="checkbox" id="ntPin" name="pinned"' + (n.pinned ? ' checked' : '') + '>' +
            '<span class="form-check-label">맨 위에 고정</span></label>' +
          (n.slack_sent ? '<span class="gw-nt-sent"><i class="fab fa-slack" aria-hidden="true"></i> 이미 슬랙으로 알린 글이에요</span>'
            : '<label class="form-check"><input class="form-check-input" type="checkbox" id="ntSlack" name="slack">' +
              '<span class="form-check-label">저장하면 슬랙으로도 알리기</span></label>') +
        '</div>' +
        '<div class="gw-lv-err" id="ntErr" role="alert"></div>' +
        '<div class="gw-nt-f-btns"><a class="btn btn-light" href="' + (edit ? BASE + '/' + NID : BASE) + '">취소</a>' +
          '<button type="submit" class="btn btn-primary" id="ntSave">' + (edit ? '저장' : '올리기') + '</button></div>' +
        '</form>';
      root.innerHTML = h;
      wireEdit(n, edit);
    }).catch(function (e) {
      root.innerHTML = '<div class="gw-empty gw-empty-box">' + esc(e.message) + '</div>';
    });
  }

  function wireEdit(n, edit) {
    var f = document.getElementById('ntForm');
    var err = document.getElementById('ntErr');
    var pick = document.getElementById('ntPick');
    var flist = document.getElementById('ntFiles');
    var newFiles = [];
    var removeIds = [];
    var dirty = false;
    f.addEventListener('input', function () { dirty = true; });
    window.addEventListener('beforeunload', function (e) { if (dirty) { e.preventDefault(); e.returnValue = ''; } });

    var cat = document.getElementById('ntCat');
    function showAud() {
      var o = cat && cat.options[cat.selectedIndex];
      var a = o && o.getAttribute('data-aud');
      document.getElementById('ntCatAud').textContent = a ? '볼 수 있는 사람: ' + a : '';
    }
    if (cat) { cat.addEventListener('change', showAud); showAud(); }

    var ta = document.getElementById('ntBody');
    function grow() { ta.style.height = 'auto'; ta.style.height = Math.min(Math.max(ta.scrollHeight + 2, 260), 900) + 'px'; }
    ta.addEventListener('input', grow); grow();

    pick.addEventListener('change', function () {
      err.textContent = '';
      Array.prototype.forEach.call(pick.files, function (file) {
        if (file.size > FILE_MAX) { err.textContent = file.name + ' — 15MB까지 올릴 수 있어요.'; return; }
        var i = newFiles.push(file) - 1;
        var li = document.createElement('li');
        li.className = 'new';
        li.innerHTML = '<i class="fas ' + fileIcon(file.name) + '" aria-hidden="true"></i><span>' + esc(file.name) + '</span>' +
          '<span class="gw-nt-size">' + size(file.size) + '</span><button type="button" class="btn btn-sm btn-link text-danger">빼기</button>';
        li.querySelector('button').addEventListener('click', function () { newFiles[i] = null; li.remove(); });
        flist.appendChild(li);
      });
      pick.value = '';
      dirty = true;
    });
    flist.addEventListener('click', function (e) {
      var b = e.target.closest('[data-frm]');
      if (!b) return;
      removeIds.push(Number(b.getAttribute('data-frm')));
      b.closest('li').remove();
      dirty = true;
    });

    document.getElementById(edit ? 'ntTitle' : (cat ? 'ntCat' : 'ntTitle')).focus();

    f.addEventListener('submit', function (e) {
      e.preventDefault();
      err.textContent = '';
      var title = f.elements.title.value.trim();
      if (cat && !cat.value) { err.textContent = '카테고리를 골라 주세요.'; cat.focus(); return; }
      if (!title) { err.textContent = '제목을 적어 주세요.'; f.elements.title.focus(); return; }
      var slack = document.getElementById('ntSlack');
      var wantSlack = slack && slack.checked;
      if (wantSlack && !window.confirm('저장한 뒤 슬랙으로도 알릴까요?')) return;
      var sb = document.getElementById('ntSave');
      sb.disabled = true;
      sb.textContent = '저장하는 중…';
      var body = { kind: KIND, title: title, body: f.elements.body.value, pinned: document.getElementById('ntPin').checked,
                   category_id: cat ? Number(cat.value) : null };
      var id = edit ? NID : 0;
      var notes = [];
      api('/api/notices' + (edit ? '/' + NID : ''), { method: 'POST', body: body })
        .then(function (j) {
          id = j.id || NID;
          dirty = false;
          // 빼기 → 올리기 → 알리기(앞 단계가 실패해도 글은 저장됨 — 남은 일을 안내)
          return removeIds.reduce(function (p, fid) {
            return p.then(function () { return api('/api/notices/files/' + fid + '/delete', { method: 'POST', body: {} }); });
          }, Promise.resolve());
        })
        .then(function () {
          var files = newFiles.filter(Boolean);
          if (!files.length) return null;
          var fd = new FormData();
          files.forEach(function (x) { fd.append('file', x, x.name); });
          return api('/api/notices/' + id + '/files', { method: 'POST', form: fd })
            .catch(function (x) { notes.push('첨부를 올리지 못했어요: ' + x.message); });
        })
        .then(function () {
          if (!wantSlack) return null;
          return api('/api/notices/' + id + '/notify', { method: 'POST', body: {} })
            .then(function () { notes.push('슬랙으로 알렸어요.'); })
            .catch(function (x) { notes.push(x.message); });
        })
        .then(function () {
          flashNext((edit ? '저장했어요.' : '글을 올렸어요.') + (notes.length ? ' ' + notes.join(' ') : ''));
          location.href = BASE + '/' + id;
        })
        .catch(function (x) {
          err.textContent = x.message;
          sb.disabled = false;
          sb.textContent = edit ? '저장' : '올리기';
          if (id && !edit) { NID = id; edit = true; }   // 글은 이미 만들어짐 — 다시 누르면 고치기로
        });
    });
  }

  if (MODE === 'view' && NID) renderView();
  else if ((MODE === 'new' || MODE === 'edit') && CAN_EDIT) renderEdit();
  else { renderList(); flashPending(); }
})();
