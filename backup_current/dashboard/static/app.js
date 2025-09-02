(function(){
  const apiKey = localStorage.getItem('apiKey') || prompt('API Key?');
  const userEmail = localStorage.getItem('userEmail') || prompt('당신의 이메일?');
  localStorage.setItem('apiKey', apiKey);
  localStorage.setItem('userEmail', userEmail);
  document.getElementById('who').textContent = userEmail;

  const thead = document.getElementById('thead');
  const tbody = document.getElementById('tbody');
  const monthEl = document.getElementById('month');
  const managerEl = document.getElementById('manager');
  const modeEl = document.getElementById('mode');

  function render(rows){
    tbody.innerHTML='';
    if(!rows || rows.length===0){ tbody.innerHTML = '<tr><td class="muted">데이터 없음</td></tr>'; return; }
    const cols = Object.keys(rows[0]);
    thead.innerHTML = '<tr>'+cols.map(c=>`<th>${c}</th>`).join('')+'</tr>';
    rows.forEach(r=>{
      const tr = document.createElement('tr');
      tr.innerHTML = cols.map(c=>`<td>${(r[c]||'').toString()}</td>`).join('');
      tbody.appendChild(tr);
    });
  }

  const socket = io({extraHeaders:{'X-API-Key': apiKey, 'X-User-Email': userEmail}});
  socket.on('connect', ()=>{
    socket.emit('auth', {apiKey, userEmail});
    socket.emit('projects:subscribe', {filters: {}});
  });
  socket.on('projects:update', payload=>{ render(payload.rows); });

  document.getElementById('refresh').addEventListener('click', async ()=>{
    const month = monthEl.value.trim();
    const manager = managerEl.value.trim();
    const mode = modeEl.value;
    const path = mode === 'settlement' ? '/api/settlement/projects' : '/api/projects';
    const url = new URL(location.origin + path);
    if(month) url.searchParams.set('month', month);
    if(manager) url.searchParams.set('manager', manager);
    const resp = await fetch(url.toString(), {headers:{'X-API-Key': apiKey, 'X-User-Email': userEmail}});
    const data = await resp.json();
    render(data.rows || []);
  });

  // ---- 신규 프로젝트 등록 폼 ----
  const companySel = document.getElementById('f_company');
  const ownerSel = document.getElementById('f_owner');
  const addrEl = document.getElementById('f_addr');
  const noteEl = document.getElementById('f_note');
  const btnCreate = document.getElementById('btnCreate');
  const msgErr = document.getElementById('createMsg');
  const msgOk = document.getElementById('createOk');

  function setErr(msg){
    msgErr.style.display = 'block';
    msgErr.textContent = msg;
    msgOk.style.display = 'none';
  }
  function setOk(msg){
    msgOk.style.display = 'block';
    msgOk.innerHTML = msg;
    msgErr.style.display = 'none';
  }

  async function loadOptions(){
    try {
      const resp = await fetch('/api/meta/options', {headers:{'X-API-Key': apiKey, 'X-User-Email': userEmail}});
      const data = await resp.json();
      const companies = data.companies || [];
      const owners = data.owners || [];

      companySel.innerHTML = '<option value="">선택하세요</option>' + companies.map(c=>`<option value="${c}">${c}</option>`).join('');
      ownerSel.innerHTML = '<option value="">선택하세요</option>' + owners.map(o=>`<option value="${o}">${o}</option>`).join('');
    } catch (e){
      setErr('옵션 로드 실패: ' + e.message);
    }
  }
  loadOptions();

  btnCreate.addEventListener('click', async ()=>{
    const company = companySel.value.trim();
    const owner = ownerSel.value.trim();
    const addr = addrEl.value.trim();

    // 간단 검증
    if(!company){ setErr('사업자를 선택하세요.'); companySel.focus(); return; }
    if(!owner){ setErr('담당자를 선택하세요.'); ownerSel.focus(); return; }
    if(!addr){ setErr('현장 주소를 입력하세요.'); addrEl.focus(); return; }

    setOk('등록 중...');
    try {
      const resp = await fetch('/api/projects/auto', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-API-Key': apiKey,
          'X-User-Email': userEmail
        },
        body: JSON.stringify({
          '사업자': company,
          '담당자': owner,
          '현장 주소': addr,
          ...(noteEl.value.trim()? {'비고': noteEl.value.trim()}: {})
        })
      });
      const data = await resp.json();
      if(data.ok){
        setOk(`✅ 등록 완료 — 프로젝트 코드: <span class="pill">${data.project_code}</span>`);
        document.getElementById('refresh').click();
        companySel.value=''; ownerSel.value=''; addrEl.value=''; noteEl.value='';
      }else{
        setErr('실패: ' + (data.error || 'Unknown error'));
      }
    } catch (e){
      setErr('실패: ' + e.message);
    }
  });
})();