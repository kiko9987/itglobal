(function(){
  // 보안 개선: sessionStorage 사용 및 토큰 기반 인증
  let accessToken = sessionStorage.getItem('accessToken');
  let userEmail = sessionStorage.getItem('userEmail');
  
  // 토큰이 없으면 로그인 페이지로 리다이렉트
  if (!accessToken || !userEmail) {
    window.location.href = '/login';
    return;
  }
  
  // 사용자 정보 표시 (XSS 방지)
  const whoElement = document.getElementById('who');
  if (whoElement) {
    whoElement.textContent = userEmail;
  }

  const thead = document.getElementById('thead');
  const tbody = document.getElementById('tbody');
  const monthEl = document.getElementById('month');
  const managerEl = document.getElementById('manager');
  const modeEl = document.getElementById('mode');

  // XSS 방지를 위한 안전한 렌더링 함수
  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }
  
  function render(rows){
    tbody.innerHTML='';
    if(!rows || rows.length===0){ 
      tbody.innerHTML = '<tr><td class="muted">데이터 없음</td></tr>'; 
      return; 
    }
    
    const cols = Object.keys(rows[0]);
    
    // 안전한 헤더 생성
    thead.innerHTML = '<tr>' + cols.map(c => `<th>${escapeHtml(c)}</th>`).join('') + '</tr>';
    
    // 안전한 행 생성
    rows.forEach(r=>{
      const tr = document.createElement('tr');
      cols.forEach(c => {
        const td = document.createElement('td');
        td.textContent = (r[c] || '').toString();
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  // 새로운 프로젝트를 테이블에 부드럽게 추가하는 함수
  async function addNewProjectToTable(projectCode) {
    try {
      // 새로 생성된 프로젝트 데이터만 가져오기
      const project = await secureApiCall(`/api/projects/${projectCode}`);
      if (project && !project.error) {
        
        // 기존 테이블이 비어있었다면 헤더부터 다시 생성
        if (tbody.innerHTML.includes('데이터 없음')) {
          const cols = Object.keys(project);
          thead.innerHTML = '<tr>' + cols.map(c => `<th>${escapeHtml(c)}</th>`).join('') + '</tr>';
          tbody.innerHTML = '';
        }
        
        // 새 행 생성
        const tr = document.createElement('tr');
        tr.style.backgroundColor = '#e8f5e8'; // 새 항목 하이라이트
        tr.style.transition = 'background-color 3s ease';
        
        const cols = Object.keys(project);
        cols.forEach(c => {
          const td = document.createElement('td');
          td.textContent = (project[c] || '').toString();
          tr.appendChild(td);
        });
        
        // 테이블 맨 위에 새 행 추가 (최신 항목이 위에 오도록)
        tbody.insertBefore(tr, tbody.firstChild);
        
        // 3초 후 하이라이트 제거
        setTimeout(() => {
          tr.style.backgroundColor = '';
        }, 3000);
        
        return true;
      }
    } catch (error) {
      console.error('새 프로젝트 데이터 로드 실패:', error);
      // 실패시 전체 테이블 새로고침으로 폴백
      document.getElementById('refresh').click();
    }
    return false;
  }

  // 안전한 API 호출 헬퍼 함수
  async function secureApiCall(url, options = {}) {
    const headers = {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers
    };
    
    try {
      const response = await fetch(url, {
        ...options,
        headers
      });
      
      if (response.status === 401) {
        // 토큰 만료 시 로그인 페이지로 리다이렉트
        sessionStorage.clear();
        window.location.href = '/login';
        return null;
      }
      
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
      
      return await response.json();
    } catch (error) {
      console.error('API 호출 오류:', error);
      throw error;
    }
  }

  const socket = io({extraHeaders:{'Authorization': `Bearer ${accessToken}`}});
  socket.on('connect', ()=>{
    socket.emit('auth', {token: accessToken});
    socket.emit('projects:subscribe', {filters: {}});
  });
  socket.on('projects:update', payload=>{ render(payload.rows); });

  document.getElementById('refresh')?.addEventListener('click', async ()=>{
    try {
      const month = monthEl?.value.trim();
      const manager = managerEl?.value.trim();
      const mode = modeEl?.value;
      const path = mode === 'settlement' ? '/api/settlement/projects' : '/api/projects';
      const url = new URL(location.origin + path);
      if(month) url.searchParams.set('month', month);
      if(manager) url.searchParams.set('manager', manager);
      
      const data = await secureApiCall(url.toString());
      if (data) {
        render(data.rows || []);
      }
    } catch (error) {
      setErr('데이터 로드 실패: ' + error.message);
    }
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
      const data = await secureApiCall('/api/meta/options');
      if (data) {
        const companies = data.companies || [];
        const owners = data.owners || [];

        companySel.innerHTML = '<option value="">선택하세요</option>' + companies.map(c=>`<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
        ownerSel.innerHTML = '<option value="">선택하세요</option>' + owners.map(o=>`<option value="${escapeHtml(o)}">${escapeHtml(o)}</option>`).join('');
      }
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
      const data = await secureApiCall('/api/projects/auto', {
        method: 'POST',
        body: JSON.stringify({
          '사업자': company,
          '담당자': owner,
          '현장 주소': addr,
          ...(noteEl.value.trim()? {'비고': noteEl.value.trim()}: {})
        })
      });
      
      if(data && data.ok){
        setOk(`✅ 등록 완료 — 프로젝트 코드: <span class="pill">${escapeHtml(data.project_code)}</span>`);
        
        // 부드러운 테이블 업데이트 (페이지 새로고침 없음)
        setTimeout(async () => {
          const success = await addNewProjectToTable(data.project_code);
          if (!success) {
            // 실패시에만 전체 새로고침
            console.log('부드러운 업데이트 실패, 전체 새로고침 실행');
          }
        }, 800); // 구글 시트 업데이트 완료를 위한 약간의 지연
        
        companySel.value=''; ownerSel.value=''; addrEl.value=''; noteEl.value='';
      }else{
        setErr('실패: ' + (data?.error || 'Unknown error'));
      }
    } catch (e){
      setErr('실패: ' + e.message);
    }
  });
})();