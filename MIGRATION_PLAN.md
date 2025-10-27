# ITG 대시보드 파이썬 이관 계획

## 현재 상황 분석

### 기존 시스템 강점
- Flask 기반 파이썬 애플리케이션 (3900+ 줄)
- 20개+ 전문 utils 모듈
- 완성도 높은 인증/보안 시스템
- Google Sheets API 연동 완료
- 고급 캐시 시스템 구축
- 감사 로그 시스템

### 주요 문제점
- 복잡한 JavaScript DOM 조작
- 캐시 동기화 문제
- 실시간 상태 관리 복잡성
- 클라이언트-서버 상태 불일치

## 이관 전략: 서버사이드 렌더링 강화

완전 재작성이 아닌 **현재 Flask 기반 유지 + 프론트엔드 서버사이드 전환**

## Phase 1: 서버사이드 렌더링 강화 (1-2일)

### 1.1 프로젝트 상태 관리 개선
```python
# app.py에 추가
@app.route('/projects')
@login_required
def project_list():
    projects_data = load_data()

    # 서버에서 모든 상태 계산
    for project in projects_data:
        project['is_cancelled'] = project.get('수금 관련 특이사항') == '공사취소'
        project['is_overdue'] = calculate_overdue_status(project)
        project['edit_permissions'] = get_edit_permissions(project, session['user'])
        project['lock_status'] = get_lock_status(project['프로젝트 코드'])

    return render_template('project_list_server.html',
                         projects=projects_data,
                         user_role=get_user_role(),
                         timestamp=datetime.now())
```

### 1.2 새로운 서버사이드 템플릿 생성
```html
<!-- templates/project_list_server.html -->
{% for project in projects %}
<tr class="{% if project.is_cancelled %}project-cancelled{% endif %}">
    <td>{{ project['프로젝트 코드'] }}</td>
    <td>
        {% if project.is_cancelled %}
            <span class="badge bg-danger">취소됨</span>
        {% endif %}
    </td>
    <td>
        <div class="project-details {% if project.is_cancelled %}cancelled-overlay{% endif %}">
            {% if project.is_cancelled %}
                <div class="cancelled-watermark">취소된 공사</div>
                <!-- 모든 입력 요소를 disabled로 렌더링 -->
                <input type="text" disabled class="form-control" value="{{ project['현장명'] }}">
            {% else %}
                <input type="text" class="form-control" value="{{ project['현장명'] }}">
            {% endif %}
        </div>
    </td>
</tr>
{% endfor %}
```

## Phase 2: JavaScript 최소화 (2-3일)

### 2.1 폼 제출을 서버사이드로 전환
```python
@app.route('/api/project/update/<project_code>', methods=['POST'])
@login_required
def update_project_server(project_code):
    # 서버에서 완전한 상태 검증
    project_data = get_project(project_code)

    if project_data.get('수금 관련 특이사항') == '공사취소':
        return jsonify({'error': '취소된 프로젝트는 수정할 수 없습니다'}), 400

    # 업데이트 처리
    result = update_project_data(project_code, request.form)

    # 성공 시 페이지 새로고침
    return redirect(url_for('project_list', success=True))
```

### 2.2 실시간 업데이트를 서버 폴링으로 대체
```python
@app.route('/api/projects/status')
@login_required
def get_projects_status():
    """페이지별 상태 확인 API"""
    projects_data = load_data()

    status = {
        'timestamp': datetime.now().isoformat(),
        'total_projects': len(projects_data),
        'cancelled_projects': [p['프로젝트 코드'] for p in projects_data
                             if p.get('수금 관련 특이사항') == '공사취소'],
        'locked_projects': get_all_locked_projects(),
        'last_update': get_last_sheet_update()
    }

    return jsonify(status)
```

## Phase 3: 캐시 시스템 개선 (1-2일)

### 3.1 서버사이드 캐시 일관성 보장
```python
from utils.smart_cache_manager import invalidate_related_cache

def update_project_with_cache_invalidation(project_code, data):
    # 업데이트 수행
    result = update_google_sheet(project_code, data)

    if result['success']:
        # 관련된 모든 캐시 무효화
        invalidate_related_cache([
            f"project_{project_code}",
            "projects_list",
            "summary_stats",
            "cancelled_projects"
        ])

    return result
```

### 3.2 페이지별 캐시 전략
```python
@app.route('/projects')
@cached(timeout=60, key_prefix='projects_page')
@login_required
def project_list_cached():
    # 캐시된 데이터로 페이지 렌더링
    projects_data = load_data()
    return render_template('project_list_server.html', projects=projects_data)
```

## Phase 4: 사용자 경험 개선 (1일)

### 4.1 진행 상황 표시
```python
@app.route('/api/operation/status/<operation_id>')
def get_operation_status(operation_id):
    """장시간 작업의 진행 상황 추적"""
    status = get_background_task_status(operation_id)
    return jsonify(status)
```

### 4.2 에러 처리 강화
```python
@app.errorhandler(Exception)
def handle_exception(e):
    # 서버에서 완전한 에러 처리
    log_error(e, request.url, session.get('user'))

    if request.path.startswith('/api/'):
        return jsonify({'error': '서버 오류가 발생했습니다'}), 500
    else:
        return render_template('error.html', error=str(e)), 500
```

## 이관 후 예상 이점

### 성능 개선
- ✅ 캐시 불일치 문제 해결
- ✅ 클라이언트-서버 상태 동기화 문제 해결
- ✅ JavaScript 복잡성 대폭 감소

### 안정성 향상
- ✅ 서버에서 완전한 상태 제어
- ✅ 데이터 일관성 보장
- ✅ 에러 처리 중앙화

### 개발 효율성
- ✅ Python 단일 언어로 개발
- ✅ 디버깅 용이성 증대
- ✅ 코드 유지보수성 향상

## 이관 일정

- **Day 1-2**: Phase 1 (서버사이드 렌더링)
- **Day 3-5**: Phase 2 (JavaScript 최소화)
- **Day 6-7**: Phase 3 (캐시 시스템 개선)
- **Day 8**: Phase 4 (UX 개선) + 테스트

**총 소요 기간: 약 8일**

## 위험 요소 및 대응책

### 위험 요소
1. 기존 데이터 손실 위험
2. 사용자 워크플로우 변경
3. 성능 일시적 저하 가능성

### 대응책
1. 데이터 백업 자동화
2. 점진적 배포 (단계별 테스트)
3. 롤백 계획 수립
4. 사용자 교육 자료 준비

## 다음 단계

1. ✅ 현재 구조 분석 완료
2. 🔄 Phase 1 구현 시작
3. ⏳ 백업 및 테스트 환경 구축
4. ⏳ 단계별 검증 진행