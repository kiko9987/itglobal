# ProjectRowAccordion.js 리팩토링 계획서

**작성일**: 2025-01-15
**우선순위**: 높음 (단기 1-2개월 내 실행 권장)
**담당**: 개발팀

---

## 1. 현재 상황 분석

### 문제점
- **파일 크기**: 5,645줄 (유지보수 불가능 수준)
- **단일 책임 원칙 위반**: 하나의 클래스가 모든 카드 렌더링/이벤트/검증 담당
- **코드 이해 난이도**: 신규 팀원이 전체 로직을 파악하는데 수일 소요
- **버그 위험도**: 한 카드 수정 시 다른 카드에 영향 가능성 높음
- **병렬 개발 불가**: 여러 명이 동시에 작업 시 conflict 필연적

### 현재 구조 (추정)
```
ProjectRowAccordion.js (5,645줄)
├── 아코디언 헤더 렌더링 (200줄)
├── 기본 정보 카드 (800줄)
├── 공사 정보 카드 (700줄)
├── 금액 정보 카드 (900줄)
├── 수금 정보 카드 (1,000줄)
├── 담당자/시공자 카드 (600줄)
├── 메모 카드 (400줄)
├── 공통 이벤트 핸들러 (500줄)
├── 데이터 검증 로직 (300줄)
└── 기타 유틸리티 (245줄)
```

---

## 2. 개선 목표

### 주요 목표
1. **단일 책임 원칙 적용**: 각 카드가 독립된 컴포넌트로 분리
2. **코드 가독성 향상**: 각 파일 500줄 이하로 유지
3. **유지보수성 개선**: 특정 카드 수정 시 다른 카드에 영향 없음
4. **병렬 개발 가능**: 여러 팀원이 동시에 다른 카드 작업 가능
5. **테스트 용이성**: 각 카드를 독립적으로 테스트 가능

### 성공 기준
- 전체 코드 라인 수는 증가할 수 있지만, **파일당 평균 400줄 이하**
- 카드 추가/수정 시 **영향 범위를 단일 파일로 제한**
- 신규 팀원이 **특정 카드 로직을 1시간 내 이해 가능**

---

## 3. 제안하는 새 구조

### 디렉토리 구조
```
dashboard/src/js/components/
├── ProjectRowAccordion.js (200줄) ← 오케스트레이터 역할만
└── accordion-cards/
    ├── BaseCard.js (150줄) ← 공통 부모 클래스
    ├── AccordionHeader.js (150줄)
    ├── BasicInfoCard.js (400줄)
    ├── ConstructionCard.js (400줄)
    ├── FinancialCard.js (500줄)
    ├── ReceivableCard.js (600줄)
    ├── ContactCard.js (400줄)
    ├── MemoCard.js (350줄)
    └── CardValidator.js (200줄) ← 검증 로직 분리
```

### 각 컴포넌트의 책임

#### ProjectRowAccordion.js (메인 오케스트레이터)
```javascript
/**
 * 역할: 아코디언 전체 관리 (렌더링 X, 조율만 O)
 * - 카드 컴포넌트 인스턴스 생성 및 관리
 * - 아코디언 열기/닫기 애니메이션
 * - 카드 간 통신 중계
 * - 전역 이벤트 바인딩 (열기/닫기 버튼)
 */
class ProjectRowAccordion {
  constructor() {
    this.cards = {
      header: new AccordionHeader(),
      basicInfo: new BasicInfoCard(),
      construction: new ConstructionCard(),
      financial: new FinancialCard(),
      receivable: new ReceivableCard(),
      contact: new ContactCard(),
      memo: new MemoCard()
    };
  }

  async openAccordion(tableRow, projectData) {
    // 1. 검증
    if (!CardValidator.validateProjectData(projectData)) return;

    // 2. 각 카드에 데이터 전달 및 렌더링 요청
    const headerHtml = await this.cards.header.render(projectData);
    const basicInfoHtml = await this.cards.basicInfo.render(projectData);
    // ... 나머지 카드들

    // 3. DOM에 삽입 및 애니메이션
    this.insertCardsIntoDom([headerHtml, basicInfoHtml, ...]);
  }
}
```

#### BaseCard.js (공통 부모 클래스)
```javascript
/**
 * 역할: 모든 카드의 공통 기능 제공
 * - 기본 렌더링 템플릿
 * - 공통 이벤트 바인딩 (저장 버튼, 취소 버튼)
 * - 로딩 상태 관리
 * - 에러 메시지 표시
 */
class BaseCard {
  constructor(cardType) {
    this.cardType = cardType;
    this.isEditing = false;
    this.originalData = null;
  }

  // 자식 클래스가 구현해야 하는 메서드
  async render(projectData) {
    throw new Error('render() must be implemented');
  }

  // 공통 메서드
  showLoading() { /* ... */ }
  hideLoading() { /* ... */ }
  showError(message) { /* ... */ }
  enableEditMode() { /* ... */ }
  disableEditMode() { /* ... */ }
}
```

#### BasicInfoCard.js (기본 정보 카드)
```javascript
/**
 * 역할: 기본 정보 카드 렌더링 및 편집
 * - 프로젝트 코드, 사업자, 담당자, 거래처 표시
 * - 인라인 편집 기능
 * - 프로젝트 코드 변경 시 특수 로직 처리
 */
class BasicInfoCard extends BaseCard {
  constructor() {
    super('basic-info');
  }

  async render(projectData) {
    // 기본 정보만 렌더링
    return `
      <div class="card basic-info-card">
        <div class="card-header">기본 정보</div>
        <div class="card-body">
          ${this.renderProjectCode(projectData['프로젝트 코드'])}
          ${this.renderCompany(projectData['사업자'])}
          ${this.renderOwner(projectData['담당자'])}
          ${this.renderClient(projectData['거래처'])}
        </div>
      </div>
    `;
  }

  renderProjectCode(code) { /* ... */ }
  renderCompany(company) { /* ... */ }
  // ...
}
```

#### FinancialCard.js (금액 정보 카드)
```javascript
/**
 * 역할: 금액 정보 카드 렌더링 및 계산
 * - 총액1, 총액2, 부가세 표시
 * - 금액 계산 로직 (AmountCalculator 사용)
 * - 콤마 포맷팅
 * - 빠른 금액 버튼 기능
 */
class FinancialCard extends BaseCard {
  constructor() {
    super('financial');
    this.calculator = new AmountCalculator();
  }

  async render(projectData) {
    const amount1 = projectData['총액 1'];
    const vatIncluded = projectData['부가세'] === 'TRUE';
    const amount2 = this.calculator.calculate(amount1, vatIncluded);

    return `
      <div class="card financial-card">
        <!-- 금액 정보만 표시 -->
      </div>
    `;
  }
}
```

#### ReceivableCard.js (수금 정보 카드)
```javascript
/**
 * 역할: 수금 정보 카드 렌더링 및 관리
 * - 수금 회차별 금액 표시
 * - 수금 진행률 계산 및 프로그레스 바
 * - 수금 모드 전환 (display/edit)
 * - 메모 툴팁 표시
 */
class ReceivableCard extends BaseCard {
  constructor() {
    super('receivable');
    this.isReceivableMode = false;
  }

  async render(projectData) {
    const receivables = this.parseReceivables(projectData);
    const progress = this.calculateProgress(receivables);

    return `
      <div class="card receivable-card">
        <!-- 수금 정보 표시 -->
      </div>
    `;
  }

  toggleReceivableMode() { /* ... */ }
  calculateProgress(receivables) { /* ... */ }
}
```

#### CardValidator.js (검증 로직)
```javascript
/**
 * 역할: 모든 카드의 데이터 검증
 * - 프로젝트 데이터 무결성 확인
 * - 필수 필드 존재 여부 검증
 * - 데이터 타입 검증
 */
class CardValidator {
  static validateProjectData(projectData) {
    const requiredFields = ['프로젝트 코드', '담당자', '거래처'];
    return requiredFields.every(field => projectData.hasOwnProperty(field));
  }

  static validateBasicInfo(data) { /* ... */ }
  static validateFinancial(data) { /* ... */ }
  static validateReceivable(data) { /* ... */ }
}
```

---

## 4. 마이그레이션 전략

### Phase 1: 준비 단계 (1주)
1. 새 디렉토리 구조 생성 (`accordion-cards/`)
2. BaseCard.js 작성 (공통 기능 정의)
3. CardValidator.js 작성 (검증 로직 분리)
4. 기존 ProjectRowAccordion.js 백업

### Phase 2: 카드별 분리 (2-3주)
**우선순위 순서** (간단한 것부터):
1. **AccordionHeader.js** (가장 단순, 헤더만)
2. **MemoCard.js** (비교적 단순, 메모만)
3. **ContactCard.js** (중간, 담당자/시공자)
4. **BasicInfoCard.js** (중요, 기본 정보)
5. **ConstructionCard.js** (중간, 공사 정보)
6. **FinancialCard.js** (복잡, 금액 계산 로직)
7. **ReceivableCard.js** (가장 복잡, 수금 모드 전환)

**각 카드 분리 작업 프로세스**:
```
1. 기존 코드에서 해당 카드 관련 메서드 식별
2. 새 파일에 BaseCard 상속 클래스 생성
3. 메서드 복사 및 정리
4. import/export 추가
5. 기존 파일에서 새 클래스 호출로 대체
6. 테스트 (아코디언 열기/닫기, 편집, 저장)
7. 기존 코드 삭제
```

### Phase 3: 통합 및 테스트 (1주)
1. ProjectRowAccordion.js를 오케스트레이터로 리팩토링
2. 모든 카드가 정상 작동하는지 통합 테스트
3. 성능 테스트 (렌더링 속도 비교)
4. 코드 리뷰 및 최종 정리

### Phase 4: 문서화 및 배포 (1주)
1. 각 카드 컴포넌트 JSDoc 주석 추가
2. 개발 가이드 작성 (새 카드 추가 방법)
3. 스테이징 환경 배포 및 QA
4. 프로덕션 배포

**총 소요 시간**: 5-6주

---

## 5. 예상 효과

### 정량적 효과
- **파일당 평균 라인 수**: 5,645줄 → 평균 350줄 (93% 감소)
- **코드 이해 시간**: 신규 개발자 기준 8시간 → 1시간 (87% 감소)
- **버그 수정 시간**: 평균 2시간 → 30분 (75% 감소)
- **병렬 개발 가능 인원**: 1명 → 4명 동시 작업 가능

### 정성적 효과
- **유지보수성**: 특정 카드 수정이 다른 카드에 영향 없음
- **테스트 용이성**: 각 카드를 독립적으로 유닛 테스트 가능
- **확장성**: 새 카드 추가 시 기존 코드 수정 불필요
- **코드 품질**: 각 파일이 명확한 책임을 가져 리뷰 용이
- **팀 협업**: conflict 최소화, 작업 분배 명확

---

## 6. 리스크 및 대응 방안

### 리스크 1: 리팩토링 중 기능 장애
**대응**:
- 기존 파일을 백업 폴더에 보관
- 카드 단위로 점진적 마이그레이션
- 각 카드 완료 시마다 철저한 테스트

### 리스크 2: 성능 저하 가능성
**대응**:
- 렌더링 성능 측정 도구 사용
- 필요 시 lazy loading 적용
- 번들 크기 모니터링 (Vite 빌드 분석)

### 리스크 3: 팀원 러닝 커브
**대응**:
- 새 구조 설명 세션 진행 (1시간)
- 간단한 예제 카드 작성 가이드 제공
- 코드 리뷰 시 새 패턴 적극 안내

### 리스크 4: 예상보다 긴 개발 기간
**대응**:
- 우선순위가 낮은 카드는 Phase 2 이후로 연기
- 핵심 카드(Basic, Financial, Receivable)만 먼저 분리
- 나머지는 점진적으로 진행

---

## 7. 체크리스트

### 시작 전 확인사항
- [ ] 팀 회의를 통해 리팩토링 일정 합의
- [ ] 기존 코드 백업 완료
- [ ] 개발/스테이징 환경 준비
- [ ] 롤백 계획 수립

### Phase 1 완료 체크리스트
- [ ] `accordion-cards/` 디렉토리 생성
- [ ] BaseCard.js 작성 및 테스트
- [ ] CardValidator.js 작성 및 테스트
- [ ] 기존 파일 백업

### Phase 2 완료 체크리스트
- [ ] AccordionHeader.js 분리 완료
- [ ] MemoCard.js 분리 완료
- [ ] ContactCard.js 분리 완료
- [ ] BasicInfoCard.js 분리 완료
- [ ] ConstructionCard.js 분리 완료
- [ ] FinancialCard.js 분리 완료
- [ ] ReceivableCard.js 분리 완료

### Phase 3 완료 체크리스트
- [ ] ProjectRowAccordion.js 오케스트레이터로 리팩토링
- [ ] 모든 카드 정상 작동 확인
- [ ] 편집 기능 정상 작동 확인
- [ ] 저장 기능 정상 작동 확인
- [ ] 성능 테스트 통과

### Phase 4 완료 체크리스트
- [ ] JSDoc 주석 작성 완료
- [ ] 개발 가이드 문서 작성
- [ ] 스테이징 배포 및 QA 완료
- [ ] 프로덕션 배포 완료

---

## 8. 참고 자료

### 현재 파일
- `dashboard/src/js/components/ProjectRowAccordion.js` (5,645줄)

### 관련 유틸리티
- `dashboard/src/js/utils/AmountCalculator.js` (금액 계산)
- `dashboard/src/js/utils/FormValidator.js` (폼 검증)
- `dashboard/src/js/components/InlineEditor.js` (인라인 편집)

### 참고 패턴
- **디자인 패턴**: Template Method Pattern (BaseCard 상속)
- **아키텍처**: Component-based Architecture
- **원칙**: Single Responsibility Principle (SRP)

---

## 9. 추가 개선 사항 (Phase 4 이후)

리팩토링 완료 후 고려할 추가 개선사항:

1. **카드 Lazy Loading**:
   - 아코디언 열 때 모든 카드를 한 번에 렌더링하지 않고, 보이는 카드부터 로드

2. **카드 상태 관리 개선**:
   - 각 카드의 편집 상태를 중앙에서 관리
   - 한 카드 편집 중일 때 다른 카드 편집 방지

3. **카드 간 데이터 동기화**:
   - 한 카드에서 데이터 변경 시 관련 카드 자동 업데이트
   - 예: 금액 변경 시 수금 진행률 자동 재계산

4. **TypeScript 마이그레이션**:
   - 타입 안정성 확보
   - IDE 자동완성 지원

---

**작성자**: Claude Code
**최종 수정**: 2025-01-15
