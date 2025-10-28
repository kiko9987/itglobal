# Manual QA 테스트 가이드

## 🎯 목적
Redis 전환 후 실제 환경에서 서비스가 정상 동작하는지 수동으로 검증

## ⚠️ 사전 준비

### 1. Redis 서버 실행 확인
```bash
# Docker 사용 시
docker ps | findstr redis

# Python 연결 확인
python -c "import redis; r = redis.Redis(host='localhost', port=6379); print(r.ping())"
# 출력: True
```

### 2. 서비스 시작
```bash
# 개발 모드
python run.py

# 또는 Gunicorn (프로덕션 모드)
gunicorn -c gunicorn.conf.py dashboard.app:app
```

### 3. 로그 확인 준비
```bash
# 별도 터미널에서 로그 모니터링
tail -f logs/app.log

# 또는 Windows PowerShell
Get-Content logs/app.log -Wait
```

---

## 📋 테스트 시나리오

### ✅ 시나리오 1: 프로젝트 데이터 캐싱 검증

**목적**: DataFrame이 pickle로 정상 직렬화/역직렬화되는지 확인

#### 1-1. 첫 로드 (API 호출)
```
1. 브라우저에서 http://localhost:5000 접속
2. 프로젝트 목록 페이지 로드
3. 로그 확인:
   [예상 로그]
   - "DataFrame 감지 - pickle 직렬화: current_sheet_data (shape: (N, M))"
   - "캐시 저장 (pickle): current_sheet_data"
```

#### 1-2. 캐시 히트 (두 번째 로드)
```
1. F5로 페이지 새로고침
2. 로그 확인:
   [예상 로그]
   - "캐시 히트 (pickle): current_sheet_data"
   - Google Sheets API 호출 없음
```

#### 1-3. 캐시 무효화 후 재로드
```
1. Redis CLI에서 캐시 삭제
   docker exec -it redis-claude-project redis-cli
   > KEYS cache:*
   > DEL cache:current_sheet_data
   > exit

2. 페이지 새로고침
3. 로그 확인:
   [예상 로그]
   - "캐시 미스: current_sheet_data"
   - Google Sheets API 호출
   - "DataFrame 감지 - pickle 직렬화"
```

#### ✅ 합격 기준
- [ ] 첫 로드 시 DataFrame pickle 직렬화 로그 확인
- [ ] 두 번째 로드 시 캐시 히트 로그 확인
- [ ] 데이터 정상 표시 (UI 깨짐 없음)
- [ ] 에러 로그 없음

---

### ✅ 시나리오 2: 동시 편집 락 검증 (다중 탭)

**목적**: Redis 분산 락이 다중 탭 환경에서 정상 동작하는지 확인

#### 2-1. 같은 사용자, 다른 탭
```
1. 브라우저 탭1: 프로젝트 "TEST001" 편집 모드 진입
   [예상 로그]
   - "새 잠금 생성: TEST001 by user@email.com"

2. 브라우저 탭2 (새 탭 열기): 같은 프로젝트 편집 시도
   [예상 결과]
   - UI 경고: "다른 탭에서 이미 편집 중입니다. 해당 탭으로 이동하세요."

   [예상 로그]
   - "잠금 획득 실패: 같은 사용자지만 다른 탭"

3. 탭1에서 편집 완료 후 저장/취소
   [예상 로그]
   - "잠금 해제: TEST001 by user@email.com"

4. 탭2에서 다시 편집 시도
   [예상 결과]
   - 정상 진입 가능
```

#### 2-2. 다른 사용자 (시뮬레이션)
```
1. 브라우저1 (사용자A): 프로젝트 편집 진입
2. 브라우저2 (시크릿 모드, 사용자B): 같은 프로젝트 편집 시도
   [예상 결과]
   - UI 경고: "사용자A님이 편집 중입니다."
   - 남은 시간 표시
```

#### ✅ 합격 기준
- [ ] 같은 사용자 다른 탭 차단 확인
- [ ] 다른 사용자 차단 확인
- [ ] 락 해제 후 정상 진입 확인
- [ ] 락 정보 (사용자명, 남은 시간) 정상 표시

---

### ✅ 시나리오 3: 락 TTL 자동 만료 및 연장

**목적**: 락 자동 만료 및 하트비트 연장이 정상 동작하는지 확인

#### 3-1. 락 자동 만료 (5분)
```
1. 프로젝트 편집 진입
2. 5분 동안 아무 작업 안 함
3. 다른 사용자가 편집 시도
   [예상 결과]
   - 정상 진입 가능 (락 자동 만료됨)

4. Redis에서 확인:
   docker exec -it redis-claude-project redis-cli
   > TTL lock:TEST001
   (integer) -2  <- 키 없음 (만료됨)
```

#### 3-2. 락 자동 연장 (하트비트)
```
1. 프로젝트 편집 진입
2. 편집 화면에서 계속 작업 중 (1-2분마다 서버에 요청)
3. 5분 경과
4. Redis에서 확인:
   > TTL lock:TEST001
   (integer) 250  <- 여전히 TTL 남아있음 (자동 연장됨)

5. 로그 확인:
   [예상 로그]
   - "잠금 연장: TEST001 by user@email.com"
```

#### ✅ 합격 기준
- [ ] 5분 방치 시 락 자동 만료 확인
- [ ] 활동 중 락 자동 연장 확인
- [ ] TTL 정상 갱신 확인

---

### ✅ 시나리오 4: Grace Period 복구

**목적**: 네트워크 지연으로 락이 만료되어도 30초 내 복구 가능한지 확인

#### 4-1. Grace Period 내 복구
```
1. 프로젝트 편집 진입
2. Redis CLI에서 락 강제 만료:
   > DEL lock:TEST001
   (integer) 1

3. Grace Period 마커 확인:
   > GET grace:TEST001
   "{\"user_email\":\"user@test.com\",\"tab_id\":\"tab123\",...}"

4. 30초 이내에 편집 화면에서 저장 버튼 클릭
   [예상 결과]
   - 정상 저장 성공

   [예상 로그]
   - "Grace Period 내 잠금 복구: TEST001"
```

#### 4-2. Grace Period 만료
```
1. 프로젝트 편집 진입
2. Redis CLI에서 락 강제 만료
3. 30초 대기
4. 저장 버튼 클릭
   [예상 결과]
   - 저장 실패
   - UI 경고: "편집 권한이 만료되었습니다"
```

#### ✅ 합격 기준
- [ ] Grace Period 내 복구 성공
- [ ] Grace Period 만료 시 실패
- [ ] 적절한 에러 메시지 표시

---

### ✅ 시나리오 5: Redis 재시작 시 Fail Fast

**목적**: Redis 장애 시 적절한 에러 처리 확인

#### 5-1. 서비스 부팅 시 Redis 없음
```
1. Redis 중지:
   docker stop redis-claude-project

2. 서비스 시작 시도:
   python run.py

   [예상 결과]
   - 서비스 시작 실패
   - 로그: "Redis 연결 실패: ..."
   - 로그: "서비스를 시작할 수 없습니다. Redis 서버를 확인하세요."
   - 프로세스 종료 (sys.exit(1))
```

#### 5-2. 서비스 실행 중 Redis 장애
```
1. 서비스 정상 실행 중
2. Redis 중지:
   docker stop redis-claude-project

3. 브라우저에서 프로젝트 목록 조회
   [예상 결과]
   - HTTP 503 Service Unavailable
   - JSON 응답: {"error": "Service temporarily unavailable"}

4. 로그 확인:
   [예상 로그]
   - "Redis 장애로 캐시 조회 실패: current_sheet_data"
   - "ServiceUnavailable: Cache unavailable: ..."
```

#### 5-3. Redis 복구 후 정상화
```
1. Redis 재시작:
   docker start redis-claude-project

2. 브라우저에서 프로젝트 목록 조회
   [예상 결과]
   - 정상 동작 (캐시 미스 → API 호출 → 캐시 저장)
```

#### ✅ 합격 기준
- [ ] 부팅 시 Redis 없으면 서비스 시작 실패
- [ ] 실행 중 Redis 장애 시 503 반환
- [ ] 명확한 에러 메시지 표시
- [ ] Redis 복구 후 정상화

---

### ✅ 시나리오 6: Optimistic Lock + Redis Lock 조합

**목적**: 버전 충돌 감지와 Redis 락이 함께 정상 동작하는지 확인

#### 6-1. 정상 저장 (버전 일치)
```
1. 프로젝트 편집 진입
2. 데이터 수정
3. 저장 버튼 클릭
   [예상 결과]
   - 저장 성공
   - 버전 증가 (_version += 1)

4. 로그 확인:
   [예상 로그]
   - "Optimistic Lock: 버전 일치 (v=N)"
   - "데이터 저장 성공"
```

#### 6-2. 버전 충돌 감지
```
1. 탭1: 프로젝트 편집 진입 (version=10)
2. 탭2 (시크릿): 같은 프로젝트 조회 (version=10)
3. 탭1: 저장 (version=10 → 11로 증가)
4. 탭2: Redis CLI에서 락 강제 해제 후 저장 시도

   [예상 결과]
   - 저장 실패
   - UI 경고: "다른 사용자가 먼저 수정했습니다. 새로고침 후 다시 시도하세요."

5. 로그 확인:
   [예상 로그]
   - "Optimistic Lock: 버전 불일치 (기대=10, 실제=11)"
```

#### ✅ 합격 기준
- [ ] 정상 저장 시 버전 증가 확인
- [ ] 버전 충돌 감지 확인
- [ ] 적절한 에러 메시지 표시
- [ ] 데이터 무결성 보장

---

### ✅ 시나리오 7: 대량 데이터 캐싱 (성능 테스트)

**목적**: 큰 DataFrame도 정상 처리되는지 확인

#### 7-1. 대용량 프로젝트 데이터
```
1. Google Sheets에 100개 이상 프로젝트 입력
2. 프로젝트 목록 조회
3. 로그 확인:
   [예상 로그]
   - "DataFrame 감지 - pickle 직렬화: current_sheet_data (shape: (100+, 30))"
   - 직렬화 시간 < 1초

4. Redis에서 캐시 크기 확인:
   > STRLEN cache:current_sheet_data
   (integer) 50000+  <- bytes 크기
```

#### 7-2. 캐시 히트 성능
```
1. 페이지 새로고침 10회
2. 응답 시간 측정
   [예상 결과]
   - 첫 로드: 2-3초 (Google Sheets API)
   - 이후: < 100ms (캐시)

3. 개발자 도구 네트워크 탭 확인:
   - Status: 200
   - Time: < 100ms
```

#### ✅ 합격 기준
- [ ] 100개 이상 프로젝트 정상 처리
- [ ] pickle 직렬화 시간 < 1초
- [ ] 캐시 히트 응답 시간 < 100ms
- [ ] 메모리 오버플로우 없음

---

### ✅ 시나리오 8: 캐시 무효화 전파

**목적**: 데이터 변경 시 캐시가 정상 무효화되는지 확인

#### 8-1. 프로젝트 수정 후 무효화
```
1. 프로젝트 "TEST001" 수정 및 저장
2. 로그 확인:
   [예상 로그]
   - "무효화 마커 설정: current_sheet_data"
   - "캐시 삭제 + 무효화 마커 설정: current_sheet_data"

3. 프로젝트 목록 페이지로 이동
4. 로그 확인:
   [예상 로그]
   - "캐시 미스: current_sheet_data"
   - Google Sheets API 재호출
   - 변경된 데이터 반영 확인
```

#### 8-2. 레이스 컨디션 방지 (무효화 마커)
```
1. 탭1: 프로젝트 목록 조회 시작 (API 호출 중)
2. 탭2: 프로젝트 수정 및 저장 (캐시 무효화)
3. 탭1: API 응답 도착하여 캐시 쓰기 시도

   [예상 결과]
   - 탭1의 오래된 데이터는 캐시에 저장되지 않음

4. 로그 확인:
   [예상 로그]
   - "캐시 쓰기 거부: current_sheet_data - 데이터 수집이 무효화 마커보다 이전"

5. 탭3: 프로젝트 목록 조회
   [예상 결과]
   - 탭2에서 수정한 최신 데이터 표시 (탭1의 오래된 데이터 아님)
```

#### ✅ 합격 기준
- [ ] 데이터 변경 시 캐시 무효화 확인
- [ ] 무효화 마커 정상 설정
- [ ] 레이스 컨디션 방지 동작 확인
- [ ] 최신 데이터 반영 확인

---

## 📊 테스트 결과 기록

### 체크리스트

| 시나리오 | 상태 | 비고 |
|---------|------|------|
| 1. 프로젝트 데이터 캐싱 | ⬜ | |
| 2. 동시 편집 락 (다중 탭) | ⬜ | |
| 3. 락 TTL 자동 만료/연장 | ⬜ | |
| 4. Grace Period 복구 | ⬜ | |
| 5. Redis 재시작 Fail Fast | ⬜ | |
| 6. Optimistic Lock 조합 | ⬜ | |
| 7. 대량 데이터 캐싱 | ⬜ | |
| 8. 캐시 무효화 전파 | ⬜ | |

### 발견된 이슈

| 이슈 | 심각도 | 설명 | 해결 방법 |
|------|--------|------|----------|
| | | | |

---

## 🔧 디버깅 팁

### Redis 상태 확인
```bash
# Docker 사용 시
docker exec -it redis-claude-project redis-cli

# 주요 명령어
> KEYS *                    # 모든 키 조회
> KEYS cache:*             # 캐시 키만 조회
> KEYS lock:*              # 락 키만 조회
> GET cache:current_sheet_data    # 캐시 값 조회
> TTL lock:TEST001         # 락 TTL 확인
> FLUSHDB                  # 전체 삭제 (주의!)
```

### 로그 필터링
```bash
# 캐시 관련 로그만 보기
grep "캐시" logs/app.log

# 락 관련 로그만 보기
grep "잠금" logs/app.log

# 에러만 보기
grep "ERROR" logs/app.log
```

### 네트워크 모니터링
```bash
# 브라우저 개발자 도구 (F12)
- Network 탭
- Preserve log 체크
- Filter: /api/
```

---

## ✅ 최종 검증

모든 시나리오 통과 후:

1. [ ] Redis 캐시 크기 확인 (DBSIZE)
2. [ ] 메모리 사용량 확인 (INFO memory)
3. [ ] 에러 로그 0건 확인
4. [ ] Google Sheets API 호출 감소 확인 (60초 TTL 동작)
5. [ ] 사용자 경험 개선 확인 (응답 속도)

---

**테스트 완료 일시**: ____________________
**테스터**: ____________________
**Redis 버전**: ____________________
**Python 버전**: ____________________

