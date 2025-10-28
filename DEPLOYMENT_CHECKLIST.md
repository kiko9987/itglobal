# Redis 마이그레이션 배포 체크리스트

## 📋 배포 전 준비사항

### ✅ 1. 코드 검증
- [ ] 모든 변경사항 커밋됨
- [ ] Git 상태 clean (`git status`)
- [ ] 브랜치: main에 머지 완료

### ✅ 2. 테스트 완료
- [ ] 유닛 테스트 통과 (pytest tests/unit/)
- [ ] 통합 테스트 통과 (pytest tests/integration/)
- [ ] Manual QA 8개 시나리오 모두 통과
- [ ] DataFrame 직렬화 테스트 통과

### ✅ 3. 환경 설정 확인
- [ ] `.env` 파일 Redis 설정 확인
- [ ] `REDIS_HOST=localhost`
- [ ] `REDIS_PORT=6379`
- [ ] `REDIS_DB=0` (또는 분리된 DB 번호)
- [ ] `REDIS_PASSWORD` 설정 (필요 시)

### ✅ 4. 의존성 설치
- [ ] `requirements.txt` 업데이트됨
- [ ] `redis==5.2.1` 포함 확인
- [ ] `pip install -r requirements.txt` 실행

---

## 🚀 배포 절차

### Step 1: Redis 서버 준비

#### Docker 사용 시
```bash
# 1. Redis 컨테이너 실행
docker run -d --name redis-claude-project \
  -p 6379:6379 \
  --restart unless-stopped \
  redis:7-alpine

# 2. 연결 확인
docker exec -it redis-claude-project redis-cli ping
# 응답: PONG

# 3. 자동 시작 설정 확인
docker inspect redis-claude-project | grep RestartPolicy
# "RestartPolicy": { "Name": "unless-stopped" }
```

#### Memurai 사용 시 (Windows)
```powershell
# 1. 서비스 상태 확인
Get-Service Memurai

# 2. 자동 시작 설정
Set-Service -Name Memurai -StartupType Automatic

# 3. 서비스 시작
Start-Service Memurai
```

#### WSL2 사용 시
```bash
# 1. 자동 시작 설정
sudo systemctl enable redis-server

# 2. 서비스 시작
sudo systemctl start redis-server

# 3. 상태 확인
sudo systemctl status redis-server
```

---

### Step 2: 기존 서비스 중지

```bash
# 1. 현재 실행 중인 프로세스 확인
# Windows PowerShell
Get-Process python | Where-Object {$_.CommandLine -like "*dashboard*"}

# Linux/WSL
ps aux | grep "dashboard.app"

# 2. 프로세스 종료
# 안전한 종료 (Ctrl+C 또는 kill)
# 강제 종료는 피하기 (데이터 손실 가능)

# 3. 포트 확인
netstat -ano | findstr :5000
netstat -ano | findstr :8000
```

---

### Step 3: 데이터 백업 (선택사항)

```bash
# 1. SQLite 백업 (projects.db)
cp instance/projects.db instance/projects.db.backup_$(date +%Y%m%d_%H%M%S)

# 2. users.json 백업 (있는 경우)
cp users.json users.json.backup_$(date +%Y%m%d_%H%M%S)

# 3. 로그 아카이브
mkdir -p logs/archive
mv logs/*.log logs/archive/ 2>/dev/null || true
```

---

### Step 4: 애플리케이션 시작

#### 개발 모드
```bash
# Flask 개발 서버
python run.py

# 로그 확인
tail -f logs/app.log
```

#### 프로덕션 모드 (권장)
```bash
# Gunicorn with workers=1 (초기 단계)
gunicorn -c gunicorn.conf.py dashboard.app:app

# 백그라운드 실행
nohup gunicorn -c gunicorn.conf.py dashboard.app:app > logs/gunicorn.log 2>&1 &

# 프로세스 확인
ps aux | grep gunicorn
```

---

### Step 5: 헬스 체크

#### 1. 서비스 응답 확인
```bash
# 기본 헬스 체크
curl http://localhost:8000/

# API 헬스 체크
curl http://localhost:8000/api/health
```

#### 2. Redis 연결 확인
```bash
# 서버 로그 확인
grep "Redis 연결 성공" logs/app.log
# 출력: Redis 연결 성공: localhost:6379 (DB: 0)

# 또는 Python으로 직접 확인
python -c "from dashboard.utils.redis_client import get_redis_client; print(get_redis_client().ping())"
# 출력: True
```

#### 3. 캐시 동작 확인
```bash
# 1. 프로젝트 목록 조회 (첫 로드)
curl http://localhost:8000/api/projects

# 2. 로그 확인
grep "DataFrame 감지" logs/app.log
# 출력: DataFrame 감지 - pickle 직렬화: current_sheet_data

# 3. 두 번째 조회 (캐시 히트)
curl http://localhost:8000/api/projects

# 4. 로그 확인
grep "캐시 히트 (pickle)" logs/app.log
# 출력: 캐시 히트 (pickle): current_sheet_data
```

#### 4. 락 동작 확인
```bash
# Redis CLI에서 락 키 확인
docker exec -it redis-claude-project redis-cli KEYS lock:*

# 편집 중인 프로젝트가 있으면 출력:
# 1) "lock:PROJECT_CODE"
```

---

### Step 6: 모니터링 설정

#### 1. 로그 모니터링 스크립트 생성
```bash
# watch_logs.sh 생성
cat > watch_logs.sh << 'EOF'
#!/bin/bash
echo "=== 실시간 로그 모니터링 ==="
echo "Ctrl+C로 종료"
echo ""
tail -f logs/app.log | grep --line-buffered -E "ERROR|WARNING|Redis|캐시|잠금"
EOF

chmod +x watch_logs.sh
./watch_logs.sh
```

#### 2. Redis 모니터링
```bash
# Redis 정보 확인 스크립트
cat > redis_monitor.sh << 'EOF'
#!/bin/bash
echo "=== Redis 상태 ==="
docker exec redis-claude-project redis-cli INFO stats | grep -E "total_commands|keyspace"
docker exec redis-claude-project redis-cli DBSIZE
docker exec redis-claude-project redis-cli INFO memory | grep used_memory_human
EOF

chmod +x redis_monitor.sh
./redis_monitor.sh
```

#### 3. 주기적 헬스 체크 (cron)
```bash
# healthcheck.sh 생성
cat > healthcheck.sh << 'EOF'
#!/bin/bash
TIMESTAMP=$(date "+%Y-%m-%d %H:%M:%S")
STATUS=$(curl -s -o /dev/null -w "%{http_code}" http://localhost:8000/)

if [ "$STATUS" -eq 200 ]; then
    echo "[$TIMESTAMP] OK - Service is running"
else
    echo "[$TIMESTAMP] ERROR - Service returned $STATUS"
    # 알림 전송 (선택사항)
fi
EOF

chmod +x healthcheck.sh

# cron 등록 (5분마다 실행)
# crontab -e
# */5 * * * * /path/to/healthcheck.sh >> /path/to/logs/healthcheck.log 2>&1
```

---

## 📊 배포 후 검증

### ✅ 즉시 확인사항 (배포 후 10분 내)

- [ ] 서비스 정상 응답 (HTTP 200)
- [ ] Redis 연결 성공 로그 확인
- [ ] 에러 로그 0건
- [ ] 프로젝트 목록 정상 표시
- [ ] 캐시 히트 로그 확인

### ✅ 1시간 모니터링

- [ ] CPU 사용률 < 50%
- [ ] 메모리 사용률 < 70%
- [ ] Redis 메모리 < 100MB
- [ ] 응답 시간 < 200ms (평균)
- [ ] 에러율 < 1%

### ✅ 1일 모니터링

- [ ] Google Sheets API 호출 감소 확인 (기대: 75회/시간 → 60회/시간)
- [ ] 캐시 히트율 > 80%
- [ ] 락 경합 없음
- [ ] 사용자 불편사항 없음

### ✅ 1주 모니터링

- [ ] 장기 안정성 확인
- [ ] 메모리 누수 없음
- [ ] Redis 캐시 크기 안정화
- [ ] workers=2 증가 검토 (CPU 60%+ 시)

---

## 🔧 트러블슈팅

### 문제 1: 서비스가 시작되지 않음

**증상**: `sys.exit(1)` 종료

**해결**:
```bash
# Redis 연결 확인
docker ps | grep redis

# Redis 재시작
docker restart redis-claude-project

# 서비스 재시작
python run.py
```

---

### 문제 2: 503 Service Unavailable

**증상**: 모든 요청이 503 반환

**해결**:
```bash
# 1. Redis 상태 확인
docker exec redis-claude-project redis-cli ping

# 2. 로그 확인
grep "ServiceUnavailable" logs/app.log

# 3. Redis 메모리 확인
docker exec redis-claude-project redis-cli INFO memory

# 4. 필요시 캐시 정리
docker exec redis-claude-project redis-cli FLUSHDB
```

---

### 문제 3: 캐시가 동작하지 않음

**증상**: 매번 Google Sheets API 호출

**해결**:
```bash
# 1. 캐시 키 확인
docker exec redis-claude-project redis-cli KEYS cache:*

# 2. 캐시 저장 로그 확인
grep "캐시 저장" logs/app.log

# 3. DataFrame 직렬화 로그 확인
grep "DataFrame 감지" logs/app.log

# 4. 에러 확인
grep "pickle 직렬화 실패" logs/app.log
```

---

### 문제 4: 락이 해제되지 않음

**증상**: 편집 모드 진입 불가

**해결**:
```bash
# 1. 락 키 확인
docker exec redis-claude-project redis-cli KEYS lock:*

# 2. 특정 락 TTL 확인
docker exec redis-claude-project redis-cli TTL lock:PROJECT_CODE

# 3. 강제 해제 (관리자만)
docker exec redis-claude-project redis-cli DEL lock:PROJECT_CODE
docker exec redis-claude-project redis-cli DEL grace:PROJECT_CODE
```

---

## 🔄 롤백 절차 (비상시)

만약 Redis 전환 후 심각한 문제가 발생하면:

### 1. 서비스 중지
```bash
# Gunicorn 프로세스 종료
pkill -f gunicorn

# 또는 PID로 종료
ps aux | grep gunicorn
kill <PID>
```

### 2. 이전 코드로 복원
```bash
# 이전 커밋으로 롤백
git log --oneline -10
git checkout <PREVIOUS_COMMIT_HASH>

# 의존성 재설치
pip install -r requirements.txt
```

### 3. 서비스 재시작
```bash
# Redis 없이 시작 (workers=1 메모리 기반)
python run.py
```

### 4. 문제 분석
```bash
# 로그 보관
cp logs/app.log logs/rollback_$(date +%Y%m%d_%H%M%S).log

# 에러 분석
grep ERROR logs/rollback_*.log
```

---

## 📈 성능 지표 수집

### Google Sheets API 호출 감소 측정

**측정 방법**:
```python
# dashboard/utils/google_sheets.py에 카운터 추가
import logging
logger = logging.getLogger(__name__)

def fetch_sheet_data(...):
    logger.info(f"[GOOGLE_SHEETS_API_CALL] Fetching data from sheet: {sheet_id}")
    # ... 기존 코드
```

**분석**:
```bash
# 1시간 동안 API 호출 횟수 확인
grep "GOOGLE_SHEETS_API_CALL" logs/app.log | grep "$(date +%Y-%m-%d\ %H):" | wc -l

# 기대값: 60-75회/시간 (60초 TTL, 프리패치 포함)
# 이전: 150회/시간 (캐시 없음)
```

---

## ✅ 배포 완료 체크리스트

### 필수 항목
- [ ] Redis 서버 자동 시작 설정
- [ ] 애플리케이션 정상 구동
- [ ] 헬스 체크 통과
- [ ] 캐시 동작 확인
- [ ] 락 동작 확인
- [ ] 모니터링 설정 완료

### 문서화
- [ ] 배포 일시 기록
- [ ] 배포 버전 (Git commit hash)
- [ ] Redis 버전 기록
- [ ] 트러블슈팅 이력 기록

### 사용자 공지
- [ ] 내부 사용자에게 변경사항 공지
- [ ] 문제 발생 시 연락처 제공
- [ ] 피드백 수집 채널 마련

---

## 📝 배포 정보

| 항목 | 값 |
|------|-----|
| 배포 일시 | ______________________ |
| Git Commit | ______________________ |
| Redis 버전 | ______________________ |
| Python 버전 | ______________________ |
| Workers 수 | 1 |
| 배포자 | ______________________ |

---

**다음 단계**: 1-2주 모니터링 후 workers=2 증가 검토
