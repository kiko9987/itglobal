# Redis 설치 및 실행 가이드

## 🎯 목표
로컬 PC에 Redis 서버를 설치하고 실행하여 통합 테스트 및 서비스 배포 준비

## 📋 설치 옵션

### ✅ 권장: Docker 사용 (가장 간단)

#### 1단계: Docker Desktop 실행
- Windows 시작 메뉴에서 "Docker Desktop" 검색 후 실행
- Docker Desktop이 완전히 시작될 때까지 대기 (1-2분)
- 트레이 아이콘이 초록색으로 변경되면 준비 완료

#### 2단계: Redis 컨테이너 실행
```bash
# PowerShell 또는 CMD에서 실행
docker run -d --name redis-claude-project -p 6379:6379 redis:7-alpine

# 설명:
# -d: 백그라운드 실행
# --name redis-claude-project: 컨테이너 이름
# -p 6379:6379: 포트 매핑
# redis:7-alpine: Redis 7 경량 이미지
```

#### 3단계: Redis 연결 확인
```bash
# Redis 컨테이너에 접속하여 ping 테스트
docker exec -it redis-claude-project redis-cli ping

# 응답: PONG (정상)
```

#### Redis 관리 명령어
```bash
# 컨테이너 상태 확인
docker ps | findstr redis

# 컨테이너 중지
docker stop redis-claude-project

# 컨테이너 시작
docker start redis-claude-project

# 컨테이너 재시작
docker restart redis-claude-project

# Redis CLI 접속
docker exec -it redis-claude-project redis-cli

# 로그 확인
docker logs redis-claude-project

# 컨테이너 삭제 (주의: 데이터 삭제됨)
docker rm -f redis-claude-project
```

---

### 대안 1: Memurai (Windows용 Redis)

Windows 전용 Redis 배포판입니다.

#### 다운로드 및 설치
1. https://www.memurai.com/get-memurai 접속
2. "Download Memurai Developer" 클릭
3. 설치 프로그램 실행 (관리자 권한 필요)
4. 설치 완료 후 자동으로 서비스 시작

#### 연결 확인
```bash
# PowerShell에서 실행
& "C:\Program Files\Memurai\memurai-cli.exe" ping

# 응답: PONG
```

#### Memurai 관리
```bash
# 서비스 상태 확인
Get-Service Memurai

# 서비스 중지
Stop-Service Memurai

# 서비스 시작
Start-Service Memurai

# 서비스 재시작
Restart-Service Memurai
```

---

### 대안 2: WSL2 + Redis

WSL2 (Windows Subsystem for Linux)를 사용합니다.

#### 설치
```bash
# WSL2에서 실행
sudo apt update
sudo apt install redis-server -y

# Redis 설정 파일 수정 (선택사항)
sudo nano /etc/redis/redis.conf

# Redis 시작
sudo service redis-server start
```

#### 연결 확인
```bash
redis-cli ping
# 응답: PONG
```

---

## 🧪 설치 검증

### 1. Python에서 Redis 연결 테스트
```bash
python -c "import redis; r = redis.Redis(host='localhost', port=6379, decode_responses=True); print(r.ping())"

# 출력: True (정상)
```

### 2. Redis 정보 확인
```bash
# Docker 사용 시
docker exec -it redis-claude-project redis-cli INFO server

# 로컬 설치 시
redis-cli INFO server
```

### 3. 간단한 Set/Get 테스트
```bash
# Docker 사용 시
docker exec -it redis-claude-project redis-cli

# 로컬 설치 시
redis-cli

# Redis CLI에서 실행
SET test_key "Hello Redis"
GET test_key
# 출력: "Hello Redis"

DEL test_key
exit
```

---

## 📊 통합 테스트 실행

Redis 서버가 실행 중이면 통합 테스트를 실행할 수 있습니다.

### 유닛 테스트 (Redis 불필요)
```bash
# 구조 및 인터페이스 테스트
pytest tests/unit/test_redis_modules_structure.py -v

# DataFrame 직렬화 테스트
pytest tests/unit/test_dataframe_serialization.py -v
```

### 통합 테스트 (Redis 필수)
```bash
# Redis 연결 확인 먼저
python -c "import redis; redis.Redis(host='localhost', port=6379).ping()"

# 통합 테스트 실행 (Fail Fast 테스트 제외)
pytest tests/integration/test_redis_integration.py -v -k "not test_redis_connection_failure_on_init"

# 전체 테스트 (주의: Fail Fast는 sys.exit(1) 호출)
pytest tests/integration/test_redis_integration.py -v
```

---

## 🔧 트러블슈팅

### 문제 1: "Connection refused" 에러
**증상**:
```
ConnectionRefusedError: [WinError 10061] 대상 컴퓨터에서 연결을 거부했습니다
```

**해결**:
1. Redis 서버가 실행 중인지 확인
   ```bash
   # Docker
   docker ps | findstr redis

   # Memurai
   Get-Service Memurai

   # WSL2
   sudo service redis-server status
   ```

2. 포트가 올바른지 확인 (기본: 6379)
   ```bash
   netstat -an | findstr 6379
   ```

3. 방화벽 설정 확인

### 문제 2: Docker Desktop이 시작되지 않음
**해결**:
1. Windows 재시작
2. Docker Desktop 재설치
3. WSL2 업데이트 필요 시 업데이트

### 문제 3: "NOAUTH Authentication required"
**증상**:
```
redis.exceptions.AuthenticationError: NOAUTH Authentication required.
```

**해결**:
1. `.env` 파일에 비밀번호 설정
   ```bash
   REDIS_PASSWORD=your_password
   ```

2. 또는 Redis 비밀번호 제거
   ```bash
   # redis.conf 수정
   # requirepass your_password  <- 주석 처리
   ```

### 문제 4: pytest에서 "SystemExit: 1" 발생
**원인**: `test_redis_connection_failure_on_init` 테스트가 실제로 `sys.exit(1)` 호출

**해결**: 해당 테스트 제외
```bash
pytest tests/integration/test_redis_integration.py -v -k "not test_redis_connection_failure_on_init"
```

---

## 📝 환경 변수 설정

`.env` 파일에서 Redis 설정을 확인/수정하세요:

```bash
# Redis 설정
REDIS_HOST=localhost
REDIS_PORT=6379
REDIS_DB=0              # 0-15 중 선택
REDIS_PASSWORD=         # 비밀번호 (선택사항)
REDIS_SOCKET_TIMEOUT=5
REDIS_SOCKET_CONNECT_TIMEOUT=5
```

**테스트용 DB 사용 권장**:
```bash
REDIS_DB=15  # 프로덕션 데이터와 분리
```

---

## 🚀 다음 단계

Redis 설치가 완료되면:

1. ✅ 통합 테스트 실행
2. ✅ Manual QA 시나리오 수행
3. ✅ 서비스 시작 및 검증
4. ✅ 모니터링 설정

상세한 Manual QA 가이드는 `MANUAL_QA_GUIDE.md`를 참조하세요.

---

## 📚 참고 자료

- [Redis 공식 문서](https://redis.io/docs/)
- [Memurai 문서](https://docs.memurai.com/)
- [Docker Hub - Redis](https://hub.docker.com/_/redis)
- [WSL2 설치 가이드](https://docs.microsoft.com/en-us/windows/wsl/install)

---

**작성 일시**: 2025-01
**작성자**: Claude Code
