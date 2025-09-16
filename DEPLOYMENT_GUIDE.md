# 🚀 IT Global Dashboard - 클라우드 배포 가이드

## 📋 개요

이 가이드는 IT Global Dashboard를 Google Cloud Platform에 안전하고 확장 가능하게 배포하는 방법을 설명합니다.

---

## 🏗️ 아키텍처 개요

```
GitHub Repository
        ↓
    GitHub Actions (CI/CD)
        ↓
    Google Container Registry
        ↓
    Cloud Run (Serverless)
        ↓
┌─────────────────────────────────┐
│ Cloud SQL (PostgreSQL)         │
│ Cloud Memorystore (Redis)      │
│ Secret Manager (환경변수)       │
│ Cloud Monitoring (모니터링)     │
└─────────────────────────────────┘
```

---

## 🛠️ 사전 준비사항

### 1. 필수 도구 설치

```bash
# Google Cloud SDK
curl https://sdk.cloud.google.com | bash
exec -l $SHELL

# Docker
sudo apt-get install docker.io

# GitHub CLI (옵션)
sudo apt install gh
```

### 2. GCP 프로젝트 설정

```bash
# 프로젝트 생성
gcloud projects create itglobal-dashboard --name="IT Global Dashboard"

# 프로젝트 설정
gcloud config set project itglobal-dashboard

# 빌링 계정 연결 (필수)
gcloud billing projects link itglobal-dashboard --billing-account=BILLING_ACCOUNT_ID
```

### 3. 필수 API 활성화

```bash
# 스크립트 실행 권한 부여
chmod +x deploy-cloud-run.sh

# 또는 수동으로 서비스 활성화
gcloud services enable \
    cloudbuild.googleapis.com \
    run.googleapis.com \
    sqladmin.googleapis.com \
    secretmanager.googleapis.com \
    monitoring.googleapis.com \
    logging.googleapis.com \
    vpcaccess.googleapis.com
```

---

## 🔐 보안 설정

### 1. 서비스 계정 생성

```bash
# 서비스 계정 생성
gcloud iam service-accounts create itglobal-dashboard-sa \
    --display-name="IT Global Dashboard Service Account"

# 필요한 역할 부여
gcloud projects add-iam-policy-binding itglobal-dashboard \
    --member="serviceAccount:itglobal-dashboard-sa@itglobal-dashboard.iam.gserviceaccount.com" \
    --role="roles/cloudsql.client"

gcloud projects add-iam-policy-binding itglobal-dashboard \
    --member="serviceAccount:itglobal-dashboard-sa@itglobal-dashboard.iam.gserviceaccount.com" \
    --role="roles/secretmanager.secretAccessor"

gcloud projects add-iam-policy-binding itglobal-dashboard \
    --member="serviceAccount:itglobal-dashboard-sa@itglobal-dashboard.iam.gserviceaccount.com" \
    --role="roles/monitoring.metricWriter"
```

### 2. Secret Manager 설정

```bash
# 데이터베이스 비밀번호 저장
gcloud secrets create cloud_sql_password --data-file=- <<< "your-secure-password"

# Flask 시크릿 키 저장
python -c "import secrets; print(secrets.token_hex(32))" | gcloud secrets create flask_secret_key --data-file=-

# Google OAuth 인증 정보 저장
gcloud secrets create google_oauth_client_id --data-file=- <<< "your-client-id"
gcloud secrets create google_oauth_client_secret --data-file=- <<< "your-client-secret"
```

---

## 🗄️ 데이터베이스 설정

### 1. Cloud SQL 인스턴스 생성

```bash
# PostgreSQL 인스턴스 생성
gcloud sql instances create itglobal-main \
    --database-version=POSTGRES_15 \
    --tier=db-g1-small \
    --region=asia-northeast3 \
    --storage-size=20GB \
    --storage-type=SSD \
    --backup-start-time=03:00 \
    --enable-bin-log \
    --deletion-protection

# 데이터베이스 생성
gcloud sql databases create itglobal_db --instance=itglobal-main

# 사용자 생성
gcloud sql users create dashboard_user \
    --instance=itglobal-main \
    --password=secure-password
```

### 2. Cloud Memorystore (Redis) 설정

```bash
# Redis 인스턴스 생성
gcloud redis instances create itglobal-cache \
    --size=1 \
    --region=asia-northeast3 \
    --redis-version=redis_7_0 \
    --network=default
```

---

## 🌐 네트워킹 설정

### 1. VPC Connector 생성

```bash
# VPC 커넥터 생성 (Cloud Run이 private 리소스에 접근)
gcloud compute networks vpc-access connectors create itglobal-vpc-connector \
    --network=default \
    --region=asia-northeast3 \
    --range=10.8.0.0/28 \
    --min-instances=2 \
    --max-instances=3
```

### 2. 방화벽 규칙 설정

```bash
# Cloud Run에서 Cloud SQL 접근 허용
gcloud compute firewall-rules create allow-cloud-sql \
    --allow tcp:5432 \
    --source-ranges 10.8.0.0/28 \
    --description "Allow Cloud Run to access Cloud SQL"
```

---

## 🚀 배포 방법

### 방법 1: 자동 배포 스크립트 사용 (권장)

```bash
# 배포 스크립트 실행
./deploy-cloud-run.sh itglobal-dashboard asia-northeast3
```

### 방법 2: 수동 배포

```bash
# 1. Docker 이미지 빌드
docker build -t gcr.io/itglobal-dashboard/itglobal-dashboard:latest .

# 2. Container Registry에 푸시
gcloud auth configure-docker
docker push gcr.io/itglobal-dashboard/itglobal-dashboard:latest

# 3. Cloud Run 서비스 배포
gcloud run deploy itglobal-dashboard \
    --image gcr.io/itglobal-dashboard/itglobal-dashboard:latest \
    --region asia-northeast3 \
    --platform managed \
    --allow-unauthenticated \
    --memory 2Gi \
    --cpu 2 \
    --min-instances 1 \
    --max-instances 10 \
    --vpc-connector itglobal-vpc-connector \
    --add-cloudsql-instances itglobal-dashboard:asia-northeast3:itglobal-main \
    --service-account itglobal-dashboard-sa@itglobal-dashboard.iam.gserviceaccount.com
```

---

## 🔄 CI/CD 파이프라인 설정

### 1. GitHub Secrets 설정

GitHub 저장소의 Settings > Secrets and variables > Actions에서 다음 시크릿을 추가:

```
GCP_PROJECT_ID: itglobal-dashboard
GCP_SA_KEY: [서비스 계정 JSON 키]
SLACK_WEBHOOK: [Slack 웹훅 URL] (옵션)
```

### 2. 서비스 계정 키 생성

```bash
# 서비스 계정 키 생성 (CI/CD용)
gcloud iam service-accounts keys create key.json \
    --iam-account=itglobal-dashboard-sa@itglobal-dashboard.iam.gserviceaccount.com

# 키 내용을 GitHub Secrets에 추가 (GCP_SA_KEY)
cat key.json | base64 -w 0
```

### 3. 자동 배포 확인

```bash
# develop 브랜치에 push하면 개발 환경에 자동 배포
git checkout -b develop
git push origin develop

# main 브랜치에 push하면 운영 환경에 자동 배포
git checkout main
git merge develop
git push origin main
```

---

## 📊 모니터링 및 로깅

### 1. Cloud Monitoring 대시보드

```bash
# 커스텀 메트릭 확인
gcloud logging read "resource.type=cloud_run_revision AND resource.labels.service_name=itglobal-dashboard" \
    --limit=50 \
    --format="table(timestamp,severity,textPayload)"
```

### 2. 알림 정책 설정

```bash
# 에러율 알림 정책 생성
gcloud alpha monitoring policies create --policy-from-file=monitoring/error-rate-policy.yaml
```

### 3. 로그 기반 메트릭

```bash
# 사용자 정의 로그 메트릭 생성
gcloud logging metrics create user_login_failures \
    --description="Failed login attempts" \
    --log-filter='resource.type="cloud_run_revision" AND textPayload:"login failed"'
```

---

## 🧪 테스트 및 검증

### 1. 헬스체크 확인

```bash
# 서비스 URL 가져오기
SERVICE_URL=$(gcloud run services describe itglobal-dashboard \
    --region=asia-northeast3 \
    --format="value(status.url)")

# 헬스체크 수행
curl -f "$SERVICE_URL/health"
curl -f "$SERVICE_URL/ready"
```

### 2. 부하 테스트

```bash
# Apache Bench를 사용한 간단한 부하 테스트
ab -n 100 -c 10 "$SERVICE_URL/"

# 또는 Locust를 사용한 상세한 부하 테스트
cd tests/performance
locust --headless --users 50 --spawn-rate 5 --run-time 2m --host "$SERVICE_URL"
```

### 3. 보안 스캔

```bash
# Trivy를 사용한 컨테이너 취약점 스캔
trivy image gcr.io/itglobal-dashboard/itglobal-dashboard:latest
```

---

## 🔧 문제 해결

### 1. 일반적인 문제들

#### 배포 실패
```bash
# 서비스 로그 확인
gcloud logs tail -s run.googleapis.com/services/itglobal-dashboard

# 빌드 로그 확인
gcloud builds list --limit=5
gcloud builds log [BUILD_ID]
```

#### 데이터베이스 연결 문제
```bash
# Cloud SQL 프록시를 통한 로컬 테스트
gcloud sql connect itglobal-main --user=dashboard_user
```

#### 메모리 부족
```bash
# 리소스 사용량 확인
gcloud run services describe itglobal-dashboard \
    --region=asia-northeast3 \
    --format="export"
```

### 2. 모니터링 명령어

```bash
# 실시간 로그 보기
gcloud logs tail -s run.googleapis.com/services/itglobal-dashboard --follow

# 메트릭 확인
gcloud monitoring metrics list --filter="resource.type=cloud_run_revision"

# 서비스 상태 확인
gcloud run services list --platform=managed --region=asia-northeast3
```

---

## 🔄 유지보수

### 1. 정기적인 작업

- **매주**: 로그 검토 및 성능 지표 확인
- **매월**: 보안 업데이트 및 종속성 업데이트
- **분기별**: 비용 최적화 검토 및 용량 계획

### 2. 백업 전략

```bash
# 데이터베이스 백업 (자동 백업이 활성화되어 있음)
gcloud sql backups list --instance=itglobal-main

# Secret Manager 백업
gcloud secrets versions list flask_secret_key
```

### 3. 스케일링 정책

```bash
# 트래픽 증가 시 인스턴스 수 조정
gcloud run services update itglobal-dashboard \
    --region=asia-northeast3 \
    --max-instances=20 \
    --min-instances=2
```

---

## 💰 비용 최적화

### 1. 리소스 사용량 모니터링

```bash
# 월별 비용 확인
gcloud billing projects describe itglobal-dashboard

# 서비스별 비용 분석
gcloud logging read "protoPayload.serviceName=run.googleapis.com" \
    --format="table(timestamp,protoPayload.resourceName)"
```

### 2. 최적화 권장사항

- **CPU 할당**: 일정한 트래픽이 있는 경우 CPU 항상 할당 고려
- **메모리**: 실제 사용량에 따라 적절히 조정
- **최소 인스턴스**: 콜드 스타트 방지와 비용의 균형점 찾기

---

## 📞 지원 및 문의

### 1. 로그 수집

문제 발생 시 다음 정보를 수집:

```bash
# 서비스 설명
gcloud run services describe itglobal-dashboard --region=asia-northeast3

# 최근 로그 (마지막 1시간)
gcloud logs read "resource.type=cloud_run_revision" \
    --freshness=1h \
    --format="table(timestamp,severity,textPayload)"

# 시스템 메트릭
gcloud monitoring metrics list --filter="metric.type=run.googleapis.com/container/cpu/utilization"
```

### 2. 긴급 상황 대응

```bash
# 서비스 중지
gcloud run services update itglobal-dashboard \
    --region=asia-northeast3 \
    --no-traffic

# 이전 리비전으로 롤백
gcloud run services update-traffic itglobal-dashboard \
    --region=asia-northeast3 \
    --to-revisions=itglobal-dashboard-[PREVIOUS_REVISION_ID]=100
```

---

## 📚 추가 리소스

- [Cloud Run 공식 문서](https://cloud.google.com/run/docs)
- [Cloud SQL 최적화 가이드](https://cloud.google.com/sql/docs/postgres/optimize-performance)
- [Secret Manager 보안 모범 사례](https://cloud.google.com/secret-manager/docs/best-practices)
- [모니터링 설정 가이드](https://cloud.google.com/monitoring/quickstart)

---

**🔐 보안 주의사항**
- 프로덕션 환경에서는 항상 최소 권한 원칙 적용
- 정기적인 보안 감사 수행
- 모든 시크릿은 Secret Manager 사용
- 네트워크 접근은 최소한으로 제한