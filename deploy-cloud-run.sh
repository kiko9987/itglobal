#!/bin/bash

# IT Global Dashboard - Cloud Run Deployment Script
# 사용법: ./deploy-cloud-run.sh [PROJECT_ID] [REGION]

set -e  # 오류 발생 시 스크립트 중단

# 색상 출력을 위한 설정
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 로그 함수들
log_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

log_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

log_warn() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# 기본 설정
PROJECT_ID=${1:-$(gcloud config get-value project)}
REGION=${2:-"asia-northeast3"}
SERVICE_NAME="itglobal-dashboard"
IMAGE_NAME="gcr.io/${PROJECT_ID}/${SERVICE_NAME}"
VERSION=$(date +%Y%m%d-%H%M%S)

log_info "=== IT Global Dashboard Cloud Run 배포 시작 ==="
log_info "프로젝트 ID: ${PROJECT_ID}"
log_info "리전: ${REGION}"
log_info "서비스명: ${SERVICE_NAME}"
log_info "이미지: ${IMAGE_NAME}:${VERSION}"

# 전제 조건 확인
check_prerequisites() {
    log_info "전제 조건 확인 중..."

    # gcloud CLI 확인
    if ! command -v gcloud &> /dev/null; then
        log_error "gcloud CLI가 설치되지 않았습니다."
        exit 1
    fi

    # Docker 확인
    if ! command -v docker &> /dev/null; then
        log_error "Docker가 설치되지 않았습니다."
        exit 1
    fi

    # 프로젝트 ID 확인
    if [[ -z "${PROJECT_ID}" ]]; then
        log_error "프로젝트 ID가 설정되지 않았습니다."
        log_error "사용법: $0 [PROJECT_ID] [REGION]"
        exit 1
    fi

    # gcloud 인증 확인
    if ! gcloud auth list --filter=status:ACTIVE --format="value(account)" | grep -q .; then
        log_error "gcloud 인증이 되지 않았습니다."
        log_error "다음 명령어로 인증하세요: gcloud auth login"
        exit 1
    fi

    log_success "전제 조건 확인 완료"
}

# Google Cloud 서비스 활성화
enable_services() {
    log_info "필요한 Google Cloud 서비스 활성화 중..."

    local services=(
        "cloudbuild.googleapis.com"
        "run.googleapis.com"
        "sql-component.googleapis.com"
        "sqladmin.googleapis.com"
        "secretmanager.googleapis.com"
        "monitoring.googleapis.com"
        "logging.googleapis.com"
        "vpcaccess.googleapis.com"
    )

    for service in "${services[@]}"; do
        log_info "서비스 활성화: ${service}"
        gcloud services enable "${service}" --project="${PROJECT_ID}" --quiet
    done

    log_success "서비스 활성화 완료"
}

# Docker 이미지 빌드
build_image() {
    log_info "Docker 이미지 빌드 중..."

    # Container Registry 인증 설정
    gcloud auth configure-docker --quiet

    # 이미지 빌드
    docker build -t "${IMAGE_NAME}:${VERSION}" -t "${IMAGE_NAME}:latest" .

    if [ $? -eq 0 ]; then
        log_success "Docker 이미지 빌드 완료"
    else
        log_error "Docker 이미지 빌드 실패"
        exit 1
    fi
}

# Container Registry에 푸시
push_image() {
    log_info "Container Registry에 이미지 푸시 중..."

    docker push "${IMAGE_NAME}:${VERSION}"
    docker push "${IMAGE_NAME}:latest"

    if [ $? -eq 0 ]; then
        log_success "이미지 푸시 완료"
    else
        log_error "이미지 푸시 실패"
        exit 1
    fi
}

# Cloud Run 서비스 배포
deploy_service() {
    log_info "Cloud Run 서비스 배포 중..."

    # 서비스 YAML 파일의 플레이스홀더 대체
    local temp_service_file="cloud-run-service-${VERSION}.yaml"
    sed "s/PROJECT_ID/${PROJECT_ID}/g; s/REGION/${REGION}/g; s/:latest/:${VERSION}/g" cloud-run-service.yaml > "${temp_service_file}"

    # Cloud Run 서비스 배포
    gcloud run services replace "${temp_service_file}" \
        --region="${REGION}" \
        --project="${PROJECT_ID}"

    # 임시 파일 삭제
    rm "${temp_service_file}"

    if [ $? -eq 0 ]; then
        log_success "Cloud Run 서비스 배포 완료"
    else
        log_error "Cloud Run 서비스 배포 실패"
        exit 1
    fi
}

# IAM 권한 설정
setup_permissions() {
    log_info "IAM 권한 설정 중..."

    # 서비스 계정 생성 (이미 존재하면 스킵)
    local service_account="${SERVICE_NAME}-sa@${PROJECT_ID}.iam.gserviceaccount.com"

    if ! gcloud iam service-accounts describe "${service_account}" --project="${PROJECT_ID}" &>/dev/null; then
        log_info "서비스 계정 생성: ${service_account}"
        gcloud iam service-accounts create "${SERVICE_NAME}-sa" \
            --display-name="IT Global Dashboard Service Account" \
            --description="Service account for IT Global Dashboard on Cloud Run" \
            --project="${PROJECT_ID}"
    fi

    # 필요한 역할 부여
    local roles=(
        "roles/cloudsql.client"
        "roles/secretmanager.secretAccessor"
        "roles/monitoring.metricWriter"
        "roles/logging.logWriter"
        "roles/storage.objectViewer"
    )

    for role in "${roles[@]}"; do
        log_info "역할 부여: ${role}"
        gcloud projects add-iam-policy-binding "${PROJECT_ID}" \
            --member="serviceAccount:${service_account}" \
            --role="${role}" \
            --quiet
    done

    log_success "IAM 권한 설정 완료"
}

# 트래픽 전환
update_traffic() {
    log_info "트래픽 전환 중..."

    gcloud run services update-traffic "${SERVICE_NAME}" \
        --to-latest \
        --region="${REGION}" \
        --project="${PROJECT_ID}"

    if [ $? -eq 0 ]; then
        log_success "트래픽 전환 완료"
    else
        log_error "트래픽 전환 실패"
        exit 1
    fi
}

# 배포 상태 확인
check_deployment() {
    log_info "배포 상태 확인 중..."

    local service_url=$(gcloud run services describe "${SERVICE_NAME}" \
        --region="${REGION}" \
        --project="${PROJECT_ID}" \
        --format="value(status.url)")

    if [[ -n "${service_url}" ]]; then
        log_success "서비스 URL: ${service_url}"

        # 헬스체크 수행
        log_info "헬스체크 수행 중..."
        if curl -f "${service_url}/health" &>/dev/null; then
            log_success "헬스체크 성공! 배포가 완료되었습니다."
        else
            log_warn "헬스체크 실패. 서비스가 아직 준비되지 않았을 수 있습니다."
            log_info "몇 분 후 다시 확인해주세요: ${service_url}/health"
        fi
    else
        log_error "서비스 URL을 가져올 수 없습니다."
        exit 1
    fi
}

# 정리 작업
cleanup() {
    log_info "정리 작업 중..."

    # 로컬 Docker 이미지 정리 (선택적)
    read -p "로컬 Docker 이미지를 삭제하시겠습니까? (y/N): " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        docker rmi "${IMAGE_NAME}:${VERSION}" "${IMAGE_NAME}:latest" 2>/dev/null || true
        log_info "로컬 Docker 이미지 삭제 완료"
    fi
}

# 메인 실행 함수
main() {
    check_prerequisites
    enable_services
    setup_permissions
    build_image
    push_image
    deploy_service
    update_traffic
    check_deployment
    cleanup

    log_success "=== 배포 완료! ==="
    echo
    echo "서비스 관리 명령어:"
    echo "  상태 확인: gcloud run services describe ${SERVICE_NAME} --region=${REGION}"
    echo "  로그 보기: gcloud logs tail -s run.googleapis.com/services/${SERVICE_NAME}"
    echo "  삭제: gcloud run services delete ${SERVICE_NAME} --region=${REGION}"
    echo
}

# 스크립트 실행
main "$@"