"""세금계산서·사업자등록증 관리 블루프린트 — PM 대시보드 (②등록증 업로드/열람 + ③계산서 요청).

슬랙 계산서 봇(#계산서_관리)·사업자등록증 Drive 저장소를 그대로 공유한다.
PM 아코디언 문서 카드에서 프로젝트별로:
  - ② 사업자등록증 업로드(Drive 저장)·열람(새 탭 Drive URL)
  - ③ 세금계산서 발행 요청 → 슬랙 #계산서_관리 카드 발송 (슬랙과 동일 코어 post_invoice_request)

게이트: 등록증 필요(거래처 아님)한데 canonical 없으면 계산서 요청 차단 → 먼저 업로드 유도.

⚠️ 경로 접두사 주의: 상태변경 POST 가 프로덕션 보안 미들웨어(CSRF)를 통과하려면 경로가
   반드시 '/api/' 로 시작해야 한다(그래야 세션 기반 _validate_api_auth 분기). 그래서
   url_prefix='/api/invoice'. (프로젝트 편집 등 기존 정상 동작 엔드포인트도 모두 /api/ 접두사.)

- 조회(status): login_required
- 업로드/요청: editor_required (편집자+)
"""
from flask import Blueprint, request, session

from ..auth import login_required, editor_required
from ..api.responses import APIResponse, APIErrorCode
from ..utils.logging_config import get_logger

logger = get_logger(__name__)

invoice_bp = Blueprint('invoice_mgmt', __name__, url_prefix='/api/invoice')


def _current_initial() -> str:
    """현재 로그인 사용자 이메일 → 이니셜 (요청자 기록용). 슬랙과 동일 매핑 (as_mgmt 재사용)."""
    try:
        from .as_mgmt import _current_initial as _ci
        return _ci()
    except Exception:
        email = (session.get('user', {}) or {}).get('email', '') or ''
        return email.split('@')[0] if email else '-'


def _license_required(code: str) -> bool:
    """등록증 검증 필수 여부 (거래처 유입은 skip). 슬랙 _is_license_required 재사용."""
    try:
        from .slack_bot import _is_license_required
        return bool(_is_license_required(code))
    except Exception as exc:
        logger.warning(f'[INVOICE] 등록증 필수 여부 조회 실패 → 필수 처리 ({code}): {exc}')
        return True


def _resolve_invoice_email(code: str) -> str:
    """세금계산서용 발행 이메일 자동채움 — 슬랙 모달과 동일 우선순위 (2026-07-28).

    ① 거래처 탭 이메일(상호명 기준, 홈택스 발행 이력 = 계산서용) **최우선**
    ② 발주처 이메일(온라인 리드=견적용) fallback (법인은 견적≠계산서 이메일 흔함).
    없으면 '' (매니저가 모달에서 직접 입력).
    """
    biz, client_email = '', ''
    try:
        from ..services.project_service import get_project_records
        for r in (get_project_records() or []):
            if (r.get('프로젝트 코드') or '').strip() == code:
                biz = str(r.get('사업자명') or '').strip()
                client_email = str(r.get('발주처 이메일') or '').strip()
                break
    except Exception as exc:
        logger.warning(f'[INVOICE] 이메일 프리필 프로젝트 조회 실패 ({code}): {exc}')
    email = ''
    if biz and biz != '-':
        try:
            from ..services.partner_status_sync import get_cached_partner_email
            email = get_cached_partner_email(biz) or ''
        except Exception as exc:
            logger.warning(f'[INVOICE] 거래처 이메일 조회 실패 (발주처로 fallback): {exc}')
    if not email:
        email = client_email
    return '' if email == '-' else email


def _vat_decided(code: str) -> bool:
    """프로젝트 부가세 필드가 채워졌는지 (미결정 상태로 계산서 발행 방지). 슬랙 검증과 동일 규칙.

    프로젝트를 못 찾거나 조회 실패 시 통과(True) — 슬랙과 동일한 안전 default.
    """
    if not code or code == '-':
        return True
    try:
        from ..services.project_service import get_project_records
        for r in (get_project_records() or []):
            if (r.get('프로젝트 코드') or '').strip() == code:
                vat_raw = r.get('부가세')
                if vat_raw in (None, '', ' '):
                    return False
                if isinstance(vat_raw, str) and not vat_raw.strip():
                    return False
                return True
        return True
    except Exception as exc:
        logger.warning(f'[INVOICE] 부가세 필드 검증 실패 (통과): {exc}')
        return True


# ─────────────────────────────────────────────────────────────
# ② 사업자등록증 — 상태 / 업로드 / 열람
# ─────────────────────────────────────────────────────────────
@invoice_bp.route('/license/status/<code>', methods=['GET'])
@login_required
def api_license_status(code):
    """등록증 상태 — 아코디언 펼침 시 lazy 로드. {exists, required, view_url}."""
    code = (code or '').strip()
    exists = False
    required = True
    view_url = None
    source = None
    email = ''
    try:
        from ..services.business_license_handler import get_license_state
        required = _license_required(code)
        st = get_license_state(code)   # own / reuse / partner(거래처 탭) / None
        exists = bool(st.get('exists'))
        view_url = st.get('view_url')
        source = st.get('source')   # 'own'|'reuse'|'partner'|None (프론트 배지 분기)
    except Exception as exc:
        logger.warning(f'[INVOICE] 등록증 상태 조회 실패 ({code}): {exc}')
    try:
        email = _resolve_invoice_email(code)  # 계산서 모달 이메일 프리필 (슬랙과 동일 우선순위)
    except Exception as exc:
        logger.warning(f'[INVOICE] 이메일 프리필 실패 ({code}): {exc}')
    return APIResponse.success(data={
        'code': code, 'exists': exists, 'required': required,
        'view_url': view_url, 'source': source, 'email': email,
    })


@invoice_bp.route('/license/upload', methods=['POST'])
@editor_required
def api_license_upload():
    """사업자등록증 업로드 (multipart: file + code[, force]).

    OCR 문서유형 게이트: is_license 아니면 저장 거부 + 경고(saved=False, warned=True).
    force=1 이면 게이트 무시하고 강제 저장 (매니저가 확인 후). 성공 시 공사확정 카드 배지 동기화.
    """
    code = (request.form.get('code') or '').strip()
    if not code:
        return APIResponse.error(
            message='프로젝트 코드가 필요합니다',
            error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
        )
    f = request.files.get('file')
    if not f or not (f.filename or '').strip():
        return APIResponse.error(
            message='파일을 선택해주세요',
            error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
        )
    force = (request.form.get('force') or '').strip().lower() in ('1', 'true', 'yes', 'on')

    try:
        file_bytes = f.read()
    except Exception as exc:
        logger.warning(f'[INVOICE] 업로드 파일 읽기 실패 ({code}): {exc}')
        return APIResponse.error(
            message='파일을 읽지 못했습니다',
            error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
        )
    if not file_bytes:
        return APIResponse.error(
            message='빈 파일입니다',
            error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
        )

    from ..services.business_license_handler import (
        _MAX_LICENSE_BYTES, save_business_license, get_license_state,
    )
    if len(file_bytes) > _MAX_LICENSE_BYTES:
        return APIResponse.error(
            message='파일이 너무 큽니다 (최대 50MB)',
            error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
        )

    filename = f.filename
    mimetype = f.mimetype or 'application/octet-stream'

    # OCR 문서유형 게이트 — 세금계산서/카드/정산문서 오저장 방지 (2026-09-01 R3883 교훈).
    # 명시적 업로드라 거부 대신 '경고 후 force' 로 우회 허용(슬랙 자동저장의 강한 거부와 구분).
    ocr_name = ''  # OCR 추출 사업자명 (저장 후 시트 자동 반영용 — 미수령 등록 건 자동 기재)
    if not force:
        try:
            from ..services.business_license_ocr import analyze_business_license
            an = analyze_business_license(file_bytes)
            ocr_name = (an.get('name') or '').strip()
            if not an.get('is_license'):
                if an.get('is_card'):
                    warn = '카드 이미지로 보입니다 (사업자등록증 아님).'
                elif an.get('doc_negative'):
                    warn = '세금계산서·정산문서로 보입니다 (사업자등록증 아님).'
                elif not an.get('has_text'):
                    warn = '문서 내용을 인식하지 못했습니다 (사진 화질을 확인해주세요).'
                else:
                    warn = '사업자등록증으로 인식되지 않습니다.'
                return APIResponse.success(data={
                    'saved': False, 'warned': True, 'warning': warn,
                    'name': an.get('name') or '', 'bno': an.get('bno') or '',
                })
        except Exception as exc:
            logger.warning(f'[INVOICE] OCR 게이트 실패 (계속 저장): {exc}')

    res = save_business_license(code, file_bytes, filename, mimetype)
    if not res.get('ok'):
        reason = res.get('reason', '')
        msg = {
            'no_project_folder': '프로젝트 문서 폴더 경로가 설정돼 있지 않습니다. 폴더 경로를 먼저 등록해주세요.',
            'invalid_file_signature': '지원하지 않는 파일 형식입니다 (이미지·PDF만 가능).',
        }.get(reason, f'저장 실패 ({reason or "알 수 없음"})')
        return APIResponse.error(
            message=msg, error_code='LICENSE_SAVE_FAILED', status_code=400,
            details={'reason': reason},
        )

    # OCR 사업자명 자동 반영 — 미수령 등록 건: 등록증 첨부 시 사업자명 자동 기재 (슬랙 스레드 첨부와 동일 정책).
    # 비어있으면 저장, 기존값과 다르면 덮어쓰지 않음(_maybe_update_business_name). → PM 업로드도 양방향 자동 기재.
    if ocr_name:
        try:
            from ..services.business_license_handler import _maybe_update_business_name
            _maybe_update_business_name(code, ocr_name)
        except Exception as exc:
            logger.warning(f'[INVOICE] 사업자명 자동 반영 실패 ({code}): {exc}')

    # 공사확정 카드 사업자등록증 배지 동기화 (슬랙=서브).
    try:
        from ..services.project_slack_notifier import refresh_project_card_license
        refresh_project_card_license(code)
    except Exception as exc:
        logger.warning(f'[INVOICE] 카드 배지 갱신 실패 ({code}): {exc}')

    view_url = None
    try:
        view_url = get_license_state(code).get('view_url')  # 저장으로 캐시 무효화됨 → 최신 재조회
    except Exception:
        pass

    logger.info(f'[INVOICE] PM 사업자등록증 업로드: {code} → {res.get("file_name")} (by={_current_initial()})')
    return APIResponse.success(data={
        'saved': True, 'exists': True, 'name': res.get('file_name'), 'view_url': view_url,
    })


@invoice_bp.route('/license/delete', methods=['POST'])
@editor_required
def api_license_delete():
    """사업자등록증 삭제 — Drive 휴지통 이동(복구 가능). body: {code}. (편집 모드 전용 UI)"""
    data = request.get_json(silent=True) or {}
    code = (data.get('code') or '').strip()
    if not code:
        return APIResponse.error(
            message='프로젝트 코드가 필요합니다',
            error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
        )
    from ..services.business_license_handler import trash_license_canonical
    res = trash_license_canonical(code)
    if not res.get('ok'):
        reason = res.get('reason', '')
        msg = {
            'no_license': '삭제할 사업자등록증이 없습니다.',
            'no_project_folder': '프로젝트 문서 폴더가 없습니다.',
        }.get(reason, f'삭제 실패 ({reason or "알 수 없음"})')
        return APIResponse.error(
            message=msg, error_code='LICENSE_DELETE_FAILED', status_code=400,
            details={'reason': reason},
        )
    # 공사확정 카드 배지 동기화 (등록증 없음으로).
    try:
        from ..services.project_slack_notifier import refresh_project_card_license
        refresh_project_card_license(code)
    except Exception as exc:
        logger.warning(f'[INVOICE] 카드 배지 갱신 실패 ({code}): {exc}')
    logger.info(f'[INVOICE] PM 사업자등록증 삭제(휴지통): {code} → {res.get("file_name")} (by={_current_initial()})')
    return APIResponse.success(data={'deleted': True, 'name': res.get('file_name')})


# ─────────────────────────────────────────────────────────────
# ③ 세금계산서 발행 요청 → 슬랙 #계산서_관리
# ─────────────────────────────────────────────────────────────
@invoice_bp.route('/request', methods=['POST'])
@editor_required
def api_invoice_request():
    """세금계산서 발행 요청. body: {code, biz, addr, amt, vat(sep/incl), email, memo}.

    게이트: 등록증 필요한데 없으면 409(LICENSE_REQUIRED) → 프론트가 업로드 유도.
            부가세 미결정이면 409(VAT_UNDECIDED). 통과 시 슬랙 카드 발송(post_invoice_request).
    """
    data = request.get_json(silent=True) or {}
    code = (data.get('code') or '').strip()
    if not code:
        return APIResponse.error(
            message='프로젝트 코드가 필요합니다',
            error_code=APIErrorCode.VALIDATION_ERROR, status_code=400,
        )
    biz = (data.get('biz') or '').strip()
    addr = (data.get('addr') or '').strip()
    amt = (data.get('amt') or '').strip()
    vat = (data.get('vat') or 'sep').strip() or 'sep'
    email = (data.get('email') or '').strip()
    memo = (data.get('memo') or '').strip()

    # 게이트 ① 등록증 (모든 유입 필수). **재사용 인식** — 자기 폴더 없어도 같은 사업자명의
    #   기존 등록증이 있으면 통과(발송 시 ensure_license 가 복사 첨부). 매칭 없으면 차단.
    try:
        from ..services.business_license_handler import get_license_state
        has_lic = bool(get_license_state(code).get('exists'))
    except Exception as exc:
        logger.warning(f'[INVOICE] 등록증 검증 실패 (통과 처리): {exc}')
        has_lic = True  # Drive 지연 시 통과 (슬랙과 동일 — 관리자 후속)
    if not has_lic:
        return APIResponse.error(
            message='사업자등록증을 먼저 업로드해주세요. (같은 사업자명의 기존 등록증이 있으면 자동 재사용됩니다)',
            error_code='LICENSE_REQUIRED', status_code=409,
        )

    # 게이트 ② 부가세 결정 여부
    if not _vat_decided(code):
        return APIResponse.error(
            message='부가세(포함/미포함)가 지정되지 않았습니다. 프로젝트 편집에서 먼저 지정해주세요.',
            error_code='VAT_UNDECIDED', status_code=409,
        )

    amt_digits = ''.join(ch for ch in amt if ch.isdigit())
    initial = _current_initial()

    try:
        from .slack_bot import post_invoice_request
        res = post_invoice_request(
            code=code, biz=biz or '-', addr=addr or '-', amt_digits=amt_digits,
            vat_val=vat, email=email or '-', memo=memo, requester_initial=initial,
            dedup_check=True,
        )
    except Exception as exc:
        logger.error(f'[INVOICE] 계산서 요청 발송 실패 ({code}): {exc}', exc_info=True)
        return APIResponse.error(
            message='계산서 요청 발송 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR, status_code=500,
        )

    if not res.get('ok'):
        reason = res.get('reason')
        if reason == 'dedup':
            return APIResponse.error(
                message='방금 같은 프로젝트의 계산서 요청이 접수됐습니다 (90초 중복 방지).',
                error_code='INVOICE_DEDUP', status_code=409,
            )
        return APIResponse.error(
            message=f'계산서 요청 발송에 실패했습니다 ({reason or "알 수 없음"}).',
            error_code='INVOICE_POST_FAILED', status_code=502,
            details={'reason': reason},
        )

    logger.info(f'[INVOICE] PM 계산서 요청: {code} ts={res.get("ts")} (by={initial})')
    return APIResponse.success(data={'ts': res.get('ts'), 'thread_url': res.get('thread_url')})
