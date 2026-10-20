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


@invoice_bp.route('/issued/<code>', methods=['GET'])
@login_required
def api_invoice_issued(code):
    """이미 발행된 계산서 요약 — PM 계산서 요청 모달용 (슬랙 모달과 같은 계산, 2026-10-01).

    {has_issued, remaining_supply(잔여 공급가, 프리필 가능할 때만), warn(안내 문구, 평문),
     active([{date, stage, kind, amt, biz}] 수정발행 대체 후보)}
    나눠 내는 공사(2026-10-20)는 split=True + payers {단계: {issued, remaining_supply, warn}} — 창에서 '누구 계산서'를
    고르면 그 사람 기준(부담액 − 이미 발행)으로 채운다. 전체 기준 warn·remaining_supply 는 비움.
    """
    code = (code or '').strip()
    data = {'has_issued': False, 'remaining_supply': None, 'warn': '', 'active': []}

    def _plain(txt: str) -> str:   # 슬랙 mrkdwn → 평문 (이모지 shortcode·볼드 제거)
        for sc, emo in ((':clipboard:', '📋'), (':warning:', '⚠️'), (':white_check_mark:', '✅')):
            txt = txt.replace(sc, emo)
        return txt.replace('*', '')

    try:
        from .slack_bot import _project_issued_invoice, _fmt_issued_warn
        iss = _project_issued_invoice(code)
        if iss:
            data['active'] = (iss.get('active') or [])[:10]
            if iss.get('issued_supply', 0) > 0 or iss.get('uncertain'):
                data['has_issued'] = True
                data['warn'] = _plain(_fmt_issued_warn(iss))
                if iss.get('issued_supply', 0) > 0 and not iss.get('uncertain'):
                    data['remaining_supply'] = int(iss['remaining_supply'])
        try:
            from ..services.payer_split import stage_payers
            payers = stage_payers(_payer_rec_cached(code) or {})
        except Exception:
            payers = {}
        if payers:
            per = {}
            for s in payers:
                one = _project_issued_invoice(code, stage=s) or {}
                per[s] = {
                    'issued': int(one.get('issued_gross') or 0),
                    'remaining_supply': int(one['remaining_supply']) if one else None,
                    'warn': _plain(_fmt_issued_warn(one)) if int(one.get('issued_gross') or 0) > 0 else '',
                }
            data.update(split=True, payers=per, warn='', remaining_supply=None)
    except Exception as exc:
        logger.warning(f'[INVOICE] 이미발행 조회 실패 ({code}): {exc}')
    return APIResponse.success(data=data)


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
        'payers': _payer_invoice_prefill(code),
    })


# ─────────────────────────────────────────────────────────────
# ④ 나눠 내는 공사 — 부담자 설정 (2026-10-19 오남프라자, 매니저도 사용)
# ─────────────────────────────────────────────────────────────
@invoice_bp.route('/partner-search', methods=['GET'])
@login_required
def api_partner_search():
    """거래처 탭 자동완성 — [부담자 설정] 사업자명 칸. {items: [{name, bno, ceo, email}]}"""
    q = (request.args.get('q') or '').strip()
    try:
        from ..services.partner_status_sync import search_partners
        items = search_partners(q)
    except Exception as exc:
        logger.warning(f'[PAYER] 거래처 검색 실패 ({q!r}): {exc}')
        items = []
    return APIResponse.success(data={'items': items})


def _payer_rec_from_sheet(code: str):
    """그 공사 행을 시트에서 바로 읽어 {필드: 값} — 칸 배정·입금 보호 판단은 캐시 말고 최신값으로."""
    import os
    from ..services.lead_service import get_sheets_manager
    sid = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sn = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    mgr = get_sheets_manager()
    row = mgr.find_row_by_project_code(sid, code, f"{sn}!A:A")
    if not row:
        return None
    rm = mgr.get_row_map(sid, sn, row) or {}
    return {f: rm.get(letter, '') for f, letter in mgr.get_field_to_letter().items()}


def _payer_rec_cached(code: str):
    """캐시된 공사 데이터에서 그 행만 {필드: 값} — [부담자 설정] 창 초기값용. 없으면 None.

    2026-10-20: 창을 열 때마다 시트를 직접 읽어(A열 전체 + 그 행 = 1.5~2.6초) 버튼 반응이 느리다는 제보.
    화면 초기값은 캐시로 충분하고(저장 시 update_project_in_cache 로 갱신), 저장(api_payers_save)은
    칸 배정·입금 보호 판단 때문에 계속 _payer_rec_from_sheet 최신값으로 다시 읽는다.
    get_project_records()는 4,200행 전체를 dict 로 바꿔 0.27초 → 아코디언마다 미리 조회하므로 한 행만 꺼냄(0.01초).
    """
    import pandas as pd
    from ..services.project_service import load_data
    df = load_data()
    if not isinstance(df, pd.DataFrame) or df.empty or '프로젝트 코드' not in df.columns:
        return None
    sub = df[df['프로젝트 코드'].astype(str).str.strip() == code]
    if sub.empty:
        return None
    row = sub.iloc[0]
    return row.where(pd.notna(row), None).to_dict()


def _payer_rows(rec: dict) -> list:
    from ..services.payer_split import stage_payers
    return [{'stage': s, 'role': p['role'], 'name': p['name'], 'biz_no': p['biz_no'],
             'email': p['email'], 'amount': p['amount'], 'method': p['method'],
             'vat_free': bool(p.get('vat_free')), 'received': p['received']}
            for s, p in stage_payers(rec).items()]


@invoice_bp.route('/payers/<code>', methods=['GET'])
@login_required
def api_payers_get(code):
    """[부담자 설정] 창 초기값 — {payers, total2, stage_values}. 캐시 우선, 없으면(막 등록한 공사) 시트."""
    code = (code or '').strip()
    rec = None
    try:
        rec = _payer_rec_cached(code)
    except Exception as exc:
        logger.debug(f'[PAYER] 캐시 조회 실패 ({code}): {exc}')
    if rec is None:
        try:
            rec = _payer_rec_from_sheet(code)
        except Exception as exc:
            logger.warning(f'[PAYER] 조회 실패 ({code}): {exc}')
            rec = None
    if rec is None:
        return APIResponse.error(message='프로젝트를 시트에서 찾지 못했습니다.',
                                 error_code=APIErrorCode.PROJECT_NOT_FOUND, status_code=404)
    from ..services.payer_split import _num, STAGES, vat_ratio
    # total1·vat_ratio: 창이 부담액을 부가세 별도로 보여 주고 입력받기 위해(2026-10-21) — 칸에는 받을 돈으로 저장
    return APIResponse.success(data={
        'code': code, 'payers': _payer_rows(rec), 'total2': _num(rec.get('총액 2')),
        'total1': _num(rec.get('총액 1')), 'vat_ratio': vat_ratio(rec),
        'stage_values': {s: _num(rec.get(s)) for s in STAGES},
    })


@invoice_bp.route('/payers/<code>', methods=['POST'])
@editor_required
def api_payers_save(code):
    """[부담자 설정] 저장 — 칸 자동 배정 → '○○ 부담자' 칸 + 칸 메모 맨 윗줄 기록.

    body: {payers: [{role, name, biz_no, email, amount, method, stage?}], confirm_total?: bool}
    400 = 입력 오류(이름·금액 누락, 입금 있는 칸에 부담자 빠짐 등), 409 PAYER_TOTAL_CHECK = 부담액 합계 ≠
    총액 2(화면에서 확인하면 confirm_total=true 로 다시 보냄). payers=[] 이면 나눠 내기 해제.
    """
    code = (code or '').strip()
    data = request.get_json(silent=True) or {}
    rows = data.get('payers')
    if not isinstance(rows, list):
        return APIResponse.error(message='부담자 목록 형식이 올바르지 않습니다.',
                                 error_code=APIErrorCode.VALIDATION_ERROR, status_code=400)
    try:
        rec = _payer_rec_from_sheet(code)
    except Exception as exc:
        logger.warning(f'[PAYER] 조회 실패 ({code}): {exc}')
        rec = None
    if rec is None:
        return APIResponse.error(message='프로젝트를 시트에서 찾지 못했습니다.',
                                 error_code=APIErrorCode.PROJECT_NOT_FOUND, status_code=404)
    from ..services.payer_split import plan_payers, save_payers, PAYER_FIELD
    plan, errors, warns = plan_payers(rows, rec)
    if errors:
        return APIResponse.error(message='\n'.join(errors),
                                 error_code=APIErrorCode.VALIDATION_ERROR, status_code=400)
    if warns and not data.get('confirm_total'):
        return APIResponse.error(message='\n'.join(warns), error_code='PAYER_TOTAL_CHECK', status_code=409)
    before = {f: str(rec.get(f) or '') for f in PAYER_FIELD.values()}
    try:
        changed = save_payers(code, plan)
    except Exception as exc:
        logger.error(f'[PAYER] 저장 실패 ({code}): {exc}', exc_info=True)
        return APIResponse.error(message='부담자 저장 중 오류가 났습니다. 잠시 후 다시 시도해 주세요.',
                                 error_code=APIErrorCode.INTERNAL_ERROR, status_code=500)
    if changed:
        try:
            from ..services.project_service import update_project_in_cache, invalidate_project_cache
            if not update_project_in_cache(code, changed):
                invalidate_project_cache()
        except Exception as exc:
            logger.warning(f'[PAYER] 캐시 갱신 실패 ({code}): {exc}')
        try:
            from ..utils.user_database import get_audit_repository
            email = (session.get('user', {}) or {}).get('email', '') or ''
            for f in PAYER_FIELD.values():
                if f in changed:
                    get_audit_repository().log_action(
                        email, 'payer_split_update', details='PM 부담자 설정', project_code=code,
                        field_name=f, old_value=before.get(f, ''), new_value=changed[f])
        except Exception as exc:
            logger.warning(f'[PAYER] 감사 로그 실패 ({code}): {exc}')
    logger.info(f'[PAYER] 부담자 설정: {code} by {_current_initial()} → '
                f'{ {s: p["name"] for s, p in plan.items()} } (바뀐 칸 {len(changed)})')
    rec2 = dict(rec)
    rec2.update({k: v for k, v in changed.items() if k in PAYER_FIELD.values()})
    return APIResponse.success(data={'fields': changed, 'payers': _payer_rows(rec2)})


def _payer_invoice_prefill(code: str) -> dict:
    """나눠 내는 공사 — 부담자별 계산서 프리필 {단계: {name, email}} (2026-10-19 오남프라자).

    사업자명은 거래처 탭 정식 상호(사업자번호로 찾음) 우선, 이메일은 부담자 칸 > 거래처 탭.
    부담자 칸이 없는 평범한 공사는 {}.
    """
    out = {}
    try:
        from ..services.payer_split import stage_payers
        from ..services.project_service import get_project_records
        from ..services.partner_status_sync import get_partner_info, get_partner_name_by_bno
        rec = next((r for r in (get_project_records() or [])
                    if str(r.get('프로젝트 코드') or '').strip() == code), None)
        for stage, p in (stage_payers(rec) if rec else {}).items():
            official = get_partner_name_by_bno(p['biz_no']) if p.get('biz_no') else None
            info = get_partner_info(official or p['name']) or {}
            out[stage] = {'name': official or p['name'],
                          'email': p.get('email') or info.get('email') or ''}
    except Exception as exc:
        logger.warning(f'[INVOICE] 부담자 프리필 실패 ({code}): {exc}')
    return out


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

    res = save_business_license(code, file_bytes, filename, mimetype, ocr_biz=ocr_name)
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

    # 등록증 부재로 거래처 정보로 대체됐던 계산서 스레드가 있으면 실제 등록증 재첨부.
    #   (계산서 요청이 등록증 업로드보다 먼저 나간 경우 — 담당자가 스레드에서 바로 열람)
    #   슬랙 파일 업로드는 느려 응답 블로킹 방지 위해 백그라운드.
    try:
        import threading as _th
        from .slack_bot import attach_license_to_pending_invoice_threads
        _th.Thread(target=attach_license_to_pending_invoice_threads,
                   args=(code,), daemon=True).start()
    except Exception as exc:
        logger.warning(f'[INVOICE] 계산서 스레드 등록증 재첨부 트리거 실패 ({code}): {exc}')

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
    """세금계산서 발행 요청. body: {code, biz, addr, amt, vat(sep/incl), email, memo, stages,
    replace?: [{date, stage, amt, biz}] (수정발행 대체 대상)}.

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
    # 계산서 발행 단계 CSV(예 '계약금,잔금') — 첨부 완료 시 시트 '발행' 자동기록용 (2026-09-07)
    _stages_in = data.get('stages')
    if isinstance(_stages_in, (list, tuple)):
        stages = ','.join(str(s).strip() for s in _stages_in if str(s).strip())
    else:
        stages = str(_stages_in or '').strip()

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

    # 수정발행 — 대체할 기존 계산서 (2026-10-01). 클라이언트 값은 서버의 현재 유효 계산서 목록과
    #   대조해 일치하는 것만 사용 (임의 값으로 메모에 취소 줄이 생기지 않게).
    replace = []
    _req_replace = data.get('replace') or []
    if isinstance(_req_replace, list) and _req_replace:
        try:
            from .slack_bot import _project_issued_invoice
            _active = (_project_issued_invoice(code) or {}).get('active') or []
            for r in _req_replace:
                if not isinstance(r, dict):
                    continue
                hit = next((it for it in _active
                            if it['stage'] == r.get('stage') and abs(it['amt'] - float(r.get('amt') or 0)) <= 1
                            and (it.get('biz') or '') == (r.get('biz') or '')
                            and (it.get('date') or '') == (r.get('date') or '')), None)
                if hit and hit not in replace:
                    replace.append(hit)
        except Exception as exc:
            logger.warning(f'[INVOICE] 수정발행 대상 확인 실패 ({code}): {exc}')

    # 금액 사전 점검 (2026-10-15, 슬랙 모달과 같은 규칙) — 부가세 포함 금액 의심·총액 초과면
    #   409 로 경고만 돌려주고, 화면에서 확인(confirm_amount=true) 후 다시 보내면 통과.
    if not data.get('confirm_amount'):
        try:
            from .slack_bot import _invoice_amount_warning
            _warn = _invoice_amount_warning(code, amt_digits, vat, has_replace=bool(replace),
                                            stage=(stages.split(',')[-1].strip() or None))
        except Exception as exc:
            logger.warning(f'[INVOICE] 금액 점검 실패 (통과): {exc}')
            _warn = ''
        if _warn:
            return APIResponse.error(message=_warn, error_code='INVOICE_AMOUNT_CHECK', status_code=409)

    try:
        from .slack_bot import post_invoice_request
        res = post_invoice_request(
            code=code, biz=biz or '-', addr=addr or '-', amt_digits=amt_digits,
            vat_val=vat, email=email or '-', memo=memo, requester_initial=initial,
            stages=stages, dedup_check=True, replace=replace,
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


@invoice_bp.route('/thread-url', methods=['GET'])
@login_required
def api_invoice_thread_url():
    """프로젝트 코드(+단계)의 슬랙 계산서 카드 스레드 링크 반환. 없으면 url=None.

    PM 수금관리 모드 세금계산서 아이콘 '바로가기'용. 발행 시 저장된 invoice_thread:{code}(해시) 조회.
    링크는 발행 요청 시점부터 저장돼 과거(그 이전) 발행 건은 url=None (프론트가 안내 처리).
    """
    code = (request.args.get('code') or '').strip()
    stage = (request.args.get('stage') or '').strip()
    if not code:
        return APIResponse.success(data={'url': None, 'code': code, 'stage': stage})
    url = None
    try:
        from ..utils.redis_client import get_redis_client
        rc = get_redis_client().redis

        def _dec(v):
            return v.decode() if isinstance(v, bytes) else v
        key = f'invoice_thread:{code}'
        if stage:
            v = rc.hget(key, stage)
            url = _dec(v) if v else None
        if not url:
            h = rc.hgetall(key) or {}
            if h:
                url = _dec(list(h.values())[0])  # 단계 매칭 없으면 대표 링크
    except Exception as exc:
        logger.warning(f'[INVOICE] 스레드 링크 조회 실패 ({code}/{stage}): {exc}')
    return APIResponse.success(data={'url': url, 'code': code, 'stage': stage})
