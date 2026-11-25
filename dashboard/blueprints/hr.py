"""그룹웨어 인사 화면 — 내 정보·구성원(인사카드)·그룹웨어 설정.

정식 오픈 전까지는 시범 명단(그룹웨어 새 화면을 쓸 수 있는 사람)만 들어온다.
API 는 /api/ 아래(CSRF 경로 규칙). 권한은 services/hr.py 가 다시 확인한다.
"""
import io
import json
from functools import wraps

from flask import Blueprint, jsonify, redirect, render_template, request, session

from dashboard.auth import login_required
from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

hr_bp = Blueprint('hr', __name__)


def _me() -> str:
    return ((session.get('user') or {}).get('email') or '').lower()


def _is_admin() -> bool:
    return (session.get('user') or {}).get('permission_level') in ('admin', 'super_admin')


def _is_field() -> bool:
    return bool((session.get('user') or {}).get('gw_field'))


def _hr_admin() -> bool:
    """HR 관리자인가 — 현장 세션은 절대 아님."""
    return not _is_field() and hr.is_hr_admin(_me())


def _my_person():
    """로그인한 사람의 인사카드 — 현장 세션은 인사카드 번호로, 그 외는 이메일로."""
    user = session.get('user') or {}
    if user.get('gw_field'):
        p = hr.get_person(int(user.get('gw_person_id') or 0))
        return p if p and p.get('status') != hr.STATUS_RESIGNED else None
    return hr.find_by_email(_me())


def gw_page(f):
    """그룹웨어 화면 — 로그인 + 새 화면을 쓸 수 있는 사람(정식 오픈 전엔 시범 명단).
    현장 세션은 링크 발급 때 이미 확인했고, 다닐 수 있는 경로는 groupware 문지기가 막는다."""
    @login_required
    @wraps(f)
    def wrapper(*a, **kw):
        if not (session.get('user') or {}).get('gw_field') and not gw.can_use(_me()):
            if request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': '아직 열리지 않은 화면입니다.'}), 403
            return redirect('/projects')
        return f(*a, **kw)
    return wrapper


def dev_preview_gate():
    """개발 중 화면(토요일 휴무·차량) 문 — 블루프린트 before_request 로 건다(11-21 부장님).
    명단(gw.DEV_PREVIEW_USERS) 밖이면 화면은 '아직 개발 중' 안내, API 는 403. 로그인 확인은 각 화면이 한다."""
    if not session.get('user') or gw.dev_preview_ok(_me()):
        return None
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'message': '아직 개발 중인 화면이에요.'}), 403
    return render_template('gw/dev_preview.html')


# 코드 버그로 보는 예외 — KeyError·IndexError 는 LookupError, UnicodeError·JSONDecodeError 는 ValueError 의
#   하위라 '없음'(404)·'입력 오류'(400)로 문구('days' 같은 키 이름)만 나가고 로그도 없었다(11-24 조용한 실패 점검).
#   서비스가 '없음'을 알릴 때는 LookupError·ValueError(또는 그 하위 SatOffError)를 직접 올리고,
#   KeyError·IndexError 를 일부러 올리는 곳은 없다(11-24 grep 확인). gw_satoff._err 도 같이 쓴다.
BUG_ERRORS = (KeyError, IndexError, UnicodeError, json.JSONDecodeError)


def _fail(e: Exception):
    if isinstance(e, BUG_ERRORS):
        logger.error(f'[HR] 처리 실패(코드 오류): {type(e).__name__}: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '처리하지 못했습니다. 잠시 후 다시 해 주세요.'}), 500
    if isinstance(e, PermissionError):
        return jsonify({'success': False, 'message': str(e)}), 403
    if isinstance(e, LookupError):
        return jsonify({'success': False, 'message': str(e)}), 404
    if isinstance(e, (ValueError, hr.SsnKeyMissing)):
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.error(f'[HR] 처리 실패: {e}', exc_info=True)
    return jsonify({'success': False, 'message': '처리하지 못했습니다. 잠시 후 다시 해 주세요.'}), 500


# ── 화면 ──────────────────────────────────────────────────────────────
@hr_bp.route('/hr/me')
@gw_page
def page_me():
    p = _my_person()
    return render_template('gw/hr_me.html', person_id=p['id'] if p else None, fields=hr.FIELDS)


@hr_bp.route('/hr/people')
@gw_page
def page_people():
    if not _hr_admin():
        return redirect('/hr/me')
    return render_template('gw/hr_people.html', fields=hr.FIELDS, statuses=hr.STATUSES)


@hr_bp.route('/gw/admin')
@gw_page
def page_admin():
    if not _is_admin():
        return redirect('/hr/me')
    return render_template('gw/admin.html')


# ── 인사 API ──────────────────────────────────────────────────────────
@hr_bp.route('/api/hr/people', methods=['GET'])
@gw_page
def api_list():
    if not _hr_admin():
        return jsonify({'success': False, 'message': '구성원 목록은 경영지원·대표만 볼 수 있습니다.'}), 403
    return jsonify({'success': True, 'people': hr.list_people(include_resigned=request.args.get('resigned') == '1')})


@hr_bp.route('/api/hr/people/<int:pid>', methods=['GET'])
@gw_page
def api_card(pid):
    me = '' if _is_field() else _me()      # 현장 세션은 관리자 권한 판단에 이메일을 쓰지 않음
    if not _hr_admin():
        mine = _my_person()
        if not mine or mine['id'] != pid:
            return jsonify({'success': False, 'message': '본인 인사카드만 볼 수 있습니다.'}), 403
    c = hr.card(pid, me)
    if not c:
        return jsonify({'success': False, 'message': '없는 구성원입니다.'}), 404
    from dashboard.services import hr_requests as hq
    mine = _my_person()
    if mine and mine['id'] == pid:
        c['can_self_edit'] = True                 # 연락처·비상연락망·서류 올리기는 본인이 바로(11-20)
        c['self_doc_types'] = hq.SELF_DOC_TYPES
        from dashboard.services import certificates as ct
        c['cert_options'] = ct.web_options()       # 그룹웨어 증명서 신청(11-20)
        c['cert_requests'] = [_cert_row(r) for r in ct.mine(mine['slack_id'])] if mine.get('slack_id') else []
    if _hr_admin():
        c['access_log'] = hr.access_log(pid)
        try:
            from dashboard.services import hr_tasks as ht
            c['tasks'] = [dict(t, title=ht.title(t), due_label=ht.due_label(t), guide=ht.KINDS[t['kind']][3])
                          for t in ht.for_person(pid) if t['kind'] in ht.KINDS]
        except Exception as e:
            logger.warning(f'[HR] 할 일 불러오기 실패 #{pid}: {e}')
    # 경력·학력·가족·서류 — 여기까지 왔으면 본인 또는 인사카드 관리자
    c.update(hr.private_sections(pid))
    return jsonify({'success': True, 'person': c})


def _can_see(pid: int) -> bool:
    if _hr_admin():
        return True
    mine = _my_person()
    return bool(mine) and mine['id'] == pid


@hr_bp.route('/api/hr/documents/<int:doc_id>', methods=['GET'])
@gw_page
def api_doc_download(doc_id):
    """서류 보기·내려받기 — 본인·인사카드 관리자. 기록은 남기지 않음(11-10 KiKO 결정)."""
    from flask import send_file
    d = hr.document_file(doc_id)
    if not d or not _can_see(d['person_id']):
        return jsonify({'success': False, 'message': '볼 수 없는 서류입니다.'}), 403 if d else 404
    import os
    if not os.path.exists(d['path']):
        return jsonify({'success': False, 'message': '파일이 서버에 없습니다. 경영지원에 문의해 주세요.'}), 404
    inline = request.args.get('view') == '1'
    resp = send_file(d['path'], as_attachment=not inline, download_name=d['file_name'], max_age=0)
    resp.headers['Cache-Control'] = 'no-store'
    return resp


@hr_bp.route('/api/hr/people/<int:pid>/documents', methods=['POST'])
@gw_page
def api_doc_upload(pid):
    f = request.files.get('file')
    if not f:
        return jsonify({'success': False, 'message': '파일을 골라 주세요.'}), 400
    try:
        hr.save_document(pid, request.form.get('doc_type', ''), f.filename or '', f.read(), _me() if not _is_field() else '')
        logger.info(f'[HR] 서류 올림 #{pid} by {_me()}')
        return jsonify({'success': True, 'documents': hr.documents(pid)})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/hr/people/<int:pid>/certificate', methods=['POST'])
@gw_page
def api_certificate(pid):
    """재직·경력증명서 바로 발급(11-20) — 인사카드 관리자만. 퇴직자·창구 요청용. PDF 를 돌려주고 인사카드 서류에도 남김."""
    if not _hr_admin():
        return jsonify({'success': False, 'message': '증명서 발급은 경영지원·대표만 할 수 있습니다.'}), 403
    from flask import send_file
    from dashboard.services import certificates as ct
    body = request.get_json(silent=True) or {}
    try:
        res = ct.issue_pdf(pid, body.get('ctype', ''), (body.get('purpose') or '').strip()[:100], _me(),
                           ssn=bool(body.get('ssn')), group_date=bool(body.get('group_date')),
                           company_key=body.get('company') or '')
    except ct.CertInputError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except hr.SsnKeyMissing:
        # 예전엔 'GW_HR_KEY 가 없어…' 원문이 400 으로 나가고 로그도 없었다(11-24 조용한 실패 점검)
        logger.error(f'[HR] 증명서 발급 실패 #{pid} — 주민번호 열쇠(GW_HR_KEY)가 서버에 없음', extra={'force_alert': True})
        return jsonify({'success': False, 'message': ct.SSN_KEY_MISSING_MSG}), 503
    except Exception as e:
        # PDF·글꼴·PNG·저장 실패, 코드 오류(KeyError 등) — 원문 대신 일반 문구 + ERROR(11-24 조용한 실패 점검)
        logger.error(f'[HR] 증명서 발급 실패 #{pid} by {_me()}: {type(e).__name__}: {e}', exc_info=True,
                     extra={'force_alert': True})
        return jsonify({'success': False, 'message': ct.ISSUE_FAILED_MSG}), 500
    logger.info(f"[HR] 증명서 발급 #{pid} {res['doc_no']} by {_me()}")
    resp = send_file(io.BytesIO(res['pdf']), mimetype='application/pdf', as_attachment=True,
                     download_name=res['file_name'], max_age=0)
    resp.headers['Cache-Control'] = 'no-store'
    resp.headers['X-Doc-No'] = res['doc_no']
    return resp


@hr_bp.route('/api/hr/documents/<int:doc_id>', methods=['DELETE'])
@gw_page
def api_doc_delete(doc_id):
    try:
        d = hr.document_file(doc_id)
        hr.delete_document(doc_id, _me() if not _is_field() else '')
        logger.info(f'[HR] 서류 지움 #{doc_id} by {_me()}')
        return jsonify({'success': True, 'documents': hr.documents(d['person_id']) if d else []})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/hr/families/<int:fid>/ssn', methods=['POST'])
@gw_page
def api_family_ssn(fid):
    try:
        return jsonify({'success': True, 'ssn': hr.reveal_family_ssn(fid, _me() if not _is_field() else '')})
    except Exception as e:
        return _fail(e)


# ── 본인 정보 변경 요청(11-20) — 내 정보에서 요청 → 경영지원 반영·반려 ─────────────
def _dm(uids, text: str) -> None:
    """전자결재 봇 DM(뒤에서) — 실패해도 화면은 그대로."""
    import threading

    def _bg():
        try:
            from dashboard.blueprints import slack_eapproval as se
            c = se.eap_client()
            if c:
                for u in uids:
                    se._notify(c, u, text)
        except Exception as e:
            logger.warning(f'[HR] 변경 요청 알림 실패: {e}')
    threading.Thread(target=_bg, daemon=True).start()


def _base_url() -> str:
    import os
    return ((os.getenv('GW_PUBLIC_URL') or '').strip() or 'https://pm.itg-aircon.com').rstrip('/')


def _requester_id() -> str:
    p = _my_person() or {}
    return (_me() if not _is_field() else '') or p.get('name') or ''


def _after_self_edit(row: dict) -> None:
    from dashboard.services import certificates as ct
    from dashboard.services import hr_requests as hq
    _dm(ct.handler_ids(), f"✏️ {row['name']} 님이 내 정보에서 {row['kind_label']}을(를) 고쳤어요 — {hq.summary(row)}\n"
                          f"보기: {_base_url()}/hr/people#{row['person_id']}")
    logger.info(f"[HR] 본인 수정 #{row['person_id']} {row['name']} {row['kind']}")


@hr_bp.route('/api/hr/people/<int:pid>/emergency', methods=['POST'])
@gw_page
def api_admin_emergency(pid):
    """비상연락망 고치기 — 경영지원·대표(11-20, 인사카드 화면). 본인은 /api/hr/me/emergency."""
    from dashboard.services import hr_requests as hq
    if not _hr_admin():
        return jsonify({'success': False, 'message': '경영지원·대표만 고칠 수 있습니다.'}), 403
    p = hr.get_person(pid)
    if not p:
        return jsonify({'success': False, 'message': '없는 구성원입니다.'}), 404
    try:
        hq.update_emergency(p, (request.get_json(silent=True) or {}).get('rows') or [], _me(), by_admin=True)
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.info(f'[HR] 비상연락망 고침 #{pid} by {_me()}')
    return jsonify({'success': True})


@hr_bp.route('/api/hr/me/<kind>', methods=['POST'])
@gw_page
def api_self_edit(kind):
    """내 정보 직접 고치기(11-20) — 연락처·비상연락망·서류 올리기만. 나머지는 경영지원(update_person)."""
    from dashboard.services import hr_requests as hq
    me = _my_person()
    if not me:
        return jsonify({'success': False, 'message': '인사카드가 연결되지 않았어요. 경영지원에 문의해 주세요.'}), 403
    try:
        if kind == 'contact':
            row = hq.update_info(me, request.get_json(silent=True) or {}, _requester_id())
        elif kind == 'emergency':
            row = hq.update_emergency(me, (request.get_json(silent=True) or {}).get('rows') or [], _requester_id())
        elif kind == 'documents':
            f = request.files.get('file')
            if not f:
                return jsonify({'success': False, 'message': '파일을 골라 주세요.'}), 400
            row = hq.upload_doc(me, request.form.get('doc_type', ''), f.filename or '', f.read(), _requester_id())
        else:
            return jsonify({'success': False, 'message': '고칠 수 없는 항목이에요.'}), 404
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    _after_self_edit(row)
    return jsonify({'success': True})


def _cert_row(r: dict) -> dict:
    from dashboard.services import certificates as ct
    return {'id': r['id'], 'ctype': r['ctype'], 'company': ct.company_label(r.get('company') or ''),
            'purpose': ct.purpose_text(r.get('purpose') or '', r.get('dest') or ''), 'status': r['status'],
            'status_label': ct.STATUS.get(r['status'], ('', r['status']))[1], 'note': r.get('handled_note') or '',
            'created_at': r['created_at']}


@hr_bp.route('/api/hr/me/cert-request', methods=['POST'])
@gw_page
def api_my_cert_request():
    """그룹웨어 증명서 신청(11-20) — 슬랙 [📜 증명서 신청]과 같은 장부·카드(신청자·경영지원 DM)."""
    from dashboard.services import certificates as ct
    me = _my_person()
    if not me:
        return jsonify({'success': False, 'message': '인사카드가 연결되지 않았어요. 경영지원에 문의해 주세요.'}), 403
    if not me.get('slack_id'):
        return jsonify({'success': False, 'message': '인사카드에 슬랙 계정이 연결돼 있어야 신청할 수 있어요. 경영지원에 말씀해 주세요.'}), 400
    try:
        vals = ct.web_vals(request.get_json(silent=True) or {})
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    from dashboard.blueprints import slack_eapproval as se
    from dashboard.blueprints.slack_cert import submit
    client = se.eap_client()
    if not client:
        return jsonify({'success': False, 'message': '지금은 신청할 수 없어요(전자결재 봇 연결 없음). 경영지원에 말씀해 주세요.'}), 503
    try:
        row = submit(client, me['slack_id'], vals)
    except Exception as e:
        logger.error(f'[HR] 증명서 신청 실패 #{me["id"]}: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '신청을 저장하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    # 경영지원 카드가 한 장도 안 갔으면 저장은 됐어도 화면에서 직접 말하라고 안내(11-24 조용한 실패 점검)
    notified = row.get('handler_notified') is not False
    out = {'success': True, 'notified': notified, 'requests': [_cert_row(r) for r in ct.mine(me['slack_id'])]}
    if not notified:
        from dashboard.blueprints.slack_cert import NOT_NOTIFIED_MSG
        out['warning'] = NOT_NOTIFIED_MSG
    return jsonify(out)


@hr_bp.route('/api/hr/me/cert-request/<int:rid>/cancel', methods=['POST'])
@gw_page
def api_my_cert_cancel(rid):
    from dashboard.services import certificates as ct
    me = _my_person()
    row = ct.get(rid)
    if not me or not me.get('slack_id') or not row or row['requester_id'] != me['slack_id']:
        return jsonify({'success': False, 'message': '없는 신청이에요.'}), 404
    new = ct.finish(rid, ct.CANCELLED, me['name'], '신청자 취소')
    if not new:
        return jsonify({'success': False, 'message': '이미 처리된 신청이에요.'}), 400
    try:
        from dashboard.blueprints import slack_eapproval as se
        from dashboard.blueprints.slack_cert import _refresh
        c = se.eap_client()
        if c:
            _refresh(c, new)
    except Exception as e:
        logger.warning(f'[HR] 증명서 취소 카드 고치기 실패 #{rid}: {e}')
    return jsonify({'success': True, 'requests': [_cert_row(r) for r in ct.mine(me['slack_id'])]})


def _sync_tasks() -> None:
    """입사일·퇴사일이 바뀌었을 수 있으니 입사·퇴사 할 일을 맞춘다(11-20) — 뒤에서, 실패해도 저장은 그대로."""
    try:
        from dashboard.blueprints.slack_hr_tasks import sync_async
        sync_async()
    except Exception as e:
        logger.warning(f'[HR] 할 일 맞추기 시작 실패: {e}')


@hr_bp.route('/api/hr/tasks/<int:tid>/done', methods=['POST'])
@gw_page
def api_task_done(tid):
    """입사·퇴사 할 일 [완료] — 인사카드 관리자. 슬랙 DM 카드도 같이 닫힘으로 바꾼다."""
    if not _hr_admin():
        return jsonify({'success': False, 'message': '경영지원·대표만 처리할 수 있습니다.'}), 403
    from dashboard.services import hr_tasks as ht
    me = hr.find_by_email(_me()) or {}
    row = ht.complete(tid, me.get('name') or _me())
    cur = row or ht.get(tid)
    if not cur:
        return jsonify({'success': False, 'message': '없는 할 일입니다.'}), 404
    if row:
        import threading

        def _bg():
            try:
                from dashboard.blueprints import slack_eapproval as se
                from dashboard.blueprints.slack_hr_tasks import _refresh
                c = se.eap_client()
                if c:
                    _refresh(c, row)
            except Exception as e:
                logger.warning(f'[HR] 할 일 슬랙 카드 고치기 실패 #{tid}: {e}')
        threading.Thread(target=_bg, daemon=True).start()
        logger.info(f'[HR] 할 일 완료 #{tid} by {_me()}')
    return jsonify({'success': True, 'tasks': ht.for_person(cur['person_id'])})


@hr_bp.route('/api/hr/people/<int:pid>', methods=['PUT'])
@gw_page
def api_update(pid):
    try:
        c = hr.update_person(pid, request.get_json(silent=True) or {}, _me())
        logger.info(f'[HR] 인사카드 수정 #{pid} by {_me()}')
        _sync_tasks()
        return jsonify({'success': True, 'person': c})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/hr/people', methods=['POST'])
@gw_page
def api_create():
    try:
        c = hr.create_person(request.get_json(silent=True) or {}, _me())
        logger.info(f'[HR] 구성원 추가 #{c["id"]} by {_me()}')
        _sync_tasks()
        return jsonify({'success': True, 'person': c})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/hr/people/<int:pid>/ssn', methods=['POST'])
@gw_page
def api_reveal_ssn(pid):
    try:
        full = hr.reveal_ssn(pid, _me())
        logger.info(f'[HR] 주민번호 전체 보기 #{pid} by {_me()}')
        return jsonify({'success': True, 'ssn': full})
    except Exception as e:
        return _fail(e)


# ── 그룹웨어 설정 API(관리자) ──────────────────────────────────────────
@hr_bp.route('/api/gw/roles', methods=['GET'])
@gw_page
def api_roles():
    if not _is_admin():
        return jsonify({'success': False, 'message': '관리자만 볼 수 있습니다.'}), 403
    return jsonify({'success': True, 'hr_admins': hr.hr_admins(), 'leave_viewers': hr.leave_viewers(),
                    'pilots': gw.list_pilots(), 'open_all': gw.is_open_all()})


@hr_bp.route('/api/gw/roles', methods=['POST'])
@gw_page
def api_set_roles():
    """{"key": "hr_admins"|"leave_viewers", "emails": [...]}"""
    if not _is_admin():
        return jsonify({'success': False, 'message': '관리자만 바꿀 수 있습니다.'}), 403
    data = request.get_json(silent=True) or {}
    key = data.get('key')
    if key not in ('hr_admins', 'leave_viewers'):
        return jsonify({'success': False, 'message': '바꿀 수 없는 항목입니다.'}), 400
    emails = sorted({str(e).strip().lower() for e in (data.get('emails') or []) if str(e).strip()})
    bad = [e for e in emails if '@' not in e]
    if bad:
        return jsonify({'success': False, 'message': f'이메일 형식이 아닙니다: {", ".join(bad)}'}), 400
    if key == 'hr_admins' and not emails:
        return jsonify({'success': False, 'message': '인사카드 관리자는 한 명 이상 있어야 합니다.'}), 400
    gw.set_setting(key, json.dumps(emails), _me())
    logger.info(f'[GW] {key} = {emails} by {_me()}')
    return jsonify({'success': True, key: emails})


# ── 현장 직원 로그인 번호·휴대폰(인사카드 관리자, 11-15) — services/gw_device.py ─────────
@hr_bp.route('/api/hr/people/<int:pid>/devices', methods=['GET'])
@gw_page
def api_devices(pid):
    from dashboard.services import gw_device
    if not _hr_admin():
        return jsonify({'success': False, 'message': '인사카드 관리자만 볼 수 있습니다.'}), 403
    p = hr.get_person(pid)
    if not p:
        return jsonify({'success': False, 'message': '없는 구성원입니다.'}), 404
    can = p.get('status') != hr.STATUS_RESIGNED and p.get('access') == hr.ACCESS_FIELD
    return jsonify({'success': True, 'can_register': can, 'devices': gw_device.list_for(pid) if can else [],
                    'idle_days': gw_device.IDLE_DAYS, 'pending': gw_device.pending_code(pid) if can else None,
                    'has_slack': bool((p.get('slack_id') or '').strip())})


@hr_bp.route('/api/hr/people/<int:pid>/device-code', methods=['POST'])
@gw_page
def api_device_code(pid):
    from dashboard.services import gw_device
    if not _hr_admin():
        return jsonify({'success': False, 'message': '인사카드 관리자만 만들 수 있습니다.'}), 403
    try:
        return jsonify({'success': True, **gw_device.issue_code(pid, _me())})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/hr/people/<int:pid>/device-code/dm', methods=['POST'])
@gw_page
def api_device_code_dm(pid):
    """방금 만든 로그인 번호를 그 직원 슬랙 DM 으로(11-20) — 단톡방 대신. 번호가 그 사람의 안 쓴 번호일 때만."""
    from dashboard.services import gw_device
    if not _hr_admin():
        return jsonify({'success': False, 'message': '인사카드 관리자만 보낼 수 있습니다.'}), 403
    p = hr.get_person(pid)
    code = ''.join(ch for ch in str((request.get_json(silent=True) or {}).get('code') or '') if ch.isdigit())
    if not p or not gw_device.is_pending(pid, code):
        return jsonify({'success': False, 'message': '이미 쓴 번호이거나 기한이 지났어요. 새 번호를 만들어 주세요.'}), 400
    sid = (p.get('slack_id') or '').strip()
    if not sid:
        return jsonify({'success': False, 'message': '인사카드에 슬랙 ID 가 없어요.'}), 400
    pend = gw_device.pending_code(pid) or {}
    until = (pend.get('expires_at') or '')[5:10].replace('-', '/')
    text = (f"📱 *그룹웨어 휴대폰 로그인 번호*: `{code[:4]} {code[4:]}`\n"
            f"{_base_url()} 를 휴대폰에서 열고 「번호로 로그인」에 넣어 주세요. "
            f"{until}까지, 한 번만 쓸 수 있어요. 한 번 로그인하면 그 휴대폰에서는 계속 열려요.\n"
            "홈 화면에 추가해 두면 앱처럼 열 수 있어요. 이 번호는 다른 사람에게 보여 주지 마세요.")
    try:
        from dashboard.blueprints.slack_leave import _client
        c = _client()
        ch = c.conversations_open(users=sid)['channel']['id']
        c.chat_postMessage(channel=ch, text=text)
    except Exception as e:
        logger.warning(f'[HR] 로그인 번호 DM 실패 #{pid}: {e}')
        return jsonify({'success': False, 'message': '슬랙으로 보내지 못했어요. 번호를 직접 1:1 로 전해 주세요.'}), 502
    logger.info(f'[HR] 로그인 번호 DM #{pid} by {_me()}')
    return jsonify({'success': True})


@hr_bp.route('/api/hr/people/<int:pid>/devices/<int:did>/revoke', methods=['POST'])
@gw_page
def api_device_revoke(pid, did):
    from dashboard.services import gw_device
    if not _hr_admin():
        return jsonify({'success': False, 'message': '인사카드 관리자만 끊을 수 있습니다.'}), 403
    if not gw_device.revoke(did, _me(), person_id=pid):
        return jsonify({'success': False, 'message': '이미 로그아웃된 휴대폰입니다.'}), 400
    return jsonify({'success': True, 'devices': gw_device.list_for(pid)})


# ── 휴가(조회) ────────────────────────────────────────────────────────
# 승인·반려는 슬랙 휴가 봇. 본인 신청·취소는 여기서도(11-15) — 아래 /api/leave/apply·cancel.
def _leave_all() -> bool:
    """전 직원 휴가를 볼 수 있나 — 인사카드 관리자·휴가 현황 열람자. 현장 세션은 아님."""
    return not _is_field() and hr.can_view_all_leave(_me())


@hr_bp.route('/leave')
@gw_page
def page_leave_me():
    p = _my_person()
    return render_template('gw/leave_me.html', person_id=p['id'] if p else None)


@hr_bp.route('/leave/team')
@gw_page
def page_leave_team():
    if not _leave_all():
        return redirect('/leave')
    return render_template('gw/leave_team.html')


# ── 쉬는 날(11-25 부장님 — 플렉스 「쉬는 날」 대체) ─────────────────────────
# 보기·넣기·고치기·지우기 = 관리자(admin·super_admin)만(11-25 부장님 '어드민만 보이게'). 다른 직원은 휴가 달력의 빨간 날로만 봄.
def _holiday_body() -> dict:
    b = request.get_json(silent=True) or {}
    return {k: b.get(k) for k in ('name', 'kind', 'start', 'end', 'yearly', 'note')}


def _holiday_actor() -> str:
    """기록에 남길 이름 — 인사카드 이름, 없으면 로그인 이름."""
    try:
        p = _my_person()
    except Exception:
        p = None
    u = session.get('user') or {}
    return (p or {}).get('name') or u.get('name') or _me()


@hr_bp.route('/leave/holidays')
@gw_page
def page_leave_holidays():
    if not _holiday_admin():
        return redirect('/leave')
    return render_template('gw/leave_holidays.html', can_edit=True)


def _holiday_admin() -> bool:
    return _is_admin() and not _is_field()


@hr_bp.route('/api/leave/holidays', methods=['GET'])
@gw_page
def api_holidays():
    from datetime import date as _date
    from dashboard.services import company_holidays as ch
    if not _holiday_admin():
        return jsonify({'success': False, 'message': '쉬는 날은 관리자만 볼 수 있어요.'}), 403
    try:
        year = int(request.args.get('year') or _date.today().year)
    except ValueError:
        return jsonify({'success': False, 'message': '연도 형식이 맞지 않아요.'}), 400
    if not (ch.YEAR_MIN <= year <= ch.YEAR_MAX):
        return jsonify({'success': False, 'message': '연도를 다시 골라 주세요.'}), 400
    try:
        data = ch.year_list(year)
    except Exception as e:
        return _fail(e)
    return jsonify(dict(data, success=True, can_edit=True, today=_date.today().isoformat()))


def _holiday_write(action: str, hid: int = 0):
    from dashboard.services import company_holidays as ch
    if not _holiday_admin():
        return jsonify({'success': False, 'message': '쉬는 날은 관리자만 넣고 고칠 수 있어요.'}), 403
    body = request.get_json(silent=True) or {}
    try:
        if body.get('preview'):
            return jsonify(dict(ch.preview(action, _holiday_body(), hid), success=True))
        expect = body.get('expect')
        res = ch.save(action, _holiday_body(), _me(), _holiday_actor(), hid,
                      expect=int(expect) if expect is not None else None)
    except Exception as e:
        return _fail(e)
    return jsonify(dict(res, success=True))


@hr_bp.route('/api/leave/holidays/add', methods=['POST'])
@gw_page
def api_holiday_add():
    return _holiday_write('add')


@hr_bp.route('/api/leave/holidays/<int:hid>/edit', methods=['POST'])
@gw_page
def api_holiday_edit(hid):
    return _holiday_write('edit', hid)


@hr_bp.route('/api/leave/holidays/<int:hid>/delete', methods=['POST'])
@gw_page
def api_holiday_delete(hid):
    return _holiday_write('delete', hid)


@hr_bp.route('/api/leave/me', methods=['GET'])
@gw_page
def api_leave_me():
    """내 휴가 + 신청 창 정보(form) + 기록마다 본인 취소 단추(cancel)."""
    from dashboard.services import leave_view, leave_web
    p = _my_person()
    if not p:
        return jsonify({'success': False, 'message': '아직 인사카드가 연결되지 않았어요. 경영지원에 말씀해 주세요.'}), 404
    user = session.get('user') or {}
    try:
        form = leave_web.form_info(user)
    except Exception as e:
        logger.warning(f'[GW휴가] 신청 창 정보 실패: {e}', exc_info=True)
        form = {'can_apply': False, 'why': '휴가 신청 정보를 불러오지 못했어요. 잠시 뒤 다시 열어 주세요.'}
    summary = leave_web.annotate(leave_view.summary(p), user)
    try:
        summary['promotions'] = leave_web.promotions_for(user)   # 연차 사용 촉진 문서(11-18)
    except Exception as e:
        logger.warning(f'[GW휴가] 촉진 정보 실패: {e}', exc_info=True)
        summary['promotions'] = []
    return jsonify({'success': True, 'leave': summary, 'form': form})


# ── 휴가(본인 신청·취소) — 슬랙 휴가 앱과 같은 장부·같은 알림 ──────────────
@hr_bp.route('/api/leave/apply', methods=['POST'])
@gw_page
def api_leave_apply():
    from dashboard.services import leave_web
    try:
        row, errs = leave_web.apply(session.get('user') or {}, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if errs:
        msg = errs.pop('_', '') or next(iter(errs.values()))
        return jsonify({'success': False, 'message': msg, 'errors': errs}), 400
    return jsonify({'success': True, 'id': row['id'], 'status': row['status'], 'days': row['days']})


@hr_bp.route('/api/leave/preview', methods=['POST'])
@gw_page
def api_leave_preview():
    """신청 창 일수 미리보기 — 날짜를 고를 때마다(고른 토요일이 일수에 드는지 미리 보이게, 11-24)."""
    from dashboard.services import leave_web
    try:
        return jsonify({'success': True, **leave_web.preview(session.get('user') or {},
                                                             request.get_json(silent=True) or {})})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/leave/<int:rid>/cancel', methods=['POST'])
@gw_page
def api_leave_cancel(rid):
    from dashboard.services import leave_web
    try:
        msg = leave_web.cancel(session.get('user') or {}, rid)
    except Exception as e:
        return _fail(e)
    if msg:
        return jsonify({'success': False, 'message': msg}), 400
    return jsonify({'success': True})


@hr_bp.route('/api/leave/<int:rid>/edit', methods=['POST'])
@gw_page
def api_leave_edit(rid):
    """본인 고치기(승인 대기)·변경 신청(승인된 시작 전 휴가) — 11-18."""
    from dashboard.services import leave_web
    try:
        row, errs, mode = leave_web.edit(session.get('user') or {}, rid, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if errs:
        msg = errs.pop('_', '') or next(iter(errs.values()))
        return jsonify({'success': False, 'message': msg, 'errors': errs}), 400
    return jsonify({'success': True, 'id': row['id'], 'status': row['status'], 'days': row['days'], 'mode': mode})


# ── 휴가(경영지원) — 고치기·취소·연차 조정·관리자 부여 휴가 등록(11-18, 플렉스 관리자 기능 맞추기) ─────
def _leave_person_name(pid: int):
    from dashboard.services import leave_view
    p = hr.get_person(pid)
    return p, (leave_view.leave_name(p) if p else None)


@hr_bp.route('/api/leave/<int:rid>/admin-edit', methods=['POST'])
@gw_page
def api_leave_admin_edit(rid):
    from dashboard.services import leave_web
    try:
        row, errs = leave_web.admin_edit(session.get('user') or {}, rid, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if errs:
        msg = errs.pop('_', '') or next(iter(errs.values()))
        return jsonify({'success': False, 'message': msg, 'errors': errs}), 400
    return jsonify({'success': True, 'id': row['id'], 'days': row['days']})


@hr_bp.route('/api/leave/<int:rid>/admin-cancel', methods=['POST'])
@gw_page
def api_leave_admin_cancel(rid):
    from dashboard.services import leave_web
    try:
        msg = leave_web.admin_cancel(session.get('user') or {}, rid)
    except Exception as e:
        return _fail(e)
    if msg:
        return jsonify({'success': False, 'message': msg}), 400
    return jsonify({'success': True})


@hr_bp.route('/api/leave/person/<int:pid>/adjust', methods=['POST'])
@gw_page
def api_leave_adjust(pid):
    from dashboard.services import leave_web
    p, name = _leave_person_name(pid)
    if not name:
        return jsonify({'success': False, 'message': '휴가 장부에 연결되지 않은 직원이에요.'}), 404
    try:
        g, err = leave_web.adjust(session.get('user') or {}, name, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if err:
        return jsonify({'success': False, 'message': err}), 400
    return jsonify({'success': True, 'days': g['days']})


@hr_bp.route('/api/leave/adjust-bulk', methods=['POST'])
@gw_page
def api_leave_adjust_bulk():
    """전 직원 연차 부여·차감(플렉스 같은 표) — 모두 맞을 때만 한꺼번에 넣는다."""
    from dashboard.services import leave_web
    try:
        made, errs = leave_web.adjust_bulk(session.get('user') or {}, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if errs:
        msg = errs.pop('_', '') or '고칠 칸이 있어요. 빨간 글을 확인해 주세요.'
        return jsonify({'success': False, 'message': msg, 'errors': errs}), 400
    return jsonify({'success': True, 'count': len(made)})


@hr_bp.route('/api/leave/person/<int:pid>/custom-grant', methods=['POST'])
@gw_page
def api_leave_custom_grant(pid):
    """맞춤 휴가 부여(보건·난임 등 일수를 먼저 주면 직원이 신청) — 11-18."""
    from dashboard.services import leave_web
    p, name = _leave_person_name(pid)
    if not name:
        return jsonify({'success': False, 'message': '휴가 장부에 연결되지 않은 직원이에요.'}), 404
    try:
        g, err = leave_web.custom_grant(session.get('user') or {}, name, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if err:
        return jsonify({'success': False, 'message': err}), 400
    return jsonify({'success': True, 'id': g['id']})


@hr_bp.route('/api/leave/custom-grant/<int:gid>/cancel', methods=['POST'])
@gw_page
def api_leave_custom_grant_cancel(gid):
    from dashboard.services import leave_web
    try:
        msg = leave_web.cancel_custom_grant(session.get('user') or {}, gid)
    except Exception as e:
        return _fail(e)
    if msg:
        return jsonify({'success': False, 'message': msg}), 400
    return jsonify({'success': True})


@hr_bp.route('/api/leave/adjust/<int:gid>/cancel', methods=['POST'])
@gw_page
def api_leave_adjust_cancel(gid):
    from dashboard.services import leave_web
    try:
        g, err = leave_web.cancel_adjust(session.get('user') or {}, gid)
    except Exception as e:
        return _fail(e)
    if err:
        return jsonify({'success': False, 'message': err}), 400
    return jsonify({'success': True})


@hr_bp.route('/api/leave/person/<int:pid>/register', methods=['POST'])
@gw_page
def api_leave_register(pid):
    from dashboard.services import leave_web
    p, name = _leave_person_name(pid)
    if not name:
        return jsonify({'success': False, 'message': '휴가 장부에 연결되지 않은 직원이에요.'}), 404
    try:
        row, errs = leave_web.admin_register(session.get('user') or {}, name, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if errs:
        msg = errs.pop('_', '') or next(iter(errs.values()))
        return jsonify({'success': False, 'message': msg, 'errors': errs}), 400
    return jsonify({'success': True, 'id': row['id'], 'days': row['days']})


@hr_bp.route('/api/leave/person/<int:pid>/settlement', methods=['GET'])
@gw_page
def api_leave_settlement(pid):
    """퇴직 정산(회계연도 기준 ↔ 입사일 기준 차이·미사용 연차) — 경영지원만, last_day=마지막 근무일."""
    from dashboard.services import leave_web
    p = hr.get_person(pid)
    if not p:
        return jsonify({'success': False, 'message': '없는 구성원입니다.'}), 404
    try:
        out = leave_web.settlement(session.get('user') or {}, p, request.args.get('last_day') or '')
    except Exception as e:
        return _fail(e)
    if not out.get('ok'):
        return jsonify({'success': False, 'message': out.get('why') or '계산하지 못했어요.'}), 400
    return jsonify({'success': True, 'settlement': out})


@hr_bp.route('/api/leave/person/<int:pid>', methods=['GET'])
@gw_page
def api_leave_person(pid):
    from dashboard.services import leave_view, leave_web
    mine = _my_person()
    if not _leave_all() and not (mine and mine['id'] == pid):
        return jsonify({'success': False, 'message': '본인 휴가만 볼 수 있습니다.'}), 403
    p = hr.get_person(pid)
    if not p:
        return jsonify({'success': False, 'message': '없는 구성원입니다.'}), 404
    admin = {'can': False}
    if _leave_all():
        try:
            admin = leave_web.admin_info(session.get('user') or {}, p)
        except Exception as e:
            logger.warning(f'[GW휴가] 경영지원 칸 정보 실패: {e}', exc_info=True)
    return jsonify({'success': True, 'leave': leave_view.summary(p), 'admin': admin})


@hr_bp.route('/api/leave/team', methods=['GET'])
@gw_page
def api_leave_team():
    from dashboard.services import leave_view
    if not _leave_all():
        return jsonify({'success': False, 'message': '전 직원 휴가 현황은 볼 수 없는 계정입니다.'}), 403
    from dashboard.services import leave_web
    return jsonify({'success': True, 'people': leave_view.team(),
                    'can_manage': bool(leave_web.manager_sid(session.get('user') or {}))})


@hr_bp.route('/api/leave/export', methods=['GET'])
@gw_page
def api_leave_export():
    """엑셀 — kind=balance(보유 현황)|usage(사용 내역)|monthly(월별 연차 사용), year. 전 직원 휴가 열람자만(11-18)."""
    from datetime import date as _date
    from flask import send_file
    from dashboard.services import leave_export
    if not _leave_all():
        return jsonify({'success': False, 'message': '전 직원 휴가는 내려받을 수 없는 계정입니다.'}), 403
    kind = request.args.get('kind') or 'balance'
    try:
        year = int(request.args.get('year') or _date.today().year)
        if not 2020 <= year <= 2100 or kind not in leave_export.KINDS:
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({'success': False, 'message': '내려받을 종류·연도를 다시 골라 주세요.'}), 400
    try:
        data, fname = leave_export.build(kind, year)
    except Exception as e:
        return _fail(e)
    logger.info(f'[GW휴가] 엑셀 내려받기 {kind} {year} ({_me()})')
    resp = send_file(io.BytesIO(data), as_attachment=True, download_name=fname, max_age=0,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp.headers['Cache-Control'] = 'no-store'
    return resp


@hr_bp.route('/api/leave/monthly', methods=['GET'])
@gw_page
def api_leave_monthly():
    """직원별 월별 연차 상세 — pid(없으면 본인), year. 본인 또는 전 직원 열람자."""
    from datetime import date as _date
    from dashboard.services import leave as lv
    from dashboard.services import leave_view
    mine = _my_person()
    try:
        pid = int(request.args.get('pid') or (mine or {}).get('id') or 0)
        year = int(request.args.get('year') or _date.today().year)
        if not 2020 <= year <= 2100:
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({'success': False, 'message': '연도를 다시 골라 주세요.'}), 400
    if not _leave_all() and not (mine and mine['id'] == pid):
        return jsonify({'success': False, 'message': '본인 휴가만 볼 수 있습니다.'}), 403
    p = hr.get_person(pid)
    if not p:
        return jsonify({'success': False, 'message': '없는 구성원입니다.'}), 404
    name = leave_view.leave_name(p)
    if not name:
        return jsonify({'success': True, 'monthly': {'year': year, 'months': [], 'opening': 0}})
    return jsonify({'success': True, 'monthly': lv.monthly_detail(name, year)})


@hr_bp.route('/api/leave/promotions', methods=['GET'])
@gw_page
def api_leave_promotions():
    """연차 촉진 탭 — 전 직원 휴가 열람자만(처리 단추는 경영지원만)."""
    from dashboard.services import leave_web
    if not _leave_all():
        return jsonify({'success': False, 'message': '연차 촉진 기록은 볼 수 없는 계정입니다.'}), 403
    try:
        return jsonify({'success': True, **leave_web.promo_list(session.get('user') or {})})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/leave/promotions/send', methods=['POST'])
@gw_page
def api_leave_promo_send():
    from dashboard.services import leave_web
    try:
        rec, err = leave_web.promo_send(session.get('user') or {}, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if err:
        return jsonify({'success': False, 'message': err}), 400
    return jsonify({'success': True, 'id': rec['id']})


@hr_bp.route('/api/leave/promotions/<int:pid>/designate', methods=['POST'])
@gw_page
def api_leave_promo_designate(pid):
    from dashboard.services import leave_web
    try:
        rec, err = leave_web.promo_designate(session.get('user') or {}, pid, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if err:
        return jsonify({'success': False, 'message': err}), 400
    return jsonify({'success': True})


@hr_bp.route('/api/leave/promotions/<int:pid>/ack', methods=['POST'])
@gw_page
def api_leave_promo_ack(pid):
    from dashboard.services import leave_web
    try:
        msg = leave_web.promo_ack(session.get('user') or {}, pid)
    except Exception as e:
        return _fail(e)
    if msg:
        return jsonify({'success': False, 'message': msg}), 400
    return jsonify({'success': True})


@hr_bp.route('/api/leave/promotions/<int:pid>/plan', methods=['POST'])
@gw_page
def api_leave_promo_plan(pid):
    from dashboard.services import leave_web
    try:
        rec, err = leave_web.promo_plan(session.get('user') or {}, pid, request.get_json(silent=True) or {})
    except Exception as e:
        return _fail(e)
    if err:
        return jsonify({'success': False, 'message': err}), 400
    return jsonify({'success': True})


@hr_bp.route('/api/leave/calendar', methods=['GET'])
@gw_page
def api_leave_calendar():
    from datetime import date as _date
    from dashboard.services import leave_view
    if not _leave_all():
        return jsonify({'success': False, 'message': '전 직원 휴가 달력은 볼 수 없는 계정입니다.'}), 403
    month = request.args.get('month') or _date.today().strftime('%Y-%m')
    try:
        y, m = (int(x) for x in month.split('-'))
        _date(y, m, 1)
    except (ValueError, TypeError):
        return jsonify({'success': False, 'message': '달 형식은 YYYY-MM 입니다.'}), 400
    try:
        from dashboard.services import sat_off
        sat = sat_off.calendar_month(y, m)          # 기술영업 토요일 휴무(11-20)
    except Exception as e:
        logger.warning(f'[GW휴가] 토요일 휴무 달력 실패: {e}')
        sat = []
    try:
        from dashboard.services import company_holidays as ch
        hol = ch.month_days(y, m)                   # 쉬는 날(11-25) — 달력에 빨간 날로
    except Exception as e:
        logger.warning(f'[GW휴가] 쉬는 날 달력 실패: {e}')
        hol = []
    return jsonify({'success': True, 'month': f'{y:04d}-{m:02d}', 'entries': leave_view.calendar(y, m), 'sat_off': sat,
                    'holidays': hol})


# ── 홈 ────────────────────────────────────────────────────────────────
@hr_bp.route('/home')
@gw_page
def page_home():
    if _is_field():
        return redirect('/leave')   # 현장·미화 직원은 앱을 열면 바로 내 휴가(11-15, 플렉스 앱처럼)
    p = _my_person()
    name = (p or {}).get('name') or (session.get('user') or {}).get('name', '')
    return render_template('gw/home.html', person_name=name, is_field=_is_field())


@hr_bp.route('/api/home', methods=['GET'])
@gw_page
def api_home():
    """결재·내 휴가·휴가자. 메일은 느릴 수 있어 /api/home/mail 로 따로."""
    from dashboard.services import gw_home, leave_view
    p = _my_person()
    out = {'success': True, 'field': _is_field()}
    try:
        out['leave'] = leave_view.summary(p) if p else None
        if out['leave']:
            out['leave'] = {k: out['leave'][k] for k in ('linked', 'balance', 'grants', 'pending', 'upcoming', 'annual_used')}
    except Exception as e:
        # 실패도 '인사카드 없음'(leave=None)과 같아 '경영지원에 연결 요청'이 떴다 → 따로 표시(11-24 조용한 실패 점검)
        logger.error(f'[GW-HOME] 휴가 요약 실패: {type(e).__name__}: {e}', exc_info=True)
        out['leave'] = None
        out['leave_error'] = True
    # 결재·승인 목록은 슬랙이 확인한 본인 ID 로만 (인사카드 slack_id 바꿔치기 방지, 2026-11-13)
    from dashboard.services.gw_approval import slack_id_status
    _sid, _sid_state = ('', 'unlinked') if _is_field() else slack_id_status(_me(), (p or {}).get('slack_id') or '')
    if not _is_field():
        try:
            out['approvals'] = gw_home.approvals(_sid)
            if out['approvals'] is not None and _sid_state == 'unavailable':
                out['approvals']['unavailable'] = True   # 연결 안 됨이 아니라 슬랙 확인이 잠깐 안 됨(11-15)
        except Exception as e:
            logger.warning(f'[GW-HOME] 결재 목록 실패: {e}')
            out['approvals'] = None
        try:
            out['off'] = gw_home.people_off()
        except Exception as e:
            logger.warning(f'[GW-HOME] 휴가자 실패: {e}')
            out['off'] = None
    try:   # 할 일 카드(11-10) — 현장은 연말정산만
        out['todos'] = gw_home.todos(p, _sid, _is_field(), _hr_admin(),
                                     out.get('approvals'))
    except Exception as e:
        logger.warning(f'[GW-HOME] 할 일 실패: {e}')
        out['todos'] = []
    return jsonify(out)


# ── 알림 🔔(11-10) — 이미 쌓인 기록에서 나와 관련된 최근 일. 현장은 휴가·연말정산만 ─────────────
@hr_bp.route('/api/gw/notify', methods=['GET'])
@gw_page
def api_notify():
    from dashboard.services import gw_notify
    return jsonify({'success': True, **gw_notify.feed(_me(), _my_person(), _is_field(), _hr_admin())})


@hr_bp.route('/api/gw/notify/seen', methods=['POST'])
@gw_page
def api_notify_seen():
    from dashboard.services import gw_notify
    return jsonify({'success': True, 'seen_at': gw_notify.mark_seen(gw_notify.who_key(_me(), _my_person(), _is_field()))})


@hr_bp.route('/api/home/sales', methods=['GET'])
@gw_page
def api_home_sales():
    """영업 숫자 칸·그래프(11-10) — 문의·방문·계약. 현장 계정은 없음."""
    from dashboard.services import gw_home
    if _is_field():
        return jsonify({'success': False, 'message': '현장 계정은 영업 숫자 칸이 없습니다.'}), 403
    try:
        return jsonify({'success': True, 'sales': gw_home.sales()})
    except Exception as e:
        # 원천(리드 시트·공사현황)을 못 읽어도 여기로 온다 — 0건으로 보이던 것(11-24 조용한 실패 점검)이라 ERROR
        logger.error(f'[GW-HOME] 영업 숫자 실패: {type(e).__name__}: {e}',
                     exc_info=not isinstance(e, gw_home.SalesSourceError))
        return jsonify({'success': False, 'message': '영업 숫자를 불러오지 못했습니다.'}), 500


@hr_bp.route('/api/home/mail', methods=['GET'])
@gw_page
def api_home_mail():
    from dashboard.services import gw_home
    if _is_field():
        return jsonify({'success': False, 'message': '현장 계정은 메일 칸이 없습니다.'}), 403
    return jsonify({'success': True, 'mail': gw_home.mail_summary(_me())})


# ── 연말정산 ──────────────────────────────────────────────────────────
def _ye_year():
    from dashboard.services import yearend as ye
    try:
        return int(request.args.get('year') or request.form.get('year') or (request.get_json(silent=True) or {}).get('year') or ye.current_year())
    except (TypeError, ValueError):
        return ye.current_year()


def _who() -> str:
    return _me() if not _is_field() else f'field:{(session.get("user") or {}).get("gw_person_id")}'


@hr_bp.route('/yearend')
@gw_page
def page_yearend():
    p = _my_person()
    return render_template('gw/yearend.html', person_id=p['id'] if p else None, is_hr_admin=_hr_admin())


@hr_bp.route('/yearend/admin')
@gw_page
def page_yearend_admin():
    if not _hr_admin():
        return redirect('/yearend')
    return render_template('gw/yearend_admin.html')


@hr_bp.route('/api/yearend/me', methods=['GET'])
@gw_page
def api_ye_me():
    from dashboard.services import yearend as ye
    p = _my_person()
    if not p:
        return jsonify({'success': False, 'message': '아직 인사카드가 연결되지 않았어요. 경영지원에 말씀해 주세요.'}), 404
    v = ye.my_view(p, _ye_year())
    v['problems'] = ye.problems(v)
    return jsonify({'success': True, 'ye': v})


@hr_bp.route('/api/yearend/me', methods=['POST'])
@gw_page
def api_ye_save():
    from dashboard.services import yearend as ye
    p = _my_person()
    if not p:
        return jsonify({'success': False, 'message': '아직 인사카드가 연결되지 않았어요. 경영지원에 말씀해 주세요.'}), 404
    try:
        ye.save_answers(p['id'], _ye_year(), request.get_json(silent=True) or {})
        v = ye.my_view(p, _ye_year())
        v['problems'] = ye.problems(v)
        return jsonify({'success': True, 'ye': v})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/yearend/me/submit', methods=['POST'])
@gw_page
def api_ye_submit():
    from dashboard.services import yearend as ye
    p = _my_person()
    try:
        ye.submit(p['id'], _ye_year(), p)
        logger.info(f'[연말정산] 제출 #{p["id"]}')
        return api_ye_me()
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/yearend/me/reopen', methods=['POST'])
@gw_page
def api_ye_reopen():
    from dashboard.services import yearend as ye
    p = _my_person()
    try:
        ye.reopen(p['id'], _ye_year())
        return api_ye_me()
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/yearend/me/documents', methods=['POST'])
@gw_page
def api_ye_upload():
    from dashboard.services import yearend as ye
    p = _my_person()
    f = request.files.get('file')
    if not p or not f:
        return jsonify({'success': False, 'message': '파일을 골라 주세요.'}), 400
    try:
        ye.upload(p['id'], _ye_year(), request.form.get('kind', ''), f.filename or '', f.read(), _who())
        return api_ye_me()
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/yearend/me/documents/<int:doc_id>', methods=['DELETE'])
@gw_page
def api_ye_delete(doc_id):
    from dashboard.services import yearend as ye
    p = _my_person()
    try:
        ye.delete_upload(p['id'], _ye_year(), doc_id, _who())
        return api_ye_me()
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/yearend/admin', methods=['GET'])
@gw_page
def api_ye_admin():
    from dashboard.services import yearend as ye
    if not _hr_admin():
        return jsonify({'success': False, 'message': '경영지원·대표만 볼 수 있습니다.'}), 403
    y = _ye_year()
    return jsonify({'success': True, 'year': y, 'season': ye.season(y), 'rows': ye.admin_rows(y)})


@hr_bp.route('/api/yearend/admin/season', methods=['POST'])
@gw_page
def api_ye_season():
    from dashboard.services import yearend as ye
    data = request.get_json(silent=True) or {}
    try:
        s = ye.set_season(_ye_year(), _me() if not _is_field() else '', status=data.get('status'),
                          due_date=data.get('due_date'), notice=data.get('notice'))
        logger.info(f'[연말정산] 설정 {s} by {_me()}')
        return jsonify({'success': True, 'season': s})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/yearend/admin/<int:pid>/status', methods=['POST'])
@gw_page
def api_ye_status(pid):
    from dashboard.services import yearend as ye
    data = request.get_json(silent=True) or {}
    try:
        sub = ye.set_status(pid, _ye_year(), data.get('status', ''), _me() if not _is_field() else '', data.get('reason', ''))
        resp = {'success': True, 'submission': sub}
        if sub['status'] == ye.RETURNED:
            p = hr.get_person(pid)
            # DM 을 못 보내도 보완 요청은 저장됨 — 화면이 '직접 알려 주세요'를 띄우도록 결과를 같이 보냄
            resp['notified'] = ye._dm((p or {}).get('slack_id') or '',
                                      f'*{_ye_year()}년 귀속 연말정산* 자료 보완을 요청드려요.\n사유: {sub["returned_reason"]}\n'
                                      '그룹웨어 → 인사 → 연말정산에서 고친 뒤 다시 제출해 주세요.')
        return jsonify(resp)
    except ye.StatusConflict as e:     # 화면을 연 뒤 직원이 상태를 바꿈 — 화면이 새로 불러옴(11-24 조용한 실패 점검)
        return jsonify({'success': False, 'message': str(e), 'conflict': True}), 409
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/yearend/admin/notify', methods=['POST'])
@gw_page
def api_ye_notify():
    from dashboard.services import yearend as ye
    try:
        return jsonify({'success': True, **ye.notify_pending(_ye_year(), _me() if not _is_field() else '')})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/yearend/admin/roster.xlsx', methods=['GET'])
@gw_page
def api_ye_roster():
    from flask import send_file
    from dashboard.services import yearend as ye
    try:
        data = ye.roster_xlsx(_ye_year(), _me() if not _is_field() else '')
    except Exception as e:
        return _fail(e)
    resp = send_file(io.BytesIO(data), as_attachment=True, download_name=f'연말정산_{_ye_year()}귀속_명단.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp.headers['Cache-Control'] = 'no-store'
    return resp


@hr_bp.route('/api/yearend/admin/bundle.zip', methods=['GET'])
@gw_page
def api_ye_bundle():
    from flask import send_file
    from dashboard.services import yearend as ye
    try:
        data, missing = ye.bundle_zip_report(_ye_year(), _me() if not _is_field() else '')
    except Exception as e:
        return _fail(e)
    resp = send_file(io.BytesIO(data), as_attachment=True, download_name=f'연말정산_{_ye_year()}귀속_자료.zip',
                     mimetype='application/zip')
    resp.headers['Cache-Control'] = 'no-store'
    resp.headers['X-Missing-Files'] = str(missing)   # 화면이 읽어 '빠진 파일 N개' 안내(11-24 조용한 실패 점검)
    return resp


@hr_bp.route('/api/yearend/admin/<int:pid>/result', methods=['POST'])
@gw_page
def api_ye_result(pid):
    from dashboard.services import yearend as ye
    f = request.files.get('file')
    if not f:
        return jsonify({'success': False, 'message': '파일을 골라 주세요.'}), 400
    try:
        ye.upload_result(pid, _ye_year(), f.filename or '', f.read(), _me() if not _is_field() else '')
        sent = ye.notify_result(pid, _ye_year())
        return jsonify({'success': True, 'notified': sent})
    except Exception as e:
        return _fail(e)
