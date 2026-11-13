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


def _fail(e: Exception):
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
    if _hr_admin():
        c['access_log'] = hr.access_log(pid)
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


@hr_bp.route('/api/hr/people/<int:pid>', methods=['PUT'])
@gw_page
def api_update(pid):
    try:
        c = hr.update_person(pid, request.get_json(silent=True) or {}, _me())
        logger.info(f'[HR] 인사카드 수정 #{pid} by {_me()}')
        return jsonify({'success': True, 'person': c})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/hr/people', methods=['POST'])
@gw_page
def api_create():
    try:
        c = hr.create_person(request.get_json(silent=True) or {}, _me())
        logger.info(f'[HR] 구성원 추가 #{c["id"]} by {_me()}')
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


# ── 휴가(조회) ────────────────────────────────────────────────────────
# 신청·승인·취소는 슬랙 휴가 봇. 여기서는 보기만.
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


@hr_bp.route('/api/leave/me', methods=['GET'])
@gw_page
def api_leave_me():
    from dashboard.services import leave_view
    p = _my_person()
    if not p:
        return jsonify({'success': False, 'message': '인사카드가 연결되지 않았습니다.'}), 404
    return jsonify({'success': True, 'leave': leave_view.summary(p)})


@hr_bp.route('/api/leave/person/<int:pid>', methods=['GET'])
@gw_page
def api_leave_person(pid):
    from dashboard.services import leave_view
    mine = _my_person()
    if not _leave_all() and not (mine and mine['id'] == pid):
        return jsonify({'success': False, 'message': '본인 휴가만 볼 수 있습니다.'}), 403
    p = hr.get_person(pid)
    if not p:
        return jsonify({'success': False, 'message': '없는 구성원입니다.'}), 404
    return jsonify({'success': True, 'leave': leave_view.summary(p)})


@hr_bp.route('/api/leave/team', methods=['GET'])
@gw_page
def api_leave_team():
    from dashboard.services import leave_view
    if not _leave_all():
        return jsonify({'success': False, 'message': '전 직원 휴가 현황은 볼 수 없는 계정입니다.'}), 403
    return jsonify({'success': True, 'people': leave_view.team()})


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
    return jsonify({'success': True, 'month': f'{y:04d}-{m:02d}', 'entries': leave_view.calendar(y, m)})


# ── 홈 ────────────────────────────────────────────────────────────────
@hr_bp.route('/home')
@gw_page
def page_home():
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
        logger.warning(f'[GW-HOME] 휴가 요약 실패: {e}')
        out['leave'] = None
    # 결재·승인 목록은 슬랙이 확인한 본인 ID 로만 (인사카드 slack_id 바꿔치기 방지, 2026-11-13)
    from dashboard.services.gw_approval import verified_slack_id
    _sid = '' if _is_field() else verified_slack_id(_me(), (p or {}).get('slack_id') or '')
    if not _is_field():
        try:
            out['approvals'] = gw_home.approvals(_sid)
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
        logger.warning(f'[GW-HOME] 영업 숫자 실패: {type(e).__name__}: {e}')
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
        return jsonify({'success': False, 'message': '인사카드가 연결되지 않았습니다.'}), 404
    v = ye.my_view(p, _ye_year())
    v['problems'] = ye.problems(v)
    return jsonify({'success': True, 'ye': v})


@hr_bp.route('/api/yearend/me', methods=['POST'])
@gw_page
def api_ye_save():
    from dashboard.services import yearend as ye
    p = _my_person()
    if not p:
        return jsonify({'success': False, 'message': '인사카드가 연결되지 않았습니다.'}), 404
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
        if sub['status'] == ye.RETURNED:
            p = hr.get_person(pid)
            ye._dm((p or {}).get('slack_id') or '',
                   f'*{_ye_year()}년 귀속 연말정산* 자료 보완을 요청드려요.\n사유: {sub["returned_reason"]}\n'
                   '그룹웨어 → 인사 → 연말정산에서 고친 뒤 다시 제출해 주세요.')
        return jsonify({'success': True, 'submission': sub})
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
        data = ye.bundle_zip(_ye_year(), _me() if not _is_field() else '')
    except Exception as e:
        return _fail(e)
    resp = send_file(io.BytesIO(data), as_attachment=True, download_name=f'연말정산_{_ye_year()}귀속_자료.zip',
                     mimetype='application/zip')
    resp.headers['Cache-Control'] = 'no-store'
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
