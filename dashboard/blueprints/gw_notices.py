"""그룹웨어 공지사항·개발 노트 화면(2026-11-25). 규칙은 services/notices.py.
상태를 바꾸는 API 는 /api/ 로 시작해야 보안 미들웨어를 통과한다. 현장 세션은 groupware 문지기가 공지사항 보기만 열어 둔다.
"""
import os

from flask import Blueprint, abort, jsonify, render_template, request, send_file, session

from dashboard.blueprints.hr import BUG_ERRORS, _is_admin, _is_field, _me, _my_person, gw_page
from dashboard.services import notices as ns
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_notices_bp = Blueprint('gw_notices', __name__)


def _viewer() -> dict:
    return ns.viewer(_my_person(), '' if _is_field() else _me(), admin=_is_admin() and not _is_field(), field=_is_field())


def _admin() -> bool:
    return _is_admin() and not _is_field()


def _author() -> dict:
    v = _viewer()
    user = session.get('user') or {}
    return {'name': v['name'] or user.get('name') or '', 'email': _me(), 'key': v['key']}


def _err(e: Exception):
    if isinstance(e, BUG_ERRORS):
        logger.error(f'[공지] 처리 실패(코드 오류): {type(e).__name__}: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '처리하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    if isinstance(e, PermissionError):
        return jsonify({'success': False, 'message': str(e)}), 403
    if isinstance(e, LookupError):
        return jsonify({'success': False, 'message': str(e)}), 404
    if isinstance(e, ValueError):
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.error(f'[공지] 처리 실패: {e}', exc_info=True)
    return jsonify({'success': False, 'message': '처리하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500


def _kind(v) -> str:
    k = (v or ns.NOTICE).strip()
    if k not in ns.KINDS:
        raise ValueError('없는 게시판이에요.')
    if k == ns.DEVNOTE and _is_field():
        raise PermissionError('볼 수 없는 게시판이에요.')
    return k


# ── 화면 ──────────────────────────────────────────────────────────────
def _page(kind: str, nid: int = 0, mode: str = 'list'):
    if mode in ('new', 'edit') and not _admin():
        abort(403)
    return render_template('gw/notices.html', kind=kind, kind_label=ns.KIND_LABEL[kind], nid=nid, mode=mode,
                           can_edit=_admin())


@gw_notices_bp.route('/notices')
@gw_page
def notices_page():
    return _page(ns.NOTICE)


@gw_notices_bp.route('/notices/<int:nid>')
@gw_page
def notice_page(nid):
    return _page(ns.NOTICE, nid, 'view')


@gw_notices_bp.route('/notices/new')
@gw_page
def notice_new():
    return _page(ns.NOTICE, 0, 'new')


@gw_notices_bp.route('/notices/<int:nid>/edit')
@gw_page
def notice_edit(nid):
    return _page(ns.NOTICE, nid, 'edit')


@gw_notices_bp.route('/devnotes')
@gw_page
def devnotes_page():
    if _is_field():
        abort(403)
    return _page(ns.DEVNOTE)


@gw_notices_bp.route('/devnotes/<int:nid>')
@gw_page
def devnote_page(nid):
    if _is_field():
        abort(403)
    return _page(ns.DEVNOTE, nid, 'view')


@gw_notices_bp.route('/devnotes/new')
@gw_page
def devnote_new():
    return _page(ns.DEVNOTE, 0, 'new')


@gw_notices_bp.route('/devnotes/<int:nid>/edit')
@gw_page
def devnote_edit(nid):
    return _page(ns.DEVNOTE, nid, 'edit')


# ── 읽기 API ─────────────────────────────────────────────────────────
@gw_notices_bp.route('/api/notices', methods=['GET'])
@gw_page
def api_list():
    try:
        kind = _kind(request.args.get('kind'))
        v = _viewer()
        cid = int(request.args.get('category') or 0) or None
        limit = min(int(request.args.get('limit') or 0), 200)
        items = ns.list_notices(kind, v, cid, limit, (request.args.get('q') or '')[:50])
        cats = ns.visible_categories(v) if kind == ns.NOTICE else []
        return jsonify({'success': True, 'kind': kind, 'items': items, 'categories': cats,
                        'unread': ns.unread_count(kind, v), 'can_edit': _admin()})
    except Exception as e:
        return _err(e)


@gw_notices_bp.route('/api/notices/<int:nid>', methods=['GET'])
@gw_page
def api_get(nid):
    try:
        v = _viewer()
        n = ns.get(nid, v, mark_read=request.args.get('peek') != '1')
        if n['kind'] == ns.DEVNOTE and _is_field():
            raise PermissionError('볼 수 없는 글이에요.')
        return jsonify({'success': True, 'notice': n, 'can_edit': _admin()})
    except Exception as e:
        return _err(e)


@gw_notices_bp.route('/api/notices/files/<int:fid>', methods=['GET'])
@gw_page
def api_file(fid):
    try:
        f = ns.file_for(fid, _viewer())
    except Exception as e:
        return _err(e)
    if not os.path.exists(f['path']):
        return jsonify({'success': False, 'message': '파일이 서버에 없어요. 경영지원에 알려 주세요.'}), 404
    inline = request.args.get('view') == '1'
    resp = send_file(f['path'], as_attachment=not inline, download_name=f['name'], max_age=0)
    resp.headers['Cache-Control'] = 'no-store'
    return resp


@gw_notices_bp.route('/api/notices/unread', methods=['GET'])
@gw_page
def api_unread():
    """홈·메뉴 — 안 읽은 공지 수와 최근 공지(현장 세션은 공지사항만)."""
    try:
        v = _viewer()
        recent = ns.list_notices(ns.NOTICE, v, limit=4)
        out = {'success': True, 'notice': ns.unread_count(ns.NOTICE, v), 'recent': recent}
        if not _is_field():
            out['devnote'] = ns.unread_count(ns.DEVNOTE, v)
        return jsonify(out)
    except Exception as e:
        return _err(e)


# ── 쓰기 API(관리자) ─────────────────────────────────────────────────
@gw_notices_bp.route('/api/notices', methods=['POST'])
@gw_page
def api_create():
    body = request.get_json(silent=True) or {}
    try:
        kind = _kind(body.get('kind'))
        res = ns.save(kind, body, _author(), _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, **res})


@gw_notices_bp.route('/api/notices/<int:nid>', methods=['POST'])
@gw_page
def api_update(nid):
    body = request.get_json(silent=True) or {}
    try:
        kind = _kind(body.get('kind'))
        res = ns.save(kind, dict(body, _edit=True), _author(), _admin(), nid)
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, **res})


@gw_notices_bp.route('/api/notices/<int:nid>/delete', methods=['POST'])
@gw_page
def api_delete(nid):
    try:
        ns.delete(nid, _admin(), _author()['name'])
    except Exception as e:
        return _err(e)
    return jsonify({'success': True})


@gw_notices_bp.route('/api/notices/<int:nid>/files', methods=['POST'])
@gw_page
def api_file_add(nid):
    fs = request.files.getlist('file')
    if not fs:
        return jsonify({'success': False, 'message': '파일을 골라 주세요.'}), 400
    try:
        out = [ns.add_file(nid, f.filename or '', f.read(), _admin()) for f in fs]
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, 'files': out})


@gw_notices_bp.route('/api/notices/files/<int:fid>/delete', methods=['POST'])
@gw_page
def api_file_delete(fid):
    try:
        ns.remove_file(fid, _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True})


@gw_notices_bp.route('/api/notices/<int:nid>/notify', methods=['POST'])
@gw_page
def api_notify(nid):
    try:
        res = ns.notify(nid, _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, **res})


# ── 카테고리(관리자) ──────────────────────────────────────────────────
@gw_notices_bp.route('/api/notices/categories', methods=['GET'])
@gw_page
def api_categories():
    if not _admin():
        return jsonify({'success': False, 'message': '관리자만 볼 수 있어요.'}), 403
    from dashboard.services import hr
    depts = sorted({(p.get('department') or '').strip() for p in hr.list_people() if (p.get('department') or '').strip()})
    return jsonify({'success': True, 'categories': ns.categories(), 'audiences': ns.AUDIENCES, 'depts': depts})


@gw_notices_bp.route('/api/notices/categories', methods=['POST'])
@gw_page
def api_category_add():
    try:
        c = ns.save_category(request.get_json(silent=True) or {}, _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, 'category': c})


@gw_notices_bp.route('/api/notices/categories/<int:cid>', methods=['POST'])
@gw_page
def api_category_edit(cid):
    try:
        c = ns.save_category(request.get_json(silent=True) or {}, _admin(), cid)
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, 'category': c})


@gw_notices_bp.route('/api/notices/categories/<int:cid>/delete', methods=['POST'])
@gw_page
def api_category_delete(cid):
    try:
        ns.delete_category(cid, _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True})
