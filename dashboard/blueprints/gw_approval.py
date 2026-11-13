"""그룹웨어 결재함·플렉스 보관함 화면과 API (2026-11-10). 규칙은 services/gw_approval.py.

현장 세션은 groupware 문지기가 이 경로를 막는다(_FIELD_PAGES·_FIELD_API 에 없음).
"""
from flask import Blueprint, Response, jsonify, render_template, request, session

from dashboard.blueprints.hr import gw_page
from dashboard.services import gw_approval as ga
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_approval_bp = Blueprint('gw_approval', __name__)


def _v() -> dict:
    return ga.viewer(((session.get('user') or {}).get('email') or '').lower())


def _filters() -> dict:
    a = request.args
    return {'box': a.get('box', 'todo'), 'q': a.get('q', '').strip()[:60], 'form': a.get('form', ''),
            'status': a.get('status', ''), 'date_from': a.get('from', '')[:10], 'date_to': a.get('to', '')[:10]}


# ── 화면 ──────────────────────────────────────────────────────────────
@gw_approval_bp.route('/approval')
@gw_page
def page_box():
    return render_template('gw/approval.html')


@gw_approval_bp.route('/approval/archive')
@gw_page
def page_archive():
    return render_template('gw/approval_archive.html')


# ── 결재함 API ────────────────────────────────────────────────────────
@gw_approval_bp.route('/api/approval/list', methods=['GET'])
@gw_page
def api_list():
    return jsonify({'success': True, **ga.list_docs(_v(), **_filters())})


@gw_approval_bp.route('/api/approval/<int:aid>', methods=['GET'])
@gw_page
def api_detail(aid):
    d = ga.detail(aid, _v())
    if not d:
        return jsonify({'success': False, 'message': '볼 수 없는 결재입니다.'}), 404
    return jsonify({'success': True, 'doc': d})


@gw_approval_bp.route('/api/approval/<int:aid>/act', methods=['POST'])
@gw_page
def api_act(aid):
    body = request.get_json(silent=True) or {}
    v = _v()
    raw_rev = body.get('rev')
    rev = raw_rev if isinstance(raw_rev, int) and not isinstance(raw_rev, bool) and raw_rev >= 0 else None
    try:
        msg = ga.act(aid, str(body.get('action') or ''), v, str(body.get('reason') or ''), rev)
    except Exception as e:
        logger.error(f'[GW결재] 처리 실패 #{aid} {body.get("action")}: {e}', exc_info=True, extra={'force_alert': True})
        return jsonify({'success': False, 'message': '처리 중 오류가 났어요. 잠시 뒤 다시 눌러 주세요.'}), 500
    if msg:
        from dashboard.blueprints.slack_eapproval import STALE_MSG
        return jsonify({'success': False, 'message': msg, 'stale': msg == STALE_MSG}), 409
    logger.info(f'[GW결재] #{aid} {body.get("action")} by {v["email"]}')
    return jsonify({'success': True, 'doc': ga.detail(aid, v)})


@gw_approval_bp.route('/api/approval/export.xlsx', methods=['GET'])
@gw_page
def api_export():
    f = _filters()
    data = ga.export_xlsx(_v(), **f)
    name = f'결재함_{f["box"]}.xlsx'
    from urllib.parse import quote
    return Response(data, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f"attachment; filename*=UTF-8''{quote(name)}",
                             'Cache-Control': 'no-store'})


# ── 플렉스 보관함 API ─────────────────────────────────────────────────
@gw_approval_bp.route('/api/approval/archive', methods=['GET'])
@gw_page
def api_archive():
    a = request.args
    try:
        page = int(a.get('page') or 1)
    except ValueError:
        page = 1
    return jsonify({'success': True, **ga.archive_list(_v(), q=a.get('q', '')[:60], form=a.get('form', ''),
                                                        date_from=a.get('from', '')[:10], date_to=a.get('to', '')[:10],
                                                        page=page)})


@gw_approval_bp.route('/api/approval/archive/<doc_no>', methods=['GET'])
@gw_page
def api_archive_detail(doc_no):
    d = ga.archive_detail(doc_no, _v())
    if not d:
        return jsonify({'success': False, 'message': '볼 수 없는 문서입니다.'}), 404
    return jsonify({'success': True, 'doc': d})
