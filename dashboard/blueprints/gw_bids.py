"""그룹웨어 「입찰 공고」 — 나라장터 공고 목록과 검토 표시 (2026-11-18). 규칙은 services/g2b_bids.py.

현장 세션은 groupware 문지기가 이 경로를 막는다(_FIELD_PAGES·_FIELD_API 에 없음).
상태 바꾸는 API 는 /api/ 로 시작해야 보안 미들웨어를 통과한다.
"""
from flask import Blueprint, jsonify, render_template, request, session

from dashboard.blueprints.hr import gw_page
from dashboard.services import g2b_bids
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_bids_bp = Blueprint('gw_bids', __name__)


def _me() -> str:
    u = session.get('user') or {}
    return (u.get('name') or u.get('email') or '').strip()


@gw_bids_bp.route('/bids')
@gw_page
def page_bids():
    return render_template('gw/bids.html')


@gw_bids_bp.route('/api/bids', methods=['GET'])
@gw_page
def api_list():
    a = request.args
    try:
        data = g2b_bids.list_bids(tab=a.get('tab', g2b_bids.ELIGIBLE), include_closed=a.get('closed') == '1',
                                  q=a.get('q', '').strip()[:40])
    except Exception as e:
        logger.error(f'[입찰공고] 목록 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '공고를 불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500
    return jsonify({'success': True, **data})


def _is_admin() -> bool:
    return (session.get('user') or {}).get('permission_level') in ('admin', 'super_admin')


# ── 우리 회사 정보(11-20) — 보기는 누구나, 고치기는 관리자 ──
@gw_bids_bp.route('/api/bids/profile', methods=['GET'])
@gw_page
def api_profile():
    p = {k: v for k, v in g2b_bids.get_profile().items() if not k.startswith('_')}
    return jsonify({'success': True, 'profile': p, 'grades': list(g2b_bids.CREDIT_GRADES), 'can_edit': _is_admin()})


@gw_bids_bp.route('/api/bids/profile', methods=['POST'])
@gw_page
def api_profile_save():
    if not _is_admin():
        return jsonify({'success': False, 'message': '우리 회사 정보는 관리자만 고칠 수 있어요.'}), 403
    try:
        p = g2b_bids.save_profile(request.get_json(silent=True) or {}, _me())
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[입찰공고] 회사 정보 저장 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '저장하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    return jsonify({'success': True, 'profile': {k: v for k, v in p.items() if not k.startswith('_')}})


@gw_bids_bp.route('/api/bids/<bid_no>/mark', methods=['POST'])
@gw_page
def api_mark(bid_no):
    body = request.get_json(silent=True) or {}
    try:
        item = g2b_bids.set_mark(bid_no[:40], str(body.get('mark') or ''), _me())
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except LookupError as e:
        return jsonify({'success': False, 'message': str(e)}), 404
    except Exception as e:
        logger.error(f'[입찰공고] 표시 실패 {bid_no}: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '저장하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    logger.info(f'[입찰공고] {bid_no} 표시={item["mark"] or "(없음)"} by {_me()}')
    return jsonify({'success': True, 'item': item})
