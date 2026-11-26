"""그룹웨어 재무 > 입출금(2026-11-26). 관리자(admin·super_admin)만 — 잔액이 보이는 화면.
규칙은 services/bank_ledger.py. 상태를 바꾸는 API 는 /api/ 로 시작(보안 미들웨어 경로 규칙).
"""
from datetime import date, timedelta

from flask import Blueprint, jsonify, redirect, render_template, request, session

from dashboard.blueprints.hr import BUG_ERRORS, _is_admin, _is_field, gw_page
from dashboard.services import bank_ledger as bl
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_finance_bp = Blueprint('gw_finance', __name__)

HIDDEN_CODES = ('농협',)   # 입출금 화면에 안 보이는 계좌(장부엔 남음)


def _admin() -> bool:
    return _is_admin() and not _is_field()


@gw_finance_bp.before_request
def _admin_only():
    if not session.get('user') or _admin():
        return None
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'message': '관리자만 볼 수 있는 화면이에요.'}), 403
    return redirect('/home')


def _by() -> str:
    u = session.get('user') or {}
    return u.get('name') or u.get('email') or ''


def _err(e: Exception):
    if isinstance(e, BUG_ERRORS):
        logger.error(f'[입출금] 처리 실패(코드 오류): {type(e).__name__}: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '처리하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    if isinstance(e, PermissionError):
        return jsonify({'success': False, 'message': str(e)}), 403
    if isinstance(e, LookupError):
        return jsonify({'success': False, 'message': str(e)}), 404
    if isinstance(e, ValueError):
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.error(f'[입출금] 처리 실패: {e}', exc_info=True)
    return jsonify({'success': False, 'message': '처리하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500


def _day(v, default: date) -> str:
    try:
        return date.fromisoformat(str(v or '')).isoformat()
    except ValueError:
        return default.isoformat()


@gw_finance_bp.route('/finance/bank')
@gw_page
def bank_page():
    return render_template('gw/finance_bank.html')


@gw_finance_bp.route('/api/finance/bank', methods=['GET'])
@gw_page
def api_list():
    try:
        today = date.today()
        start = _day(request.args.get('start'), today - timedelta(days=6))
        end = _day(request.args.get('end'), today)
        if start > end:
            start, end = end, start
        acct = int(request.args.get('account') or 0) or None
        rows = bl.list_tx(start, end, acct, request.args.get('direction') or '',
                          request.args.get('unassigned') == '1', (request.args.get('q') or '').strip()[:40])
        # 농협(N통장)은 SB 가 관리하지 않는 계좌(11-26 부장님) — 화면·합계에서 뺀다.
        # 장부엔 그대로 남김: 자금일보 초안이 전자결재 N 출금과 짝지을 때 씀(services/fund_daily.py)
        rows = [r for r in rows if r.get('acct_code') not in HIDDEN_CODES]
        # 자금일보에서 아침마다 옮겨 오는 잔액 줄은 거래가 아니라 거래 내역엔 숨김 — 구분 '잔액 맞추기'를 고르면 보임
        if request.args.get('direction') != bl.BAL:
            rows = [r for r in rows if r.get('source') != bl.REPORT]
        tot = {'in': sum(r['amount'] for r in rows if r['direction'] == bl.IN),
               'out': sum(r['amount'] for r in rows if r['direction'] == bl.OUT),
               'unassigned': sum(1 for r in rows if r['direction'] in (bl.IN, bl.OUT) and not (r['memo'] or ''))}
        return jsonify({'success': True, 'start': start, 'end': end, 'rows': rows, 'total': tot,
                        'accounts': [a for a in bl.summary(end) if a['code'] not in HIDDEN_CODES]})
    except Exception as e:
        return _err(e)


@gw_finance_bp.route('/api/finance/bank', methods=['POST'])
@gw_page
def api_add():
    try:
        tid = bl.add_manual(request.get_json(silent=True) or {}, _by(), _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, 'id': tid})


@gw_finance_bp.route('/api/finance/bank/<int:tid>', methods=['POST'])
@gw_page
def api_update(tid):
    try:
        bl.update(tid, request.get_json(silent=True) or {}, _by(), _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True})


@gw_finance_bp.route('/api/finance/bank/<int:tid>/delete', methods=['POST'])
@gw_page
def api_delete(tid):
    try:
        bl.delete(tid, _by(), _admin())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True})
