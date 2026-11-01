"""전자결재 나머지 양식 (2026-11-01) — 양식 = 입력 칸 + 결재선 정의, 창·읽기·검사·카드는 공용.

플렉스 1년치(1,311건) 기준으로 남은 양식:
  업무요청형(① 경영지원 처리 완료 한 단계): 계약서 작성 요청 · 보증보험 신청 · 카드결제 요청 ·
                                          법인서류 요청 · 법인카드 지출결의서(월별)
  지출형(① 승인 → ② 이체 완료):           개인경비 청구
자금집행·비품 신청은 따로 만든 창(e_approval.new_request_view / supply_request_view)을 쓴다.

칸 값은 approvals.extra['f'] 에 {칸 key: 값}. 금액 칸은 amount, 현장은 project_code 열에도 넣는다.
설계: docs/e_approval_design.md 4-2·4-4~4-9
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple

FORM_CONTRACT, FORM_INSURANCE, FORM_CARDPAY = 'contract', 'insurance', 'cardpay'
FORM_CORPDOCS, FORM_CARDEXP, FORM_EXPENSE = 'corpdocs', 'cardexp', 'expense'

A_FORM_PROJECT = 'eap_fproject'   # 양식 창 현장 external_select (고르면 공사현황 값으로 채움)
CB_FORM = 'eap_form_submit'

CARD_COMPANIES = ['국민', '신한', '삼성', '현대', '롯데', '하나', '우리', '비씨', '농협', '씨티', '기타']
GUARANTEE_TYPES = ['입찰보증', '계약보증', '선금보증', '하자보증', '지급보증', '공사이행보증']
CORP_DOC_TYPES = ['인감증명서', '사용인감계', '법인등기사항전부증명서', '주주명부', '재무제표', '부가가치세증명원',
                  '국세완납증명서', '지방세완납증명서', '4대보험완납증명서', '4대보험가입자명부', '산업재해율확인서',
                  '중소기업확인서', '신용평가보고서', '고용보험완납증명서', '건설업등록증', '면허수첩', '기타']


@dataclass
class F:
    """입력 칸. kind: text·multiline·money·number·date·select·multiselect·project."""
    key: str
    label: str
    kind: str = 'text'
    required: bool = True
    options: List[str] = field(default_factory=list)
    placeholder: str = ''
    hint: str = ''
    auto: str = ''          # 현장을 고르면 채울 공사현황 칸(또는 특수값)
    card: bool = True       # 카드에 보일지


@dataclass
class Form:
    key: str
    label: str
    emoji: str
    kind: str               # 'task'(업무요청형) / 'spend'(지출형)
    fields: List[F]
    amount_key: str = ''    # 금액 열에 넣을 칸
    menu: bool = True
    project_card: bool = False   # 공사확정 카드 [📋 결재 요청]에 보일지
    note: str = ''          # 창 맨 위 안내


FORMS: Dict[str, Form] = {}


def _add(form: Form) -> None:
    FORMS[form.key] = form


_add(Form(FORM_CONTRACT, '계약서 작성 요청', '📑', 'task', project_card=True, amount_key='total',
          note='현장을 고르면 총액·인수자·연락처가 공사현황 값으로 채워져요. 사업자등록증은 공사 카드 스레드에 있어야 해요.',
          fields=[
              F('project', '현장', 'project'),
              F('deposit', '계약금 (VAT 포함, 원)', 'money', required=False),
              F('deposit_date', '계약금 입금 일자', 'date', required=False),
              F('mid', '중도금 (VAT 포함, 원)', 'money', required=False),
              F('mid_date', '중도금 입금 일자', 'date', required=False),
              F('balance', '잔금 (VAT 포함, 원)', 'money', required=False),
              F('balance_date', '잔금 입금 일자', 'date', required=False),
              F('total', '총 금액 (VAT 포함, 원)', 'money', auto='총액 2'),
              F('receiver', '인수자명', auto='발주처 담당자'),
              F('receiver_phone', '인수자 연락처', auto='발주처 연락처'),
              F('delivery', '납기 및 제품인도 기간', 'date', required=False),
              F('send', '계약서 발송 방법', 'select', options=['이메일', '카카오톡', '대면']),
              F('contact', '이메일 또는 전화번호', auto='발주처 이메일'),
              F('memo', '비고', 'multiline', required=False,
                hint='계약·선급·하자 보증보험이 필요하면 보증금율·기간을 적어 주세요.'),
          ]))

_add(Form(FORM_INSURANCE, '보증보험 신청', '🛡️', 'task', project_card=True, amount_key='total',
          fields=[
              F('kind', '보증 종류', 'select', options=GUARANTEE_TYPES),
              F('project', '현장', 'project'),
              F('total', '총 금액 (VAT 포함, 원)', 'money', auto='총액 2'),
              F('rate', '이행율 (%)', 'number', placeholder='예: 10'),
              F('start', '보증 시작일', 'date'),
              F('years', '보증 기간 (년)', 'number', placeholder='예: 2'),
              F('send', '전달 방법', 'select', options=['이메일', '담당자 직접 전달']),
              F('email', '이메일 주소', required=False, auto='발주처 이메일', hint='전달 방법이 이메일일 때'),
              F('memo', '비고', 'multiline', required=False),
          ]))

_add(Form(FORM_CARDPAY, '카드결제 요청', '💳', 'task', project_card=True, amount_key='total',
          note='고객 카드로 공사 대금을 결제하는 요청이에요. 수수료 3%를 더한 금액으로 결제합니다.',
          fields=[
              F('project', '현장', 'project'),
              F('card_co', '카드사', 'select', options=CARD_COMPANIES),
              F('card_no', '카드번호', placeholder='예: 1234-5678-9012-3456'),
              F('expiry', '유효기간 (MM/YY)', placeholder='예: 08/29'),
              F('sales', '원 매출 (VAT 포함, 원)', 'money'),
              F('installment', '할부 개월', 'number', placeholder='일시불이면 0'),
              F('memo', '비고', 'multiline', required=False),
          ]))

_add(Form(FORM_CORPDOCS, '법인서류 요청', '🏢', 'task',
          fields=[
              F('grn', '사업자 구분', 'select', options=['G (글로벌)', 'R (글로벌그룹)', 'N (사업자 없음)']),
              F('project', '현장', 'project', required=False),
              F('docs', '서류 종류', 'multiselect', options=CORP_DOC_TYPES),
              F('year', '발행 연도', 'select', required=False, options=[]),   # 열 때 올해부터 채움
              F('other', '기타 선택 시 서류 이름', required=False),
              F('reason', '사유', 'multiline'),
              F('cert_memo', '증명서 기재 내용', 'multiline', required=False),
              F('due', '제출 마감일', 'date'),
              F('receive', '수령 방법', 'select', options=['이메일', '등기']),
              F('dest', '이메일 또는 등기 주소'),
          ]))

_add(Form(FORM_CARDEXP, '법인카드 지출결의서', '🧾', 'task',
          note='지난달 법인카드 사용분이에요. 영수증은 카드가 올라간 뒤 스레드에 올려 주세요.',
          fields=[
              F('card_no', '카드번호', auto='@card', placeholder='예: 4518-****-****-1234'),
              F('start', '사용 시작일', 'date', auto='@prev_month_start'),
              F('end', '사용 종료일', 'date', auto='@prev_month_end'),
              F('memo', '비고', 'multiline', required=False),
          ]))

_add(Form(FORM_EXPENSE, '개인경비 청구', '🙋', 'spend', amount_key='amount',
          note='본인 돈으로 먼저 쓴 비용을 돌려받는 요청이에요. 계좌는 카드에 보이지 않고 경영지원에게만 보여요.',
          fields=[
              F('use_date', '사용일자', 'date'),
              F('project', '현장', 'project'),
              F('cost', '원가 분류', 'select', options=[]),   # COST_CATEGORIES 로 채움
              F('amount', '청구 금액 (원)', 'money'),
              F('reason', '사유', 'multiline'),
              F('bank', '은행', auto='@my_bank', card=False),
              F('account', '계좌번호', auto='@my_account', card=False, hint='계좌번호는 숫자와 하이픈으로'),
          ]))


TASK_FORMS = tuple(k for k, f in FORMS.items() if f.kind == 'task')
SPEND_FORMS = tuple(k for k, f in FORMS.items() if f.kind == 'spend')


def get(form_key: str) -> Optional[Form]:
    return FORMS.get(form_key)


def menu_options() -> List[dict]:
    return [{'text': {'type': 'plain_text', 'text': f'{f.emoji} {f.label}', 'emoji': True}, 'value': f.key}
            for f in FORMS.values() if f.menu]


def project_card_options() -> List[dict]:
    return [{'text': {'type': 'plain_text', 'text': f'{f.emoji} {f.label}', 'emoji': True}, 'value': f.key}
            for f in FORMS.values() if f.project_card]


# ── 자동 채움 ─────────────────────────────────────────────────────────
def prev_month(today: Optional[date] = None) -> Tuple[date, date]:
    t = today or date.today()
    end = t.replace(day=1) - timedelta(days=1)
    return end.replace(day=1), end


def _num_str(v) -> str:
    s = str(v or '').strip()
    if s in ('', '-', 'None', 'nan'):
        return ''
    try:
        n = int(float(s.replace(',', '')))
        return str(n) if n else ''
    except ValueError:
        return s


def prefill(form: Form, rec: Optional[dict], memory: Optional[dict] = None,
            today: Optional[date] = None) -> Dict[str, str]:
    """auto 칸 값 — rec=공사현황 레코드, memory={'card','bank','account'}(본인 기억)."""
    mem = memory or {}
    start, end = prev_month(today)
    out = {}
    for f in form.fields:
        if not f.auto:
            continue
        if f.auto == '@prev_month_start':
            out[f.key] = start.isoformat()
        elif f.auto == '@prev_month_end':
            out[f.key] = end.isoformat()
        elif f.auto == '@card':
            out[f.key] = mem.get('card', '')
        elif f.auto == '@my_bank':
            out[f.key] = mem.get('bank', '')
        elif f.auto == '@my_account':
            out[f.key] = mem.get('account', '')
        elif rec:
            v = rec.get(f.auto)
            out[f.key] = _num_str(v) if f.kind == 'money' else ('' if str(v or '').strip() in ('-', 'None') else str(v).strip())
    return {k: v for k, v in out.items() if v}


# ── 창 ────────────────────────────────────────────────────────────────
def _plain(t: str) -> dict:
    return {'type': 'plain_text', 'text': t, 'emoji': True}


def _opt(v: str) -> dict:
    return {'text': _plain(v[:75]), 'value': v[:150]}


def _field_options(form: Form, f: F, today: Optional[date] = None) -> List[str]:
    if f.key == 'year' and form.key == FORM_CORPDOCS:
        y = (today or date.today()).year
        return [f'{x}년' for x in range(y, 2018, -1)]
    if f.key == 'cost' and form.key == FORM_EXPENSE:
        from dashboard.services.purchase_invoice_assign import COST_CATEGORIES
        return [label for _v, label in COST_CATEGORIES]
    return f.options


def _element(form: Form, f: F, value: str, today: Optional[date] = None) -> dict:
    aid = f.key
    if f.kind == 'project':
        el = {'type': 'external_select', 'action_id': A_FORM_PROJECT, 'min_query_length': 0,
              'placeholder': _plain('코드·상호·주소로 검색 (비우면 내 담당 현장)')}
        if value:
            el['initial_option'] = {'text': _plain(value[:75]), 'value': value}
        return el
    if f.kind == 'date':
        el = {'type': 'datepicker', 'action_id': aid}
        if value:
            el['initial_date'] = value
        return el
    if f.kind in ('select', 'multiselect'):
        opts = [_opt(o) for o in _field_options(form, f, today)]
        el = {'type': 'static_select' if f.kind == 'select' else 'multi_static_select',
              'action_id': aid, 'options': opts, 'placeholder': _plain(f.label)}
        vals = value if isinstance(value, list) else ([value] if value else [])
        picked = [o for o in opts if o['value'] in vals]
        if picked:
            if f.kind == 'select':
                el['initial_option'] = picked[0]
            else:
                el['initial_options'] = picked
        return el
    el = {'type': 'plain_text_input', 'action_id': aid, 'multiline': f.kind == 'multiline'}
    if f.placeholder:
        el['placeholder'] = _plain(f.placeholder)
    if value:
        el['initial_value'] = str(value)
    return el


def form_view(form_key: str, meta: dict, values: Optional[dict] = None, project_label: str = '',
              today: Optional[date] = None) -> dict:
    """양식 창. meta: {'form', 'project'(고정이면), 'pv'(채운 횟수 — 바뀌면 칸 block_id 를 바꿔 새 값이 보이게)}."""
    form = FORMS[form_key]
    values = values or {}
    pv = int(meta.get('pv') or 0)
    sfx = f'_{pv}' if pv else ''
    blocks: List[dict] = []
    if form.note:
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': form.note}]})
    for f in form.fields:
        if f.kind == 'project' and meta.get('project'):
            blocks.append({'type': 'section', 'text': {'type': 'mrkdwn',
                                                       'text': f'*현장* {project_label or meta["project"]}'}})
            continue
        v = values.get(f.key, '')
        b = {'type': 'input', 'block_id': f.key + (sfx if f.auto and f.kind != 'project' else ''),
             'optional': not f.required, 'label': _plain(f.label), 'element': _element(form, f, v, today)}
        if f.kind == 'project':
            b['dispatch_action'] = True
            if v and values.get('_project_label'):
                b['element']['initial_option'] = {'text': _plain(values['_project_label'][:75]), 'value': v}
        if f.hint:
            b['hint'] = _plain(f.hint)
        blocks.append(b)
    return {
        'type': 'modal', 'callback_id': CB_FORM,
        'title': _plain(form.label[:24]),
        'submit': _plain('요청'), 'close': _plain('취소'),
        'private_metadata': json.dumps({**meta, 'form': form_key}),
        'blocks': blocks,
    }


def _state_value(values: dict, f: F):
    from dashboard.services.e_approval import field_block_id
    bid = field_block_id(values, f.key)
    st = values.get(bid) or {}
    if f.kind == 'project':
        el = st.get(A_FORM_PROJECT) or {}
        return ((el.get('selected_option') or {}).get('value') or '').strip()
    el = st.get(f.key) or {}
    if f.kind == 'date':
        return el.get('selected_date') or ''
    if f.kind == 'select':
        return (el.get('selected_option') or {}).get('value') or ''
    if f.kind == 'multiselect':
        return [o.get('value') for o in (el.get('selected_options') or [])]
    return (el.get('value') or '').strip()


def read_form(view: dict) -> Tuple[str, Dict]:
    """(form_key, {칸: 값}) — 고정 현장은 meta 에서."""
    meta = json.loads(view.get('private_metadata') or '{}')
    form = FORMS[meta['form']]
    values = view['state']['values']
    out = {f.key: _state_value(values, f) for f in form.fields}
    if meta.get('project'):
        out['project'] = meta['project']
    return form.key, out


def _parse_int(s) -> int:
    d = re.sub(r'[^\d]', '', str(s or ''))
    return int(d) if d else 0


def validate(form_key: str, vals: Dict) -> Dict[str, str]:
    """{칸 key: 문구} — 슬랙 block_id 로 바꾸는 건 호출부(field_block_id)."""
    form = FORMS[form_key]
    errs = {}
    for f in form.fields:
        v = vals.get(f.key)
        if f.required and (not v or (isinstance(v, list) and not v)):
            errs[f.key] = f'{f.label}을(를) 넣어 주세요.' if f.kind not in ('select', 'multiselect', 'project', 'date') \
                else f'{f.label}을(를) 골라 주세요.'
            continue
        if v and f.kind == 'money' and _parse_int(v) <= 0:
            errs[f.key] = '금액을 숫자로 넣어 주세요.'
        if v and f.kind == 'number' and not re.fullmatch(r'\d+(\.\d+)?', str(v).replace(',', '').strip()):
            errs[f.key] = '숫자로 넣어 주세요.'
    if form_key == FORM_CARDPAY:
        if vals.get('card_no') and len(re.sub(r'\D', '', vals['card_no'])) < 14:
            errs['card_no'] = '카드번호를 숫자로 넣어 주세요.'
        if vals.get('expiry') and not re.fullmatch(r'\s*(0[1-9]|1[0-2])\s*/\s*\d{2}\s*', vals['expiry']):
            errs['expiry'] = 'MM/YY 로 넣어 주세요. 예: 08/29'
    if form_key == FORM_EXPENSE and vals.get('account') and len(re.sub(r'\D', '', vals['account'])) < 8:
        errs['account'] = '계좌번호를 숫자로 넣어 주세요.'
    if form_key == FORM_INSURANCE and vals.get('send') == '이메일' and not vals.get('email'):
        errs['email'] = '이메일로 받으려면 주소를 넣어 주세요.'
    if form_key == FORM_CORPDOCS and '기타' in (vals.get('docs') or []) and not vals.get('other'):
        errs['other'] = '기타를 골랐으면 서류 이름을 적어 주세요.'
    return errs


def amount_of(form_key: str, vals: Dict) -> int:
    """금액 열 값 — 카드결제는 원 매출 + 수수료 3%."""
    if form_key == FORM_CARDPAY:
        sales = _parse_int(vals.get('sales'))
        return sales + round(sales * 0.03)
    form = FORMS[form_key]
    return _parse_int(vals.get(form.amount_key)) if form.amount_key else 0


def grn_of(form_key: str, vals: Dict) -> str:
    g = str(vals.get('grn') or '')[:1].upper()
    if g in ('G', 'R', 'N'):
        return g
    c = str(vals.get('project') or '')[:1].upper()
    return c if c in ('G', 'R') else ''


def doc_title(form_key: str, vals: Dict, requester_name: str, today: Optional[date] = None) -> str:
    """카드 머리에 붙일 이름 — 계약서는 G/R, 법인카드는 'YYYY년 MM월 …_이름 매니저'."""
    form = FORMS[form_key]
    if form_key == FORM_CARDEXP:
        s = vals.get('start') or prev_month(today)[0].isoformat()
        return f'{s[:4]}년 {s[5:7]}월 법인카드지출결의서_{requester_name} 매니저'
    g = grn_of(form_key, vals)
    if form_key in (FORM_CONTRACT, FORM_INSURANCE, FORM_CARDPAY) and g in ('G', 'R'):
        return f'{g}{form.label}'
    return form.label


# ── 카드 ──────────────────────────────────────────────────────────────
def _fmt(f: F, v) -> str:
    if isinstance(v, list):
        return ', '.join(v)
    if f.kind == 'money':
        n = _parse_int(v)
        return f'{n:,}원' if n else ''
    if f.kind == 'date' and v:
        return str(v).replace('-', '.')
    if f.kind == 'select' and f.key == 'grn':
        return str(v)
    return str(v or '')


def card_lines(form_key: str, vals: Dict) -> List[str]:
    """카드 본문 줄(현장 줄은 엔진이 따로)."""
    form = FORMS[form_key]
    lines = []
    for f in form.fields:
        if f.kind == 'project' or not f.card:
            continue
        s = _fmt(f, vals.get(f.key))
        if s:
            lines.append(f'{f.label.split(" (")[0]}: {s}')
    if form_key == FORM_CARDPAY:
        sales = _parse_int(vals.get('sales'))
        fee = round(sales * 0.03)
        lines.append(f'수수료(3%): {fee:,}원 · *결제 금액: {sales + fee:,}원*')
    if form_key == FORM_EXPENSE:
        lines.append('이체 계좌: 본인 계좌 (경영지원 DM에서 확인)')
    return lines


def summary(form_key: str, vals: Dict) -> str:
    """DM 한 줄 요약에 쓸 '무엇'."""
    if form_key == FORM_CORPDOCS:
        return ', '.join(vals.get('docs') or [])[:60] or '법인서류'
    if form_key == FORM_INSURANCE:
        return vals.get('kind') or '보증보험'
    if form_key == FORM_CARDPAY:
        return f"{vals.get('card_co') or ''} 카드결제".strip()
    if form_key == FORM_CARDEXP:
        return f"{vals.get('start', '')[:7]} 사용분"
    if form_key == FORM_EXPENSE:
        return (vals.get('reason') or '개인경비')[:40]
    return FORMS[form_key].label
