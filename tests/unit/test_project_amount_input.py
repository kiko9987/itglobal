"""슬랙 [공사 확정]·[공사 수정] 창 금액 칸 입력 검사 (2026-11-05 G4258-MW 40조 원 요청 카드)."""
import dashboard.blueprints.slack_bot as sb


def test_appended_to_prefilled_amount_is_rejected():
    # 미리 채워진 4,000,000 뒤에 3,700,000 을 이어 씀 → 예전엔 40,000,003,700,000원으로 읽음
    assert '형식' in sb._project_amount_input_error('4,000,0003,700,000')


def test_valid_amounts_pass():
    for raw in ('3,700,000', '3700000', ' 3,700,000원 ', '15,000,000', '0', '', None, '200,000,000'):
        assert sb._project_amount_input_error(raw) == '', raw


def test_bad_format_rejected():
    for raw in ('370만', '3,70,000', '3.700.000', '1,0000', 'abc'):
        assert '형식' in sb._project_amount_input_error(raw), raw


def test_over_one_billion_rejected():
    assert '10억' in sb._project_amount_input_error('40000003700000')   # 쉼표 없이 이어 쓴 경우
    assert '10억' in sb._project_amount_input_error('1,000,000,001')
    assert sb._project_amount_input_error('1,000,000,000') == ''
