"""
프로젝트 코드 자동 생성 유틸리티
Google Sheets 수식 규칙과 동일하게 구현
"""

def generate_project_code(row_number, company, manager):
    """
    프로젝트 코드 생성

    규칙 (Google Sheets 수식):
    =IF($B2890="플렌트","P"&TEXT(ROW()-1,"0000"),
       IF($B2890="글로벌","G"&TEXT(ROW()-1,"0000"),
          IF($B2890="글로벌그룹","R"&TEXT(ROW()-1,"0000"),)))
    &SWITCH($C2890,
            "박용구","-YG","박정우","-JW","강성환","-SH","박민우","-MW",
            "이근혁","-GH","김호중","-HJ","아이티","-IT","김단이","-DN",
            "권태훈","-TH","주영민","-YM","심장원","-SJW","빈승정","-SJ",
            "박민재","-MJ","조성헌","-JSH","황해승","-HS","강민석","-MS","")

    Args:
        row_number (int): 행번호 (2부터 시작)
        company (str): 사업자 (플렌트/글로벌/글로벌그룹)
        manager (str): 담당자 이름

    Returns:
        str: 프로젝트 코드 (예: G2890-SH)
    """

    # 1. 사업자에 따른 Prefix
    company_prefix = {
        '플렌트': 'P',
        '글로벌': 'G',
        '글로벌그룹': 'R'
    }

    prefix = company_prefix.get(company, '')
    if not prefix:
        return ''  # 사업자가 유효하지 않으면 빈 코드

    # 2. 행번호 4자리 (ROW()-1 규칙 반영)
    # Google Sheets 수식: TEXT(ROW()-1,"0000")
    # row_number는 실제 시트 행번호이므로 -1 적용
    number_part = f"{max(row_number - 1, 0):04d}"

    # 3. 담당자에 따른 Suffix
    manager_suffix = {
        '박용구': '-YG',
        '박정우': '-JW',
        '강성환': '-SH',
        '박민우': '-MW',
        '이근혁': '-GH',
        '김호중': '-HJ',
        '아이티': '-IT',
        '김단이': '-DN',
        '권태훈': '-TH',
        '주영민': '-YM',
        '심장원': '-SJW',
        '빈승정': '-SJ',
        '박민재': '-MJ',
        '조성헌': '-JSH',
        '황해승': '-HS',
        '강민석': '-MS',
        '강정권': '-JK',
        '이상덕': '-SD',
    }

    suffix = manager_suffix.get(manager, '')

    # 4. 최종 프로젝트 코드
    project_code = f"{prefix}{number_part}{suffix}"

    return project_code


def parse_project_code(project_code):
    """
    프로젝트 코드를 분석하여 구성 요소 추출

    Args:
        project_code (str): 프로젝트 코드 (예: G2890-SH)

    Returns:
        dict: {
            'prefix': 'G',
            'number': 2890,
            'suffix': '-SH',
            'company': '글로벌',
            'manager': '강성환'
        }
    """
    if not project_code:
        return None

    try:
        # Prefix (첫 글자)
        prefix = project_code[0]

        # 사업자 역매핑
        prefix_to_company = {
            'P': '플렌트',
            'G': '글로벌',
            'R': '글로벌그룹'
        }
        company = prefix_to_company.get(prefix, '')

        # Suffix (- 이후)
        parts = project_code.split('-')
        suffix = f"-{parts[1]}" if len(parts) > 1 else ''

        # 담당자 역매핑
        suffix_to_manager = {
            '-YG': '박용구',
            '-JW': '박정우',
            '-SH': '강성환',
            '-MW': '박민우',
            '-GH': '이근혁',
            '-HJ': '김호중',
            '-IT': '아이티',
            '-DN': '김단이',
            '-TH': '권태훈',
            '-YM': '주영민',
            '-SJW': '심장원',
            '-SJ': '빈승정',
            '-MJ': '박민재',
            '-JSH': '조성헌',
            '-HS': '황해승',
            '-MS': '강민석',
            '-JK': '강정권',
            '-SD': '이상덕',
        }
        manager = suffix_to_manager.get(suffix, '')

        # 숫자 부분 (Prefix 다음부터 - 전까지)
        number_str = parts[0][1:]  # P2890 → 2890
        number = int(number_str)

        return {
            'prefix': prefix,
            'number': number,
            'suffix': suffix,
            'company': company,
            'manager': manager
        }
    except Exception as e:
        return None


def should_update_project_code(current_code, row_number, company, manager):
    """
    프로젝트 코드를 업데이트해야 하는지 확인

    Args:
        current_code (str): 현재 프로젝트 코드
        row_number (int): 행번호
        company (str): 새로운 사업자
        manager (str): 새로운 담당자

    Returns:
        bool: 업데이트 필요 여부
    """
    parsed = parse_project_code(current_code)
    if not parsed:
        return False  # 파싱 실패 시 기존 코드 유지 (알 수 없는 형식)

    # 담당자 파싱 실패 시 (매핑에 없는 담당자) 기존 코드 유지
    # suffix만 비어있고 나머지는 정상인 경우 사업자 변경만 체크
    if not parsed['manager']:
        # 사업자만 변경되었는지 확인
        return parsed['company'] != company

    # 정상적으로 파싱된 경우: 사업자 또는 담당자가 변경되었는지 확인
    return parsed['company'] != company or parsed['manager'] != manager
