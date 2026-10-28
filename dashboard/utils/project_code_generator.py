"""
프로젝트 코드 유틸리티 — 사업자·담당자를 바꿀 때 코드 다시 만들기.

2026-10-28 3차 정리:
  - 접미사·앞 글자 표는 staff_directory(설정 owner_suffix_map·company_prefix_map + 관리자 화면 값) 한 곳에서.
    예전엔 이 파일에 따로 박힌 표가 있어 새 사람(이미영 등)이 빠지고 퇴사자가 남았다.
  - 번호는 **지금 코드의 번호를 그대로** 쓴다. 예전엔 '시트 행번호 − 1'로 다시 매겨서, 새 공사 등록(DB 일련번호)과
    행 위치가 어긋난 공사는 담당자만 바꿔도 번호가 바뀔 수 있었다.
"""


def _maps():
    from dashboard.services.staff_directory import suffix_map, company_prefix_map
    return company_prefix_map(), suffix_map()


def generate_project_code(row_number, company, manager):
    """(옛 호환) 행번호 기준 코드 — 새 코드는 regenerate_project_code 를 쓴다.

    Returns: 'G2890-SH' 형식, 사업자를 모르면 ''.
    """
    prefixes, suffixes = _maps()
    prefix = prefixes.get(company, '')
    if not prefix:
        return ''
    sfx = suffixes.get(manager, '')
    return f"{prefix}{max(int(row_number) - 1, 0):04d}{'-' + sfx if sfx else ''}"


def parse_project_code(project_code):
    """'G2890-SH' → {'prefix','number','suffix','company','manager'}. 형식이 다르면 None."""
    if not project_code:
        return None
    try:
        prefixes, suffixes = _maps()
        prefix = project_code[0]
        company = {v: k for k, v in prefixes.items()}.get(prefix, '')
        parts = project_code.split('-')
        suffix = f"-{parts[1]}" if len(parts) > 1 else ''
        manager = {v.upper(): k for k, v in reversed(list(suffixes.items()))}.get(suffix[1:].upper(), '') if suffix else ''
        number = int(parts[0][1:])
        return {'prefix': prefix, 'number': number, 'suffix': suffix, 'company': company, 'manager': manager}
    except Exception:
        return None


def regenerate_project_code(current_code, company, manager):
    """사업자·담당자에 맞춰 코드의 앞 글자·접미사만 바꾸고 번호는 그대로.

    Returns: (새 코드, 오류 문구). 바꿀 필요 없으면 (current_code, '').
      - 지금 코드를 못 읽으면 그대로 둔다.
      - 담당자의 접미사가 표에 없으면 바꾸지 않고 오류 문구 — 접미사 없는 코드가 생기는 것을 막음.
    """
    parsed = parse_project_code(current_code)
    if not parsed:
        return current_code, ''
    prefixes, suffixes = _maps()
    new_prefix = prefixes.get(company or '', '') or parsed['prefix']
    manager = (manager or '').strip()
    if manager and manager != parsed['manager']:
        sfx = suffixes.get(manager, '')
        if not sfx:
            return current_code, (f"담당자 '{manager}'의 공사 코드 접미사가 등록되어 있지 않아 저장하지 않았습니다. "
                                  "관리자에게 접미사 등록을 요청해 주세요.")
        new_suffix = f'-{sfx}'
    else:
        new_suffix = parsed['suffix']
    new_code = f"{new_prefix}{parsed['number']:04d}{new_suffix}"
    return new_code, ''


def should_update_project_code(current_code, row_number, company, manager):
    """(옛 호환) 사업자 또는 담당자가 코드와 달라졌는가."""
    parsed = parse_project_code(current_code)
    if not parsed:
        return False
    if not parsed['manager']:
        return parsed['company'] != company
    return parsed['company'] != company or parsed['manager'] != manager
