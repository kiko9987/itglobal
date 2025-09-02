# ITG Sheets Dashboard — Final Draft

- 기본 검색: **전체 보기**
- 정산 모듈: **본인 담당 행만 보기** (`/api/settlement/projects`)
- 신규 프로젝트: **자동 코드 생성** (`/api/projects/auto`)
- 프런트: 목록 + **신규 프로젝트 등록 폼(사업자/담당자 드롭다운 + 검증)**
- 설정: `.env` 또는 `dashboard/credentials.json`

## Quickstart
```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python dashboard/app.py
```
Open http://localhost:5000
