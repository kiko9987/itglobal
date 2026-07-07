# itgfolder:// 프로토콜 설치 가이드 (직원용)

관리 사이트(pm.itg-aircon.com)에서 프로젝트 **문서 폴더 링크**를 클릭하면 각자 자기 PC의 **탐색기**로 폴더가 열리도록 하는 설치입니다.

## 필요한 파일

1. `install-itg-folder.reg`
2. `open-itg-folder.vbs`

관리자로부터 두 파일을 전달받으세요.

## 사전 요구사항

- **Google Drive Desktop** 설치 및 로그인 완료
- Drive Desktop이 **G: 드라이브**에 마운트되어 있을 것 (기본값)
- **Windows 10/11**

Drive Desktop 미설치 상태에서는 이 프로토콜이 작동하지 않습니다.

## 설치 절차 (30초)

### 1단계 — VBS 파일 배치

`C:\ITG\` 폴더가 없으면 만들고, 그 안에 `open-itg-folder.vbs`를 넣으세요.

**경로가 반드시 이 그대로여야 합니다:**
```
C:\ITG\open-itg-folder.vbs
```

### 2단계 — 프로토콜 등록

`install-itg-folder.reg` 파일을 **더블클릭**하세요.

Windows 경고창이 뜨면:
1. "레지스트리 편집기" 경고 → **[예]**
2. "성공적으로 추가되었습니다" → **[확인]**

### 3단계 — 브라우저 재시작

Chrome/Edge를 완전히 종료했다가 다시 열어주세요. (이미 열려 있는 창은 새로고침만으로 부족합니다)

## 사용 방법

1. 관리 사이트(pm.itg-aircon.com) 로그인
2. 프로젝트 상세 → **문서 폴더** 영역의 파란색 링크 클릭
3. 브라우저 팝업: "이 사이트에서 **itgfolder**를 열려고 합니다"
   - **[itgfolder 열기]** 클릭
   - 원하면 "*pm.itg-aircon.com의 요청을 항상 허용*" 체크 → 다음부터 팝업 안 뜸
4. Windows 탐색기가 자동으로 프로젝트 폴더를 엽니다

## 문제 해결

### 링크를 클릭해도 아무 반응이 없어요
- Google Drive Desktop이 실행 중인지 확인 (트레이 아이콘)
- Drive Desktop이 **G: 드라이브**에 마운트됐는지 확인 (탐색기에서 G: 드라이브 확인)
- 브라우저 팝업 자체가 안 뜨면 → 프로토콜 등록 실패. `install-itg-folder.reg`를 다시 실행

### "이 사이트에서 itgfolder를 열려고 합니다" 팝업이 안 떠요
- `C:\ITG\open-itg-folder.vbs` 파일이 있는지 확인
- `install-itg-folder.reg`를 관리자 권한으로 다시 실행
- 브라우저 완전 재시작

### 탐색기가 열리긴 하는데 "폴더를 찾을 수 없다"고 나와요
- 해당 프로젝트의 폴더가 Drive Desktop에 아직 동기화되지 않은 상태입니다. Drive Desktop 트레이 아이콘 → 동기화 상태 확인
- 프로젝트가 방금 등록된 경우 몇 분 대기 후 재시도

### cmd/PowerShell 창이 잠깐 뜨다 사라져요
- 정상 아닙니다. VBS 파일이 아니라 이전 버전의 설정이 남아있을 수 있으니 관리자에게 문의

## 제거 방법

향후 사용을 중단하려면:

1. `install-itg-folder.reg` 파일을 **메모장으로 열어**서 마지막 줄을 확인
2. Windows 실행(`Win+R`) → `regedit`
3. `HKEY_CLASSES_ROOT\itgfolder` 키를 우클릭 → **삭제**
4. `C:\ITG\` 폴더 삭제

## 작동 원리 (참고)

- 관리 사이트의 폴더 링크는 실제로 `itgfolder://{폴더ID}` 형식
- 클릭 시 Windows가 등록된 프로토콜을 확인하고 `open-itg-folder.vbs`를 호출
- VBS가 폴더 ID를 파싱해 `explorer.exe G:\.shortcut-targets-by-id\{ID}` 실행
- Drive Desktop이 미러링한 폴더가 탐색기에서 열림

서버는 이 과정에 관여하지 않습니다. 모든 것이 각자 PC에서 실행되므로 인터넷 연결이 끊겨도 Drive Desktop이 동기화한 파일은 접근 가능합니다.
