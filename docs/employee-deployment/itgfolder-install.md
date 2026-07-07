# itgfolder:// 프로토콜 설치 가이드 (직원용)

관리 사이트(pm.itg-aircon.com)에서 프로젝트 **문서 폴더 링크**를 클릭하면 각자 자기 PC의 **탐색기**로 폴더가 열리도록 하는 설치입니다.

## 사전 요구사항

- **Google Drive Desktop** 설치 및 로그인 완료
- Drive Desktop이 **G: 드라이브**에 마운트되어 있을 것 (기본값)
- **Windows 10/11**

Drive Desktop 미설치 상태에서는 이 프로토콜이 작동하지 않습니다.

## 설치 (10초)

### 1단계 — 설치 파일 실행

관리자로부터 `install-itg-folder.bat` 파일을 전달받으세요.

파일을 **더블클릭**하세요.

Windows SmartScreen이 뜨면:
1. **[추가 정보]** 클릭
2. **[실행]** 클릭

검은 창이 뜨고 "설치 완료" 메시지가 나온 뒤 아무 키나 누르면 창이 닫힙니다.

관리자 권한 불필요. 현재 로그인한 사용자에게만 등록됩니다.

### 2단계 — 브라우저 재시작

Chrome/Edge를 **완전히 종료**했다가 다시 열어주세요. (이미 열려 있는 창은 새로고침만으로 부족합니다)

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
- 브라우저 팝업 자체가 안 뜨면 → 프로토콜 등록 실패. `install-itg-folder.bat`을 다시 실행

### "이 사이트에서 itgfolder를 열려고 합니다" 팝업이 안 떠요
- `install-itg-folder.bat`을 다시 실행
- 브라우저 완전 재시작

### 탐색기가 열리긴 하는데 "폴더를 찾을 수 없다"고 나와요
- 해당 프로젝트의 폴더가 Drive Desktop에 아직 동기화되지 않은 상태입니다. Drive Desktop 트레이 아이콘 → 동기화 상태 확인
- 프로젝트가 방금 등록된 경우 몇 분 대기 후 재시도

### cmd/PowerShell 창이 잠깐 뜨다 사라져요
- 정상 아닙니다. 설치가 잘못됐을 수 있으니 관리자에게 문의

## 제거 방법

향후 사용을 중단하려면:

1. Windows 실행(`Win+R`) → `regedit`
2. `HKEY_CURRENT_USER\Software\Classes\itgfolder` 키를 우클릭 → **삭제**
3. `C:\ITG\` 폴더 삭제

## 설치되는 내용 (참고)

`install-itg-folder.bat`이 자동으로 처리하는 것:

- `C:\ITG\open-itg-folder.vbs` 배치 — URL에서 폴더 ID 파싱 후 탐색기 실행하는 스크립트
- `HKEY_CURRENT_USER\Software\Classes\itgfolder` 프로토콜 등록 — 클릭 시 위 VBS 호출

## 작동 원리 (참고)

- 관리 사이트의 폴더 링크는 실제로 `itgfolder://{폴더ID}` 형식
- 클릭 시 Windows가 등록된 프로토콜을 확인하고 `open-itg-folder.vbs`를 호출
- VBS가 폴더 ID를 파싱해 `explorer.exe G:\.shortcut-targets-by-id\{ID}` 실행
- Drive Desktop이 미러링한 폴더가 탐색기에서 열림

서버는 이 과정에 관여하지 않습니다. 모든 것이 각자 PC에서 실행되므로 인터넷 연결이 끊겨도 Drive Desktop이 동기화한 파일은 접근 가능합니다.
