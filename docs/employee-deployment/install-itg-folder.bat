@echo off
setlocal EnableDelayedExpansion
title ITG 폴더 프로토콜 설치

REM ==============================================================
REM  itgfolder:// 프로토콜 원클릭 설치 (관리자 권한 불필요)
REM  더블클릭으로 실행.
REM  - C:\ITG\open-itg-folder.vbs 배치 (base64 embed → certutil decode)
REM  - HKCU\Software\Classes\itgfolder 프로토콜 등록
REM ==============================================================

echo.
echo [1/3] C:\ITG 폴더 준비...
if not exist "C:\ITG" mkdir "C:\ITG"

echo [2/3] open-itg-folder.vbs 배치...
set "B64=%TEMP%\itg_vbs.b64"
(
echo JyBpdGdmb2xkZXI6Ly8g7ZSE66Gc7Yag7L2cIO2VuOuTpOufrAonIOu4jOudvOyasOyggOyXkOyE
echo nCBpdGdmb2xkZXI6Ly9GT0xERVJfSUQg7YG066atIOyLnCDtmLjstpzrkKguCicgV2luZG93c+qw
echo gCB3c2NyaXB0LmV4ZeuhnCDsi6TtlontlZjrr4DroZwg7LC9IOyViCDrnLguCicKJyDrj5nsnpE6
echo IFVSTOyXkOyEnCBmb2xkZXIgSUQg7LaU7LacIO2bhAonICAgZXhwbG9yZXIuZXhlICJHOlwuc2hv
echo cnRjdXQtdGFyZ2V0cy1ieS1pZFx7Rk9MREVSX0lEfSIKJyDroZwgR29vZ2xlIERyaXZlIERlc2t0
echo b3DsnbQg66+465+s66eB7ZWcIO2PtOuNlOulvCDtg5Dsg4nquLDroZwg7Je8LgoKT3B0aW9uIEV4
echo cGxpY2l0CgpEaW0gdXJsLCBmb2xkZXJJZCwgZHJpdmVQYXRoLCBzaGVsbAp1cmwgPSBXU2NyaXB0
echo LkFyZ3VtZW50cygwKQoKJyBwcmVmaXgg7KCc6rGwOiAiaXRnZm9sZGVyOi8vIiwgIml0Z2ZvbGRl
echo cjovIiwgIml0Z2ZvbGRlcjoiIOyInOycvOuhnCDsi5zrj4QKSWYgTGVmdChMQ2FzZSh1cmwpLCAx
echo MikgPSAiaXRnZm9sZGVyOi8vIiBUaGVuCiAgICBmb2xkZXJJZCA9IE1pZCh1cmwsIDEzKQpFbHNl
echo SWYgTGVmdChMQ2FzZSh1cmwpLCAxMSkgPSAiaXRnZm9sZGVyOi8iIFRoZW4KICAgIGZvbGRlcklk
echo ID0gTWlkKHVybCwgMTIpCkVsc2VJZiBMZWZ0KExDYXNlKHVybCksIDEwKSA9ICJpdGdmb2xkZXI6
echo IiBUaGVuCiAgICBmb2xkZXJJZCA9IE1pZCh1cmwsIDExKQpFbHNlCiAgICBmb2xkZXJJZCA9IHVy
echo bApFbmQgSWYKCicgdHJhaWxpbmcgc2xhc2gg7KCc6rGwICjruIzrnbzsmrDsoIDsl5Ag65Sw6528
echo IOu2meq4sOuPhCDtlagpCkRvIFdoaWxlIFJpZ2h0KGZvbGRlcklkLCAxKSA9ICIvIgogICAgZm9s
echo ZGVySWQgPSBMZWZ0KGZvbGRlcklkLCBMZW4oZm9sZGVySWQpIC0gMSkKTG9vcAoKSWYgTGVuKGZv
echo bGRlcklkKSA9IDAgVGhlbgogICAgTXNnQm94ICLtj7TrjZQgSUTqsIAg67mE7Ja07J6I7Iq164uI
echo 64ukLiIsIHZiQ3JpdGljYWwsICJJVEcgRm9sZGVyIgogICAgV1NjcmlwdC5RdWl0IDEKRW5kIElm
echo Cgpkcml2ZVBhdGggPSAiRzpcLnNob3J0Y3V0LXRhcmdldHMtYnktaWRcIiAmIGZvbGRlcklkCgpT
echo ZXQgc2hlbGwgPSBDcmVhdGVPYmplY3QoIldTY3JpcHQuU2hlbGwiKQpzaGVsbC5SdW4gImV4cGxv
echo cmVyLmV4ZSAiIiIgJiBkcml2ZVBhdGggJiAiIiIiLCAxLCBGYWxzZQo=
) > "%B64%"
certutil -decode "%B64%" "C:\ITG\open-itg-folder.vbs" >nul
if errorlevel 1 (
    echo   [실패] VBS 파일 생성 실패. 관리자 권한으로 재시도하거나 문의.
    del "%B64%" 2>nul
    pause
    exit /b 1
)
del "%B64%"

echo [3/3] 프로토콜 등록 (HKCU)...
reg add "HKCU\Software\Classes\itgfolder" /ve /d "URL:ITG Folder Protocol" /f >nul
reg add "HKCU\Software\Classes\itgfolder" /v "URL Protocol" /t REG_SZ /d "" /f >nul
reg add "HKCU\Software\Classes\itgfolder\DefaultIcon" /ve /d "explorer.exe,1" /f >nul
reg add "HKCU\Software\Classes\itgfolder\shell\open\command" /ve /d "wscript.exe \"C:\ITG\open-itg-folder.vbs\" \"%%1\"" /f >nul

echo.
echo ============================================
echo   설치 완료
echo ============================================
echo   VBS  : C:\ITG\open-itg-folder.vbs
echo   레지스트리: HKCU\Software\Classes\itgfolder
echo.
echo   Chrome/Edge를 완전히 종료 후 재실행하세요.
echo   그 후 관리 사이트에서 문서 폴더 링크 클릭 → 탐색기 자동 실행.
echo ============================================
echo.
pause
endlocal
