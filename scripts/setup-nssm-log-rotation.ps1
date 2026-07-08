# NSSM 로그 자동 회전 설정 (관리자 권한 PowerShell에서 실행)
#
# 이유: NSSM이 리다이렉트하는 service_stdout.log / service_stderr.log가
#   무제한 증가해 서버 스토리지 폭주 위험. 매일 자정 + 20MB 초과 시 회전.
#
# 실행 방법:
#   1) 시작 → PowerShell 우클릭 → "관리자 권한으로 실행"
#   2) 아래 명령 붙여넣기:
#      Set-ExecutionPolicy -Scope Process Bypass
#      & "C:\Users\SECOM\Desktop\ITG-Project\Claude Project\scripts\setup-nssm-log-rotation.ps1"
#   3) 검증 결과 확인

$ErrorActionPreference = 'Stop'
$nssm = "$env:LOCALAPPDATA\Microsoft\WinGet\Links\nssm.exe"

if (-not (Test-Path $nssm)) {
    Write-Error "NSSM 실행 파일을 찾을 수 없음: $nssm"
    exit 1
}

Write-Output "==============================================="
Write-Output "NSSM 로그 회전 설정 (ITGFlask)"
Write-Output "==============================================="

# 설정 적용
& $nssm set ITGFlask AppRotateFiles 1        # 회전 활성화
& $nssm set ITGFlask AppRotateOnline 1       # 서비스 재시작 없이 회전
& $nssm set ITGFlask AppRotateSeconds 86400  # 매 24시간마다 (하루 1회)
& $nssm set ITGFlask AppRotateBytes 20971520 # 또는 20MB 초과 시

Write-Output ""
Write-Output "=== 검증 ==="
Write-Output "AppRotateFiles:   $(& $nssm get ITGFlask AppRotateFiles)  (1=활성)"
Write-Output "AppRotateOnline:  $(& $nssm get ITGFlask AppRotateOnline)  (1=재시작 없이)"
Write-Output "AppRotateSeconds: $(& $nssm get ITGFlask AppRotateSeconds) (86400=24시간)"
Write-Output "AppRotateBytes:   $(& $nssm get ITGFlask AppRotateBytes)   (20971520=20MB)"

Write-Output ""
Write-Output "회전 파일명 예시:"
Write-Output "  service_stdout.log        # 현재 활성"
Write-Output "  service_stdout-2026-07-08_120000.log  # 회전된 것"

Write-Output ""
Write-Output "설정 완료. 서비스는 재시작 불필요. 다음 회전 조건 도달 시 자동 회전."
