' itgfolder:// 프로토콜 핸들러
' 브라우저에서 itgfolder://FOLDER_ID 클릭 시 호출됨.
' Windows가 wscript.exe로 실행하므로 창 안 뜸.
'
' 동작: URL에서 folder ID 추출 후
'   explorer.exe "G:\.shortcut-targets-by-id\{FOLDER_ID}"
' 로 Google Drive Desktop이 미러링한 폴더를 탐색기로 염.

Option Explicit

Dim url, folderId, drivePath, shell
url = WScript.Arguments(0)

' prefix 제거: "itgfolder://", "itgfolder:/", "itgfolder:" 순으로 시도
If Left(LCase(url), 12) = "itgfolder://" Then
    folderId = Mid(url, 13)
ElseIf Left(LCase(url), 11) = "itgfolder:/" Then
    folderId = Mid(url, 12)
ElseIf Left(LCase(url), 10) = "itgfolder:" Then
    folderId = Mid(url, 11)
Else
    folderId = url
End If

' trailing slash 제거 (브라우저에 따라 붙기도 함)
Do While Right(folderId, 1) = "/"
    folderId = Left(folderId, Len(folderId) - 1)
Loop

If Len(folderId) = 0 Then
    MsgBox "폴더 ID가 비어있습니다.", vbCritical, "ITG Folder"
    WScript.Quit 1
End If

drivePath = "G:\.shortcut-targets-by-id\" & folderId

Set shell = CreateObject("WScript.Shell")
shell.Run "explorer.exe """ & drivePath & """", 1, False
