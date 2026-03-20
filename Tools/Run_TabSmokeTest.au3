; Smoke test: mo app va click lan luot tat ca tab

Global Const $APP_EXE = @ScriptDir & "\..\Main_compiled.exe"
Global Const $APP_TITLE = "PDF to Word Fixer Pro v6.1"
Global Const $LOG_PATH = @ScriptDir & "\..\Tests\Logs\TabSmokeTest.log"

DirCreate(@ScriptDir & "\..\Tests\Logs")
FileDelete($LOG_PATH)

Func _WriteLog($sMsg)
    Local $hFile = FileOpen($LOG_PATH, 1 + 8)
    If $hFile <> -1 Then
        FileWriteLine($hFile, $sMsg)
        FileClose($hFile)
    EndIf
EndFunc

If Not FileExists($APP_EXE) Then
    _WriteLog("FAIL: Khong tim thay file build: " & $APP_EXE)
    Exit 1
EndIf

Local $hWnd = WinGetHandle($APP_TITLE)
If $hWnd = "" Then
    Run('"' & $APP_EXE & '"')
    If Not WinWait($APP_TITLE, "", 10) Then
        _WriteLog("FAIL: App khong mo duoc trong 10 giay.")
        Exit 2
    EndIf
    $hWnd = WinGetHandle($APP_TITLE)
EndIf

If $hWnd = "" Then
    _WriteLog("FAIL: Khong lay duoc handle cua cua so app.")
    Exit 3
EndIf

WinActivate($hWnd)
If Not WinWaitActive($hWnd, "", 5) Then
    _WriteLog("FAIL: Khong kich hoat duoc cua so app.")
    Exit 4
EndIf

Local $aPos = WinGetPos($hWnd)
If @error Or Not IsArray($aPos) Then
    _WriteLog("FAIL: Khong lay duoc vi tri cua so.")
    Exit 5
EndIf

Local $aTabs[9][2] = [ _
    ["PDF Fix", 60], _
    ["Format", 135], _
    ["Tools", 205], _
    ["TOC", 275], _
    ["Copy Style", 365], _
    ["Advanced", 455], _
    ["Quick Utils", 545], _
    ["Smart Fix", 625], _
    ["AI Format", 700] _
]

Local $iHeaderY = 155

_WriteLog("START: Tab smoke test")

For $i = 0 To UBound($aTabs) - 1
    Local $iX = $aPos[0] + $aTabs[$i][1]
    Local $iY = $aPos[1] + $iHeaderY

    MouseClick("left", $iX, $iY, 1, 0)
    Sleep(400)

    If Not WinExists($hWnd) Then
        _WriteLog("FAIL: App bi dong/crash khi chuyen tab '" & $aTabs[$i][0] & "'.")
        Exit 10 + $i
    EndIf

    _WriteLog("PASS: " & $aTabs[$i][0])
Next

_WriteLog("PASS: Da chuyen het 9 tab ma khong thay app bi dong.")
Exit 0
