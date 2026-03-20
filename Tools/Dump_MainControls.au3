; Liet ke child controls trong cua so chinh de phuc vu smoke test

Global Const $APP_TITLE = "PDF to Word Fixer Pro v6.1"

If Not WinWait($APP_TITLE, "", 10) Then
    ConsoleWrite("FAIL: Khong tim thay cua so app." & @CRLF)
    Exit 1
EndIf

Local $hWnd = WinGetHandle($APP_TITLE)
If $hWnd = "" Then
    ConsoleWrite("FAIL: Khong lay duoc handle." & @CRLF)
    Exit 2
EndIf

Local $sList = WinGetClassList($hWnd)
ConsoleWrite("WINDOW_CLASSES" & @CRLF & $sList & @CRLF)

Local $aClasses = StringSplit(StringStripCR($sList), @LF, 2)
For $i = 0 To UBound($aClasses) - 1
    Local $sClass = $aClasses[$i]
    If $sClass = "" Then ContinueLoop

    Local $hCtrl = ControlGetHandle($hWnd, "", $sClass)
    Local $sText = ControlGetText($hWnd, "", $sClass)
    Local $aPos = ControlGetPos($hWnd, "", $sClass)
    Local $sPos = ""
    If IsArray($aPos) Then
        $sPos = $aPos[0] & "," & $aPos[1] & "," & $aPos[2] & "," & $aPos[3]
    EndIf
    ConsoleWrite($sClass & " | text=" & $sText & " | pos=" & $sPos & @CRLF)
Next
