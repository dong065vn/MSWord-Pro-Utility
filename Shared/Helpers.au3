; ============================================
; HELPERS.AU3 - Utility Functions
; ============================================

#include-once

; Xac dinh thu muc goc project cho ca Main.au3 va cac script trong Tests\
Func _GetProjectRoot()
    If FileExists(@ScriptDir & "\Resources") Then Return @ScriptDir
    If FileExists(@ScriptDir & "\..\Resources") Then Return @ScriptDir & "\.."
    Return @ScriptDir
EndFunc

; Lay duong dan toi thu muc Resources
Func _GetResourcesDir()
    Return _GetProjectRoot() & "\Resources"
EndFunc

; Lay nguon icon uu tien de set cho GUI/tray.
; Khi chay script, uu tien file ico trong project.
; Khi da compile, dung chinh exe de Windows lay du bo icon embedded.
Func _GetAppIconSource()
    Local $sProjectIcon = _GetProjectRoot() & "\app_icons\app_icon_rounded.ico"
    If FileExists($sProjectIcon) Then Return $sProjectIcon
    Local $sResourcesIcon = _GetResourcesDir() & "\icon.ico"
    If FileExists($sResourcesIcon) Then Return $sResourcesIcon
    Return @ScriptFullPath
EndFunc

Func _MakeButtonSquare($idButton)
    Local $hButton = GUICtrlGetHandle($idButton)
    If $hButton = 0 Then Return False
    DllCall("uxtheme.dll", "long", "SetWindowTheme", "hwnd", $hButton, "wstr", "", "wstr", "")
    Return Not @error
EndFunc

Func _CreateSquareButton($sText, $iX, $iY, $iW, $iH, $iStyle = $BS_FLAT, $iBgColor = -1)
    Local $btn
    If $iBgColor <> -1 Then
        $btn = GUICtrlCreateLabel($sText, $iX, $iY, $iW, $iH, _
            BitOR($SS_CENTER, $SS_CENTERIMAGE, $SS_NOTIFY))
        GUICtrlSetBkColor($btn, $iBgColor)
        GUICtrlSetColor($btn, 0xFFFFFF)
    Else
        $btn = GUICtrlCreateButton($sText, $iX, $iY, $iW, $iH, BitOR($iStyle, $BS_FLAT))
        _MakeButtonSquare($btn)
    EndIf
    GUICtrlSetFont($btn, 9, 400, 0, "Segoe UI")
    Return $btn
EndFunc

; Lay duong dan file hotkey dung chung
Func _GetHotkeyIniPath()
    Return _GetResourcesDir() & "\StyleHotkeys.ini"
EndFunc

; Lay duong dan thu muc backup hotkey
Func _GetHotkeyBackupDir()
    Return _GetProjectRoot() & "\HotkeyBackups"
EndFunc

; Cap nhat progress label
Func _UpdateProgress($sMsg)
    If Not IsDeclared("g_lblProgress") Then Return
    GUICtrlSetData($g_lblProgress, $sMsg)
EndFunc

; Cap nhat status label
Func _SetStatus($sMsg, $iColor = 0x27AE60)
    If Not IsDeclared("g_lblStatus") Then Return
    GUICtrlSetData($g_lblStatus, $sMsg)
    GUICtrlSetColor($g_lblStatus, $iColor)
EndFunc

; Thuc thi an toan (chong double-click)
Func _SafeExecute($sFuncName)
    If $g_bProcessing Then Return

    ; Connection/UI refresh handlers are short and can show Word/UAC dialogs.
    ; Running them without process wrapping avoids lifecycle noise during startup/connect.
    If $sFuncName = "_RefreshWordDocsList" Then Return Call($sFuncName)

    $g_bProcessing = True
    GUISetCursor(15, 1, $g_hGUI)
    _Process_Start($sFuncName, 0, "")
    Local $vResult = Call($sFuncName)
    If @error Then
        _Process_Fail("Khong goi duoc ham: " & $sFuncName)
    ElseIf $g_sProcessName = $sFuncName Then
        _Process_Done()
    EndIf
    GUISetCursor(2, 0, $g_hGUI)
    $g_bProcessing = False
    Return $vResult
EndFunc

Func _SafeExecute1($sFuncName, $vArg1)
    If $g_bProcessing Then Return
    $g_bProcessing = True
    GUISetCursor(15, 1, $g_hGUI)
    _Process_Start($sFuncName, 0, "")
    Local $vResult = Call($sFuncName, $vArg1)
    If @error Then
        _Process_Fail("Khong goi duoc ham: " & $sFuncName)
    ElseIf $g_sProcessName = $sFuncName Then
        _Process_Done()
    EndIf
    GUISetCursor(2, 0, $g_hGUI)
    $g_bProcessing = False
    Return $vResult
EndFunc

; Undo action
Func _UndoAction()
    If Not _CheckConnection() Then Return
    $g_oDoc.Undo()
    _UpdateProgress("Da Undo!")
EndFunc

; Log to preview
Func _LogPreview($sMsg)
    GUICtrlSetData($g_editPreview, $sMsg)
EndFunc

; Append to preview
Func _AppendPreview($sMsg)
    If Not IsDeclared("g_editPreview") Then Return
    Local $sCurrent = GUICtrlRead($g_editPreview)
    If $sCurrent = "" Then
        GUICtrlSetData($g_editPreview, $sMsg)
    Else
        GUICtrlSetData($g_editPreview, $sCurrent & @CRLF & $sMsg)
    EndIf
EndFunc

; Lay ten file khong co phan mo rong
Func _GetFileBaseName($sFileName)
    Local $iPos = StringInStr($sFileName, ".", 0, -1)
    If $iPos > 0 Then Return StringLeft($sFileName, $iPos - 1)
    Return $sFileName
EndFunc

Func _ShowSafetyOptionsDialog()
    Local $hPopup = GUICreate("Tuy chon an toan", 460, 330, -1, -1, BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU), -1, $g_hGUI)
    GUISetBkColor(0xF5F5F5, $hPopup)

    GUICtrlCreateLabel("TUY CHON DOM / PROCESS AN TOAN", 18, 14, 420, 24)
    GUICtrlSetFont(-1, 11, 700)
    GUICtrlSetColor(-1, 0x2C3E50)

    GUICtrlCreateGroup(" Pham vi mac dinh ", 18, 48, 424, 58)
    Local $chkPreferSelection = GUICtrlCreateCheckbox(" Uu tien vung chon khi tool ho tro scope tu dong", 32, 72, 380, 22)
    GUICtrlSetState($chkPreferSelection, ($g_bDomPreferSelection ? $GUI_CHECKED : $GUI_UNCHECKED))
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    GUICtrlCreateGroup(" Bo qua noi dung nhay cam ", 18, 112, 424, 92)
    Local $chkSkipTables = GUICtrlCreateCheckbox(" Bo qua bang", 32, 136, 150, 22)
    Local $chkSkipEquations = GUICtrlCreateCheckbox(" Bo qua cong thuc", 230, 136, 170, 22)
    Local $chkSkipLinks = GUICtrlCreateCheckbox(" Bo qua hyperlink", 32, 164, 170, 22)
    Local $chkSkipFields = GUICtrlCreateCheckbox(" Bo qua field", 230, 164, 170, 22)
    GUICtrlSetState($chkSkipTables, ($g_bDomSkipTables ? $GUI_CHECKED : $GUI_UNCHECKED))
    GUICtrlSetState($chkSkipEquations, ($g_bDomSkipEquations ? $GUI_CHECKED : $GUI_UNCHECKED))
    GUICtrlSetState($chkSkipLinks, ($g_bDomSkipLinks ? $GUI_CHECKED : $GUI_UNCHECKED))
    GUICtrlSetState($chkSkipFields, ($g_bDomSkipFields ? $GUI_CHECKED : $GUI_UNCHECKED))
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    GUICtrlCreateGroup(" Theo doi tien trinh ", 18, 210, 424, 58)
    Local $chkDryRun = GUICtrlCreateCheckbox(" Preview / Dry-run: tinh toan va log, khong sua noi dung", 32, 232, 330, 22)
    Local $chkSummary = GUICtrlCreateCheckbox(" Ghi summary vao khung log sau moi chuc nang", 32, 254, 330, 22)
    GUICtrlSetState($chkDryRun, ($g_bDomDryRun ? $GUI_CHECKED : $GUI_UNCHECKED))
    GUICtrlSetState($chkSummary, ($g_bDomShowSummary ? $GUI_CHECKED : $GUI_UNCHECKED))
    GUICtrlCreateGroup("", -99, -99, 1, 1)

    Local $btnSave = _CreateSquareButton("Luu", 248, 286, 90, 30, $BS_FLAT)
    Local $btnCancel = _CreateSquareButton("Huy", 350, 286, 90, 30, $BS_FLAT)
    GUISetState(@SW_SHOW, $hPopup)

    While 1
        Local $iMsg = GUIGetMsg()
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnCancel
                GUIDelete($hPopup)
                Return
            Case $btnSave
                $g_bDomPreferSelection = (GUICtrlRead($chkPreferSelection) = $GUI_CHECKED)
                $g_bDomSkipTables = (GUICtrlRead($chkSkipTables) = $GUI_CHECKED)
                $g_bDomSkipEquations = (GUICtrlRead($chkSkipEquations) = $GUI_CHECKED)
                $g_bDomSkipLinks = (GUICtrlRead($chkSkipLinks) = $GUI_CHECKED)
                $g_bDomSkipFields = (GUICtrlRead($chkSkipFields) = $GUI_CHECKED)
                $g_bDomDryRun = (GUICtrlRead($chkDryRun) = $GUI_CHECKED)
                $g_bDomShowSummary = (GUICtrlRead($chkSummary) = $GUI_CHECKED)
                GUIDelete($hPopup)
                _SetStatus("Da luu tuy chon an toan", 0x27AE60)
                _UpdateProgress("Safe opts: dry-run=" & $g_bDomDryRun & ", skip tables=" & $g_bDomSkipTables & ", skip math=" & $g_bDomSkipEquations)
                Return
        EndSwitch
    WEnd
EndFunc

; NOTE: _BackupDocument(), _SaveDocument(), _ShowHelp() are defined in Dialogs.au3

