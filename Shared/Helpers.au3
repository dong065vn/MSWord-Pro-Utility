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
    GUICtrlSetData($g_lblProgress, $sMsg)
EndFunc

; Cap nhat status label
Func _SetStatus($sMsg, $iColor = 0x27AE60)
    GUICtrlSetData($g_lblStatus, $sMsg)
    GUICtrlSetColor($g_lblStatus, $iColor)
EndFunc

; Thuc thi an toan (chong double-click)
Func _SafeExecute($sFuncName)
    If $g_bProcessing Then Return
    $g_bProcessing = True
    GUISetCursor(15, 1, $g_hGUI)
    Call($sFuncName)
    GUISetCursor(2, 0, $g_hGUI)
    $g_bProcessing = False
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
    Local $sCurrent = GUICtrlRead($g_editPreview)
    GUICtrlSetData($g_editPreview, $sCurrent & @CRLF & $sMsg)
EndFunc

; Lay ten file khong co phan mo rong
Func _GetFileBaseName($sFileName)
    Local $iPos = StringInStr($sFileName, ".", 0, -1)
    If $iPos > 0 Then Return StringLeft($sFileName, $iPos - 1)
    Return $sFileName
EndFunc

; NOTE: _BackupDocument(), _SaveDocument(), _ShowHelp() are defined in Dialogs.au3
