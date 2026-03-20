; ============================================
; UIHELPERS.AU3 - GUI Utilities
; ============================================

#include-once

; Tao button voi style
Func _CreateButton($sText, $iX, $iY, $iW, $iH, $iBgColor = -1, $iFontSize = 9, $iFontWeight = 400)
    Local $btn = GUICtrlCreateButton($sText, $iX, $iY, $iW, $iH)
    If $iBgColor <> -1 Then GUICtrlSetBkColor($btn, $iBgColor)
    If $iFontSize <> 9 Or $iFontWeight <> 400 Then
        GUICtrlSetFont($btn, $iFontSize, $iFontWeight)
    EndIf
    Return $btn
EndFunc

; Tao checkbox da check
Func _CreateCheckbox($sText, $iX, $iY, $iW, $iH, $bChecked = False)
    Local $chk = GUICtrlCreateCheckbox($sText, $iX, $iY, $iW, $iH)
    If $bChecked Then GUICtrlSetState($chk, $GUI_CHECKED)
    Return $chk
EndFunc

; Tao combo box voi data
Func _CreateCombo($iX, $iY, $iW, $iH, $sData, $sDefault = "")
    Local $cbo = GUICtrlCreateCombo("", $iX, $iY, $iW, $iH, $CBS_DROPDOWNLIST)
    GUICtrlSetData($cbo, $sData, $sDefault)
    Return $cbo
EndFunc

; Tao input voi gia tri mac dinh
Func _CreateInput($sDefault, $iX, $iY, $iW, $iH)
    Return GUICtrlCreateInput($sDefault, $iX, $iY, $iW, $iH)
EndFunc

; Tao label
Func _CreateLabel($sText, $iX, $iY, $iW, $iH)
    Return GUICtrlCreateLabel($sText, $iX, $iY, $iW, $iH)
EndFunc

; Tao group
Func _CreateGroup($sText, $iX, $iY, $iW, $iH)
    Return GUICtrlCreateGroup($sText, $iX, $iY, $iW, $iH)
EndFunc

; Dong group
Func _EndGroup()
    GUICtrlCreateGroup("", -99, -99, 1, 1)
EndFunc
