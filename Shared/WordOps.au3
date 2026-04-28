; ============================================
; WORDOPS.AU3 - Word Operations Chung
; ============================================

#include-once

; Find & Replace helper
Func _DoReplace($oFind, $sFind, $sReplace)
    If Not IsObj($oFind) Then Return 0
    If $g_bDomDryRun Then
        _Process_AddSkipped()
        Return 0
    EndIf
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    Local $vResult = $oFind.Execute($sFind, False, False, False, False, False, True, 1, False, $sReplace, $WD_REPLACE_ALL)
    If @error Then
        _Process_AddError()
        Return 0
    EndIf
    If $vResult Then _Process_AddChanged()
    Return ($vResult ? 1 : 0)
EndFunc

; Fix line spacing cho range
Func _FixLineSpacingRange($oRange, $fLineSpacing)
    If Not IsObj($oRange) Then Return
    If _Dom_ShouldSkipRange($oRange) Then
        _Process_AddSkipped()
        Return
    EndIf
    If $g_bDomDryRun Then
        _Process_AddSkipped()
        Return
    EndIf
    _UpdateProgress("Dang fix cach dong " & $fLineSpacing & "...")

    Local $oParaFormat = $oRange.ParagraphFormat
    If Not IsObj($oParaFormat) Then Return
    
    ; Reset spacing before/after if checked
    If GUICtrlRead($g_chkFixSpacingBefore) = $GUI_CHECKED Then
        $oParaFormat.SpaceBefore = 0
        $oParaFormat.SpaceAfter = 0
    EndIf

    ; Apply line spacing
    $oParaFormat.LineSpacingRule = $WD_LINE_SPACE_MULTIPLE
    $oParaFormat.LineSpacing = 12 * $fLineSpacing
EndFunc

; Lay alignment constant tu string
Func _GetAlignmentConst($sAlign)
    Switch $sAlign
        Case "Justify"
            Return $WD_ALIGN_JUSTIFY
        Case "Left"
            Return $WD_ALIGN_LEFT
        Case "Center"
            Return $WD_ALIGN_CENTER
        Case "Right"
            Return $WD_ALIGN_RIGHT
    EndSwitch
    Return $WD_ALIGN_JUSTIFY
EndFunc
