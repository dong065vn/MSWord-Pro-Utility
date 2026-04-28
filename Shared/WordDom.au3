; ============================================
; WORDDOM.AU3 - Safe Word DOM helpers
; ============================================

#include-once

Func _Dom_IsObject($oValue)
    Return IsObj($oValue)
EndFunc

Func _Dom_HasDocument()
    Return IsObj($g_oWord) And IsObj($g_oDoc)
EndFunc

Func _Dom_GetSelectionRange()
    If Not IsObj($g_oWord) Then Return 0
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return 0
    If $oSel.Type = 1 Then Return 0
    Local $oRange = $oSel.Range
    If Not IsObj($oRange) Then Return 0
    If $oRange.Start = $oRange.End Then Return 0
    Return $oRange
EndFunc

Func _Dom_GetScopeRange($sScope = "")
    If Not IsObj($g_oDoc) Then Return 0
    Local $sMode = StringLower(StringStripWS($sScope, 3))
    If $sMode = "" And $g_bDomPreferSelection Then $sMode = "selection"

    If $sMode = "selection" Or $sMode = "vung chon" Then
        Local $oSelectionRange = _Dom_GetSelectionRange()
        If IsObj($oSelectionRange) Then Return $oSelectionRange
        If $sScope <> "" Then Return 0
    EndIf

    Return $g_oDoc.Content
EndFunc

Func _Dom_GetScopeLabel($oRange)
    If Not IsObj($oRange) Then Return "khong hop le"
    If IsObj($g_oDoc) And IsObj($g_oDoc.Content) Then
        If $oRange.Start = $g_oDoc.Content.Start And $oRange.End = $g_oDoc.Content.End Then Return "toan bo tai lieu"
    EndIf
    Return "vung chon"
EndFunc

Func _Dom_IsRangeUsable($oRange)
    If Not IsObj($oRange) Then Return False
    Local $iStart = $oRange.Start
    If @error Then Return False
    Local $iEnd = $oRange.End
    If @error Then Return False
    Return ($iEnd >= $iStart)
EndFunc

Func _Dom_RangeOverlaps($oScopeRange, $iStart, $iEnd)
    If Not IsObj($oScopeRange) Then Return False
    Return Not ($iEnd <= $oScopeRange.Start Or $iStart >= $oScopeRange.End)
EndFunc

Func _Dom_ShouldSkipRange($oRange)
    If Not IsObj($oRange) Then Return True

    If $g_bDomSkipTables Then
        Local $oTables = $oRange.Tables
        If IsObj($oTables) And $oTables.Count > 0 Then Return True
    EndIf

    If $g_bDomSkipFields Then
        Local $oFields = $oRange.Fields
        If IsObj($oFields) And $oFields.Count > 0 Then Return True
    EndIf

    If $g_bDomSkipLinks Then
        Local $oLinks = $oRange.Hyperlinks
        If IsObj($oLinks) And $oLinks.Count > 0 Then Return True
    EndIf

    If $g_bDomSkipEquations Then
        Local $oMaths = $oRange.OMaths
        If IsObj($oMaths) And $oMaths.Count > 0 Then Return True
        Local $oInlineShapes = $oRange.InlineShapes
        If IsObj($oInlineShapes) Then
            For $i = 1 To $oInlineShapes.Count
                Local $oShape = $oInlineShapes.Item($i)
                If IsObj($oShape) And _Dom_IsMathInlineShape($oShape) Then Return True
            Next
        EndIf
    EndIf

    Return False
EndFunc

Func _Dom_IsMathInlineShape($oShape)
    If Not IsObj($oShape) Then Return False
    Local $sProgId = ""
    $sProgId = $oShape.OLEFormat.ProgID
    If @error Then Return False
    $sProgId = StringLower($sProgId)
    Return StringInStr($sProgId, "equation") Or StringInStr($sProgId, "mathtype")
EndFunc

Func _Dom_SafeFindReplace($oRange, $sFind, $sReplace, $bMatchWildcards = False, $bMatchCase = False, $bWholeWord = False)
    If Not _Dom_IsRangeUsable($oRange) Then Return 0
    If $g_bDomDryRun Then Return 0

    Local $oFind = $oRange.Find
    If Not IsObj($oFind) Then Return SetError(1, 0, 0)

    For $iRetry = 1 To 3
        $oFind.ClearFormatting()
        $oFind.Replacement.ClearFormatting()
        $oFind.Text = $sFind
        $oFind.Replacement.Text = $sReplace
        $oFind.Forward = True
        $oFind.Wrap = 0
        $oFind.Format = False
        $oFind.MatchCase = $bMatchCase
        $oFind.MatchWholeWord = $bWholeWord
        $oFind.MatchWildcards = $bMatchWildcards
        Local $vResult = $oFind.Execute(Default, Default, Default, Default, Default, Default, Default, Default, Default, Default, 2)
        If Not @error Then Return ($vResult ? 1 : 0)
        Sleep(150 * $iRetry)
    Next

    _Process_AddError()
    Return SetError(2, 0, 0)
EndFunc

Func _Dom_ForEachParagraphStart($oRange)
    If Not _Dom_IsRangeUsable($oRange) Then Return 0
    Local $oParas = $oRange.Paragraphs
    If Not IsObj($oParas) Then Return 0
    Return $oParas.Count
EndFunc

Func _Dom_SetRangeText($oRange, $sText)
    If Not _Dom_IsRangeUsable($oRange) Then Return False
    If _Dom_ShouldSkipRange($oRange) Then
        _Process_AddSkipped()
        Return False
    EndIf
    If $g_bDomDryRun Then
        _Process_AddSkipped()
        Return False
    EndIf
    $oRange.Text = $sText
    If @error Then
        _Process_AddError()
        Return False
    EndIf
    Return True
EndFunc

Func _Dom_GetCollectionCount($oCollection)
    If Not IsObj($oCollection) Then Return 0
    Local $iCount = $oCollection.Count
    If @error Then Return 0
    Return $iCount
EndFunc

Func _Dom_SafeDelete($oObject)
    If Not IsObj($oObject) Then Return False
    If $g_bDomDryRun Then
        _Process_AddSkipped()
        Return False
    EndIf
    $oObject.Delete()
    If @error Then
        _Process_AddError()
        Return False
    EndIf
    Return True
EndFunc
