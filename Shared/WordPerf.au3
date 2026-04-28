; ============================================
; WORDPERF.AU3 - Word COM performance helpers
; ============================================

#include-once

Global Const $PERF_CTX_NAME = 0
Global Const $PERF_CTX_SCREEN_UPDATING = 1
Global Const $PERF_CTX_DISPLAY_ALERTS = 2
Global Const $PERF_CTX_STATUS_BAR = 3
Global Const $PERF_CTX_UNDO_STARTED = 4

Func _Perf_BeginWordBatch($sName = "")
    Local $aCtx[5]
    $aCtx[$PERF_CTX_NAME] = $sName
    $aCtx[$PERF_CTX_SCREEN_UPDATING] = ""
    $aCtx[$PERF_CTX_DISPLAY_ALERTS] = ""
    $aCtx[$PERF_CTX_STATUS_BAR] = ""
    $aCtx[$PERF_CTX_UNDO_STARTED] = False

    If Not IsObj($g_oWord) Then Return $aCtx

    $aCtx[$PERF_CTX_SCREEN_UPDATING] = $g_oWord.ScreenUpdating
    If Not @error Then $g_oWord.ScreenUpdating = False

    $aCtx[$PERF_CTX_DISPLAY_ALERTS] = $g_oWord.DisplayAlerts
    If Not @error Then $g_oWord.DisplayAlerts = 0

    $aCtx[$PERF_CTX_STATUS_BAR] = $g_oWord.StatusBar
    If Not @error And $sName <> "" Then $g_oWord.StatusBar = $sName

    If IsObj($g_oWord.UndoRecord) Then
        $g_oWord.UndoRecord.StartCustomRecord($sName = "" ? "PDF to Word Fixer" : $sName)
        If Not @error Then $aCtx[$PERF_CTX_UNDO_STARTED] = True
    EndIf

    Return $aCtx
EndFunc

Func _Perf_EndWordBatch(ByRef $aCtx)
    If Not IsObj($g_oWord) Or Not IsArray($aCtx) Then Return

    If $aCtx[$PERF_CTX_UNDO_STARTED] And IsObj($g_oWord.UndoRecord) Then
        $g_oWord.UndoRecord.EndCustomRecord()
    EndIf

    If $aCtx[$PERF_CTX_STATUS_BAR] <> "" Then $g_oWord.StatusBar = $aCtx[$PERF_CTX_STATUS_BAR]
    If $aCtx[$PERF_CTX_DISPLAY_ALERTS] <> "" Then $g_oWord.DisplayAlerts = $aCtx[$PERF_CTX_DISPLAY_ALERTS]
    If $aCtx[$PERF_CTX_SCREEN_UPDATING] <> "" Then $g_oWord.ScreenUpdating = $aCtx[$PERF_CTX_SCREEN_UPDATING]
EndFunc

Func _Perf_ShouldRenderStep($iCurrent, $iTotal, $iEvery = 25)
    If $iCurrent <= 0 Then Return True
    If $iTotal > 0 And $iCurrent >= $iTotal Then Return True
    If $iEvery <= 1 Then Return True
    If Mod($iCurrent, $iEvery) = 0 Then Return True

    If $iTotal > 0 Then
        Local $iPct = Int(($iCurrent / $iTotal) * 100)
        If Mod($iPct, 5) = 0 And $iPct <> $g_iProcessLastRenderPct Then Return True
    EndIf

    Return False
EndFunc

Func _Perf_GetTextSnapshot($oRange)
    If Not IsObj($oRange) Then Return SetError(1, 0, "")
    Local $sText = $oRange.Text
    If @error Then Return SetError(2, 0, "")
    Return $sText
EndFunc

Func _Perf_CollectionCount($oCollection)
    If Not IsObj($oCollection) Then Return 0
    Local $iCount = $oCollection.Count
    If @error Then Return 0
    Return $iCount
EndFunc

Func _Perf_NormalizeArrayCapacity(ByRef $aArray, ByRef $iCapacity, $iNeeded, $iGrowBy = 64)
    If $iGrowBy < 1 Then $iGrowBy = 64
    If $iCapacity >= $iNeeded Then Return $iCapacity

    Local $iNewCapacity = $iCapacity
    If $iNewCapacity < 1 Then $iNewCapacity = $iGrowBy
    While $iNewCapacity < $iNeeded
        $iNewCapacity += $iGrowBy
    WEnd

    ReDim $aArray[$iNewCapacity]
    $iCapacity = $iNewCapacity
    Return $iCapacity
EndFunc

Func _Perf_RetrySetRangeText($oRange, $sText, $iRetries = 3)
    If Not IsObj($oRange) Then Return False
    For $iRetry = 1 To $iRetries
        $oRange.Text = $sText
        If Not @error Then Return True
        Sleep(120 * $iRetry)
    Next
    Return False
EndFunc

Func _Perf_RetryDeleteRange($oRange, $iRetries = 3)
    If Not IsObj($oRange) Then Return False
    For $iRetry = 1 To $iRetries
        $oRange.Delete()
        If Not @error Then Return True
        Sleep(120 * $iRetry)
    Next
    Return False
EndFunc
