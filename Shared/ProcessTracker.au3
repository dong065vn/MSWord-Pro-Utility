; ============================================
; PROCESSTRACKER.AU3 - Shared progress/process tracking
; ============================================

#include-once

Func _Process_Start($sName, $iTotal = 0, $sScope = "")
    $g_sProcessName = $sName
    $g_iProcessTotal = $iTotal
    $g_iProcessCurrent = 0
    $g_iProcessChanged = 0
    $g_iProcessSkipped = 0
    $g_iProcessErrors = 0
    $g_sProcessScope = $sScope
    $g_sProcessLastMessage = "Bat dau"
    $g_hProcessTimer = TimerInit()
    $g_iProcessLastRender = -1
    $g_iProcessLastRenderPct = -1

    _SetStatus("Dang chay: " & $sName, 0x3498DB)
    _Process_Render("Bat dau")
EndFunc

Func _Process_Step($sMessage, $iCurrent = -1, $iChanged = -1, $iSkipped = -1, $iErrors = -1)
    If $iCurrent >= 0 Then $g_iProcessCurrent = $iCurrent
    If $iChanged >= 0 Then $g_iProcessChanged = $iChanged
    If $iSkipped >= 0 Then $g_iProcessSkipped = $iSkipped
    If $iErrors >= 0 Then $g_iProcessErrors = $iErrors
    $g_sProcessLastMessage = $sMessage

    If _Process_ShouldRenderStep($g_iProcessCurrent) Then
        _Process_Render($sMessage)
        _Process_PumpGui()
    EndIf
EndFunc

Func _Process_AddChanged($iDelta = 1)
    $g_iProcessChanged += $iDelta
EndFunc

Func _Process_AddSkipped($iDelta = 1)
    $g_iProcessSkipped += $iDelta
EndFunc

Func _Process_AddError($iDelta = 1)
    $g_iProcessErrors += $iDelta
EndFunc

Func _Process_Done($sSummary = "")
    If $sSummary = "" Then $sSummary = _Process_BuildSummary()
    _Process_Render($sSummary)
    _SetStatus("Hoan tat: " & $g_sProcessName, 0x27AE60)
    If $g_bDomShowSummary Then _AppendPreview("[DONE] " & $g_sProcessName & " - " & $sSummary)
EndFunc

Func _Process_Fail($sMessage)
    $g_iProcessErrors += 1
    _Process_Render("Loi: " & $sMessage)
    _SetStatus("Loi: " & $g_sProcessName, 0xE74C3C)
    _AppendPreview("[FAIL] " & $g_sProcessName & " - " & $sMessage)
EndFunc

Func _Process_BuildSummary()
    Local $sSummary = "processed=" & $g_iProcessCurrent
    If $g_iProcessTotal > 0 Then $sSummary &= "/" & $g_iProcessTotal
    $sSummary &= ", changed=" & $g_iProcessChanged
    $sSummary &= ", skipped=" & $g_iProcessSkipped
    $sSummary &= ", errors=" & $g_iProcessErrors
    If $g_sProcessScope <> "" Then $sSummary &= ", scope=" & $g_sProcessScope
    If $g_hProcessTimer <> 0 Then
        Local $fElapsed = Round(TimerDiff($g_hProcessTimer) / 1000, 2)
        $sSummary &= ", elapsed=" & $fElapsed & "s"
        If $g_iProcessCurrent > 0 And $fElapsed > 0 Then $sSummary &= ", rate=" & Round($g_iProcessCurrent / $fElapsed, 1) & "/s"
    EndIf
    Return $sSummary
EndFunc

Func _Process_Render($sMessage)
    Local $sPrefix = $g_sProcessName
    If $sPrefix = "" Then $sPrefix = "Process"

    Local $sText = $sPrefix & ": " & $sMessage
    If $g_iProcessTotal > 0 Then
        Local $iPct = Int(($g_iProcessCurrent / $g_iProcessTotal) * 100)
        If $iPct > 100 Then $iPct = 100
        $sText &= " [" & $g_iProcessCurrent & "/" & $g_iProcessTotal & " - " & $iPct & "%]"
    ElseIf $g_iProcessCurrent > 0 Then
        $sText &= " [" & $g_iProcessCurrent & "]"
    EndIf
    $sText &= " - sua " & $g_iProcessChanged & ", bo qua " & $g_iProcessSkipped & ", loi " & $g_iProcessErrors

    _UpdateProgress($sText)
    $g_iProcessLastRender = $g_iProcessCurrent
    If $g_iProcessTotal > 0 Then $g_iProcessLastRenderPct = Int(($g_iProcessCurrent / $g_iProcessTotal) * 100)
EndFunc

Func _Process_PumpGui()
    If Not IsHWnd($g_hGUI) Then Return
    Local $iMsg = GUIGetMsg()
    If $iMsg = $GUI_EVENT_CLOSE Then Exit
EndFunc

Func _Process_ShouldRenderStep($iCurrent)
    If $iCurrent <= 0 Then Return True
    If $g_iProcessTotal > 0 And $iCurrent >= $g_iProcessTotal Then Return True
    If $g_iProcessLastRender < 0 Then Return True
    If $g_iProcessRenderEvery <= 1 Then Return True
    If ($iCurrent - $g_iProcessLastRender) >= $g_iProcessRenderEvery Then Return True

    If $g_iProcessTotal > 0 Then
        Local $iPct = Int(($iCurrent / $g_iProcessTotal) * 100)
        If Mod($iPct, 5) = 0 And $iPct <> $g_iProcessLastRenderPct Then Return True
    EndIf

    Return False
EndFunc
