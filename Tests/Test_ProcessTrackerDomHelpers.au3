; ============================================
; TEST_PROCESS_TRACKER_DOM_HELPERS.AU3
; ============================================

#include "..\Config.au3"
#include "..\Core\WordConnection.au3"
#include "..\Shared\Helpers.au3"
#include "..\Shared\WordPerf.au3"
#include "..\Shared\ProcessTracker.au3"
#include "..\Shared\WordDom.au3"

Global $g_iPassed = 0
Global $g_iFailed = 0

Func _Assert($bCondition, $sName)
    If $bCondition Then
        ConsoleWrite("PASS: " & $sName & @CRLF)
        $g_iPassed += 1
    Else
        ConsoleWrite("FAIL: " & $sName & @CRLF)
        $g_iFailed += 1
    EndIf
EndFunc

Func _Run()
    $g_hGUI = GUICreate("ProcessTrackerDomHelpers Test", 520, 160)
    $g_lblStatus = GUICtrlCreateLabel("", 10, 10, 500, 20)
    $g_lblProgress = GUICtrlCreateLabel("", 10, 40, 500, 20)
    $g_editPreview = GUICtrlCreateEdit("", 10, 70, 500, 70)

    _Process_Start("Unit Process", 4, "test")
    _Assert($g_sProcessName = "Unit Process", "Process name is set")
    _Assert($g_iProcessTotal = 4, "Process total is set")

    _Process_Step("Step 2", 2, 1, 0, 0)
    _Assert($g_iProcessCurrent = 2, "Process current updates")
    _Assert($g_iProcessChanged = 1, "Process changed updates")
    _Assert(StringInStr(GUICtrlRead($g_lblProgress), "2/4") > 0, "Progress label contains count")

    _Process_AddSkipped()
    _Process_AddError()
    _Assert($g_iProcessSkipped = 1, "Process skipped increments")
    _Assert($g_iProcessErrors = 1, "Process errors increments")

    Local $sSummary = _Process_BuildSummary()
    _Assert(StringInStr($sSummary, "processed=2/4") > 0, "Summary has processed count")
    _Assert(StringInStr($sSummary, "changed=1") > 0, "Summary has changed count")
    _Assert(StringInStr($sSummary, "skipped=1") > 0, "Summary has skipped count")
    _Assert(StringInStr($sSummary, "errors=1") > 0, "Summary has error count")

    _Process_Done("done summary")
    _Assert(StringInStr(GUICtrlRead($g_editPreview), "[DONE] Unit Process") > 0, "Done is appended to preview")

    $g_oWord = 0
    $g_oDoc = 0
    _Assert(_Dom_HasDocument() = False, "DOM reports missing document")
    _Assert(IsObj(_Dom_GetScopeRange("document")) = False, "DOM scope returns non-object without document")
    _Assert(_Dom_IsRangeUsable(0) = False, "DOM rejects invalid range")
    _Assert(_Dom_GetCollectionCount(0) = 0, "DOM collection count handles invalid object")

    GUIDelete($g_hGUI)

    ConsoleWrite("Passed: " & $g_iPassed & @CRLF)
    ConsoleWrite("Failed: " & $g_iFailed & @CRLF)
    If $g_iFailed > 0 Then Exit 1
    Exit 0
EndFunc

_Run()
