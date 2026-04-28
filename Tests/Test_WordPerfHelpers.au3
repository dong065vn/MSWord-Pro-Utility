; ============================================
; TEST_WORD_PERF_HELPERS.AU3
; ============================================

#include "..\Config.au3"
#include "..\Shared\WordPerf.au3"

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
    $g_oWord = 0
    Local $aCtx = _Perf_BeginWordBatch("unit")
    _Assert(IsArray($aCtx), "Batch context is an array without Word")
    _Assert($aCtx[$PERF_CTX_NAME] = "unit", "Batch context stores name")
    _Perf_EndWordBatch($aCtx)

    $g_iProcessTotal = 100
    $g_iProcessLastRenderPct = -1
    _Assert(_Perf_ShouldRenderStep(0, 100, 25), "Throttle renders initial step")
    _Assert(Not _Perf_ShouldRenderStep(1, 100, 25), "Throttle skips small step")
    _Assert(_Perf_ShouldRenderStep(25, 100, 25), "Throttle renders interval")
    _Assert(_Perf_ShouldRenderStep(100, 100, 25), "Throttle renders final step")

    Local $aValues[1]
    Local $iCapacity = 1
    $aValues[0] = "keep"
    _Perf_NormalizeArrayCapacity($aValues, $iCapacity, 5, 2)
    _Assert($iCapacity >= 5, "Array capacity grows")
    _Assert($aValues[0] = "keep", "Array capacity keeps existing data")

    _Assert(_Perf_CollectionCount(0) = 0, "Collection count handles non-object")

    ConsoleWrite("Passed: " & $g_iPassed & @CRLF)
    ConsoleWrite("Failed: " & $g_iFailed & @CRLF)
    If $g_iFailed > 0 Then Exit 1
    Exit 0
EndFunc

_Run()
