; ============================================
; TEST_SMART_FIX_PRO_PIPELINE.AU3
; ============================================

#include "..\Config.au3"
#include "..\Core\WordConnection.au3"
#include "..\Shared\Helpers.au3"
#include "..\Shared\WordPerf.au3"
#include "..\Shared\ProcessTracker.au3"
#include "..\Shared\WordDom.au3"
#include "..\Shared\WordOps.au3"
#include "..\Modules\PDFFix.au3"
#include "..\Modules\SmartFix.au3"

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

Func _MakeOptions($bDryRun, $bFakeNumbering = False)
    Local $aOptions[8]
    $aOptions[$SMARTPRO_OPT_DRYRUN] = $bDryRun
    $aOptions[$SMARTPRO_OPT_SCOPE_SELECTION] = True
    $aOptions[$SMARTPRO_OPT_SKIP_TABLES] = True
    $aOptions[$SMARTPRO_OPT_SKIP_EQUATIONS] = True
    $aOptions[$SMARTPRO_OPT_SKIP_LINKS] = True
    $aOptions[$SMARTPRO_OPT_SKIP_FIELDS] = True
    $aOptions[$SMARTPRO_OPT_FIX_LAYOUT] = False
    $aOptions[$SMARTPRO_OPT_FAKE_NUMBERING] = $bFakeNumbering
    Return $aOptions
EndFunc

Func _SetupGui()
    $g_hGUI = GUICreate("SmartFixProPipeline Test", 600, 240)
    $g_lblStatus = GUICtrlCreateLabel("", 10, 10, 560, 20)
    $g_lblProgress = GUICtrlCreateLabel("", 10, 35, 560, 20)
    $g_editPreview = GUICtrlCreateEdit("", 10, 65, 560, 150)
EndFunc

Func _Run()
    _SetupGui()

    Local $aOptions = _MakeOptions(True, True)
    Local $sFixed = _SmartPro_FixText("1.  Demo" & ChrW(160) & "  A" & ChrW(8212) & "B", $aOptions)
    _Assert($sFixed = "Demo A-B", "Text fixes normalize spacing, dash, NBSP, fake numbering")
    _Assert(_StripUnnecessaryLeadingNumbering("1.2 Heading" & @CR) = "1.2 Heading" & @CR, "Heading numbering is preserved")

    $g_oWord = ObjCreate("Word.Application")
    If Not IsObj($g_oWord) Then
        ConsoleWrite("FAIL: Cannot create Word.Application" & @CRLF)
        Exit 1
    EndIf

    $g_oWord.Visible = False
    $g_oDoc = $g_oWord.Documents.Add()
    $g_oDoc.Range(0, 0).Text = "A  B" & Chr(11) & "C" & ChrW(8211) & "D" & @CRLF & _
        "1. Fake item" & @CRLF & _
        "1.2 Real heading" & @CRLF

    Local $sBefore = $g_oDoc.Content.Text
    Local $aPreviewOptions = _MakeOptions(True, True)
    _SmartPro_Run($aPreviewOptions, False)
    _Assert($g_oDoc.Content.Text = $sBefore, "Preview Pro does not modify document")
    _Assert(StringInStr(GUICtrlRead($g_editPreview), "Total candidates:") > 0, "Preview Pro writes candidate report")

    Local $aApplyOptions = _MakeOptions(False, True)
    $g_bDomSkipTables = True
    $g_bDomSkipEquations = True
    $g_bDomSkipLinks = True
    $g_bDomSkipFields = True
    Local $iChanged = _SmartPro_ApplyTextFixesToParagraphs($g_oDoc.Content, $aApplyOptions)
    Local $sAfter = $g_oDoc.Content.Text
    _Assert($iChanged >= 1, "Apply text fixes changes dirty paragraphs")
    _Assert(StringInStr($sAfter, "A B C-D") > 0, "Apply text fixes normalizes PDF artifacts")
    _Assert(StringInStr($sAfter, "1.2 Real heading") > 0, "Apply text fixes preserves multi-level heading numbering")

    $g_oDoc.Close(False)
    $g_oWord.Quit()
    $g_oDoc = 0
    $g_oWord = 0
    GUIDelete($g_hGUI)

    ConsoleWrite("Passed: " & $g_iPassed & @CRLF)
    ConsoleWrite("Failed: " & $g_iFailed & @CRLF)
    If $g_iFailed > 0 Then Exit 1
    Exit 0
EndFunc

_Run()
