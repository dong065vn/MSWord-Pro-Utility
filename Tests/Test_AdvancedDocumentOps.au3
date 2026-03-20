#include "..\Config.au3"
#include "..\Core\WordConnection.au3"
#include "..\Shared\Helpers.au3"
#include "..\Modules\Advanced.au3"
#include "..\Modules\SmartFix.au3"
#include "..\Modules\TOC.au3"

Global $g_sTestLog = ""
Global Const $TEST_DIR = @ScriptDir & "\Artifacts\DocumentOps"

Func _Log($sMsg)
    ConsoleWrite($sMsg & @CRLF)
    $g_sTestLog &= $sMsg & @CRLF
EndFunc

Func _Assert($bCondition, $sPass, $sFail)
    If $bCondition Then
        _Log("PASS: " & $sPass)
        Return True
    EndIf
    _Log("FAIL: " & $sFail)
    Return False
EndFunc

Func _WriteLogFile()
    Local $hFile = FileOpen(@ScriptDir & "\Logs\Test_AdvancedDocumentOps.out.txt", 2 + 128)
    If $hFile <> -1 Then
        FileWrite($hFile, $g_sTestLog)
        FileClose($hFile)
    EndIf
EndFunc

Func _CreateUiStubs()
    $g_hGUI = GUICreate("Test UI", 400, 240, -1, -1, $WS_POPUP)
    $g_lblProgress = GUICtrlCreateLabel("", 10, 10, 360, 20)
    $g_lblStatus = GUICtrlCreateLabel("", 10, 35, 360, 20)
    $g_editPreview = GUICtrlCreateEdit("", 10, 60, 360, 150)
EndFunc

Func _PrepareWord()
    $g_oWord = ObjGet("", "Word.Application")
    If @error Or Not IsObj($g_oWord) Then $g_oWord = ObjCreate("Word.Application")
    If @error Or Not IsObj($g_oWord) Then Return False
    $g_oWord.Visible = True
    Return True
EndFunc

Func _ResetArtifacts()
    DirCreate($TEST_DIR)
    Local $aFiles = _FileListToArray($TEST_DIR, "*", 1)
    If @error Or Not IsArray($aFiles) Then Return
    For $i = 1 To $aFiles[0]
        FileDelete($TEST_DIR & "\" & $aFiles[$i])
    Next
EndFunc

Func _CreateDocFile($sPath, $sText)
    Local $oDoc = $g_oWord.Documents.Add()
    If Not IsObj($oDoc) Then Return False
    $oDoc.Range(0, 0).Text = $sText
    $oDoc.SaveAs2($sPath, 16)
    $oDoc.Close(0)
    Return FileExists($sPath)
EndFunc

Func _TestMergeSplitCompareProtect()
    Local $bOk = True
    Local $sMergeSource = $TEST_DIR & "\merge_source.docx"
    Local $sCompareSource = $TEST_DIR & "\compare_source.docx"
    Local $sSplitPath = $TEST_DIR & "\split_result.docx"

    _CreateDocFile($sMergeSource, "MERGE_BLOCK" & @CRLF)
    _CreateDocFile($sCompareSource, "BASE" & @CRLF & "CHANGED" & @CRLF)

    $g_oDoc = $g_oWord.Documents.Add()
    $g_oDoc.Range(0, 0).Text = "START" & @CRLF & "BODY" & @CRLF

    Local $bMerge = _MergeDocumentFromPath($sMergeSource, "1")
    $bOk = _Assert($bMerge And StringInStr($g_oDoc.Content.Text, "MERGE_BLOCK") > 0, "Merge chen duoc noi dung file", "Merge khong chen duoc noi dung") And $bOk

    Local $oRange = $g_oDoc.Range(0, 5)
    Local $bSplit = _SplitRangeToPath($oRange, $sSplitPath, True)
    $bOk = _Assert($bSplit And FileExists($sSplitPath), "Split tao duoc file moi", "Split khong tao duoc file moi") And $bOk
    $bOk = _Assert(StringInStr($g_oDoc.Content.Text, "START") = 0, "Split xoa duoc phan da tach khoi file goc", "Split khong xoa phan da tach khoi file goc") And $bOk

    $g_oDoc.Range(0, 0).Text = "BASE" & @CRLF & "CURRENT" & @CRLF
    Local $oCompareDoc = _CompareDocumentsByPath($sCompareSource)
    $bOk = _Assert(IsObj($oCompareDoc), "Compare tao duoc document so sanh", "Compare khong tao duoc document so sanh") And $bOk
    If IsObj($oCompareDoc) Then
        $bOk = _Assert($oCompareDoc.Revisions.Count > 0, "Compare phat hien duoc thay doi", "Compare khong co revision nao") And $bOk
        $oCompareDoc.Close(0)
        If IsObj($g_oDoc) Then $g_oDoc.Activate()
    EndIf

    Local $bProtected = _SetDocumentProtection(3, "")
    $bOk = _Assert($bProtected, "Protect bat duoc che do bao ve", "Protect that bai") And $bOk
    Local $bUnprotected = _RemoveDocumentProtection("")
    $bOk = _Assert($bUnprotected, "Unprotect bo duoc bao ve", "Unprotect that bai") And $bOk

    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _TestBatchSmartFix()
    Local $bOk = True
    Local $sBatch1 = $TEST_DIR & "\batch1.docx"
    Local $sBatch2 = $TEST_DIR & "\batch2.docx"
    _CreateDocFile($sBatch1, "A" & ChrW(8212) & "B" & @CRLF & "Quote " & ChrW(8220) & "x" & ChrW(8221))
    _CreateDocFile($sBatch2, "vi-" & @CRLF & "du  test")

    Local $oDocOld = $g_oDoc
    Local $aBatch[2] = [$sBatch1, $sBatch2]
    For $i = 0 To UBound($aBatch) - 1
        Local $oDoc = $g_oWord.Documents.Open($aBatch[$i])
        If Not IsObj($oDoc) Then Return False
        $g_oDoc = $oDoc
        _BatchSmartFix()
        $oDoc.Save()
        $oDoc.Close(0)
    Next
    $g_oDoc = $oDocOld

    Local $oCheck1 = $g_oWord.Documents.Open($sBatch1, False, True)
    Local $oCheck2 = $g_oWord.Documents.Open($sBatch2, False, True)
    If Not IsObj($oCheck1) Or Not IsObj($oCheck2) Then Return False

    $bOk = _Assert(StringInStr($oCheck1.Content.Text, ChrW(8212)) = 0 And StringInStr($oCheck1.Content.Text, ChrW(8220)) = 0, "BatchSmartFix sua duoc dash va smart quotes", "BatchSmartFix chua sua het dash/quotes") And $bOk
    $bOk = _Assert(StringInStr($oCheck2.Content.Text, "vi-" & @CR) = 0 And StringInStr($oCheck2.Content.Text, "  ") = 0, "BatchSmartFix sua duoc hyphenation va double spaces", "BatchSmartFix chua sua het hyphenation/spaces") And $bOk

    $oCheck1.Close(0)
    $oCheck2.Close(0)
    Return $bOk
EndFunc

Func _Run()
    DirCreate(@ScriptDir & "\Logs")
    _ResetArtifacts()
    _CreateUiStubs()

    _Log("=== TEST ADVANCED DOCUMENT OPS ===")
    If Not _PrepareWord() Then
        _Log("FAIL: Khong the khoi dong Word")
        _WriteLogFile()
        Exit 1
    EndIf

    Local $bOps = _TestMergeSplitCompareProtect()
    Local $bBatch = _TestBatchSmartFix()

    If $bOps And $bBatch Then
        _Log("=== ALL PASS ===")
        _WriteLogFile()
        Exit 0
    EndIf

    _Log("=== TEST FAILED ===")
    _WriteLogFile()
    Exit 2
EndFunc

_Run()
