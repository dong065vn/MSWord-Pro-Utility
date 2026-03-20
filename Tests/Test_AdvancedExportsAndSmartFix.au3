#include "..\Config.au3"
#include "..\Core\WordConnection.au3"
#include "..\Shared\Helpers.au3"
#include "..\Modules\Advanced.au3"
#include "..\Modules\SmartFix.au3"
#include "..\Modules\TOC.au3"

Global $g_sTestLog = ""
Global Const $TEST_OUTPUT_DIR = @ScriptDir & "\Artifacts\AdvancedSmartFix"

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
    Local $sLogPath = @ScriptDir & "\Logs\Test_AdvancedExportsAndSmartFix.out.txt"
    Local $hFile = FileOpen($sLogPath, 2 + 128)
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

Func _CleanupOutputDir()
    DirCreate($TEST_OUTPUT_DIR)
    Local $aFiles = _FileListToArray($TEST_OUTPUT_DIR, "*", 1)
    If @error Or Not IsArray($aFiles) Then Return

    For $i = 1 To $aFiles[0]
        FileDelete($TEST_OUTPUT_DIR & "\" & $aFiles[$i])
    Next
EndFunc

Func _TestSmartFixFunctions()
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    Local $sText = "vi-" & @CRLF & "du  co  hai  khoang trang" & @CRLF & _
        "A" & ChrW(8212) & "B" & @CRLF & "Quote " & ChrW(8220) & "x" & ChrW(8221)
    $g_oDoc.Range(0, 0).Text = $sText

    _FixHyphenation()
    _FixDashes()
    _BatchSmartFix()

    Local $sResult = $g_oDoc.Content.Text
    _Log("RESULT TEXT: " & StringReplace(StringReplace($sResult, @CR, "{CR}"), @LF, "{LF}"))
    Local $bOk = True
    $bOk = _Assert(StringInStr($sResult, "vi-" & @CR) = 0, "Hyphenation da duoc xoa", "Hyphenation van con") And $bOk
    $bOk = _Assert(StringInStr($sResult, ChrW(8212)) = 0, "Em dash da duoc doi", "Em dash van con") And $bOk
    $bOk = _Assert(StringInStr($sResult, "  ") = 0, "Double spaces da duoc don", "Double spaces van con") And $bOk
    $bOk = _Assert(StringInStr($sResult, ChrW(8220)) = 0 And StringInStr($sResult, ChrW(8221)) = 0, "Smart quotes da duoc doi", "Smart quotes van con") And $bOk

    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _TestAdvancedExports()
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    $g_oDoc.Range(0, 0).Text = "Heading export" & @CRLF & "Noi dung export test." & @CRLF
    $g_oDoc.Paragraphs.Item(1).Range.Style = "Heading 1"

    Local $iDocCountBefore = $g_oWord.Documents.Count
    Local $sOriginalText = $g_oDoc.Content.Text
    Local $bOk = True

    Local $aExports[4][2] = [ _
        ["export_test.pdf", $WD_EXPORT_PDF], _
        ["export_test.html", $WD_FORMAT_HTML], _
        ["export_test.txt", $WD_FORMAT_TEXT], _
        ["export_test.rtf", $WD_FORMAT_RTF] _
    ]

    For $i = 0 To UBound($aExports) - 1
        Local $sPath = $TEST_OUTPUT_DIR & "\" & $aExports[$i][0]
        If FileExists($sPath) Then FileDelete($sPath)

        Local $bExport = _ExportCurrentDocumentToPath($sPath, $aExports[$i][1])
        $bOk = _Assert($bExport And FileExists($sPath), "Export tao duoc " & $aExports[$i][0], "Export that bai: " & $aExports[$i][0]) And $bOk
    Next

    $bOk = _Assert($g_oWord.Documents.Count = $iDocCountBefore, "Export khong de lai document tam", "Export de lai document tam dang mo") And $bOk
    $bOk = _Assert($g_oDoc.Content.Text = $sOriginalText, "Export khong doi noi dung document goc", "Export lam thay doi noi dung document goc") And $bOk

    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _Run()
    DirCreate(@ScriptDir & "\Logs")
    _CleanupOutputDir()
    _CreateUiStubs()

    _Log("=== TEST ADVANCED EXPORTS + SMART FIX ===")

    If Not _PrepareWord() Then
        _Log("FAIL: Khong the khoi dong Word")
        _WriteLogFile()
        Exit 1
    EndIf

    Local $bSmartFix = _TestSmartFixFunctions()
    Local $bExports = _TestAdvancedExports()

    If $bSmartFix And $bExports Then
        _Log("=== ALL PASS ===")
        _WriteLogFile()
        Exit 0
    EndIf

    _Log("=== TEST FAILED ===")
    _WriteLogFile()
    Exit 2
EndFunc

_Run()
