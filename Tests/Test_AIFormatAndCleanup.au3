#include "..\Config.au3"
#include "..\Core\WordConnection.au3"
#include "..\Shared\Helpers.au3"
#include "..\Modules\Advanced.au3"
#include "..\Modules\AIFormat.au3"
#include "..\Modules\TOC.au3"

Global $g_sTestLog = ""

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
    Local $hFile = FileOpen(@ScriptDir & "\Logs\Test_AIFormatAndCleanup.out.txt", 2 + 128)
    If $hFile <> -1 Then
        FileWrite($hFile, $g_sTestLog)
        FileClose($hFile)
    EndIf
EndFunc

Func _CreateUiStubs()
    $g_hGUI = GUICreate("Test UI", 500, 260, -1, -1, $WS_POPUP)
    $g_lblProgress = GUICtrlCreateLabel("", 10, 10, 420, 20)
    $g_lblStatus = GUICtrlCreateLabel("", 10, 35, 420, 20)
    $g_editPreview = GUICtrlCreateEdit("", 10, 60, 470, 150)
    $g_chkAIScopeSelection = GUICtrlCreateCheckbox("Selection", 10, 220, 90, 20)
    $g_chkAIScopeAll = GUICtrlCreateCheckbox("All", 110, 220, 90, 20)
    GUICtrlSetState($g_chkAIScopeAll, $GUI_CHECKED)
EndFunc

Func _PrepareWord()
    $g_oWord = ObjGet("", "Word.Application")
    If @error Or Not IsObj($g_oWord) Then $g_oWord = ObjCreate("Word.Application")
    If @error Or Not IsObj($g_oWord) Then Return False
    $g_oWord.Visible = True
    Return True
EndFunc

Func _GetParaStyleName($oPara)
    If Not IsObj($oPara) Then Return ""
    Local $oStyle = $oPara.Range.Style
    If IsObj($oStyle) Then Return $oStyle.NameLocal
    Return ""
EndFunc

Func _TestAIFormatCore()
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    Local $sText = "# Heading One" & @CRLF & _
        "Paragraph with **bold** and `code`." & @CRLF & _
        "Spacing   test" & @CRLF & _
        "Quote " & ChrW(8220) & "x" & ChrW(8221) & " and dash " & ChrW(8212) & @CRLF
    $g_oDoc.Range(0, 0).Text = $sText

    _AI_ConvertHeadings()
    _AI_ConvertBold()
    _AI_ConvertInlineCode()
    _AI_FixExtraSpaces()
    _AI_FixEncoding()

    Local $bOk = True
    Local $oPara1 = $g_oDoc.Paragraphs.Item(1)
    $bOk = _Assert(_GetParaStyleName($oPara1) = "Heading 1", "AI convert headings gan dung style", "AI convert headings khong gan Heading 1") And $bOk
    $bOk = _Assert(StringInStr($oPara1.Range.Text, "#") = 0, "AI convert headings xoa marker #", "AI convert headings con lai marker #") And $bOk

    Local $sBody = $g_oDoc.Paragraphs.Item(2).Range.Text
    $bOk = _Assert(StringInStr($sBody, "**") = 0, "AI convert bold xoa marker **", "AI convert bold con marker **") And $bOk
    $bOk = _Assert(StringInStr($sBody, "`") = 0, "AI convert inline code xoa marker backtick", "AI convert inline code con marker backtick") And $bOk

    Local $oBoldRange = $g_oDoc.Range(27, 31)
    $bOk = _Assert($oBoldRange.Font.Bold = True, "AI convert bold ap dung Bold", "AI convert bold chua ap dung Bold") And $bOk

    Local $sAll = $g_oDoc.Content.Text
    $bOk = _Assert(StringInStr($sAll, "   ") = 0, "AI fix extra spaces don duoc nhieu space", "AI fix extra spaces chua don duoc") And $bOk
    $bOk = _Assert(StringInStr($sAll, ChrW(8220)) = 0 And StringInStr($sAll, ChrW(8221)) = 0, "AI fix encoding doi duoc smart quotes", "AI fix encoding con smart quotes") And $bOk
    $bOk = _Assert(StringInStr($sAll, ChrW(8212)) = 0, "AI fix encoding doi duoc em dash", "AI fix encoding con em dash") And $bOk

    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _TestCleanDocumentCore()
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    $g_oDoc.Range(0, 0).Text = "Visit OpenAI." & @CRLF & "Second line." & @CRLF
    $g_oDoc.Hyperlinks.Add($g_oDoc.Range(7, 13), "https://openai.com")
    $g_oDoc.Comments.Add($g_oDoc.Range(0, 5), "Test comment")

    $g_oDoc.TrackRevisions = True
    $g_oDoc.Range(0, 0).InsertBefore("Changed ")
    $g_oDoc.TrackRevisions = False

    Local $bClean = _CleanDocumentCore()
    Local $bOk = True
    $bOk = _Assert($bClean, "CleanDocumentCore chay thanh cong", "CleanDocumentCore that bai") And $bOk
    $bOk = _Assert($g_oDoc.Comments.Count = 0, "CleanDocumentCore xoa comments", "CleanDocumentCore con comments") And $bOk
    $bOk = _Assert($g_oDoc.Hyperlinks.Count = 0, "CleanDocumentCore xoa hyperlinks", "CleanDocumentCore con hyperlinks") And $bOk
    $bOk = _Assert($g_oDoc.Revisions.Count = 0, "CleanDocumentCore accept revisions", "CleanDocumentCore con revisions") And $bOk

    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _Run()
    DirCreate(@ScriptDir & "\Logs")
    _CreateUiStubs()
    _Log("=== TEST AIFORMAT + CLEANUP ===")

    If Not _PrepareWord() Then
        _Log("FAIL: Khong the khoi dong Word")
        _WriteLogFile()
        Exit 1
    EndIf

    Local $bAI = _TestAIFormatCore()
    Local $bClean = _TestCleanDocumentCore()

    If $bAI And $bClean Then
        _Log("=== ALL PASS ===")
        _WriteLogFile()
        Exit 0
    EndIf

    _Log("=== TEST FAILED ===")
    _WriteLogFile()
    Exit 2
EndFunc

_Run()
