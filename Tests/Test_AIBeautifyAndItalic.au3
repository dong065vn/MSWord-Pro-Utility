#include "..\Config.au3"
#include "..\Core\WordConnection.au3"
#include "..\Shared\Helpers.au3"
#include "..\Modules\AIFormat.au3"

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
    Local $hFile = FileOpen(@ScriptDir & "\Logs\Test_AIBeautifyAndItalic.out.txt", 2 + 128)
    If $hFile <> -1 Then
        FileWrite($hFile, $g_sTestLog)
        FileClose($hFile)
    EndIf
EndFunc

Func _CreateUiStubs()
    $g_hGUI = GUICreate("Test UI", 520, 260, -1, -1, $WS_POPUP)
    $g_lblProgress = GUICtrlCreateLabel("", 10, 10, 460, 20)
    $g_lblStatus = GUICtrlCreateLabel("", 10, 35, 460, 20)
    $g_editPreview = GUICtrlCreateEdit("", 10, 60, 490, 150)
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

Func _CreateParagraph($sText)
    Local $oRange = $g_oDoc.Range($g_oDoc.Content.End - 1, $g_oDoc.Content.End - 1)
    $oRange.InsertAfter($sText & @CRLF)
EndFunc

Func _GetParagraphIndexContaining($sNeedle)
    For $i = 1 To $g_oDoc.Paragraphs.Count
        Local $oPara = $g_oDoc.Paragraphs.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        If StringInStr($oPara.Range.Text, $sNeedle) > 0 Then Return $i
    Next
    Return 0
EndFunc

Func _RangeContainsItalicText($sNeedle)
    Local $oFindRange = $g_oDoc.Content.Duplicate
    With $oFindRange.Find
        .ClearFormatting()
        .Text = $sNeedle
        .MatchWildcards = False
    EndWith
    If $oFindRange.Find.Execute() Then
        Return ($oFindRange.Font.Italic = True)
    EndIf
    Return False
EndFunc

Func _TestBeautifyAndItalic()
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    $g_oDoc.Range(0, 0).Text = ""
    _CreateParagraph("*italic star* in paragraph")
    _CreateParagraph("This keeps **bold** intact")
    _CreateParagraph("_italic underscore_ too")
    _CreateParagraph("-")
    _CreateParagraph("")
    _CreateParagraph("")
    _CreateParagraph("---")
    _CreateParagraph("# Heading Pretty")
    _CreateParagraph("- list item one")
    _CreateParagraph("- list item two")
    _CreateParagraph("Normal paragraph after list")

    _AI_ConvertHeadings()
    _AI_ConvertBold()
    _AI_ConvertBullets()
    _AI_ConvertItalic()
    _AI_CleanupVisual()
    _AI_BeautifyDocument()

    Local $bOk = True
    Local $sAll = $g_oDoc.Content.Text

    $bOk = _Assert(StringInStr($sAll, "*italic star*") = 0, "AI convert italic xoa marker *", "AI convert italic con marker *") And $bOk
    $bOk = _Assert(StringInStr($sAll, "_italic underscore_") = 0, "AI convert italic xoa marker _", "AI convert italic con marker _") And $bOk
    $bOk = _Assert(_RangeContainsItalicText("italic star"), "AI convert italic ap dung italic cho *text*", "AI convert italic chua ap dung italic cho *text*") And $bOk
    $bOk = _Assert(_RangeContainsItalicText("italic underscore"), "AI convert italic ap dung italic cho _text_", "AI convert italic chua ap dung italic cho _text_") And $bOk
    $bOk = _Assert(StringInStr($sAll, "**bold**") = 0, "AI convert italic khong lam hong xu ly bold", "AI convert italic lam hong xu ly bold") And $bOk

    $bOk = _Assert(StringInStr($sAll, @CRLF & "-" & @CRLF) = 0, "AI cleanup visual xoa orphan bullet", "AI cleanup visual con orphan bullet") And $bOk
    $bOk = _Assert(StringInStr($sAll, "---") = 0, "AI cleanup visual xoa markdown separator", "AI cleanup visual con markdown separator") And $bOk
    $bOk = _Assert(StringInStr($sAll, @CRLF & @CRLF & @CRLF) = 0, "AI cleanup visual don double empty lines", "AI cleanup visual con nhieu dong trong lien tiep") And $bOk

    Local $iHeadingIndex = _GetParagraphIndexContaining("Heading Pretty")
    $bOk = _Assert($iHeadingIndex > 0, "Tim duoc heading de verify beautify", "Khong tim thay heading sau khi beautify") And $bOk
    If $iHeadingIndex > 0 Then
        Local $oHeadingPara = $g_oDoc.Paragraphs.Item($iHeadingIndex)
        $bOk = _Assert($oHeadingPara.Format.FirstLineIndent = 0 And $oHeadingPara.Format.LeftIndent = 0, "AI beautify bo thut dau dong cho heading", "AI beautify chua bo thut dau dong cho heading") And $bOk
    EndIf

    Local $iListFormatted = 0
    For $i = 1 To $g_oDoc.Paragraphs.Count
        Local $oPara = $g_oDoc.Paragraphs.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        If StringInStr($oPara.Range.Text, "list item") > 0 And $oPara.Range.ListFormat.ListType <> 0 Then
            If $oPara.Format.LeftIndent > 0 Then $iListFormatted += 1
        EndIf
    Next
    $bOk = _Assert($iListFormatted >= 2, "AI beautify ap dung indent cho list paragraphs", "AI beautify chua ap dung indent cho list paragraphs") And $bOk

    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _Run()
    DirCreate(@ScriptDir & "\Logs")
    _CreateUiStubs()
    _Log("=== TEST AIFORMAT BEAUTIFY + ITALIC ===")

    If Not _PrepareWord() Then
        _Log("FAIL: Khong the khoi dong Word")
        _WriteLogFile()
        Exit 1
    EndIf

    Local $bOk = _TestBeautifyAndItalic()
    If $bOk Then
        _Log("=== ALL PASS ===")
        _WriteLogFile()
        Exit 0
    EndIf

    _Log("=== TEST FAILED ===")
    _WriteLogFile()
    Exit 2
EndFunc

_Run()
