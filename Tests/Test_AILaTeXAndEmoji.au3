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
    Local $hFile = FileOpen(@ScriptDir & "\Logs\Test_AILaTeXAndEmoji.out.txt", 2 + 128)
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

Func _FindTextRange($sNeedle)
    Local $oRange = $g_oDoc.Content.Duplicate
    With $oRange.Find
        .ClearFormatting()
        .Text = $sNeedle
        .MatchWildcards = False
    EndWith
    If $oRange.Find.Execute() Then Return $oRange
    Return 0
EndFunc

Func _FindParagraphContaining($sNeedle)
    For $i = 1 To $g_oDoc.Paragraphs.Count
        Local $oPara = $g_oDoc.Paragraphs.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        If StringInStr($oPara.Range.Text, $sNeedle) > 0 Then Return $oPara
    Next
    Return 0
EndFunc

Func _TestLaTeXAndEmoji()
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    Local $sText = "Inline math $x+y$ stays inside sentence." & @CRLF & _
        "$$E=mc^2$$" & @CRLF & _
        "Emoji line 😀 keeps text and removes icons ❤"
    $g_oDoc.Range(0, 0).Text = $sText

    _AI_ConvertLaTeX()
    _AI_NormalizeAllMath(False)
    _AI_RemoveEmoji()

    Local $bOk = True
    Local $sAll = $g_oDoc.Content.Text
    $bOk = _Assert(StringInStr($sAll, "$") = 0, "AI convert LaTeX xoa marker dollar", "AI convert LaTeX con marker dollar") And $bOk
    $bOk = _Assert(StringInStr($sAll, ChrW(128512)) = 0 And StringInStr($sAll, ChrW(10084)) = 0, "AI remove emoji xoa duoc emoji", "AI remove emoji con emoji") And $bOk
    $bOk = _Assert(StringInStr($sAll, "Emoji line") > 0 And StringInStr($sAll, "keeps text") > 0, "AI remove emoji giu nguyen text thuong", "AI remove emoji lam mat text thuong") And $bOk

    Local $oMathInline = _FindTextRange("x+y")
    $bOk = _Assert(IsObj($oMathInline), "Tim duoc inline formula de verify", "Khong tim thay inline formula sau convert") And $bOk
    If IsObj($oMathInline) Then
        $bOk = _Assert($oMathInline.Font.Name = "Cambria Math" And $oMathInline.Font.Italic = True And $oMathInline.Font.Size = 12, "AI convert LaTeX format dung inline formula", "AI convert LaTeX chua format dung inline formula") And $bOk
    EndIf

    Local $oSentencePrefix = _FindTextRange("Inline math ")
    $bOk = _Assert(IsObj($oSentencePrefix), "Tim duoc prefix text de verify", "Khong tim thay prefix text sau convert") And $bOk
    If IsObj($oSentencePrefix) Then
        $bOk = _Assert($oSentencePrefix.Font.Name <> "Cambria Math", "AI convert LaTeX khong doi font ca cau", "AI convert LaTeX dang doi font ca paragraph") And $bOk
    EndIf

    Local $oBlockFormula = _FindParagraphContaining("E=mc^2")
    If Not IsObj($oBlockFormula) Then
        _Log("DEBUG: CONTENT=" & StringReplace(StringReplace($sAll, @CR, "{CR}"), @LF, "{LF}"))
        For $i = 1 To $g_oDoc.Paragraphs.Count
            Local $oDbgPara = $g_oDoc.Paragraphs.Item($i)
            If IsObj($oDbgPara) Then
                _Log("DEBUG PARA[" & $i & "]=" & StringReplace(StringReplace($oDbgPara.Range.Text, @CR, "{CR}"), @LF, "{LF}"))
            EndIf
        Next
    EndIf
    $bOk = _Assert(IsObj($oBlockFormula), "Tim duoc block formula de verify", "Khong tim thay block formula") And $bOk
    If IsObj($oBlockFormula) Then
        $bOk = _Assert($oBlockFormula.Range.Font.Name = "Cambria Math" And $oBlockFormula.Range.Font.Size = 12 And $oBlockFormula.Format.Alignment = $WD_ALIGN_CENTER, "AI convert LaTeX format dung block formula", "AI convert LaTeX chua format dung block formula") And $bOk
    EndIf

    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _Run()
    DirCreate(@ScriptDir & "\Logs")
    _CreateUiStubs()
    _Log("=== TEST AILATEX + EMOJI ===")

    If Not _PrepareWord() Then
        _Log("FAIL: Khong the khoi dong Word")
        _WriteLogFile()
        Exit 1
    EndIf

    Local $bOk = _TestLaTeXAndEmoji()
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
