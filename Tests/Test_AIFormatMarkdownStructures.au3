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
    Local $hFile = FileOpen(@ScriptDir & "\Logs\Test_AIFormatMarkdownStructures.out.txt", 2 + 128)
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

Func _TestMarkdownStructures()
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    Local $sText = "```" & @CRLF & _
        "print('hi')" & @CRLF & _
        "```" & @CRLF & _
        "- bullet one" & @CRLF & _
        "• bullet two" & @CRLF & _
        "1. first item" & @CRLF & _
        "[OpenAI](https://openai.com)" & @CRLF & _
        "| Col1 | Col2 |" & @CRLF & _
        "| --- | --- |" & @CRLF & _
        "| A | B |" & @CRLF
    $g_oDoc.Range(0, 0).Text = $sText

    _Log("STEP: before code blocks")
    _AI_ConvertCodeBlocks()
    _Log("STEP: after code blocks")
    _AI_ConvertBullets()
    _Log("STEP: after bullets")
    _AI_ConvertNumberedLists()
    _Log("STEP: after numbered")
    _AI_ConvertLinks()
    _Log("STEP: after links")
    _AI_ConvertTables()
    _Log("STEP: after tables")

    Local $bOk = True
    Local $sAll = $g_oDoc.Content.Text

    $bOk = _Assert(StringInStr($sAll, "```") = 0, "AI convert code blocks xoa marker ```", "AI convert code blocks con marker ```") And $bOk
    $bOk = _Assert(StringInStr($sAll, "[OpenAI]") = 0 And StringInStr($sAll, "https://openai.com") = 0, "AI convert links xoa markdown link", "AI convert links con markdown link") And $bOk

    Local $bHasConsolas = False
    Local $oParas = $g_oDoc.Paragraphs
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        If StringInStr($oPara.Range.Text, "print('hi')") > 0 Then
            $bHasConsolas = ($oPara.Range.Font.Name = "Consolas")
            ExitLoop
        EndIf
    Next
    $bOk = _Assert($bHasConsolas, "AI convert code blocks ap dung font code", "AI convert code blocks chua ap dung font code") And $bOk

    Local $iBulletParas = 0, $iNumberParas = 0
    Local $bMergedBulletAndNumber = False
    For $i = 1 To $g_oDoc.Paragraphs.Count
        Local $oPara2 = $g_oDoc.Paragraphs.Item($i)
        If Not IsObj($oPara2) Then ContinueLoop
        Local $iListType = $oPara2.Range.ListFormat.ListType
        _Log("PARA[" & $i & "] LISTTYPE=" & $iListType & " TEXT=" & StringReplace(StringReplace($oPara2.Range.Text, @CR, "{CR}"), @LF, "{LF}"))
        If StringInStr($oPara2.Range.Text, "bullet") > 0 And StringInStr($oPara2.Range.Text, "first item") > 0 Then $bMergedBulletAndNumber = True
        If $iListType <> 0 Then
            If StringInStr($oPara2.Range.Text, "bullet") > 0 Then $iBulletParas += 1
            If StringInStr($oPara2.Range.Text, "first item") > 0 Then $iNumberParas += 1
        EndIf
    Next
    $bOk = _Assert($iBulletParas >= 2, "AI convert bullets tao list cho ca 2 dong bullet", "AI convert bullets chua tao du list") And $bOk
    $bOk = _Assert($iNumberParas >= 1, "AI convert numbered list tao duoc numbered list", "AI convert numbered list that bai") And $bOk
    $bOk = _Assert(Not $bMergedBulletAndNumber, "AI convert lists giu paragraph tach biet", "AI convert lists lam dính bullet va numbered item") And $bOk

    $bOk = _Assert($g_oDoc.Tables.Count >= 1, "AI convert tables tao duoc Word table", "AI convert tables khong tao duoc table") And $bOk
    If $g_oDoc.Tables.Count >= 1 Then
        Local $oTable = $g_oDoc.Tables.Item(1)
        $bOk = _Assert($oTable.Rows.Count >= 2 And $oTable.Columns.Count = 2, "AI convert tables tao dung kich thuoc co ban", "AI convert tables sai kich thuoc") And $bOk
    EndIf

    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _Run()
    DirCreate(@ScriptDir & "\Logs")
    _CreateUiStubs()
    _Log("=== TEST AIFORMAT MARKDOWN STRUCTURES ===")

    If Not _PrepareWord() Then
        _Log("FAIL: Khong the khoi dong Word")
        _WriteLogFile()
        Exit 1
    EndIf

    Local $bOk = _TestMarkdownStructures()
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
