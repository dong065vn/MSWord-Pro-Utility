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
    Local $hFile = FileOpen(@ScriptDir & "\Logs\Test_AIPreviewCounts.out.txt", 2 + 128)
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

Func _ClosePreviewDialog()
    If WinExists("Preview") Then
        WinActivate("Preview")
        Send("{ENTER}")
    EndIf
EndFunc

Func _TestPreviewCounts()
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    Local $sText = "# H1" & @CRLF & _
        "## H2" & @CRLF & _
        "Paragraph with **bold one** and **bold two**." & @CRLF & _
        "[OpenAI](https://openai.com)" & @CRLF & _
        "```" & @CRLF & _
        "code sample" & @CRLF & _
        "```" & @CRLF
    $g_oDoc.Range(0, 0).Text = $sText

    AdlibRegister("_ClosePreviewDialog", 250)
    _AI_PreviewChanges()
    AdlibUnRegister("_ClosePreviewDialog")

    Local $sPreview = GUICtrlRead($g_editPreview)
    Local $bOk = True
    $bOk = _Assert(StringInStr($sPreview, "Headings (##): ~2") > 0, "AI preview dem dung so heading", "AI preview dem sai heading") And $bOk
    $bOk = _Assert(StringInStr($sPreview, "Bold (**): ~2") > 0, "AI preview dem dung so bold", "AI preview dem sai bold") And $bOk
    $bOk = _Assert(StringInStr($sPreview, "Code blocks (```): ~1") > 0, "AI preview dem dung so code block", "AI preview dem sai code block") And $bOk
    $bOk = _Assert(StringInStr($sPreview, "Links [...]: ~1") > 0, "AI preview dem dung so link", "AI preview dem sai link") And $bOk

    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _Run()
    DirCreate(@ScriptDir & "\Logs")
    _CreateUiStubs()
    _Log("=== TEST AIPREVIEW COUNTS ===")

    If Not _PrepareWord() Then
        _Log("FAIL: Khong the khoi dong Word")
        _WriteLogFile()
        Exit 1
    EndIf

    Local $bOk = _TestPreviewCounts()
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
