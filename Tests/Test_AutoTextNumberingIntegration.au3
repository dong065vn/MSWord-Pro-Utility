#include "..\Config.au3"
#include "..\Core\WordConnection.au3"
#include "..\Shared\Helpers.au3"
#include "..\Shared\WordOps.au3"
#include "..\GUI\HotkeyDialogs.au3"
#include "..\Modules\Format.au3"
#include "..\Modules\TOC.au3"
#include "..\Modules\StyleHotkey.au3"
#include "..\Modules\CopyStyle.au3"
#include "..\Modules\Advanced.au3"
#include "..\Modules\AutoTextNumbering.au3"

Global $g_sTestLog = ""
Global Const $TEST_DIR = @ScriptDir & "\Artifacts\AutoTextNumbering"
Global Const $LIVE_LOG_PATH = @ScriptDir & "\Logs\Test_AutoTextNumberingIntegration.out.txt"

Func _Log($sMsg)
    ConsoleWrite($sMsg & @CRLF)
    $g_sTestLog &= $sMsg & @CRLF
    Local $hFile = FileOpen($LIVE_LOG_PATH, 1 + 128)
    If $hFile <> -1 Then
        FileWrite($hFile, $sMsg & @CRLF)
        FileClose($hFile)
    EndIf
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
    Local $hFile = FileOpen(@ScriptDir & "\Logs\Test_AutoTextNumberingIntegration.out.txt", 2 + 128)
    If $hFile <> -1 Then
        FileWrite($hFile, $g_sTestLog)
        FileClose($hFile)
    EndIf
EndFunc

Func _CreateUiStubs()
    $g_hGUI = GUICreate("Test UI", 420, 260, -1, -1, $WS_POPUP)
    $g_lblProgress = GUICtrlCreateLabel("", 10, 10, 380, 20)
    $g_lblStatus = GUICtrlCreateLabel("", 10, 35, 380, 20)
    $g_editPreview = GUICtrlCreateEdit("", 10, 60, 390, 170)
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

Func _CloseBlockingMsgBoxes()
    WinClose("[TITLE:Hoan tat]")
    WinClose("[TITLE:Khong tim thay]")
    WinClose("[TITLE:Khong co style]")
    WinClose("[TITLE:Mau khong hop le]")
EndFunc

Func _GetParagraphText($oDoc, $iIndex)
    If Not IsObj($oDoc) Then Return ""
    Local $sText = $oDoc.Paragraphs.Item($iIndex).Range.Text
    If StringRight($sText, 2) = @CRLF Then Return StringTrimRight($sText, 2)
    If StringRight($sText, 1) = @CR Or StringRight($sText, 1) = @LF Then Return StringTrimRight($sText, 1)
    Return $sText
EndFunc

Func _TestStyleOnlyRenumber($sCaptionStyle)
    Local $bOk = True
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    Local $sDocPath = $TEST_DIR & "\style_only.docx"
    $g_oDoc.Range(0, 0).Text = _
        "Hinh 1.7 Anh thu hai" & @CRLF & _
        "Hinh 1.3 Anh thu nhat" & @CRLF & _
        "Bang 1.9 Bang tong hop" & @CRLF & _
        "Hinh 1.8 Khong dung style caption" & @CRLF

    $g_oDoc.Paragraphs.Item(1).Range.Style = $sCaptionStyle
    $g_oDoc.Paragraphs.Item(2).Range.Style = $sCaptionStyle
    $g_oDoc.Paragraphs.Item(3).Range.Style = $sCaptionStyle

    Local $sNormalStyle = _ResolvePreferredStyleName(_GetParagraphStylesDataFromDoc($g_oDoc), "Normal")
    If $sNormalStyle = "" Then $sNormalStyle = "Normal"
    $g_oDoc.Paragraphs.Item(4).Range.Style = $sNormalStyle

    AdlibRegister("_CloseBlockingMsgBoxes", 300)
    _ApplyAutoTextNumberingFromSample("Hinh 1.1 Anh edge computing", $g_oDoc, $sCaptionStyle, True, True)
    Sleep(900)
    AdlibUnRegister("_CloseBlockingMsgBoxes")

    $bOk = _Assert(_GetParagraphText($g_oDoc, 1) = "Hinh 1.1 Anh thu hai", _
        "Renumber dung paragraph Hinh caption dau tien", _
        "Renumber sai paragraph Hinh caption dau tien") And $bOk
    $bOk = _Assert(_GetParagraphText($g_oDoc, 2) = "Hinh 1.2 Anh thu nhat", _
        "Renumber dung paragraph Hinh caption thu hai", _
        "Renumber sai paragraph Hinh caption thu hai") And $bOk
    $bOk = _Assert(_GetParagraphText($g_oDoc, 3) = "Bang 1.9 Bang tong hop", _
        "Khong nham sang caption Bang", _
        "Bi doi nham caption Bang") And $bOk
    $bOk = _Assert(_GetParagraphText($g_oDoc, 4) = "Hinh 1.8 Khong dung style caption", _
        "StyleOnly bo qua paragraph cung nhom nhung khac style", _
        "StyleOnly van doi nham paragraph khac style") And $bOk

    $g_oDoc.SaveAs2($sDocPath, 16)
    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _TestApplyStyleForMatchingParagraph($sCaptionStyle)
    Local $bOk = True
    $g_oDoc = $g_oWord.Documents.Add()
    If Not IsObj($g_oDoc) Then Return False

    Local $sDocPath = $TEST_DIR & "\apply_style.docx"
    $g_oDoc.Range(0, 0).Text = _
        "Hinh 1.9 Anh chua gan style" & @CRLF & _
        "Hinh 1.4 Anh da la caption" & @CRLF

    Local $sNormalStyle = _ResolvePreferredStyleName(_GetParagraphStylesDataFromDoc($g_oDoc), "Normal")
    If $sNormalStyle = "" Then $sNormalStyle = "Normal"
    $g_oDoc.Paragraphs.Item(1).Range.Style = $sNormalStyle
    $g_oDoc.Paragraphs.Item(2).Range.Style = $sCaptionStyle

    AdlibRegister("_CloseBlockingMsgBoxes", 300)
    _ApplyAutoTextNumberingFromSample("Hinh 1.1 Anh edge computing", $g_oDoc, $sCaptionStyle, False, True)
    Sleep(900)
    AdlibUnRegister("_CloseBlockingMsgBoxes")

    $bOk = _Assert(_GetParagraphText($g_oDoc, 1) = "Hinh 1.1 Anh chua gan style", _
        "Bo loc style tat thi renumber duoc paragraph chua gan style", _
        "Khong renumber duoc paragraph chua gan style") And $bOk
    $bOk = _Assert(_GetParagraphText($g_oDoc, 2) = "Hinh 1.2 Anh da la caption", _
        "Renumber tiep theo dung thu tu trong toan bo nhom", _
        "Thu tu renumber toan nhom bi sai") And $bOk
    $bOk = _Assert(_ParagraphUsesStyleName($g_oDoc.Paragraphs.Item(1), $sCaptionStyle), _
        "Paragraph match duoc gan style caption sau khi danh so", _
        "Paragraph match khong duoc gan style caption") And $bOk

    $g_oDoc.SaveAs2($sDocPath, 16)
    $g_oDoc.Close(0)
    Return $bOk
EndFunc

Func _Run()
    DirCreate(@ScriptDir & "\Logs")
    FileDelete($LIVE_LOG_PATH)
    _ResetArtifacts()
    _CreateUiStubs()

    _Log("=== TEST AUTO TEXT NUMBERING INTEGRATION ===")
    If Not _PrepareWord() Then
        _Log("FAIL: Khong the khoi dong Word")
        _WriteLogFile()
        Exit 1
    EndIf

    Local $sStyles = _GetParagraphStylesDataFromDoc($g_oWord.Documents.Add())
    _Log("INFO: Da tao doc tam de lay style")
    $g_oWord.ActiveDocument.Close(0)
    Local $sCaptionStyle = _ResolvePreferredStyleName($sStyles, "Caption")
    If $sCaptionStyle = "" Then $sCaptionStyle = "Caption"
    _Log("INFO: Caption style = " & $sCaptionStyle)

    Local $bStyleOnly = _TestStyleOnlyRenumber($sCaptionStyle)
    Local $bApplyStyle = _TestApplyStyleForMatchingParagraph($sCaptionStyle)

    If $bStyleOnly And $bApplyStyle Then
        _Log("=== ALL PASS ===")
        _WriteLogFile()
        Exit 0
    EndIf

    _Log("=== TEST FAILED ===")
    _WriteLogFile()
    Exit 2
EndFunc

_Run()
