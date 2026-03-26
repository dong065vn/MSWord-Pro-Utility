#include "..\Config.au3"
#include "..\Shared\Helpers.au3"
#include "..\Modules\Advanced.au3"

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
    Local $hFile = FileOpen(@ScriptDir & "\Logs\Test_AdvancedListTextDetection.out.txt", 2 + 128)
    If $hFile <> -1 Then
        FileWrite($hFile, $g_sTestLog)
        FileClose($hFile)
    EndIf
EndFunc

Func _RunTests()
    Local $bOk = True

    $bOk = _Assert(_NormalizePlainTextNumberingLine("1.	Tieu de") = "1. Tieu de", _
        "Nhan dien numbering dang 1.<tab>text", _
        "Khong nhan dien numbering dang 1.<tab>text") And $bOk

    $bOk = _Assert(_NormalizePlainTextNumberingLine(" 2 )   Noi dung ") = " 2) Noi dung", _
        "Nhan dien numbering dang 2) text", _
        "Khong nhan dien numbering dang 2) text") And $bOk

    $bOk = _Assert(_NormalizePlainTextNumberingLine("3 -	Muc con") = "3- Muc con", _
        "Nhan dien numbering dang 3 - text", _
        "Khong nhan dien numbering dang 3 - text") And $bOk

    $bOk = _Assert(_NormalizePlainTextBulletLine("-	Chu dau dong") = "- Chu dau dong", _
        "Nhan dien bullet dang - text", _
        "Khong nhan dien bullet dang - text") And $bOk

    $bOk = _Assert(_NormalizePlainTextBulletLine("•    Bullet dac biet") = "• Bullet dac biet", _
        "Nhan dien bullet dang ky tu dac biet", _
        "Khong nhan dien bullet dang ky tu dac biet") And $bOk

    $bOk = _Assert(_NormalizePlainTextBulletLine("- [x] Checkbox item") = "[x] Checkbox item", _
        "Nhan dien bullet checkbox tu PDF", _
        "Khong nhan dien bullet checkbox tu PDF") And $bOk

    $bOk = _Assert(_ParagraphMatchesListMode(3, "numbering"), _
        "Phan biet dung ListType numbering", _
        "Khong phan biet duoc ListType numbering") And $bOk

    $bOk = _Assert(_ParagraphMatchesListMode(2, "bullet"), _
        "Phan biet dung ListType bullet", _
        "Khong phan biet duoc ListType bullet") And $bOk

    $bOk = _Assert(Not _ParagraphMatchesListMode(2, "numbering"), _
        "Khong nham bullet thanh numbering", _
        "Nhan dien nham bullet thanh numbering") And $bOk

    Return $bOk
EndFunc

DirCreate(@ScriptDir & "\Logs")
_Log("=== TEST ADVANCED LIST TEXT DETECTION ===")

Local $bOk = _RunTests()
If $bOk Then
    _Log("=== ALL PASS ===")
    _WriteLogFile()
    Exit 0
EndIf

_Log("=== TEST FAILED ===")
_WriteLogFile()
Exit 2
