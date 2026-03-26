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
    Local $hFile = FileOpen(@ScriptDir & "\Logs\Test_ToolsAutoTextNumberingHelpers.out.txt", 2 + 128)
    If $hFile <> -1 Then
        FileWrite($hFile, $g_sTestLog)
        FileClose($hFile)
    EndIf
EndFunc

Func _RunTests()
    Local $bOk = True
    Local $sLabel = "", $sFixedPrefix = "", $sSeparator = "", $sTitle = ""
    Local $iStartNumber = 0

    $bOk = _Assert(_ParseAutoTextCaptionSample("Hinh 1.1 Anh edge computing", $sLabel, $sFixedPrefix, $iStartNumber, $sSeparator, $sTitle), _
        "Parse duoc caption mau Hinh 1.1", _
        "Khong parse duoc caption mau Hinh 1.1") And $bOk
    $bOk = _Assert($sLabel = "Hinh" And $sFixedPrefix = "1" And $iStartNumber = 1 And $sSeparator = " " And $sTitle = "Anh edge computing", _
        "Tach dung label/prefix/number/title", _
        "Tach sai label/prefix/number/title") And $bOk

    Local $sLeading = "", $sNumberBlock = "", $sEol = ""
    $bOk = _Assert(_ParseAutoTextCaptionLine("  Hinh 1.4: Anh MEC" & @CR, $sLeading, $sLabel, $sNumberBlock, $sSeparator, $sTitle, $sEol), _
        "Parse duoc dong caption thuc te", _
        "Khong parse duoc dong caption thuc te") And $bOk
    $bOk = _Assert($sLeading = "  " And $sLabel = "Hinh" And $sNumberBlock = "1.4" And $sSeparator = ": " And $sTitle = "Anh MEC" And $sEol = @CR, _
        "Lay dung thong tin caption thuc te", _
        "Lay sai thong tin caption thuc te") And $bOk

    $bOk = _Assert(_AutoTextCaptionMatchesGroup("Hinh", "Hinh", "1.9", "1"), _
        "Nhan dien dung nhom Hinh 1.x", _
        "Khong nhan dien duoc nhom Hinh 1.x") And $bOk
    $bOk = _Assert(Not _AutoTextCaptionMatchesGroup("Hinh", "Hinh", "2.1", "1"), _
        "Khong nham Hinh 2.x vao nhom Hinh 1.x", _
        "Nhan dien nham Hinh 2.x vao nhom Hinh 1.x") And $bOk

    $bOk = _Assert(_BuildAutoTextCaptionNumber("1", 3) = "1.3", _
        "Dung dung so moi cho nhom 1.x", _
        "Dung sai so moi cho nhom 1.x") And $bOk
    $bOk = _Assert(_BuildAutoTextCaptionNumber("", 5) = "5", _
        "Dung dung so moi cho nhom mot cap", _
        "Dung sai so moi cho nhom mot cap") And $bOk

    $bOk = _Assert(_GetAutoTextNumberingSampleForType("Bang") = "Bang 1.1 Bang ket qua thuc nghiem", _
        "Co sample mac dinh cho Bang", _
        "Khong co sample mac dinh cho Bang") And $bOk

    Return $bOk
EndFunc

DirCreate(@ScriptDir & "\Logs")
_Log("=== TEST TOOLS AUTO TEXT NUMBERING HELPERS ===")

Local $bOk = _RunTests()
If $bOk Then
    _Log("=== ALL PASS ===")
    _WriteLogFile()
    Exit 0
EndIf

_Log("=== TEST FAILED ===")
_WriteLogFile()
Exit 2
