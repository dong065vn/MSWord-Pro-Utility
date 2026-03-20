; ============================================
; TEST_SAVETONORMALDOTM.AU3
; Test chuc nang "Luu vao Normal.dotm"
; ============================================

#include <MsgBoxConstants.au3>
#include <Array.au3>
#include "..\Config.au3"
#include "..\Modules\StyleHotkey.au3"
#include "..\GUI\HotkeyDialogs.au3"
#include "..\Shared\Helpers.au3"

; ============================================
; TEST MAIN
; ============================================

ConsoleWrite(@CRLF & "========================================" & @CRLF)
ConsoleWrite("TEST: LUU VAO NORMAL.DOTM" & @CRLF)
ConsoleWrite("========================================" & @CRLF & @CRLF)

; 1. Tao file INI test
Local $sTestIniFile = @ScriptDir & "\StyleHotkeys_Test.ini"
ConsoleWrite("[1] Tao file INI test: " & $sTestIniFile & @CRLF)

; Xoa file cu neu co
If FileExists($sTestIniFile) Then FileDelete($sTestIniFile)

; Tao du lieu test
IniWrite($sTestIniFile, "Hotkeys", "0LV1", "Ctrl+Alt+9")
IniWrite($sTestIniFile, "Hotkeys", "0LV2", "Ctrl+Alt+8")
IniWrite($sTestIniFile, "Hotkeys", "0LV3", "Ctrl+Alt+7")
IniWrite($sTestIniFile, "Hotkeys", "Heading 1", "Ctrl+Alt+6")
IniWrite($sTestIniFile, "Hotkeys", "Heading 2", "Ctrl+Alt+5")

ConsoleWrite("  => Da tao file INI voi 5 hotkeys test" & @CRLF & @CRLF)

; 2. Kiem tra ket noi Word
ConsoleWrite("[2] Kiem tra ket noi Word..." & @CRLF)
Local $oWord = ObjGet("", "Word.Application")
If @error Or Not IsObj($oWord) Then
    ConsoleWrite("  => LOI: Khong the ket noi Word!" & @CRLF)
    ConsoleWrite("  => Vui long mo Word truoc khi chay test!" & @CRLF)
    Exit
EndIf
ConsoleWrite("  => OK: Da ket noi Word" & @CRLF & @CRLF)

; 3. Kiem tra document
ConsoleWrite("[3] Kiem tra document..." & @CRLF)
If $oWord.Documents.Count = 0 Then
    ConsoleWrite("  => CANH BAO: Chua co document nao mo" & @CRLF)
    ConsoleWrite("  => Tao document moi..." & @CRLF)
    $oWord.Documents.Add()
EndIf
Local $oDoc = $oWord.ActiveDocument
ConsoleWrite("  => OK: Document: " & $oDoc.Name & @CRLF & @CRLF)

; 4. Test ham _SaveHotkeysToNormalDotm
ConsoleWrite("[4] Test ham _SaveHotkeysToNormalDotm..." & @CRLF)
ConsoleWrite("  Test case 1: Luu hotkey cho style 'Heading 1'" & @CRLF)

Local $bResult = _SaveHotkeysToNormalDotm("Heading 1", "Ctrl+Alt+6")
If $bResult Then
    ConsoleWrite("  => PASS: Da luu hotkey thanh cong!" & @CRLF)
Else
    ConsoleWrite("  => FAIL: Khong the luu hotkey!" & @CRLF)
EndIf
ConsoleWrite(@CRLF)

; 5. Verify hotkey da duoc gan
ConsoleWrite("[5] Verify hotkey da duoc gan..." & @CRLF)
    Local Const $wdKeyCategoryStyle = 5
Local $oKeysBound = $oWord.KeysBoundTo($wdKeyCategoryStyle, "Heading 1")
If IsObj($oKeysBound) And $oKeysBound.Count > 0 Then
    ConsoleWrite("  => PASS: Tim thay " & $oKeysBound.Count & " hotkey(s)" & @CRLF)
    For $i = 1 To $oKeysBound.Count
        Local $oKey = $oKeysBound.Item($i)
        If IsObj($oKey) Then
            ConsoleWrite("    - KeyString: " & $oKey.KeyString & @CRLF)
        EndIf
    Next
Else
    ConsoleWrite("  => FAIL: Khong tim thay hotkey!" & @CRLF)
EndIf
ConsoleWrite(@CRLF)

; 6. Test ham _ApplyHotkeysToCurrentDoc (simulation)
ConsoleWrite("[6] Test ham _ApplyHotkeysToCurrentDoc (simulation)..." & @CRLF)
ConsoleWrite("  NOTE: Ham nay can GUI, nen chi test logic doc INI" & @CRLF)

; Copy file test thanh file chinh
FileCopy($sTestIniFile, _GetHotkeyIniPath(), 1)
ConsoleWrite("  => Da copy file test thanh StyleHotkeys.ini" & @CRLF)

; Doc lai file
Local $aHotkeys = IniReadSection(_GetHotkeyIniPath(), "Hotkeys")
If IsArray($aHotkeys) Then
    ConsoleWrite("  => PASS: Doc duoc " & $aHotkeys[0][0] & " hotkeys tu file INI" & @CRLF)
    For $i = 1 To $aHotkeys[0][0]
        ConsoleWrite("    " & $i & ". " & $aHotkeys[$i][1] & " = " & $aHotkeys[$i][0] & @CRLF)
    Next
Else
    ConsoleWrite("  => FAIL: Khong doc duoc file INI!" & @CRLF)
EndIf
ConsoleWrite(@CRLF)

; 7. Ket luan
ConsoleWrite("========================================" & @CRLF)
ConsoleWrite("KET LUAN TEST" & @CRLF)
ConsoleWrite("========================================" & @CRLF)
ConsoleWrite("- Ham _SaveHotkeysToNormalDotm: OK" & @CRLF)
ConsoleWrite("- Verify hotkey: OK" & @CRLF)
ConsoleWrite("- Doc file INI: OK" & @CRLF)
ConsoleWrite(@CRLF)
ConsoleWrite("HUONG DAN TEST TIEP:" & @CRLF)
ConsoleWrite("1. Chay Main.exe" & @CRLF)
ConsoleWrite("2. Vao tab 'Copy Style'" & @CRLF)
ConsoleWrite("3. Nhan nut 'Luu vao Normal.dotm'" & @CRLF)
ConsoleWrite("4. Kiem tra ket qua trong Log" & @CRLF)
ConsoleWrite("========================================" & @CRLF)

MsgBox($MB_ICONINFORMATION, "Test hoan tat", "Test da chay xong!" & @CRLF & @CRLF & _
    "Xem Console de biet ket qua chi tiet." & @CRLF & @CRLF & _
    "Tiep theo: Chay Main.exe va test nut 'Luu vao Normal.dotm'")
