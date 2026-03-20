; ============================================
; TEST_HOTKEYFLOW.AU3 - Test script cho Hotkey Management
; Muc tieu: Verify flow Backup → Edit → Apply → Verify
; ============================================

#include "..\Config.au3"
#include "..\Core\WordConnection.au3"
#include "..\Shared\Helpers.au3"
#include "..\Modules\StyleHotkey.au3"
#include "..\GUI\HotkeyDialogs.au3"

; === TEST CONFIGURATION ===
Global Const $TEST_STYLE_NAME = "Heading 1"
Global Const $TEST_HOTKEY = "Ctrl+Alt+9"
Global $g_sTestLog = ""

; === MAIN TEST RUNNER ===
Func _RunHotkeyTests()
    ConsoleWrite("=== BAT DAU TEST HOTKEY FLOW ===" & @CRLF)
    _LogTest("Test started at: " & @HOUR & ":" & @MIN & ":" & @SEC)
    
    ; Test 1: Ket noi Word
    If Not _Test_ConnectWord() Then
        _LogTest("FAIL: Khong the ket noi Word")
        Return False
    EndIf
    
    ; Test 2: Backup hotkeys
    If Not _Test_BackupHotkeys() Then
        _LogTest("FAIL: Backup that bai")
        Return False
    EndIf
    
    ; Test 3: Parse hotkey string
    If Not _Test_ParseHotkey() Then
        _LogTest("FAIL: Parse hotkey that bai")
        Return False
    EndIf
    
    ; Test 4: Apply hotkey to Word
    If Not _Test_ApplyHotkey() Then
        _LogTest("FAIL: Apply hotkey that bai")
        Return False
    EndIf
    
    ; Test 5: Verify hotkey in Word
    If Not _Test_VerifyHotkey() Then
        _LogTest("FAIL: Verify hotkey that bai")
        Return False
    EndIf
    
    ; Test 6: Restore from backup
    If Not _Test_RestoreBackup() Then
        _LogTest("FAIL: Restore that bai")
        Return False
    EndIf
    
    _LogTest("=== TAT CA TEST PASS ===")
    _SaveTestLog()
    
    MsgBox($MB_ICONINFORMATION, "Test Complete", "Tat ca test da PASS!" & @CRLF & "Xem log: TestLog.txt")
    Return True
EndFunc

; === TEST 1: Connect to Word ===
Func _Test_ConnectWord()
    _LogTest("[TEST 1] Ket noi Word...")
    
    $g_oWord = ObjGet("", "Word.Application")
    If @error Or Not IsObj($g_oWord) Then
        _LogTest("  Khong tim thay Word dang chay, tao instance moi...")
        $g_oWord = ObjCreate("Word.Application")
        If @error Or Not IsObj($g_oWord) Then
            _LogTest("  FAIL: Khong the tao Word instance")
            Return False
        EndIf
        $g_oWord.Visible = True
    EndIf
    
    ; Tao document test neu chua co
    If $g_oWord.Documents.Count = 0 Then
        $g_oDoc = $g_oWord.Documents.Add()
    Else
        $g_oDoc = $g_oWord.ActiveDocument
    EndIf
    
    If Not IsObj($g_oDoc) Then
        _LogTest("  FAIL: Khong co document")
        Return False
    EndIf
    
    _LogTest("  PASS: Ket noi thanh cong - Doc: " & $g_oDoc.Name)
    Return True
EndFunc

; === TEST 2: Backup Hotkeys ===
Func _Test_BackupHotkeys()
    _LogTest("[TEST 2] Backup hotkeys...")
    
    Local $sBackupName = _BackupHotkeysNow()
    If $sBackupName = "" Then
        _LogTest("  FAIL: Backup that bai")
        Return False
    EndIf
    
    _LogTest("  PASS: Backup thanh cong - File: " & $sBackupName)
    Return True
EndFunc

; === TEST 3: Parse Hotkey String ===
Func _Test_ParseHotkey()
    _LogTest("[TEST 3] Parse hotkey string...")
    
    Local $aKeys = _ParseHotkeyToWordKeys($TEST_HOTKEY)
    If UBound($aKeys) < 4 Then
        _LogTest("  FAIL: Parse tra ve mang khong hop le")
        Return False
    EndIf
    
    ; Kiem tra co it nhat 1 key khac wdNoKey (255)
    Local $bHasKey = False
    For $i = 0 To 3
        If $aKeys[$i] <> 255 Then
            $bHasKey = True
            ExitLoop
        EndIf
    Next
    
    If Not $bHasKey Then
        _LogTest("  FAIL: Khong co key nao duoc parse")
        Return False
    EndIf
    
    _LogTest("  PASS: Parse thanh cong - Keys: [" & $aKeys[0] & ", " & $aKeys[1] & ", " & $aKeys[2] & ", " & $aKeys[3] & "]")
    Return True
EndFunc

; === TEST 4: Apply Hotkey to Word ===
Func _Test_ApplyHotkey()
    _LogTest("[TEST 4] Apply hotkey to Word...")
    
    Local $bResult = _ApplyStyleHotkeyViaWord($g_oDoc, $TEST_STYLE_NAME, $TEST_HOTKEY)
    If Not $bResult Then
        _LogTest("  FAIL: Apply hotkey that bai")
        Return False
    EndIf
    
    _LogTest("  PASS: Apply hotkey thanh cong")
    Return True
EndFunc

; === TEST 5: Verify Hotkey in Word ===
Func _Test_VerifyHotkey()
    _LogTest("[TEST 5] Verify hotkey in Word...")
    
    Local Const $wdKeyCategoryStyle = 5
    Local $oKeysBound = $g_oWord.KeysBoundTo($wdKeyCategoryStyle, $TEST_STYLE_NAME)
    
    If Not IsObj($oKeysBound) Or $oKeysBound.Count = 0 Then
        _LogTest("  FAIL: Khong tim thay hotkey trong Word")
        Return False
    EndIf
    
    _LogTest("  PASS: Tim thay " & $oKeysBound.Count & " hotkey(s)")
    For $i = 1 To $oKeysBound.Count
        Local $oKey = $oKeysBound.Item($i)
        If IsObj($oKey) Then
            _LogTest("    - KeyString: " & $oKey.KeyString)
        EndIf
    Next
    
    Return True
EndFunc

; === TEST 6: Restore from Backup ===
Func _Test_RestoreBackup()
    _LogTest("[TEST 6] Restore from backup...")
    
    ; Tim backup moi nhat
    Local $sBackupDir = _GetHotkeyBackupDir()
    Local $aFiles = _FileListToArray($sBackupDir, "StyleHotkeys_*.ini", 1)
    
    If @error Or Not IsArray($aFiles) Or $aFiles[0] = 0 Then
        _LogTest("  SKIP: Khong co backup de restore")
        Return True
    EndIf
    
    ; Lay file moi nhat (cuoi mang)
    Local $sLatestBackup = $sBackupDir & "\" & $aFiles[$aFiles[0]]
    Local $bResult = _RestoreHotkeysFromBackup($sLatestBackup)
    
    If Not $bResult Then
        _LogTest("  FAIL: Restore that bai")
        Return False
    EndIf
    
    _LogTest("  PASS: Restore thanh cong tu: " & $aFiles[$aFiles[0]])
    Return True
EndFunc

; === HELPER: Log Test ===
Func _LogTest($sMessage)
    ConsoleWrite($sMessage & @CRLF)
    $g_sTestLog &= $sMessage & @CRLF
EndFunc

; === HELPER: Save Test Log ===
Func _SaveTestLog()
    Local $sLogPath = @ScriptDir & "\..\TestLog_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & ".txt"
    Local $hFile = FileOpen($sLogPath, 2 + 128)
    If $hFile <> -1 Then
        FileWrite($hFile, $g_sTestLog)
        FileClose($hFile)
        _LogTest("Log saved: " & $sLogPath)
    EndIf
EndFunc

; === RUN TESTS ===
_RunHotkeyTests()
