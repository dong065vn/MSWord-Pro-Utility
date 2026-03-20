; ============================================
; TEST_INTEGRATION.AU3 - Verify Hotkey Module Integration
; Muc tieu: Kiem tra xem module da duoc tich hop dung chua
; ============================================

#include "..\Config.au3"

; === TEST CONFIGURATION ===
Global $g_sTestLog = ""
Global $g_iTestsPassed = 0
Global $g_iTestsFailed = 0

; === MAIN TEST RUNNER ===
Func _RunIntegrationTests()
    ConsoleWrite("=== BAT DAU TEST INTEGRATION ===" & @CRLF)
    _LogTest("Test started at: " & @HOUR & ":" & @MIN & ":" & @SEC)
    
    ; Test 1: Check includes
    If Not _Test_CheckIncludes() Then
        _LogTest("FAIL: Includes khong day du")
    Else
        $g_iTestsPassed += 1
    EndIf
    
    ; Test 2: Check Config variables
    If Not _Test_CheckConfigVars() Then
        _LogTest("FAIL: Config variables khong day du")
    Else
        $g_iTestsPassed += 1
    EndIf
    
    ; Test 3: Check functions exist
    If Not _Test_CheckFunctionsExist() Then
        _LogTest("FAIL: Functions khong ton tai")
    Else
        $g_iTestsPassed += 1
    EndIf
    
    ; Test 4: Check file structure
    If Not _Test_CheckFileStructure() Then
        _LogTest("FAIL: File structure khong day du")
    Else
        $g_iTestsPassed += 1
    EndIf
    
    ; Summary
    _LogTest("")
    _LogTest("=== TEST SUMMARY ===")
    _LogTest("Passed: " & $g_iTestsPassed)
    _LogTest("Failed: " & $g_iTestsFailed)
    _LogTest("Total: " & ($g_iTestsPassed + $g_iTestsFailed))
    
    If $g_iTestsFailed = 0 Then
        _LogTest("=== TAT CA TEST PASS ===")
        _LogTest("Module da duoc tich hop thanh cong!")
    Else
        _LogTest("=== CO LOI XAY RA ===")
        _LogTest("Vui long kiem tra lai!")
    EndIf
    
    _SaveTestLog()
    
    Local $sMsg = "INTEGRATION TEST COMPLETE!" & @CRLF & @CRLF
    $sMsg &= "Passed: " & $g_iTestsPassed & @CRLF
    $sMsg &= "Failed: " & $g_iTestsFailed & @CRLF & @CRLF
    
    If $g_iTestsFailed = 0 Then
        $sMsg &= "✅ Module da duoc tich hop thanh cong!" & @CRLF
        $sMsg &= "San sang de compile va test thuc te!"
        MsgBox($MB_ICONINFORMATION, "Test Complete", $sMsg)
    Else
        $sMsg &= "❌ Co loi xay ra!" & @CRLF
        $sMsg &= "Xem log: IntegrationTestLog.txt"
        MsgBox($MB_ICONWARNING, "Test Failed", $sMsg)
    EndIf
    
    Return ($g_iTestsFailed = 0)
EndFunc

; === TEST 1: Check Includes ===
Func _Test_CheckIncludes()
    _LogTest("[TEST 1] Kiem tra includes trong Main.au3...")
    
    Local $sMainFile = @ScriptDir & "\..\Main.au3"
    If Not FileExists($sMainFile) Then
        _LogTest("  FAIL: Khong tim thay Main.au3")
        $g_iTestsFailed += 1
        Return False
    EndIf
    
    Local $sContent = FileRead($sMainFile)
    
    ; Check required includes
    Local $aRequiredIncludes[2] = [ _
        "Modules\StyleHotkey.au3", _
        "GUI\HotkeyDialogs.au3" _
    ]
    
    Local $bAllFound = True
    For $i = 0 To UBound($aRequiredIncludes) - 1
        If Not StringInStr($sContent, $aRequiredIncludes[$i]) Then
            _LogTest("  FAIL: Thieu include: " & $aRequiredIncludes[$i])
            $bAllFound = False
        Else
            _LogTest("  PASS: Tim thay include: " & $aRequiredIncludes[$i])
        EndIf
    Next
    
    If Not $bAllFound Then
        $g_iTestsFailed += 1
        Return False
    EndIf
    
    _LogTest("  PASS: Tat ca includes da co")
    Return True
EndFunc

; === TEST 2: Check Config Variables ===
Func _Test_CheckConfigVars()
    _LogTest("[TEST 2] Kiem tra Config variables...")
    
    Local $sConfigFile = @ScriptDir & "\..\Config.au3"
    If Not FileExists($sConfigFile) Then
        _LogTest("  FAIL: Khong tim thay Config.au3")
        $g_iTestsFailed += 1
        Return False
    EndIf
    
    Local $sContent = FileRead($sConfigFile)
    
    ; Check required variables
    Local $aRequiredVars[3] = [ _
        "$g_btnApplyHotkeys", _
        "$g_btnBackupHotkeys", _
        "$g_btnRefreshHotkeys" _
    ]
    
    Local $bAllFound = True
    For $i = 0 To UBound($aRequiredVars) - 1
        If Not StringInStr($sContent, $aRequiredVars[$i]) Then
            _LogTest("  FAIL: Thieu variable: " & $aRequiredVars[$i])
            $bAllFound = False
        Else
            _LogTest("  PASS: Tim thay variable: " & $aRequiredVars[$i])
        EndIf
    Next
    
    If Not $bAllFound Then
        $g_iTestsFailed += 1
        Return False
    EndIf
    
    _LogTest("  PASS: Tat ca variables da co")
    Return True
EndFunc

; === TEST 3: Check Functions Exist ===
Func _Test_CheckFunctionsExist()
    _LogTest("[TEST 3] Kiem tra functions ton tai...")
    
    ; Check StyleHotkey.au3
    Local $sStyleHotkeyFile = @ScriptDir & "\..\Modules\StyleHotkey.au3"
    If Not FileExists($sStyleHotkeyFile) Then
        _LogTest("  FAIL: Khong tim thay StyleHotkey.au3")
        $g_iTestsFailed += 1
        Return False
    EndIf
    
    Local $sContent = FileRead($sStyleHotkeyFile)
    
    Local $aRequiredFuncs[5] = [ _
        "_ApplyStyleHotkeyViaWord", _
        "_ParseHotkeyToWordKeys", _
        "_ValidateHotkeyFormat", _
        "_BackupHotkeysNow", _
        "_RestoreHotkeysFromBackup" _
    ]
    
    Local $bAllFound = True
    For $i = 0 To UBound($aRequiredFuncs) - 1
        If Not StringInStr($sContent, "Func " & $aRequiredFuncs[$i]) Then
            _LogTest("  FAIL: Thieu function: " & $aRequiredFuncs[$i])
            $bAllFound = False
        Else
            _LogTest("  PASS: Tim thay function: " & $aRequiredFuncs[$i])
        EndIf
    Next
    
    If Not $bAllFound Then
        $g_iTestsFailed += 1
        Return False
    EndIf
    
    _LogTest("  PASS: Tat ca functions da co")
    Return True
EndFunc

; === TEST 4: Check File Structure ===
Func _Test_CheckFileStructure()
    _LogTest("[TEST 4] Kiem tra file structure...")
    
    Local $aRequiredFiles[4] = [ _
        @ScriptDir & "\..\Modules\StyleHotkey.au3", _
        @ScriptDir & "\..\GUI\HotkeyDialogs.au3", _
        @ScriptDir & "\..\Tools\INTEGRATION_GUIDE.md", _
        @ScriptDir & "\..\Resources\StyleHotkeys.ini" _
    ]
    
    Local $bAllFound = True
    For $i = 0 To UBound($aRequiredFiles) - 1
        If Not FileExists($aRequiredFiles[$i]) Then
            _LogTest("  FAIL: Thieu file: " & $aRequiredFiles[$i])
            $bAllFound = False
        Else
            _LogTest("  PASS: Tim thay file: " & $aRequiredFiles[$i])
        EndIf
    Next
    
    If Not $bAllFound Then
        $g_iTestsFailed += 1
        Return False
    EndIf
    
    _LogTest("  PASS: Tat ca files da co")
    Return True
EndFunc

; === HELPER: Log Test ===
Func _LogTest($sMessage)
    ConsoleWrite($sMessage & @CRLF)
    $g_sTestLog &= $sMessage & @CRLF
EndFunc

; === HELPER: Save Test Log ===
Func _SaveTestLog()
    Local $sLogPath = @ScriptDir & "\..\IntegrationTestLog_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & ".txt"
    Local $hFile = FileOpen($sLogPath, 2 + 128)
    If $hFile <> -1 Then
        FileWrite($hFile, $g_sTestLog)
        FileClose($hFile)
        _LogTest("Log saved: " & $sLogPath)
    EndIf
EndFunc

; === RUN TESTS ===
_RunIntegrationTests()
