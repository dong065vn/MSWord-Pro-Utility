; Focused per-tab GUI smoke tests

Global Const $APP_EXE = @ScriptDir & "\..\Main_compiled.exe"
Global Const $APP_TITLE = "PDF to Word Fixer Pro v6.1"
Global Const $LOG_DIR = @ScriptDir & "\..\Tests\Logs\Focused"
Global $g_hWnd = ""
Global $g_aPos = 0

DirCreate($LOG_DIR)

If $CmdLine[0] < 1 Then Exit 2
Local $sCase = $CmdLine[1]
Local $sLogPath = $LOG_DIR & "\" & $sCase & ".log"
FileDelete($sLogPath)

Func _WriteLog($sMsg)
    Local $hFile = FileOpen($sLogPath, 1 + 8)
    If $hFile <> -1 Then
        FileWriteLine($hFile, $sMsg)
        FileClose($hFile)
    EndIf
EndFunc

Func _Fail($sMsg, $iCode = 1)
    _WriteLog("FAIL: " & $sMsg)
    Exit $iCode
EndFunc

Func _DismissDialogs()
    For $iLoop = 1 To 6
        Local $aWins = WinList("[CLASS:#32770]")
        Local $bClosed = False
        For $i = 1 To $aWins[0][0]
            If $aWins[$i][0] = "" Then ContinueLoop
            If Not WinExists($aWins[$i][1]) Then ContinueLoop
            WinActivate($aWins[$i][1])
            Sleep(150)
            WinClose($aWins[$i][1])
            Sleep(200)
            $bClosed = True
        Next
        If Not $bClosed Then ExitLoop
    Next
    If $g_hWnd <> "" Then
        WinActivate($g_hWnd)
        Sleep(250)
    EndIf
EndFunc

Func _ClickApp($iX, $iY, $iSleep = 700)
    MouseClick("left", $g_aPos[0] + $iX, $g_aPos[1] + $iY, 1, 0)
    Sleep($iSleep)
EndFunc

Func _EnsureAlive($sStep)
    If Not WinExists($g_hWnd) Then _Fail("App dong/crash tai buoc: " & $sStep, 100)
EndFunc

If Not FileExists($APP_EXE) Then _Fail("Khong tim thay file build.", 3)

Local $oWord = ObjGet("", "Word.Application")
If @error Or Not IsObj($oWord) Then $oWord = ObjCreate("Word.Application")
If Not IsObj($oWord) Then _Fail("Khong mo duoc Word.", 4)
$oWord.Visible = True
If $oWord.Documents.Count = 0 Then $oWord.Documents.Add()

Local $oDoc = $oWord.ActiveDocument
$oDoc.Range(0, 0).Text = _
    "Heading smoke" & @CRLF & _
    "This is **BoldFocus** text." & @CRLF & _
    "DashFocus A" & ChrW(8212) & "B" & @CRLF & _
    "Hyphen focus vi-" & @CRLF & "du" & @CRLF & _
    "Normal paragraph." & @CRLF
$oDoc.Paragraphs.Item(1).Range.Style = "Heading 1"

If $oWord.Documents.Count < 2 Then
    Local $oDoc2 = $oWord.Documents.Add()
    $oDoc2.Range(0, 0).Text = "Second document." & @CRLF
    $oDoc.Activate()
EndIf

; Khoi dong app trong trang thai sach de tranh state tu lan test truoc
ProcessClose("Main_compiled.exe")
Sleep(800)

$g_hWnd = WinGetHandle($APP_TITLE)
If $g_hWnd = "" Then
    Run('"' & $APP_EXE & '"')
    If Not WinWait($APP_TITLE, "", 10) Then _Fail("App khong mo duoc.", 5)
    $g_hWnd = WinGetHandle($APP_TITLE)
EndIf
If $g_hWnd = "" Then _Fail("Khong lay duoc handle app.", 6)
WinActivate($g_hWnd)
If Not WinWaitActive($g_hWnd, "", 5) Then _Fail("Khong kich hoat duoc app.", 7)
$g_aPos = WinGetPos($g_hWnd)
If @error Or Not IsArray($g_aPos) Then _Fail("Khong lay duoc vi tri app.", 8)

_WriteLog("START: " & $sCase)

; Connect
_ClickApp(508, 60, 800)
_DismissDialogs()
_ClickApp(35, 60, 800)
_DismissDialogs()
_EnsureAlive("connect")
_WriteLog("PASS: connect")

Switch $sCase
    Case "pdf_fix"
        _ClickApp(60, 155)
        _ClickApp(655, 380)
        _DismissDialogs()
        _EnsureAlive("pdf help")
        _ClickApp(85, 380)
        _DismissDialogs()
        _EnsureAlive("pdf fix selected")
        _WriteLog("PASS: pdf help")
        _WriteLog("PASS: pdf fix selected")

    Case "format"
        _ClickApp(135, 155)
        _ClickApp(365, 288)
        _EnsureAlive("format preset")
        _ClickApp(240, 288)
        _DismissDialogs()
        _EnsureAlive("format apply selection")
        _WriteLog("PASS: format preset")
        _WriteLog("PASS: format apply selection")

    Case "tools"
        _ClickApp(205, 155)
        _ClickApp(115, 412)
        _DismissDialogs()
        _EnsureAlive("tools wordcount")
        _ClickApp(535, 412)
        _DismissDialogs()
        _EnsureAlive("tools stats")
        _WriteLog("PASS: tools wordcount")
        _WriteLog("PASS: tools stats")

    Case "toc"
        _WriteLog("STEP: open toc tab")
        _ClickApp(275, 155)
        _WriteLog("STEP: click create toc")
        _ClickApp(115, 202)
        _DismissDialogs()
        _EnsureAlive("toc create")
        _WriteLog("STEP: click update toc")
        _ClickApp(245, 202)
        _DismissDialogs()
        _EnsureAlive("toc update")
        _WriteLog("STEP: click delete toc")
        _ClickApp(360, 202)
        _DismissDialogs()
        _EnsureAlive("toc delete")
        _WriteLog("STEP: click preview toc")
        _ClickApp(495, 287)
        _DismissDialogs()
        _EnsureAlive("toc preview")
        _WriteLog("PASS: toc create/update/delete/preview")

    Case "copy_style"
        _WriteLog("STEP: open copy style tab")
        _ClickApp(365, 155)
        _WriteLog("STEP: refresh source")
        _ClickApp(670, 247)
        _DismissDialogs()
        _EnsureAlive("copy style refresh source")
        _WriteLog("STEP: refresh target")
        _ClickApp(670, 307)
        _DismissDialogs()
        _EnsureAlive("copy style refresh target")
        _WriteLog("STEP: preview styles")
        _ClickApp(550, 435)
        _DismissDialogs()
        _EnsureAlive("copy style preview")
        _WriteLog("STEP: open backup hotkeys")
        _ClickApp(255, 482)
        _DismissDialogs()
        _EnsureAlive("copy style backup dialog")
        _WriteLog("PASS: copy style actions")

    Case "advanced"
        _ClickApp(455, 155)
        _ClickApp(570, 200)
        _DismissDialogs()
        _EnsureAlive("advanced headings")
        _ClickApp(255, 300)
        _DismissDialogs()
        _EnsureAlive("advanced convert case")
        _WriteLog("PASS: advanced actions")

    Case "quick_utils"
        _WriteLog("STEP: open quick utils tab")
        _ClickApp(545, 155)
        _WriteLog("STEP: insert date")
        _ClickApp(435, 175)
        _DismissDialogs()
        _EnsureAlive("quick utils insert date")
        _WriteLog("STEP: show doc info")
        _ClickApp(620, 270)
        _DismissDialogs()
        _EnsureAlive("quick utils show info")
        _WriteLog("STEP: toggle bold")
        _ClickApp(470, 285)
        _DismissDialogs()
        _EnsureAlive("quick utils bold")
        _WriteLog("PASS: quick utils actions")

    Case "smart_fix"
        _WriteLog("STEP: open smart fix tab")
        _ClickApp(625, 155)
        _WriteLog("STEP: analyze")
        _ClickApp(90, 268)
        _DismissDialogs()
        _EnsureAlive("smart analyze")
        _WriteLog("STEP: fix hyphen")
        _ClickApp(90, 355)
        _DismissDialogs()
        _EnsureAlive("smart hyphen")
        _WriteLog("STEP: fix dash")
        _ClickApp(145, 387)
        _DismissDialogs()
        _EnsureAlive("smart dash")
        _WriteLog("PASS: smart fix actions")

    Case "ai_format"
        _ClickApp(700, 155)
        _ClickApp(223, 257)
        _DismissDialogs()
        _EnsureAlive("ai bold")
        _ClickApp(454, 467)
        _DismissDialogs()
        _EnsureAlive("ai preview")
        _WriteLog("PASS: ai format actions")

    Case Else
        _Fail("Khong ho tro case: " & $sCase, 9)
EndSwitch

_WriteLog("PASS: completed")
Exit 0
