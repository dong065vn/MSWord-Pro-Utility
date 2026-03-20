; Runtime smoke test: ket noi Word va kich hoat 1 handler dai dien cho moi tab

Global Const $APP_EXE = @ScriptDir & "\..\Main_compiled.exe"
Global Const $APP_TITLE = "PDF to Word Fixer Pro v6.1"
Global Const $LOG_PATH = @ScriptDir & "\..\Tests\Logs\RuntimeSmokeTest.log"

DirCreate(@ScriptDir & "\..\Tests\Logs")
FileDelete($LOG_PATH)

Func _WriteLog($sMsg)
    Local $hFile = FileOpen($LOG_PATH, 1 + 8)
    If $hFile <> -1 Then
        FileWriteLine($hFile, $sMsg)
        FileClose($hFile)
    EndIf
EndFunc

If Not FileExists($APP_EXE) Then
    _WriteLog("FAIL: Khong tim thay file build.")
    Exit 1
EndIf

Local $oWord = ObjGet("", "Word.Application")
If @error Or Not IsObj($oWord) Then
    $oWord = ObjCreate("Word.Application")
EndIf
If Not IsObj($oWord) Then
    _WriteLog("FAIL: Khong mo duoc Word.")
    Exit 2
EndIf

$oWord.Visible = True
If $oWord.Documents.Count = 0 Then $oWord.Documents.Add()
Local $oDoc = $oWord.ActiveDocument
$oDoc.Range(0, 0).Text = _
    "Heading demo" & @CRLF & _
    "This is **BoldSmoke** text." & @CRLF & _
    "DashSmoke A" & ChrW(8212) & "B" & @CRLF & _
    "Hyphen smoke vi-" & @CRLF & "du" & @CRLF & _
    "Normal smoke paragraph." & @CRLF
$oDoc.Paragraphs.Item(1).Range.Style = "Heading 1"
If $oWord.Documents.Count < 2 Then
    Local $oDoc2 = $oWord.Documents.Add()
    $oDoc2.Range(0, 0).Text = "Second smoke document." & @CRLF
    $oDoc.Activate()
EndIf

Local $hWnd = WinGetHandle($APP_TITLE)
If $hWnd = "" Then
    Run('"' & $APP_EXE & '"')
    If Not WinWait($APP_TITLE, "", 10) Then
        _WriteLog("FAIL: App khong mo duoc.")
        Exit 3
    EndIf
    $hWnd = WinGetHandle($APP_TITLE)
EndIf

If $hWnd = "" Then
    _WriteLog("FAIL: Khong lay duoc handle app.")
    Exit 4
EndIf

WinActivate($hWnd)
If Not WinWaitActive($hWnd, "", 5) Then
    _WriteLog("FAIL: Khong kich hoat duoc app.")
    Exit 5
EndIf

Local $aPos = WinGetPos($hWnd)
If @error Or Not IsArray($aPos) Then
    _WriteLog("FAIL: Khong lay duoc vi tri app.")
    Exit 6
EndIf

Func _ClickApp($iX, $iY)
    MouseClick("left", $aPos[0] + $iX, $aPos[1] + $iY, 1, 0)
    Sleep(500)
EndFunc

Func _EnsureAlive($sStep)
    If Not WinExists($hWnd) Then
        _WriteLog("FAIL: App dong/crash tai buoc: " & $sStep)
        Exit 100
    EndIf
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
    WinActivate($hWnd)
    Sleep(250)
EndFunc

Func _RunAction($sTab, $iTabX, $iTabY, $sAction, $iBtnX, $iBtnY, $bDismiss = True)
    _ClickApp($iTabX, $iTabY)
    _EnsureAlive("chuyen tab " & $sTab)
    _ClickApp($iBtnX, $iBtnY)
    Sleep(900)
    If $bDismiss Then _DismissDialogs()
    _EnsureAlive($sTab & " -> " & $sAction)
    _WriteLog("PASS: " & $sTab & " -> " & $sAction)
EndFunc

_WriteLog("START: Runtime smoke test")

; Ket noi Word
_ClickApp(508, 60) ; Lam moi
_DismissDialogs()
_EnsureAlive("Lam moi Word docs")
_ClickApp(35, 60) ; Tu dong
_DismissDialogs()
_EnsureAlive("Tu dong ket noi Word")
_WriteLog("PASS: Connection -> Refresh + Auto connect")

; Mỗi tab một handler đại diện
_RunAction("PDF Fix", 60, 155, "Help", 635, 360)
_RunAction("PDF Fix", 60, 155, "Fix selected", 35, 360)
_RunAction("Format", 135, 155, "Preset VN", 315, 270, False)
_RunAction("Format", 135, 155, "Apply selection", 175, 270)
_RunAction("Tools", 205, 155, "Thong ke tong hop", 470, 398)
_RunAction("Tools", 205, 155, "Word count", 50, 398)

_ClickApp(275, 155)
_ClickApp(50, 185) ; Tao muc luc
Sleep(1200)
_DismissDialogs()
_EnsureAlive("TOC -> Tao muc luc")
_WriteLog("PASS: TOC -> Tao muc luc")
_ClickApp(195, 185) ; Cap nhat
Sleep(800)
_DismissDialogs()
_EnsureAlive("TOC -> Cap nhat")
_WriteLog("PASS: TOC -> Cap nhat")
_ClickApp(310, 185) ; Xoa muc luc
Sleep(800)
_DismissDialogs()
_EnsureAlive("TOC -> Xoa muc luc")
_WriteLog("PASS: TOC -> Xoa muc luc")
_RunAction("TOC", 275, 155, "Xem truoc TOC", 450, 273)

_ClickApp(365, 155)
_EnsureAlive("chuyen tab Copy Style")
_ClickApp(620, 233)
Sleep(800)
_DismissDialogs()
_EnsureAlive("Copy Style -> Refresh source")
_ClickApp(620, 293)
Sleep(800)
_DismissDialogs()
_EnsureAlive("Copy Style -> Refresh target")
_WriteLog("PASS: Copy Style -> Refresh source/target")
_RunAction("Copy Style", 365, 155, "Preview styles", 505, 415)
_RunAction("Copy Style", 365, 155, "Backup hotkeys", 185, 465)

_RunAction("Advanced", 455, 155, "Liet ke Heading", 510, 185)
_RunAction("Advanced", 455, 155, "Convert case", 200, 285)

$oDoc.Activate()
$oWord.Selection.EndKey(6)
_ClickApp(545, 155)
_ClickApp(435, 175) ; Insert Date
Sleep(1000)
_DismissDialogs()
_EnsureAlive("Quick Utils -> Insert date")
_WriteLog("PASS: Quick Utils -> Insert date")
_RunAction("Quick Utils", 545, 155, "Thong tin chi tiet", 520, 255)
_RunAction("Quick Utils", 545, 155, "Toggle bold", 360, 285)

_RunAction("Smart Fix", 625, 155, "Phan tich", 40, 253, False)
_ClickApp(625, 155)
_ClickApp(40, 185) ; Fix Hyphenation
Sleep(900)
_DismissDialogs()
_EnsureAlive("Smart Fix -> Fix Hyphenation")
_WriteLog("PASS: Smart Fix -> Fix Hyphenation")
_ClickApp(625, 155)
_ClickApp(40, 217) ; Fix Dashes
Sleep(900)
_DismissDialogs()
_EnsureAlive("Smart Fix -> Fix Dashes")
_WriteLog("PASS: Smart Fix -> Fix Dashes")

$oDoc.Activate()
$oDoc.Paragraphs.Item(2).Range.Select()
_ClickApp(700, 155)
_ClickApp(178, 242) ; Bold
Sleep(1200)
_DismissDialogs()
_EnsureAlive("AI Format -> Convert bold")
_WriteLog("PASS: AI Format -> Convert bold")
_RunAction("AI Format", 700, 155, "Preview", 414, 452)

_WriteLog("PASS: Hoan tat runtime smoke test qua 9 tab.")
Exit 0
