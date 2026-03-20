; Deep runtime assertions cho cac module chinh

Global Const $APP_EXE = @ScriptDir & "\..\Main_compiled.exe"
Global Const $APP_TITLE = "PDF to Word Fixer Pro v6.1"
Global Const $LOG_PATH = @ScriptDir & "\DeepRuntimeAssertions.log"
Global Const $ARTIFACT_DIR = @ScriptDir & "\..\Tests\Artifacts"
Global $g_hWnd = ""
Global $g_aPos = 0

DirCreate(@ScriptDir & "\..\Tests\Logs")
DirCreate($ARTIFACT_DIR)
FileDelete($LOG_PATH)

Func _WriteLog($sMsg)
    Local $hFile = FileOpen($LOG_PATH, 1 + 8)
    If $hFile <> -1 Then
        FileWriteLine($hFile, $sMsg)
        FileClose($hFile)
    EndIf
EndFunc

Func _Fail($sMsg, $iCode)
    _WriteLog("FAIL: " & $sMsg)
    Exit $iCode
EndFunc

Func _DismissDialogs()
    For $iLoop = 1 To 8
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
EndFunc

Func _EnsureAlive($hWnd, $sStep)
    If Not WinExists($hWnd) Then _Fail("App dong/crash tai buoc: " & $sStep, 100)
EndFunc

If Not FileExists($APP_EXE) Then _Fail("Khong tim thay file build.", 1)

Local $oWord = ObjGet("", "Word.Application")
If @error Or Not IsObj($oWord) Then $oWord = ObjCreate("Word.Application")
If Not IsObj($oWord) Then _Fail("Khong mo duoc Word.", 2)

$oWord.Visible = True

Local $sDoc1 = $ARTIFACT_DIR & "\DeepSmoke_Source.docx"
Local $sDoc2 = $ARTIFACT_DIR & "\DeepSmoke_Target.docx"

If FileExists($sDoc1) Then FileDelete($sDoc1)
If FileExists($sDoc2) Then FileDelete($sDoc2)

Local $oDoc1 = $oWord.Documents.Add()
$oDoc1.Range(0, 0).Text = _
    "Heading Test" & @CRLF & _
    "This is **BoldSample** text with markdown." & @CRLF & _
    "DashSample A" & ChrW(8212) & "B and line break hyphen test vi-" & @CRLF & "du." & @CRLF & _
    "Normal paragraph for formatting." & @CRLF
$oDoc1.Paragraphs.Item(1).Range.Style = "Heading 1"
$oDoc1.Tables.Add($oDoc1.Range($oDoc1.Content.End - 1, $oDoc1.Content.End - 1), 2, 2)
$oDoc1.SaveAs($sDoc1)

Local $oDoc2 = $oWord.Documents.Add()
$oDoc2.Range(0, 0).Text = "Target document for copy style." & @CRLF
$oDoc2.SaveAs($sDoc2)

$oDoc1.Activate()
$oWord.Selection.HomeKey(6) ; wdStory

Local $hWnd = WinGetHandle($APP_TITLE)
If $hWnd = "" Then
    Run('"' & $APP_EXE & '"')
    If Not WinWait($APP_TITLE, "", 10) Then _Fail("App khong mo duoc.", 3)
    $hWnd = WinGetHandle($APP_TITLE)
EndIf
If $hWnd = "" Then _Fail("Khong lay duoc handle app.", 4)
 $g_hWnd = $hWnd

WinActivate($hWnd)
If Not WinWaitActive($hWnd, "", 5) Then _Fail("Khong kich hoat duoc app.", 5)
Local $aPos = WinGetPos($hWnd)
If @error Or Not IsArray($aPos) Then _Fail("Khong lay duoc vi tri app.", 6)
$g_aPos = $aPos

Func _ClickApp($iX, $iY)
    MouseClick("left", $g_aPos[0] + $iX, $g_aPos[1] + $iY, 1, 0)
    Sleep(700)
EndFunc

Func _RunAndDismiss($hWnd, $sStep, $iTabX, $iTabY, $iBtnX, $iBtnY)
    _ClickApp($iTabX, $iTabY)
    _EnsureAlive($hWnd, "chuyen tab " & $sStep)
    _ClickApp($iBtnX, $iBtnY)
    Sleep(1000)
    _DismissDialogs()
    WinActivate($g_hWnd)
    Sleep(200)
    _EnsureAlive($hWnd, $sStep)
    _WriteLog("PASS: " & $sStep)
EndFunc

_WriteLog("START: Deep runtime assertions")

; Connection
_ClickApp(508, 60)
_DismissDialogs()
_ClickApp(35, 60)
_DismissDialogs()
_EnsureAlive($hWnd, "connect")
_WriteLog("PASS: Connection -> Refresh + Auto connect")

; PDF Fix: help va fix selected
_RunAndDismiss($hWnd, "PDF Fix -> Help", 60, 155, 655, 380)
$oDoc1.Activate()
$oDoc1.Paragraphs.Item(3).Range.Select()
_RunAndDismiss($hWnd, "PDF Fix -> Fix selected", 60, 155, 85, 380)

; Format: preset va apply selection
_ClickApp(135, 155)
_ClickApp(365, 288) ; Preset VN
$oDoc1.Activate()
$oDoc1.Paragraphs.Item(4).Range.Select()
_ClickApp(240, 288) ; Apply selection
Sleep(1200)
_DismissDialogs()
_EnsureAlive($hWnd, "Format -> Apply selection")
Local $oFmtRange = $oDoc1.Paragraphs.Item(4).Range
If $oFmtRange.Font.Name <> "Times New Roman" Then _Fail("Format khong ap dung font mong doi.", 20)
_WriteLog("PASS: Format -> Preset VN + Apply selection")

; Tools: word count + stats
_RunAndDismiss($hWnd, "Tools -> Word count", 205, 155, 115, 412)
_RunAndDismiss($hWnd, "Tools -> Show stats", 205, 155, 535, 412)

; TOC: create/update/delete + assert
$oDoc1.Activate()
_ClickApp(275, 155)
_ClickApp(115, 202) ; Tao muc luc
Sleep(1200)
_DismissDialogs()
_EnsureAlive($hWnd, "TOC -> Create")
If $oDoc1.TablesOfContents.Count < 1 Then _Fail("TOC khong duoc tao.", 30)
_WriteLog("PASS: TOC -> Create")

_ClickApp(245, 202) ; Cap nhat
Sleep(800)
_DismissDialogs()
If $oDoc1.TablesOfContents.Count < 1 Then _Fail("TOC bi mat sau khi update.", 31)
_WriteLog("PASS: TOC -> Update")

_ClickApp(360, 202) ; Xoa
Sleep(800)
_DismissDialogs()
If $oDoc1.TablesOfContents.Count <> 0 Then _Fail("TOC khong bi xoa.", 32)
_WriteLog("PASS: TOC -> Delete")

_RunAndDismiss($hWnd, "TOC -> Preview", 275, 155, 495, 287)

; Copy Style: refresh + preview + backup dialog
$oDoc1.Activate()
_RunAndDismiss($hWnd, "Copy Style -> Refresh source", 365, 155, 670, 247)
_RunAndDismiss($hWnd, "Copy Style -> Refresh target", 365, 155, 670, 307)
_RunAndDismiss($hWnd, "Copy Style -> Preview styles", 365, 155, 550, 435)
_RunAndDismiss($hWnd, "Copy Style -> Backup dialog", 365, 155, 255, 482)

; Advanced: list headings
_RunAndDismiss($hWnd, "Advanced -> List headings", 455, 155, 570, 200)

; Quick Utils: insert date + show info
$oDoc1.Activate()
$oWord.Selection.EndKey(6)
Local $sBeforeDate = $oDoc1.Content.Text
_ClickApp(545, 155)
_ClickApp(450, 175) ; Insert date
Sleep(1000)
_DismissDialogs()
If $oDoc1.Content.Text = $sBeforeDate Then _Fail("Quick Utils khong chen ngay vao tai lieu.", 40)
_WriteLog("PASS: Quick Utils -> Insert date")
_RunAndDismiss($hWnd, "Quick Utils -> Show doc info", 545, 155, 620, 270)

; Smart Fix: analyze + fix dashes
_RunAndDismiss($hWnd, "Smart Fix -> Analyze", 625, 155, 90, 268)
If StringInStr($oDoc1.Content.Text, ChrW(8212)) = 0 Then
    $oDoc1.Range($oDoc1.Content.End - 1, $oDoc1.Content.End - 1).InsertBefore("DashAgain A" & ChrW(8212) & "B" & @CRLF)
EndIf
_ClickApp(625, 155)
_ClickApp(90, 355) ; Fix Hyphenation
Sleep(800)
_DismissDialogs()
_EnsureAlive($hWnd, "Smart Fix -> Fix Hyphenation")
If StringInStr($oDoc1.Content.Text, "vi-" & @CRLF & "du") <> 0 Then _Fail("Smart Fix khong sua hyphenation.", 50)
_WriteLog("PASS: Smart Fix -> Fix Hyphenation")

_ClickApp(625, 155)
_ClickApp(145, 387) ; Fix Dashes (button lower row centered on wide button)
Sleep(800)
_DismissDialogs()
_EnsureAlive($hWnd, "Smart Fix -> Fix Dashes")
If StringInStr($oDoc1.Content.Text, ChrW(8212)) <> 0 Or StringInStr($oDoc1.Content.Text, ChrW(8211)) <> 0 Then
    _Fail("Smart Fix khong sua em/en dash.", 51)
EndIf
_WriteLog("PASS: Smart Fix -> Fix Dashes")

; AI Format: convert bold assertion + preview
$oDoc1.Activate()
Local $oBoldPara = $oDoc1.Paragraphs.Item(2).Range
$oBoldPara.Select()
_ClickApp(700, 155)
_ClickApp(223, 257) ; **Bold**
Sleep(1200)
_DismissDialogs()
_EnsureAlive($hWnd, "AI Format -> Convert bold")
If StringInStr($oDoc1.Paragraphs.Item(2).Range.Text, "**") <> 0 Then _Fail("AI Format khong xoa duoc bold markdown markers.", 60)
_WriteLog("PASS: AI Format -> Convert bold")
_RunAndDismiss($hWnd, "AI Format -> Preview", 700, 155, 454, 467)

_WriteLog("PASS: Hoan tat deep runtime assertions.")
Exit 0
