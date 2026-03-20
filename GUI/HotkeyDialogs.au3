; ============================================
; HOTKEYDIALOGS.AU3 - GUI Dialogs cho quan ly Hotkeys
; Cac dialog: Backup/Restore, Edit Inline, Apply Hotkeys
; ============================================

#include-once
#include "..\Shared\Helpers.au3"

; ============================================
; BACKUP & RESTORE DIALOG
; ============================================

; === FUNCTION: Show Backup/Restore Hotkeys Dialog ===
Func _ShowBackupHotkeysDialog()
    Local $sIniFile = _GetHotkeyIniPath()
    Local $sBackupDir = _GetHotkeyBackupDir()
    
    ; Tao thu muc backup neu chua co
    If Not FileExists($sBackupDir) Then DirCreate($sBackupDir)
    
    ; Tao popup window
    Local $hPopup = GUICreate("Sao luu & Khoi phuc Hotkeys", 500, 450, -1, -1, _
        BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU))
    GUISetBkColor(0xF5F5F5)
    
    ; Tieu de
    GUICtrlCreateLabel("QUAN LY SAO LUU PHIM TAT", 15, 10, 470, 25)
    GUICtrlSetFont(-1, 12, 700)
    GUICtrlSetColor(-1, 0x2C3E50)
    
    ; Thong tin file hien tai
    _CreateBackupGroup($hPopup, " File Hotkeys hien tai ", 15, 40, 470, 70)
    Local $sCurrentInfo = _GetHotkeyFileInfo($sIniFile)
    GUICtrlCreateLabel($sCurrentInfo, 25, 60, 450, 40)
    GUICtrlSetFont(-1, 9)
    
    ; Danh sach backup
    _CreateBackupGroup($hPopup, " Danh sach ban sao luu ", 15, 115, 470, 200)
    Local $listBackups = GUICtrlCreateListView("Ten file|Ngay tao|So hotkeys", 25, 135, 450, 150, _
        BitOR($LVS_REPORT, $LVS_SINGLESEL, $LVS_SHOWSELALWAYS))
    _GUICtrlListView_SetExtendedListViewStyle($listBackups, _
        BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
    _GUICtrlListView_SetColumnWidth($listBackups, 0, 200)
    _GUICtrlListView_SetColumnWidth($listBackups, 1, 130)
    _GUICtrlListView_SetColumnWidth($listBackups, 2, 90)
    
    ; Load danh sach backup
    _LoadBackupList($listBackups, $sBackupDir)
    
    ; Nut chuc nang
    Local $btnBackupNow = GUICtrlCreateButton("Sao luu ngay", 25, 295, 110, 35)
    GUICtrlSetBkColor(-1, 0x27AE60)
    GUICtrlSetFont(-1, 9, 600)
    
    Local $btnRestore = GUICtrlCreateButton("Khoi phuc", 145, 295, 100, 35)
    GUICtrlSetBkColor(-1, 0x3498DB)
    GUICtrlSetFont(-1, 9, 600)
    
    Local $btnDelete = GUICtrlCreateButton("Xoa backup", 255, 295, 100, 35)
    GUICtrlSetBkColor(-1, 0xE74C3C)
    GUICtrlSetFont(-1, 9, 600)
    
    Local $btnRefreshList = GUICtrlCreateButton("Lam moi", 365, 295, 100, 35)
    GUICtrlSetFont(-1, 9, 600)
    
    ; Huong dan
    _CreateBackupGroup($hPopup, " Huong dan ", 15, 340, 470, 60)
    GUICtrlCreateLabel("- Sao luu ngay: Tao ban backup moi voi timestamp", 25, 358, 450, 16)
    GUICtrlCreateLabel("- Khoi phuc: Chon 1 backup trong danh sach roi nhan Khoi phuc", 25, 376, 450, 16)
    
    ; Nut dong
    Local $btnClose = GUICtrlCreateButton("Dong", 390, 410, 95, 30)
    
    GUISetState(@SW_SHOW, $hPopup)
    
    ; Event loop
    While 1
        Local $iMsg = GUIGetMsg()
        
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnClose
                ExitLoop
                
            Case $btnBackupNow
                Local $sResult = _BackupHotkeysNow($sIniFile, $sBackupDir)
                If $sResult <> "" Then
                    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da sao luu thanh cong!" & @CRLF & @CRLF & "File: " & $sResult, 0, $hPopup)
                    _LoadBackupList($listBackups, $sBackupDir)
                Else
                    MsgBox($MB_ICONWARNING, "Loi", "Khong the sao luu! Kiem tra file StyleHotkeys.ini", 0, $hPopup)
                EndIf
                
            Case $btnRestore
                Local $iSelIdx = _GUICtrlListView_GetSelectedIndices($listBackups, False)
                If $iSelIdx = "" Or $iSelIdx = -1 Then
                    MsgBox($MB_ICONWARNING, "Chua chon", "Vui long chon 1 ban backup trong danh sach!", 0, $hPopup)
                    ContinueLoop
                EndIf
                Local $sBackupName = _GUICtrlListView_GetItemText($listBackups, $iSelIdx, 0)
                Local $sBackupPath = $sBackupDir & "\" & $sBackupName
                
                If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Xac nhan", _
                    "Khoi phuc tu ban backup:" & @CRLF & $sBackupName & @CRLF & @CRLF & _
                    "File hotkeys hien tai se bi ghi de. Tiep tuc?", 0, $hPopup) = $IDYES Then
                    
                    If _RestoreHotkeysFromBackup($sBackupPath, $sIniFile) Then
                        MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da khoi phuc hotkeys thanh cong!", 0, $hPopup)
                    Else
                        MsgBox($MB_ICONERROR, "Loi", "Khong the khoi phuc! File backup co the bi hong.", 0, $hPopup)
                    EndIf
                EndIf
                
            Case $btnDelete
                Local $iDelIdx = _GUICtrlListView_GetSelectedIndices($listBackups, False)
                If $iDelIdx = "" Or $iDelIdx = -1 Then
                    MsgBox($MB_ICONWARNING, "Chua chon", "Vui long chon 1 ban backup de xoa!", 0, $hPopup)
                    ContinueLoop
                EndIf
                Local $sDelName = _GUICtrlListView_GetItemText($listBackups, $iDelIdx, 0)
                Local $sDelPath = $sBackupDir & "\" & $sDelName
                
                If MsgBox($MB_YESNO + $MB_ICONWARNING, "Xac nhan xoa", _
                    "Xoa ban backup:" & @CRLF & $sDelName & @CRLF & @CRLF & _
                    "Hanh dong nay khong the hoan tac!", 0, $hPopup) = $IDYES Then
                    
                    If FileDelete($sDelPath) Then
                        MsgBox($MB_ICONINFORMATION, "Da xoa", "Da xoa ban backup!", 0, $hPopup)
                        _LoadBackupList($listBackups, $sBackupDir)
                    Else
                        MsgBox($MB_ICONERROR, "Loi", "Khong the xoa file!", 0, $hPopup)
                    EndIf
                EndIf
                
            Case $btnRefreshList
                _LoadBackupList($listBackups, $sBackupDir)
        EndSwitch
    WEnd
    
    GUIDelete($hPopup)
EndFunc

; Helper: Tao group cho dialog backup
Func _CreateBackupGroup($hParent, $sTitle, $x, $y, $w, $h)
    GUICtrlCreateGroup($sTitle, $x, $y, $w, $h)
    GUICtrlSetFont(-1, 9, 600)
EndFunc

; Helper: Load danh sach backup vao ListView
Func _LoadBackupList($hListView, $sBackupDir)
    _GUICtrlListView_DeleteAllItems($hListView)
    
    If Not FileExists($sBackupDir) Then Return
    
    Local $aFiles = _FileListToArray($sBackupDir, "StyleHotkeys_*.ini", 1)
    If @error Or Not IsArray($aFiles) Then Return
    
    For $i = 1 To $aFiles[0]
        Local $sFileName = $aFiles[$i]
        Local $sFilePath = $sBackupDir & "\" & $sFileName
        
        ; Lay ngay tao
        Local $aTime = FileGetTime($sFilePath, 1)
        Local $sTime = ""
        If IsArray($aTime) Then
            $sTime = $aTime[2] & "/" & $aTime[1] & "/" & $aTime[0] & " " & $aTime[3] & ":" & $aTime[4]
        EndIf
        
        ; Dem so hotkeys trong file backup
        Local $aHotkeys = IniReadSection($sFilePath, "Hotkeys")
        Local $iCount = 0
        If IsArray($aHotkeys) Then $iCount = $aHotkeys[0][0]
        
        GUICtrlCreateListViewItem($sFileName & "|" & $sTime & "|" & $iCount, $hListView)
    Next
EndFunc

; ============================================
; INLINE HOTKEY EDITING DIALOG
; ============================================

; === FUNCTION: Edit Hotkey Inline (Nhap truc tiep tren ListView) ===
Func _EditHotkeyInline($hListView, $iRowIndex, ByRef $aAllStyles, $sIniFile, $hParentGUI)
    If $iRowIndex < 0 Or $iRowIndex >= UBound($aAllStyles) Then Return False
    
    Local $sStyleName = $aAllStyles[$iRowIndex][0]
    Local $sCurrentHotkey = $aAllStyles[$iRowIndex][3]
    
    ; Tao dialog nhap hotkey
    Local $hInputDlg = GUICreate("Nhap phim tat cho: " & $sStyleName, 450, 200, -1, -1, _
        BitOR($WS_POPUP, $WS_CAPTION, $WS_SYSMENU), -1, $hParentGUI)
    GUISetBkColor(0xF5F5F5)
    
    GUICtrlCreateLabel("Style: " & $sStyleName, 15, 15, 420, 20)
    GUICtrlSetFont(-1, 9, 600)
    
    GUICtrlCreateLabel("Nhap phim tat (VD: Ctrl+1, Alt+H, Ctrl+Shift+B):", 15, 45, 420, 20)
    
    Local $inputHotkey = GUICtrlCreateInput($sCurrentHotkey, 15, 70, 320, 25)
    GUICtrlSetFont(-1, 10)
    
    Local $btnCapture = GUICtrlCreateButton("BAT PHIM", 345, 68, 90, 28)
    GUICtrlSetBkColor(-1, 0x3498DB)
    GUICtrlSetFont(-1, 9, 600)
    GUICtrlSetTip(-1, "Nhan de bat phim tat tu ban phim")
    
    ; Huong dan
    GUICtrlCreateGroup(" Huong dan ", 15, 105, 420, 50)
    GUICtrlCreateLabel("- Nhap truc tiep hoac nhan 'BAT PHIM' roi nhan to hop phim", 25, 123, 400, 16)
    GUICtrlCreateLabel("- Phai co it nhat 1 modifier: Ctrl, Alt, hoac Shift", 25, 139, 400, 16)
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    Local $btnOK = GUICtrlCreateButton("OK", 250, 165, 90, 28, $BS_DEFPUSHBUTTON)
    GUICtrlSetBkColor(-1, 0x27AE60)
    GUICtrlSetFont(-1, 9, 600)
    
    Local $btnCancel = GUICtrlCreateButton("Huy", 350, 165, 80, 28)
    
    GUISetState(@SW_SHOW, $hInputDlg)
    
    Local $bCapturing = False
    Local $sNewHotkey = $sCurrentHotkey
    
    ; Event loop
    While 1
        Local $iMsg = GUIGetMsg()
        
        ; Xu ly capture hotkey
        If $bCapturing Then
            Local $sCaptured = _CaptureHotkeyPress()
            If $sCaptured <> "" Then
                GUICtrlSetData($inputHotkey, $sCaptured)
                $bCapturing = False
                GUICtrlSetData($btnCapture, "BAT PHIM")
                GUICtrlSetBkColor($btnCapture, 0x3498DB)
            EndIf
            Sleep(10)
        EndIf
        
        Switch $iMsg
            Case $GUI_EVENT_CLOSE, $btnCancel
                GUIDelete($hInputDlg)
                Return False
                
            Case $btnCapture
                If Not $bCapturing Then
                    $bCapturing = True
                    GUICtrlSetData($btnCapture, "Nhan phim...")
                    GUICtrlSetBkColor($btnCapture, 0xE74C3C)
                    GUICtrlSetData($inputHotkey, "")
                Else
                    $bCapturing = False
                    GUICtrlSetData($btnCapture, "BAT PHIM")
                    GUICtrlSetBkColor($btnCapture, 0x3498DB)
                EndIf
                
            Case $btnOK
                $sNewHotkey = StringStripWS(GUICtrlRead($inputHotkey), 3)
                
                ; Validate hotkey format
                If $sNewHotkey <> "" Then
                    If Not _ValidateHotkeyFormat($sNewHotkey) Then
                        MsgBox($MB_ICONWARNING, "Loi dinh dang", _
                            "Phim tat khong hop le!" & @CRLF & @CRLF & _
                            "Dinh dang dung: Ctrl+X, Alt+X, Shift+X, Ctrl+Shift+X..." & @CRLF & _
                            "Trong do X co the la: A-Z, 0-9, F1-F12, hoac ky tu dac biet", 0, $hInputDlg)
                        ContinueLoop
                    EndIf
                EndIf
                
                ; Cap nhat ListView va mang
                _GUICtrlListView_SetItemText($hListView, $iRowIndex, $sNewHotkey, 3)
                $aAllStyles[$iRowIndex][3] = $sNewHotkey
                
                ; Luu vao INI
                If $sNewHotkey <> "" Then
                    IniWrite($sIniFile, "Hotkeys", $sStyleName, $sNewHotkey)
                    _UpdateProgress("Da gan phim tat: " & $sStyleName & " = " & $sNewHotkey)
                Else
                    IniDelete($sIniFile, "Hotkeys", $sStyleName)
                    _UpdateProgress("Da xoa phim tat cua: " & $sStyleName)
                EndIf
                
                GUIDelete($hInputDlg)
                Return True
        EndSwitch
    WEnd
EndFunc

; ============================================
; APPLY HOTKEYS DIALOG
; ============================================

; ============================================
; APPLY HOTKEYS TO NORMAL.DOTM - MAIN FUNCTION
; ============================================

; === FUNCTION: Apply Hotkeys to Normal.dotm (ULTIMATE VERSION) ===
; Tac dung: Luu tat ca hotkeys tu file INI vao Normal.dotm
; Dua tren: tailieuhuongdan.txt - Su dung Word COM API
; Tra ve: So hotkeys da luu thanh cong
Func _ApplyHotkeysToCurrentDoc()
    ; 1. Kiem tra ket noi Word
    If Not _CheckConnection() Then Return

    ; 2. Kiem tra file INI co ton tai
    Local $sIniFile = _GetHotkeyIniPath()
    If Not FileExists($sIniFile) Then
        MsgBox($MB_ICONWARNING, "Chua co hotkey", "Chua co phim tat nao duoc luu!" & @CRLF & _
            "Hay vao 'CHON STYLE DE COPY' de gan phim tat truoc.")
        Return
    EndIf

    ; 3. Doc danh sach hotkeys tu INI
    Local $aHotkeys = IniReadSection($sIniFile, "Hotkeys")
    If @error Or Not IsArray($aHotkeys) Or $aHotkeys[0][0] = 0 Then
        MsgBox($MB_ICONWARNING, "Chua co hotkey", "Chua co phim tat nao duoc luu!")
        Return
    EndIf

    ; 4. Hien thi danh sach hotkeys truoc khi ap dung
    Local $sHotkeyList = "=== DANH SACH PHIM TAT SE LUU VAO NORMAL.DOTM ===" & @CRLF & @CRLF
    $sHotkeyList &= "Tong so: " & $aHotkeys[0][0] & " hotkeys" & @CRLF & @CRLF
    For $i = 1 To $aHotkeys[0][0]
        $sHotkeyList &= "  " & $i & ". " & $aHotkeys[$i][1] & " = " & $aHotkeys[$i][0] & @CRLF
    Next
    $sHotkeyList &= @CRLF & "QUAN TRONG:"
    $sHotkeyList &= @CRLF & "- Phim tat se duoc luu vao Normal.dotm (template toan cuc)"
    $sHotkeyList &= @CRLF & "- Se hoat dong voi MOI document co style nay"
    $sHotkeyList &= @CRLF & "- Khong can ap dung lai cho tung file"

    If MsgBox($MB_YESNO + $MB_ICONQUESTION, "Xac nhan luu vao Normal.dotm", $sHotkeyList & @CRLF & @CRLF & "Ban co muon tiep tuc?") <> $IDYES Then
        Return
    EndIf

    ; 5. Bat dau qua trinh luu
    ConsoleWrite(@CRLF & "========================================" & @CRLF)
    ConsoleWrite("BAT DAU LUU HOTKEYS VAO NORMAL.DOTM" & @CRLF)
    ConsoleWrite("========================================" & @CRLF)
    ConsoleWrite("Tong so hotkeys: " & $aHotkeys[0][0] & @CRLF)
    ConsoleWrite("========================================" & @CRLF & @CRLF)
    
    _UpdateProgress("Dang luu hotkeys vao Normal.dotm...")
    
    ; 6. Luu tung hotkey bang ham _SaveHotkeysToNormalDotm
    Local $iSuccess = 0
    Local $iFailed = 0
    Local $aFailedList[1][2] ; [StyleName, Reason]
    Local $iFailedCount = 0
    
    For $i = 1 To $aHotkeys[0][0]
        Local $sStyleName = $aHotkeys[$i][0]
        Local $sHotkey = $aHotkeys[$i][1]
        
        ; Cap nhat progress
        _UpdateProgress("Dang xu ly " & $i & "/" & $aHotkeys[0][0] & ": " & $sStyleName)
        
        ConsoleWrite("--- Hotkey " & $i & "/" & $aHotkeys[0][0] & " ---" & @CRLF)
        ConsoleWrite("Style: " & $sStyleName & @CRLF)
        ConsoleWrite("Hotkey: " & $sHotkey & @CRLF)
        
        ; Validate hotkey format truoc
        If Not _ValidateHotkeyFormat($sHotkey) Then
            $iFailed += 1
            ReDim $aFailedList[$iFailedCount + 1][2]
            $aFailedList[$iFailedCount][0] = $sStyleName
            $aFailedList[$iFailedCount][1] = "Dinh dang hotkey khong hop le: '" & $sHotkey & "'"
            $iFailedCount += 1
            ConsoleWrite("=> THAT BAI: Dinh dang hotkey khong hop le!" & @CRLF & @CRLF)
            ContinueLoop
        EndIf
        
        ; Goi ham luu vao Normal.dotm
        If _SaveHotkeysToNormalDotm($sStyleName, $sHotkey) Then
            $iSuccess += 1
            ConsoleWrite("=> THANH CONG!" & @CRLF & @CRLF)
        Else
            $iFailed += 1
            ReDim $aFailedList[$iFailedCount + 1][2]
            $aFailedList[$iFailedCount][0] = $sStyleName
            $aFailedList[$iFailedCount][1] = "Style khong ton tai trong document hien tai"
            $iFailedCount += 1
            ConsoleWrite("=> THAT BAI!" & @CRLF & @CRLF)
        EndIf
    Next
    
    ; 7. Tao bao cao ket qua
    ConsoleWrite("========================================" & @CRLF)
    ConsoleWrite("KET QUA LUU HOTKEYS VAO NORMAL.DOTM" & @CRLF)
    ConsoleWrite("========================================" & @CRLF)
    ConsoleWrite("Thanh cong: " & $iSuccess & "/" & $aHotkeys[0][0] & @CRLF)
    ConsoleWrite("That bai: " & $iFailed & "/" & $aHotkeys[0][0] & @CRLF)
    ConsoleWrite("========================================" & @CRLF & @CRLF)
    
    Local $sResult = "=== KET QUA LUU VAO NORMAL.DOTM ===" & @CRLF & @CRLF
    $sResult &= "Thanh cong: " & $iSuccess & "/" & $aHotkeys[0][0] & " hotkeys" & @CRLF
    
    If $iFailed > 0 Then
        $sResult &= "That bai: " & $iFailed & " hotkeys" & @CRLF & @CRLF
        $sResult &= "=== DANH SACH LOI ===" & @CRLF
        For $j = 0 To $iFailedCount - 1
            $sResult &= "  " & ($j+1) & ". " & $aFailedList[$j][0] & @CRLF
            $sResult &= "     => " & $aFailedList[$j][1] & @CRLF
        Next
        
        $sResult &= @CRLF & "=== NGUYEN NHAN CHINH ===" & @CRLF
        $sResult &= "- Cac style tren KHONG TON TAI trong document hien tai" & @CRLF
        $sResult &= "- Ham can style phai co trong document de copy sang Normal.dotm" & @CRLF & @CRLF
        
        $sResult &= "=== GIAI PHAP ===" & @CRLF
        $sResult &= "CACH 1 (Khuyến nghị):" & @CRLF
        $sResult &= "  1. Mo file co cac style do (file nguon ban dau)" & @CRLF
        $sResult &= "  2. Chay lai 'Luu vao Normal.dotm'" & @CRLF & @CRLF
        $sResult &= "CACH 2:" & @CRLF
        $sResult &= "  1. Vao 'CHON STYLE DE COPY'" & @CRLF
        $sResult &= "  2. Chon file nguon co cac style" & @CRLF
        $sResult &= "  3. Gan hotkey cho cac style" & @CRLF
        $sResult &= "  4. Copy sang file dich (se tu dong luu vao Normal.dotm)" & @CRLF & @CRLF
        $sResult &= "CACH 3:" & @CRLF
        $sResult &= "  - Xoa cac hotkey loi khoi file StyleHotkeys.ini" & @CRLF
        $sResult &= "  - Chi giu lai hotkey cho cac style da co trong document" & @CRLF
    EndIf
    
    If $iSuccess > 0 Then
        $sResult &= @CRLF & "=== DANH SACH PHIM TAT DA LUU THANH CONG ===" & @CRLF
        Local $iSuccessCount = 0
        For $i = 1 To $aHotkeys[0][0]
            ; Chi hien thi nhung hotkey thanh cong
            Local $bIsSuccess = True
            For $j = 0 To $iFailedCount - 1
                If $aFailedList[$j][0] = $aHotkeys[$i][0] Then
                    $bIsSuccess = False
                    ExitLoop
                EndIf
            Next
            
            If $bIsSuccess Then
                $iSuccessCount += 1
                $sResult &= "  " & $iSuccessCount & ". Nhan " & $aHotkeys[$i][1] & " de ap dung style '" & $aHotkeys[$i][0] & "'" & @CRLF
            EndIf
        Next
        
        $sResult &= @CRLF & "=== LUU Y QUAN TRONG ===" & @CRLF
        $sResult &= "- Phim tat da duoc luu vao Normal.dotm (template toan cuc)" & @CRLF
        $sResult &= "- Se hoat dong voi MOI document co style nay" & @CRLF
        $sResult &= "- Khong can ap dung lai cho tung file" & @CRLF
        $sResult &= "- Neu mo document moi, phim tat van hoat dong!" & @CRLF
    EndIf

    ; 8. Hien thi ket qua
    _LogPreview($sResult)
    _UpdateProgress("Da luu " & $iSuccess & " hotkeys vao Normal.dotm!")
    
    If $iFailed = 0 Then
        MsgBox($MB_ICONINFORMATION, "Hoan thanh", "Da luu thanh cong " & $iSuccess & " phim tat vao Normal.dotm!" & @CRLF & @CRLF & _
            "Phim tat se hoat dong voi MOI document co style nay!" & @CRLF & _
            "Khong can ap dung lai cho tung file!" & @CRLF & @CRLF & _
            "=> Xem chi tiet trong tab 'PDF Fix' > khung 'Xem truoc / Log' (phia duoi)")
    Else
        Local $sErrorMsg = "Da luu thanh cong: " & $iSuccess & "/" & $aHotkeys[0][0] & " hotkeys" & @CRLF
        $sErrorMsg &= "That bai: " & $iFailed & " hotkeys" & @CRLF & @CRLF
        $sErrorMsg &= "NGUYEN NHAN:" & @CRLF
        $sErrorMsg &= "- Cac style loi KHONG TON TAI trong document hien tai" & @CRLF & @CRLF
        $sErrorMsg &= "GIAI PHAP:" & @CRLF
        $sErrorMsg &= "1. Mo file co cac style do" & @CRLF
        $sErrorMsg &= "2. Chay lai 'Luu vao Normal.dotm'" & @CRLF & @CRLF
        $sErrorMsg &= "=> XEM CHI TIET LOI:" & @CRLF
        $sErrorMsg &= "   Tab 'PDF Fix' > Khung 'Xem truoc / Log'"
        
        MsgBox($MB_ICONWARNING, "Hoan thanh co loi", $sErrorMsg)
    EndIf
EndFunc
