; ============================================
; NORMALDOTMMANAGER.AU3
; Module tich hop Normal.dotm Backup vao Main App
; ============================================

#include-once
#include "..\Tools\NormalDotmBackup.au3"

; ============================================
; PUBLIC FUNCTIONS (goi tu EventLoop)
; ============================================

; Handler: Backup Normal.dotm
Func _OnBackupNormalDotm()
    If Not IsObj($g_oWord) Then
        _UpdateProgress("Loi: Chua ket noi Word!")
        Return
    EndIf
    
    _UpdateProgress("Dang backup Normal.dotm...")
    
    Local $sBackupPath = _BackupNormalDotm($g_oWord)
    If @error Then
        _UpdateProgress("Loi: Backup that bai!")
        MsgBox($MB_ICONERROR, "Loi", "Khong the backup Normal.dotm!")
        Return
    EndIf
    
    _UpdateProgress("Backup thanh cong: " & $sBackupPath)
    
    Local $iResponse = MsgBox($MB_YESNO + $MB_ICONINFORMATION, "Backup thanh cong", _
        "Normal.dotm da duoc backup!" & @CRLF & @CRLF & _
        "Location: " & $sBackupPath & @CRLF & @CRLF & _
        "Mo folder backup?")
    
    If $iResponse = $IDYES Then
        ShellExecute($sBackupPath)
    EndIf
EndFunc

; Handler: Restore Normal.dotm
Func _OnRestoreNormalDotm()
    If Not IsObj($g_oWord) Then
        _UpdateProgress("Loi: Chua ket noi Word!")
        Return
    EndIf
    
    ; Lay danh sach backups
    Local $aBackups = _ListBackups()
    If Not IsArray($aBackups) Or $aBackups[0] = 0 Then
        MsgBox($MB_ICONWARNING, "Canh bao", "Chua co backup nao!" & @CRLF & @CRLF & _
            "Vui long tao backup truoc.")
        Return
    EndIf
    
    ; Tao GUI chon backup
    Local $sSelectedBackup = _ShowBackupSelector($aBackups)
    If $sSelectedBackup = "" Then Return
    
    _UpdateProgress("Dang restore tu: " & $sSelectedBackup)
    
    ; Xac nhan
    Local $iResponse = MsgBox($MB_YESNO + $MB_ICONWARNING, "Xac nhan Restore", _
        "Restore Normal.dotm tu backup:" & @CRLF & @CRLF & _
        $sSelectedBackup & @CRLF & @CRLF & _
        "CANH BAO:" & @CRLF & _
        "- Normal.dotm hien tai se bi ghi de" & @CRLF & _
        "- Word se bi dong" & @CRLF & _
        "- Backup hien tai se duoc luu truoc" & @CRLF & @CRLF & _
        "Tiep tuc?")
    
    If $iResponse = $IDNO Then
        _UpdateProgress("User huy restore")
        Return
    EndIf
    
    ; Restore
    Local $bResult = _RestoreNormalDotm($g_oWord, $sSelectedBackup)
    If @error Or Not $bResult Then
        _UpdateProgress("Loi: Restore that bai!")
    Else
        _UpdateProgress("Restore thanh cong!")
    EndIf
EndFunc

; Handler: Scan Normal.dotm
Func _OnScanNormalDotm()
    If Not IsObj($g_oWord) Then
        _UpdateProgress("Loi: Chua ket noi Word!")
        Return
    EndIf
    
    _UpdateProgress("Dang scan Normal.dotm...")
    
    Local $aStyles = _ScanNormalDotmStyles($g_oWord)
    If @error Then
        _UpdateProgress("Loi: Khong the scan styles!")
        Return
    EndIf
    
    ; Dem loai styles
    Local $iParagraph = 0, $iCharacter = 0, $iTable = 0, $iList = 0, $iBuiltIn = 0, $iCustom = 0
    For $i = 1 To $aStyles[0][0]
        Switch $aStyles[$i][2]
            Case "Paragraph"
                $iParagraph += 1
            Case "Character"
                $iCharacter += 1
            Case "Table"
                $iTable += 1
            Case "List"
                $iList += 1
        EndSwitch
        
        If $aStyles[$i][5] = "Yes" Then
            $iBuiltIn += 1
        Else
            $iCustom += 1
        EndIf
    Next
    
    _UpdateProgress("Scan hoan tat: " & $aStyles[0][0] & " styles")
    
    Local $sMsg = "=== NORMAL.DOTM SCAN RESULT ===" & @CRLF & @CRLF
    $sMsg &= "Total Styles: " & $aStyles[0][0] & @CRLF & @CRLF
    $sMsg &= "By Type:" & @CRLF
    $sMsg &= "  Paragraph: " & $iParagraph & @CRLF
    $sMsg &= "  Character: " & $iCharacter & @CRLF
    $sMsg &= "  Table: " & $iTable & @CRLF
    $sMsg &= "  List: " & $iList & @CRLF & @CRLF
    $sMsg &= "By Origin:" & @CRLF
    $sMsg &= "  Built-in: " & $iBuiltIn & @CRLF
    $sMsg &= "  Custom: " & $iCustom & @CRLF & @CRLF
    $sMsg &= "Export to file?"
    
    Local $iResponse = MsgBox($MB_YESNOCANCEL + $MB_ICONINFORMATION, "Scan Result", $sMsg)
    
    If $iResponse = $IDYES Then
        ; Export TXT
        Local $sFile = FileSaveDialog("Export to TXT", @ScriptDir, "Text files (*.txt)", 16, "normal_dotm_styles.txt")
        If Not @error Then
            _ExportStylesToFile($aStyles, $sFile)
            ShellExecute($sFile)
        EndIf
    ElseIf $iResponse = $IDNO Then
        ; Export CSV
        Local $sFile = FileSaveDialog("Export to CSV", @ScriptDir, "CSV files (*.csv)", 16, "normal_dotm_styles.csv")
        If Not @error Then
            _ExportStylesToCSV($aStyles, $sFile)
            ShellExecute($sFile)
        EndIf
    EndIf
EndFunc

; Handler: Mo Backup Folder
Func _OnOpenBackupFolder()
    _EnsureBackupFolder()
    ShellExecute($BACKUP_FOLDER)
    _UpdateProgress("Mo folder: " & $BACKUP_FOLDER)
EndFunc

; ============================================
; PRIVATE HELPERS
; ============================================

; Hien thi GUI chon backup
Func _ShowBackupSelector($aBackups)
    Local $hGUI = GUICreate("Chon Backup", 500, 400)
    
    GUICtrlCreateLabel("Chon backup de restore:", 10, 10, 480, 20)
    
    Local $listBackups = GUICtrlCreateListView("Backup Name|Date|Styles", 10, 40, 480, 300, _
        $LVS_REPORT + $LVS_SINGLESEL, $LVS_EX_FULLROWSELECT + $LVS_EX_GRIDLINES)
    _GUICtrlListView_SetColumnWidth($listBackups, 0, 250)
    _GUICtrlListView_SetColumnWidth($listBackups, 1, 120)
    _GUICtrlListView_SetColumnWidth($listBackups, 2, 100)
    
    ; Populate list
    For $i = 1 To $aBackups[0]
        Local $sBackupPath = $BACKUP_FOLDER & "\" & $aBackups[$i]
        Local $sManifest = $sBackupPath & "\MANIFEST.txt"
        
        ; Doc so styles
        Local $sStyleCount = "N/A"
        If FileExists($sManifest) Then
            Local $hFile = FileOpen($sManifest, 0)
            While True
                Local $sLine = FileReadLine($hFile)
                If @error Then ExitLoop
                If StringInStr($sLine, "styles)") Then
                    Local $aMatch = StringRegExp($sLine, "\((\d+) styles\)", 1)
                    If Not @error Then $sStyleCount = $aMatch[0]
                    ExitLoop
                EndIf
            WEnd
            FileClose($hFile)
        EndIf
        
        ; Lay date
        Local $sDate = StringRegExpReplace($aBackups[$i], "NormalDotm_(\d{8})_(\d{6})", "$1")
        
        _GUICtrlListView_AddItem($listBackups, $aBackups[$i])
        _GUICtrlListView_AddSubItem($listBackups, $i - 1, $sDate, 1)
        _GUICtrlListView_AddSubItem($listBackups, $i - 1, $sStyleCount, 2)
    Next
    
    Local $btnOK = GUICtrlCreateButton("OK", 300, 350, 90, 30)
    Local $btnCancel = GUICtrlCreateButton("Cancel", 400, 350, 90, 30)
    
    GUISetState(@SW_SHOW)
    
    Local $sResult = ""
    While True
        Local $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $GUI_EVENT_CLOSE, $btnCancel
                ExitLoop
            Case $btnOK
                Local $iSelected = _GUICtrlListView_GetSelectedIndices($listBackups)
                If $iSelected <> "" Then
                    $sResult = _GUICtrlListView_GetItemText($listBackups, $iSelected)
                    ExitLoop
                Else
                    MsgBox($MB_ICONWARNING, "Canh bao", "Vui long chon backup!")
                EndIf
        EndSwitch
    WEnd
    
    GUIDelete($hGUI)
    Return $sResult
EndFunc

; ============================================
; QUICK ACTIONS (goi nhanh tu menu)
; ============================================

; Quick backup (1 click, khong hoi)
Func _QuickBackupNormalDotm()
    If Not IsObj($g_oWord) Then Return
    
    _UpdateProgress("Quick backup Normal.dotm...")
    Local $sBackupPath = _BackupNormalDotm($g_oWord)
    
    If @error Then
        _UpdateProgress("Loi: Backup that bai!")
    Else
        _UpdateProgress("Backup thanh cong: " & $sBackupPath)
        ; Khong hien MsgBox, chi log
    EndIf
EndFunc
