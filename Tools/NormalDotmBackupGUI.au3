; ============================================
; NORMALDOTMBACKUPGUI.AU3
; GUI Tool: Backup & Restore Normal.dotm
; ============================================

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <MsgBoxConstants.au3>
#include "NormalDotmBackup.au3"

; === GLOBALS ===
Global $g_oWord = 0
Global $g_hGUI = 0
Global $g_listBackups = 0
Global $g_editLog = 0
Global $g_lblStatus = 0

; === MAIN ===
_Main()

Func _Main()
    ; Tao GUI
    $g_hGUI = GUICreate("Normal.dotm Backup Tool v1.0", 800, 600)
    
    ; Status bar
    $g_lblStatus = GUICtrlCreateLabel("Status: Chua ket noi Word", 10, 10, 780, 20)
    GUICtrlSetColor(-1, 0xFF0000)
    
    ; Group: Actions
    GUICtrlCreateGroup("Actions", 10, 40, 380, 200)
    Local $btnConnect = GUICtrlCreateButton("1. Ket noi Word", 20, 60, 150, 30)
    Local $btnScan = GUICtrlCreateButton("2. Quick Scan", 20, 100, 150, 30)
    Local $btnBackup = GUICtrlCreateButton("3. Full Backup", 20, 140, 150, 30)
    Local $btnRestore = GUICtrlCreateButton("4. Restore", 20, 180, 150, 30)
    
    Local $btnExportTxt = GUICtrlCreateButton("Export TXT", 200, 60, 150, 30)
    Local $btnExportCSV = GUICtrlCreateButton("Export CSV", 200, 100, 150, 30)
    Local $btnOpenFolder = GUICtrlCreateButton("Mo Backup Folder", 200, 140, 150, 30)
    Local $btnRefresh = GUICtrlCreateButton("Refresh List", 200, 180, 150, 30)
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    ; Group: Backups List
    GUICtrlCreateGroup("Backups", 400, 40, 390, 200)
    $g_listBackups = GUICtrlCreateListView("Backup Name|Date|Styles", 410, 60, 370, 170, _
        $LVS_REPORT + $LVS_SINGLESEL, $LVS_EX_FULLROWSELECT + $LVS_EX_GRIDLINES)
    _GUICtrlListView_SetColumnWidth($g_listBackups, 0, 200)
    _GUICtrlListView_SetColumnWidth($g_listBackups, 1, 100)
    _GUICtrlListView_SetColumnWidth($g_listBackups, 2, 60)
    GUICtrlCreateGroup("", -99, -99, 1, 1)
    
    ; Log
    GUICtrlCreateLabel("Log:", 10, 250, 780, 20)
    $g_editLog = GUICtrlCreateEdit("", 10, 270, 780, 300, $ES_READONLY + $WS_VSCROLL + $ES_AUTOVSCROLL)
    GUICtrlSetFont(-1, 9, 400, 0, "Consolas")
    
    ; Footer
    Local $btnHelp = GUICtrlCreateButton("Help", 650, 575, 70, 25)
    Local $btnExit = GUICtrlCreateButton("Exit", 725, 575, 70, 25)
    
    GUISetState(@SW_SHOW)
    
    _Log("=== NORMAL.DOTM BACKUP TOOL ===")
    _Log("Vui long ket noi Word de bat dau")
    _Log("")
    
    ; Event loop
    While True
        Local $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $GUI_EVENT_CLOSE, $btnExit
                ExitLoop
                
            Case $btnConnect
                _OnConnect()
                
            Case $btnScan
                _OnScan()
                
            Case $btnBackup
                _OnBackup()
                
            Case $btnRestore
                _OnRestore()
                
            Case $btnExportTxt
                _OnExportTxt()
                
            Case $btnExportCSV
                _OnExportCSV()
                
            Case $btnOpenFolder
                _OnOpenFolder()
                
            Case $btnRefresh
                _OnRefreshList()
                
            Case $btnHelp
                _OnHelp()
        EndSwitch
    WEnd
    
    GUIDelete()
EndFunc

; ============================================
; EVENT HANDLERS
; ============================================

Func _OnConnect()
    _Log("[ACTION] Ket noi Word...")
    
    $g_oWord = ObjGet("", "Word.Application")
    If @error Or Not IsObj($g_oWord) Then
        _Log("[ERROR] Khong the ket noi Word!")
        _Log("Vui long mo Word truoc!")
        GUICtrlSetData($g_lblStatus, "Status: Loi - Khong the ket noi Word")
        GUICtrlSetColor($g_lblStatus, 0xFF0000)
        MsgBox($MB_ICONERROR, "Loi", "Vui long mo Word truoc!")
        Return
    EndIf
    
    _Log("[SUCCESS] Da ket noi Word")
    _Log("Normal.dotm path: " & $g_oWord.NormalTemplate.FullName)
    GUICtrlSetData($g_lblStatus, "Status: Da ket noi Word")
    GUICtrlSetColor($g_lblStatus, 0x00AA00)
    
    ; Load backups list
    _OnRefreshList()
EndFunc

Func _OnScan()
    If Not IsObj($g_oWord) Then
        MsgBox($MB_ICONWARNING, "Canh bao", "Vui long ket noi Word truoc!")
        Return
    EndIf
    
    _Log("")
    _Log("[ACTION] Quick Scan...")
    
    Local $aStyles = _ScanNormalDotmStyles($g_oWord)
    If @error Then
        _Log("[ERROR] Khong the quet styles!")
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
    
    _Log("[RESULT] Total Styles: " & $aStyles[0][0])
    _Log("  - Paragraph: " & $iParagraph)
    _Log("  - Character: " & $iCharacter)
    _Log("  - Table: " & $iTable)
    _Log("  - List: " & $iList)
    _Log("  - Built-in: " & $iBuiltIn)
    _Log("  - Custom: " & $iCustom)
    
    Local $sMsg = "=== SCAN RESULT ===" & @CRLF & @CRLF
    $sMsg &= "Total Styles: " & $aStyles[0][0] & @CRLF & @CRLF
    $sMsg &= "By Type:" & @CRLF
    $sMsg &= "  Paragraph: " & $iParagraph & @CRLF
    $sMsg &= "  Character: " & $iCharacter & @CRLF
    $sMsg &= "  Table: " & $iTable & @CRLF
    $sMsg &= "  List: " & $iList & @CRLF & @CRLF
    $sMsg &= "By Origin:" & @CRLF
    $sMsg &= "  Built-in: " & $iBuiltIn & @CRLF
    $sMsg &= "  Custom: " & $iCustom
    
    MsgBox($MB_ICONINFORMATION, "Scan Result", $sMsg)
EndFunc

Func _OnBackup()
    If Not IsObj($g_oWord) Then
        MsgBox($MB_ICONWARNING, "Canh bao", "Vui long ket noi Word truoc!")
        Return
    EndIf
    
    _Log("")
    _Log("[ACTION] Full Backup...")
    
    Local $sBackupPath = _BackupNormalDotm($g_oWord)
    If @error Then
        _Log("[ERROR] Backup that bai!")
        MsgBox($MB_ICONERROR, "Loi", "Backup that bai!")
        Return
    EndIf
    
    _Log("[SUCCESS] Backup thanh cong!")
    _Log("Location: " & $sBackupPath)
    
    MsgBox($MB_ICONINFORMATION, "Backup thanh cong", _
        "Normal.dotm da duoc backup!" & @CRLF & @CRLF & _
        "Location: " & $sBackupPath)
    
    ; Refresh list
    _OnRefreshList()
    
    ; Mo folder
    ShellExecute($sBackupPath)
EndFunc

Func _OnRestore()
    If Not IsObj($g_oWord) Then
        MsgBox($MB_ICONWARNING, "Canh bao", "Vui long ket noi Word truoc!")
        Return
    EndIf
    
    ; Lay backup duoc chon
    Local $iSelected = _GUICtrlListView_GetSelectedIndices($g_listBackups)
    If $iSelected = "" Then
        MsgBox($MB_ICONWARNING, "Canh bao", "Vui long chon backup de restore!")
        Return
    EndIf
    
    Local $sBackupName = _GUICtrlListView_GetItemText($g_listBackups, $iSelected)
    
    _Log("")
    _Log("[ACTION] Restore from: " & $sBackupName)
    
    Local $iResponse = MsgBox($MB_YESNO + $MB_ICONWARNING, "Xac nhan Restore", _
        "Restore Normal.dotm tu backup:" & @CRLF & @CRLF & _
        $sBackupName & @CRLF & @CRLF & _
        "Canh bao: Normal.dotm hien tai se bi ghi de!" & @CRLF & _
        "(Backup hien tai se duoc luu truoc khi restore)" & @CRLF & @CRLF & _
        "Tiep tuc?")
    
    If $iResponse = $IDNO Then
        _Log("[INFO] User huy restore")
        Return
    EndIf
    
    Local $bResult = _RestoreNormalDotm($g_oWord, $sBackupName)
    If @error Or Not $bResult Then
        _Log("[ERROR] Restore that bai!")
    Else
        _Log("[SUCCESS] Restore thanh cong!")
    EndIf
EndFunc

Func _OnExportTxt()
    If Not IsObj($g_oWord) Then
        MsgBox($MB_ICONWARNING, "Canh bao", "Vui long ket noi Word truoc!")
        Return
    EndIf
    
    _Log("")
    _Log("[ACTION] Export to TXT...")
    
    Local $aStyles = _ScanNormalDotmStyles($g_oWord)
    If @error Then
        _Log("[ERROR] Khong the quet styles!")
        Return
    EndIf
    
    Local $sFile = FileSaveDialog("Export to TXT", @ScriptDir, "Text files (*.txt)", 16, "styles_export.txt")
    If @error Then Return
    
    _ExportStylesToFile($aStyles, $sFile)
    _Log("[SUCCESS] Export thanh cong: " & $sFile)
    
    ShellExecute($sFile)
EndFunc

Func _OnExportCSV()
    If Not IsObj($g_oWord) Then
        MsgBox($MB_ICONWARNING, "Canh bao", "Vui long ket noi Word truoc!")
        Return
    EndIf
    
    _Log("")
    _Log("[ACTION] Export to CSV...")
    
    Local $aStyles = _ScanNormalDotmStyles($g_oWord)
    If @error Then
        _Log("[ERROR] Khong the quet styles!")
        Return
    EndIf
    
    Local $sFile = FileSaveDialog("Export to CSV", @ScriptDir, "CSV files (*.csv)", 16, "styles_export.csv")
    If @error Then Return
    
    _ExportStylesToCSV($aStyles, $sFile)
    _Log("[SUCCESS] Export thanh cong: " & $sFile)
    
    ShellExecute($sFile)
EndFunc

Func _OnOpenFolder()
    _EnsureBackupFolder()
    ShellExecute($BACKUP_FOLDER)
    _Log("[INFO] Mo folder: " & $BACKUP_FOLDER)
EndFunc

Func _OnRefreshList()
    _Log("[INFO] Refresh backups list...")
    
    _GUICtrlListView_DeleteAllItems($g_listBackups)
    
    Local $aBackups = _ListBackups()
    If Not IsArray($aBackups) Then
        _Log("[INFO] Chua co backup nao")
        Return
    EndIf
    
    For $i = 1 To $aBackups[0]
        Local $sBackupPath = $BACKUP_FOLDER & "\" & $aBackups[$i]
        Local $sManifest = $sBackupPath & "\MANIFEST.txt"
        
        ; Doc so styles tu manifest
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
        
        ; Lay date tu ten folder
        Local $sDate = StringRegExpReplace($aBackups[$i], "NormalDotm_(\d{8})_(\d{6})", "$1")
        
        _GUICtrlListView_AddItem($g_listBackups, $aBackups[$i])
        _GUICtrlListView_AddSubItem($g_listBackups, $i - 1, $sDate, 1)
        _GUICtrlListView_AddSubItem($g_listBackups, $i - 1, $sStyleCount, 2)
    Next
    
    _Log("[SUCCESS] Load " & $aBackups[0] & " backup(s)")
EndFunc

Func _OnHelp()
    Local $sHelp = "=== NORMAL.DOTM BACKUP TOOL ===" & @CRLF & @CRLF
    $sHelp &= "CHUC NANG:" & @CRLF
    $sHelp &= "1. Ket noi Word - Ket noi voi Word dang chay" & @CRLF
    $sHelp &= "2. Quick Scan - Quet nhanh thong tin styles" & @CRLF
    $sHelp &= "3. Full Backup - Backup toan bo Normal.dotm" & @CRLF
    $sHelp &= "4. Restore - Restore tu backup da chon" & @CRLF & @CRLF
    $sHelp &= "BACKUP BAO GOM:" & @CRLF
    $sHelp &= "- Danh sach styles" & @CRLF
    $sHelp &= "- Chi tiet format cua tung style" & @CRLF
    $sHelp &= "- Page setup settings" & @CRLF
    $sHelp &= "- Keyboard shortcuts (hotkeys)" & @CRLF
    $sHelp &= "- File Normal.dotm copy" & @CRLF & @CRLF
    $sHelp &= "LUU Y:" & @CRLF
    $sHelp &= "- Backup duoc luu tai: " & $BACKUP_FOLDER & @CRLF
    $sHelp &= "- Restore se dong Word va ghi de Normal.dotm" & @CRLF
    $sHelp &= "- Backup hien tai se duoc luu truoc khi restore"
    
    MsgBox($MB_ICONINFORMATION, "Help", $sHelp)
EndFunc

; ============================================
; HELPERS
; ============================================

Func _Log($sMsg)
    ConsoleWrite($sMsg & @CRLF)
    GUICtrlSetData($g_editLog, $sMsg & @CRLF, 1)
EndFunc
