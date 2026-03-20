; ============================================
; NORMALDOTMBACKUP.AU3
; Tool chuyen dung: Quet, Sao luu, Backup Normal.dotm
; ============================================

#include-once
#include <Array.au3>
#include <File.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>

; === COM ERROR HANDLER ===
Global $g_oBackupError = ObjEvent("AutoIt.Error", "_BackupComErrorHandler")

Func _BackupComErrorHandler($oError)
    ConsoleWrite("[COM ERROR] Number: 0x" & Hex($oError.number) & @CRLF)
    ConsoleWrite("[COM ERROR] Description: " & $oError.description & @CRLF)
    ConsoleWrite("[COM ERROR] WinDescription: " & $oError.windescription & @CRLF)
    ConsoleWrite("[COM ERROR] Source: " & $oError.source & @CRLF)
    ConsoleWrite("[COM ERROR] ScriptLine: " & $oError.scriptline & @CRLF)
    Return SetError(1, 0, 0)
EndFunc

; === CONSTANTS ===
Global Const $BACKUP_FOLDER = @ScriptDir & "\NormalDotmBackups"
Global Const $wdOrganizerObjectStyles = 0
Global Const $wdStyleTypeParagraph = 1
Global Const $wdStyleTypeCharacter = 2
Global Const $wdStyleTypeTable = 3
Global Const $wdStyleTypeList = 4

; ============================================
; MAIN FUNCTIONS
; ============================================

; Tao backup folder neu chua co
Func _EnsureBackupFolder()
    If Not FileExists($BACKUP_FOLDER) Then
        DirCreate($BACKUP_FOLDER)
        ConsoleWrite("[INFO] Tao folder backup: " & $BACKUP_FOLDER & @CRLF)
    EndIf
    Return FileExists($BACKUP_FOLDER)
EndFunc

; Quet toan bo styles trong Normal.dotm
Func _ScanNormalDotmStyles($oWord)
    If Not IsObj($oWord) Then Return SetError(1, 0, 0)
    
    ConsoleWrite(@CRLF & "=== QUET STYLES TRONG NORMAL.DOTM ===" & @CRLF)
    
    ; Lay Normal template
    Local $oNormalTemplate = $oWord.NormalTemplate
    If Not IsObj($oNormalTemplate) Then
        ConsoleWrite("[ERROR] Khong the truy cap Normal.dotm!" & @CRLF)
        Return SetError(2, 0, 0)
    EndIf
    
    ; Quet styles (with error handling)
    Local $oStyles = 0
    If IsObj($oNormalTemplate) Then
        $oStyles = $oNormalTemplate.Styles
    EndIf
    
    If Not IsObj($oStyles) Or @error Then
        ConsoleWrite("[ERROR] Khong the truy cap Styles collection!" & @CRLF)
        ConsoleWrite("[ERROR] COM Error: " & @error & @CRLF)
        Return SetError(3, 0, 0)
    EndIf
    
    Local $iCount = 0
    If IsObj($oStyles) Then
        $iCount = $oStyles.Count
    EndIf
    
    If $iCount = 0 Then
        ConsoleWrite("[WARNING] Khong co styles nao!" & @CRLF)
        Local $aEmpty[1][6]
        $aEmpty[0][0] = 0
        Return $aEmpty
    EndIf
    
    ConsoleWrite("[INFO] Tong so styles: " & $iCount & @CRLF)
    
    ; Tao array chua thong tin styles
    Local $aStyles[$iCount + 1][6]
    $aStyles[0][0] = $iCount ; Row count
    $aStyles[0][1] = "Name"
    $aStyles[0][2] = "Type"
    $aStyles[0][3] = "BaseStyle"
    $aStyles[0][4] = "NextStyle"
    $aStyles[0][5] = "BuiltIn"
    
    Local $iRow = 1
    For $i = 1 To $iCount
        Local $oStyle = 0
        ; Try-catch style access
        If IsObj($oStyles) Then
            $oStyle = $oStyles.Item($i)
        EndIf
        
        If IsObj($oStyle) Then
            ; Get name safely
            Local $sName = ""
            If IsObj($oStyle) Then
                $sName = $oStyle.NameLocal
            EndIf
            $aStyles[$iRow][1] = $sName
            
            ; Get type safely
            Local $iType = 0
            If IsObj($oStyle) Then
                $iType = $oStyle.Type
            EndIf
            $aStyles[$iRow][2] = _GetStyleTypeName($iType)
            
            ; BaseStyle (co the null) - with error handling
            Local $sBaseStyle = ""
            If IsObj($oStyle) Then
                Local $oBaseStyle = $oStyle.BaseStyle
                If IsObj($oBaseStyle) And Not @error Then
                    $sBaseStyle = $oBaseStyle.NameLocal
                EndIf
            EndIf
            $aStyles[$iRow][3] = $sBaseStyle
            
            ; NextParagraphStyle (chi co voi paragraph styles) - with error handling
            Local $sNextStyle = ""
            If IsObj($oStyle) And $iType = $wdStyleTypeParagraph Then
                Local $oNextStyle = $oStyle.NextParagraphStyle
                If IsObj($oNextStyle) And Not @error Then
                    $sNextStyle = $oNextStyle.NameLocal
                EndIf
            EndIf
            $aStyles[$iRow][4] = $sNextStyle
            
            ; BuiltIn - with error handling
            Local $bBuiltIn = False
            If IsObj($oStyle) Then
                $bBuiltIn = $oStyle.BuiltIn
            EndIf
            $aStyles[$iRow][5] = $bBuiltIn ? "Yes" : "No"
            
            $iRow += 1
        EndIf
    Next
    
    ConsoleWrite("[SUCCESS] Da quet xong " & ($iRow - 1) & " styles" & @CRLF)
    Return $aStyles
EndFunc

; Export styles ra file text
Func _ExportStylesToFile($aStyles, $sFilePath)
    If Not IsArray($aStyles) Then Return SetError(1, 0, False)
    
    Local $hFile = FileOpen($sFilePath, 2) ; Overwrite mode
    If $hFile = -1 Then Return SetError(2, 0, False)
    
    ; Header
    FileWriteLine($hFile, "=== NORMAL.DOTM STYLES BACKUP ===")
    FileWriteLine($hFile, "Date: " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC)
    FileWriteLine($hFile, "Total Styles: " & $aStyles[0][0])
    FileWriteLine($hFile, "")
    FileWriteLine($hFile, StringFormat("%-40s | %-15s | %-20s | %-20s | %-8s", "Name", "Type", "BaseStyle", "NextStyle", "BuiltIn"))
    FileWriteLine($hFile, StringRepeat("-", 120))
    
    ; Data
    For $i = 1 To $aStyles[0][0]
        FileWriteLine($hFile, StringFormat("%-40s | %-15s | %-20s | %-20s | %-8s", _
            $aStyles[$i][1], $aStyles[$i][2], $aStyles[$i][3], $aStyles[$i][4], $aStyles[$i][5]))
    Next
    
    FileClose($hFile)
    ConsoleWrite("[SUCCESS] Export styles thanh cong: " & $sFilePath & @CRLF)
    Return True
EndFunc

; Backup toan bo Normal.dotm (styles + settings)
Func _BackupNormalDotm($oWord, $sBackupName = "")
    If Not IsObj($oWord) Then Return SetError(1, 0, False)
    
    _EnsureBackupFolder()
    
    ; Tao ten backup
    If $sBackupName = "" Then
        $sBackupName = "NormalDotm_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC
    EndIf
    
    Local $sBackupPath = $BACKUP_FOLDER & "\" & $sBackupName
    DirCreate($sBackupPath)
    
    ConsoleWrite(@CRLF & "=== BACKUP NORMAL.DOTM ===" & @CRLF)
    ConsoleWrite("[INFO] Backup path: " & $sBackupPath & @CRLF)
    
    ; 1. Backup Styles
    ConsoleWrite(@CRLF & "[1/4] Backup Styles..." & @CRLF)
    Local $aStyles = _ScanNormalDotmStyles($oWord)
    If @error Then
        ConsoleWrite("[ERROR] Khong the quet styles!" & @CRLF)
        Return SetError(2, 0, False)
    EndIf
    
    _ExportStylesToFile($aStyles, $sBackupPath & "\styles.txt")
    
    ; 2. Backup Style Definitions (chi tiet format)
    ConsoleWrite(@CRLF & "[2/4] Backup Style Definitions..." & @CRLF)
    _BackupStyleDefinitions($oWord, $sBackupPath & "\style_definitions.txt")
    
    ; 3. Backup Page Setup
    ConsoleWrite(@CRLF & "[3/4] Backup Page Setup..." & @CRLF)
    _BackupPageSetup($oWord, $sBackupPath & "\page_setup.txt")
    
    ; 4. Backup Hotkeys
    ConsoleWrite(@CRLF & "[4/4] Backup Hotkeys..." & @CRLF)
    _BackupHotkeys($oWord, $sBackupPath & "\hotkeys.txt")
    
    ; 5. Tao file Normal.dotm copy
    ConsoleWrite(@CRLF & "[BONUS] Copy file Normal.dotm..." & @CRLF)
    Local $sNormalPath = $oWord.NormalTemplate.FullName
    FileCopy($sNormalPath, $sBackupPath & "\Normal.dotm", 1)
    
    ; Tao manifest
    _CreateBackupManifest($sBackupPath, $aStyles[0][0])
    
    ConsoleWrite(@CRLF & "[SUCCESS] Backup hoan tat!" & @CRLF)
    ConsoleWrite("Backup location: " & $sBackupPath & @CRLF)
    
    Return $sBackupPath
EndFunc

; Backup chi tiet format cua tung style
Func _BackupStyleDefinitions($oWord, $sFilePath)
    Local $oNormalTemplate = $oWord.NormalTemplate
    If Not IsObj($oNormalTemplate) Then Return False
    
    Local $oStyles = 0
    If IsObj($oNormalTemplate) Then
        $oStyles = $oNormalTemplate.Styles
    EndIf
    If Not IsObj($oStyles) Or @error Then Return False
    
    Local $hFile = FileOpen($sFilePath, 2)
    If $hFile = -1 Then Return False
    
    FileWriteLine($hFile, "=== STYLE DEFINITIONS (FORMAT DETAILS) ===")
    FileWriteLine($hFile, "Date: " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC)
    FileWriteLine($hFile, "")
    
    Local $iCount = 0
    If IsObj($oStyles) Then
        $iCount = $oStyles.Count
    EndIf
    
    For $i = 1 To $iCount
        Local $oStyle = 0
        If IsObj($oStyles) Then
            $oStyle = $oStyles.Item($i)
        EndIf
        
        If IsObj($oStyle) Then
            FileWriteLine($hFile, "")
            FileWriteLine($hFile, StringRepeat("=", 80))
            
            Local $sName = ""
            If IsObj($oStyle) Then
                $sName = $oStyle.NameLocal
            EndIf
            FileWriteLine($hFile, "STYLE: " & $sName)
            FileWriteLine($hFile, StringRepeat("=", 80))
            
            Local $iType = 0
            If IsObj($oStyle) Then
                $iType = $oStyle.Type
            EndIf
            FileWriteLine($hFile, "Type: " & _GetStyleTypeName($iType))
            
            Local $bBuiltIn = False
            If IsObj($oStyle) Then
                $bBuiltIn = $oStyle.BuiltIn
            EndIf
            FileWriteLine($hFile, "BuiltIn: " & ($bBuiltIn ? "Yes" : "No"))
            
            ; Chi lay format cho paragraph styles
            If $iType = $wdStyleTypeParagraph Then
                Local $oPara = 0
                Local $oFont = 0
                
                If IsObj($oStyle) Then
                    $oPara = $oStyle.ParagraphFormat
                    $oFont = $oStyle.Font
                EndIf
                
                If IsObj($oPara) And Not @error Then
                    FileWriteLine($hFile, "")
                    FileWriteLine($hFile, "--- Paragraph Format ---")
                    FileWriteLine($hFile, "  Alignment: " & $oPara.Alignment)
                    FileWriteLine($hFile, "  FirstLineIndent: " & $oPara.FirstLineIndent & " pt")
                    FileWriteLine($hFile, "  LeftIndent: " & $oPara.LeftIndent & " pt")
                    FileWriteLine($hFile, "  RightIndent: " & $oPara.RightIndent & " pt")
                    FileWriteLine($hFile, "  SpaceBefore: " & $oPara.SpaceBefore & " pt")
                    FileWriteLine($hFile, "  SpaceAfter: " & $oPara.SpaceAfter & " pt")
                    FileWriteLine($hFile, "  LineSpacing: " & $oPara.LineSpacing)
                    FileWriteLine($hFile, "  LineSpacingRule: " & $oPara.LineSpacingRule)
                EndIf
                
                If IsObj($oFont) And Not @error Then
                    FileWriteLine($hFile, "")
                    FileWriteLine($hFile, "--- Font Format ---")
                    FileWriteLine($hFile, "  Name: " & $oFont.Name)
                    FileWriteLine($hFile, "  Size: " & $oFont.Size & " pt")
                    FileWriteLine($hFile, "  Bold: " & $oFont.Bold)
                    FileWriteLine($hFile, "  Italic: " & $oFont.Italic)
                    FileWriteLine($hFile, "  Underline: " & $oFont.Underline)
                    FileWriteLine($hFile, "  Color: " & $oFont.Color)
                EndIf
            EndIf
        EndIf
    Next
    
    FileClose($hFile)
    ConsoleWrite("[SUCCESS] Export style definitions: " & $sFilePath & @CRLF)
    Return True
EndFunc

; Backup Page Setup
Func _BackupPageSetup($oWord, $sFilePath)
    Local $oNormalTemplate = $oWord.NormalTemplate
    If Not IsObj($oNormalTemplate) Then Return False
    
    ; Tao document tam de lay page setup
    Local $oTempDoc = $oWord.Documents.Add($oNormalTemplate.FullName)
    If Not IsObj($oTempDoc) Then Return False
    
    Local $oPageSetup = $oTempDoc.PageSetup
    Local $hFile = FileOpen($sFilePath, 2)
    If $hFile = -1 Then
        $oTempDoc.Close(False)
        Return False
    EndIf
    
    FileWriteLine($hFile, "=== PAGE SETUP (NORMAL.DOTM) ===")
    FileWriteLine($hFile, "Date: " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC)
    FileWriteLine($hFile, "")
    
    If IsObj($oPageSetup) Then
        FileWriteLine($hFile, "TopMargin: " & $oPageSetup.TopMargin & " pt")
        FileWriteLine($hFile, "BottomMargin: " & $oPageSetup.BottomMargin & " pt")
        FileWriteLine($hFile, "LeftMargin: " & $oPageSetup.LeftMargin & " pt")
        FileWriteLine($hFile, "RightMargin: " & $oPageSetup.RightMargin & " pt")
        FileWriteLine($hFile, "PageWidth: " & $oPageSetup.PageWidth & " pt")
        FileWriteLine($hFile, "PageHeight: " & $oPageSetup.PageHeight & " pt")
        FileWriteLine($hFile, "Orientation: " & $oPageSetup.Orientation)
        FileWriteLine($hFile, "PaperSize: " & $oPageSetup.PaperSize)
        FileWriteLine($hFile, "Gutter: " & $oPageSetup.Gutter & " pt")
        FileWriteLine($hFile, "HeaderDistance: " & $oPageSetup.HeaderDistance & " pt")
        FileWriteLine($hFile, "FooterDistance: " & $oPageSetup.FooterDistance & " pt")
    EndIf
    
    FileClose($hFile)
    $oTempDoc.Close(False)
    
    ConsoleWrite("[SUCCESS] Export page setup: " & $sFilePath & @CRLF)
    Return True
EndFunc

; Backup Hotkeys
Func _BackupHotkeys($oWord, $sFilePath)
    Local $hFile = FileOpen($sFilePath, 2)
    If $hFile = -1 Then Return False
    
    FileWriteLine($hFile, "=== HOTKEYS (NORMAL.DOTM) ===")
    FileWriteLine($hFile, "Date: " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC)
    FileWriteLine($hFile, "")
    FileWriteLine($hFile, StringFormat("%-40s | %-20s", "Style Name", "Hotkey"))
    FileWriteLine($hFile, StringRepeat("-", 70))
    
    ; Quet hotkeys cho styles
    Local $oNormalTemplate = $oWord.NormalTemplate
    If Not IsObj($oNormalTemplate) Then
        FileClose($hFile)
        Return False
    EndIf
    
    Local $oStyles = 0
    If IsObj($oNormalTemplate) Then
        $oStyles = $oNormalTemplate.Styles
    EndIf
    If Not IsObj($oStyles) Or @error Then
        FileClose($hFile)
        Return False
    EndIf
    
    Local Const $wdKeyCategoryStyle = 5
    Local $iCount = 0
    If IsObj($oStyles) Then
        $iCount = $oStyles.Count
    EndIf
    
    For $i = 1 To $iCount
        Local $oStyle = 0
        If IsObj($oStyles) Then
            $oStyle = $oStyles.Item($i)
        EndIf
        
        If IsObj($oStyle) Then
            Local $sStyleName = ""
            If IsObj($oStyle) Then
                $sStyleName = $oStyle.NameLocal
            EndIf
            
            Local $oKeysBound = $oWord.KeysBoundTo($wdKeyCategoryStyle, $sStyleName)
            
            If IsObj($oKeysBound) And Not @error And $oKeysBound.Count > 0 Then
                For $j = 1 To $oKeysBound.Count
                    Local $oKey = $oKeysBound.Item($j)
                    If IsObj($oKey) And Not @error Then
                        FileWriteLine($hFile, StringFormat("%-40s | %-20s", $sStyleName, $oKey.KeyString))
                    EndIf
                Next
            EndIf
        EndIf
    Next
    
    FileClose($hFile)
    ConsoleWrite("[SUCCESS] Export hotkeys: " & $sFilePath & @CRLF)
    Return True
EndFunc

; Tao manifest file
Func _CreateBackupManifest($sBackupPath, $iStyleCount)
    Local $hFile = FileOpen($sBackupPath & "\MANIFEST.txt", 2)
    If $hFile = -1 Then Return False
    
    FileWriteLine($hFile, "=== NORMAL.DOTM BACKUP MANIFEST ===")
    FileWriteLine($hFile, "")
    FileWriteLine($hFile, "Backup Date: " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC)
    FileWriteLine($hFile, "Computer: " & @ComputerName)
    FileWriteLine($hFile, "User: " & @UserName)
    FileWriteLine($hFile, "")
    FileWriteLine($hFile, "--- Contents ---")
    FileWriteLine($hFile, "1. styles.txt - List of all styles (" & $iStyleCount & " styles)")
    FileWriteLine($hFile, "2. style_definitions.txt - Detailed format of each style")
    FileWriteLine($hFile, "3. page_setup.txt - Page setup settings")
    FileWriteLine($hFile, "4. hotkeys.txt - Keyboard shortcuts for styles")
    FileWriteLine($hFile, "5. Normal.dotm - Copy of Normal template file")
    FileWriteLine($hFile, "")
    FileWriteLine($hFile, "--- Restore Instructions ---")
    FileWriteLine($hFile, "1. Close all Word instances")
    FileWriteLine($hFile, "2. Copy Normal.dotm to: %APPDATA%\Microsoft\Templates\")
    FileWriteLine($hFile, "3. Or use restore function in tool")
    
    FileClose($hFile)
    Return True
EndFunc

; ============================================
; RESTORE FUNCTIONS
; ============================================

; List tat ca backups
Func _ListBackups()
    _EnsureBackupFolder()
    
    Local $aBackups = _FileListToArray($BACKUP_FOLDER, "*", $FLTA_FOLDERS)
    If @error Then Return SetError(1, 0, 0)
    
    ConsoleWrite(@CRLF & "=== DANH SACH BACKUPS ===" & @CRLF)
    For $i = 1 To $aBackups[0]
        ConsoleWrite($i & ". " & $aBackups[$i] & @CRLF)
    Next
    
    Return $aBackups
EndFunc

; Restore Normal.dotm tu backup
Func _RestoreNormalDotm($oWord, $sBackupName)
    If Not IsObj($oWord) Then Return SetError(1, 0, False)
    
    Local $sBackupPath = $BACKUP_FOLDER & "\" & $sBackupName
    If Not FileExists($sBackupPath) Then
        ConsoleWrite("[ERROR] Backup khong ton tai: " & $sBackupPath & @CRLF)
        Return SetError(2, 0, False)
    EndIf
    
    ConsoleWrite(@CRLF & "=== RESTORE NORMAL.DOTM ===" & @CRLF)
    ConsoleWrite("[INFO] Restore from: " & $sBackupPath & @CRLF)
    
    ; Kiem tra file Normal.dotm trong backup
    Local $sBackupNormalDotm = $sBackupPath & "\Normal.dotm"
    If Not FileExists($sBackupNormalDotm) Then
        ConsoleWrite("[ERROR] Khong tim thay Normal.dotm trong backup!" & @CRLF)
        Return SetError(3, 0, False)
    EndIf
    
    ; Lay duong dan Normal.dotm hien tai
    Local $sCurrentNormalPath = $oWord.NormalTemplate.FullName
    ConsoleWrite("[INFO] Current Normal.dotm: " & $sCurrentNormalPath & @CRLF)
    
    ; Backup Normal.dotm hien tai truoc khi restore
    Local $sPreRestoreBackup = $sCurrentNormalPath & ".pre-restore." & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC
    FileCopy($sCurrentNormalPath, $sPreRestoreBackup, 1)
    ConsoleWrite("[INFO] Pre-restore backup: " & $sPreRestoreBackup & @CRLF)
    
    ; Dong Word de co the copy file
    Local $iResponse = MsgBox($MB_YESNO + $MB_ICONWARNING, "Restore Normal.dotm", _
        "De restore Normal.dotm, can dong tat ca Word instances." & @CRLF & @CRLF & _
        "Backup hien tai da duoc luu tai:" & @CRLF & $sPreRestoreBackup & @CRLF & @CRLF & _
        "Tiep tuc?")
    
    If $iResponse = $IDNO Then
        ConsoleWrite("[INFO] User huy restore" & @CRLF)
        Return False
    EndIf
    
    ; Dong Word
    ConsoleWrite("[INFO] Dong Word..." & @CRLF)
    $oWord.Quit(0) ; 0 = wdDoNotSaveChanges
    Sleep(2000)
    
    ; Copy file
    ConsoleWrite("[INFO] Copy Normal.dotm..." & @CRLF)
    FileCopy($sBackupNormalDotm, $sCurrentNormalPath, 1)
    
    If FileExists($sCurrentNormalPath) Then
        ConsoleWrite("[SUCCESS] Restore thanh cong!" & @CRLF)
        MsgBox($MB_ICONINFORMATION, "Restore thanh cong", _
            "Normal.dotm da duoc restore!" & @CRLF & @CRLF & _
            "Vui long khoi dong lai Word de ap dung thay doi.")
        Return True
    Else
        ConsoleWrite("[ERROR] Restore that bai!" & @CRLF)
        Return SetError(4, 0, False)
    EndIf
EndFunc

; ============================================
; HELPER FUNCTIONS
; ============================================

; Lay ten loai style
Func _GetStyleTypeName($iType)
    Switch $iType
        Case $wdStyleTypeParagraph
            Return "Paragraph"
        Case $wdStyleTypeCharacter
            Return "Character"
        Case $wdStyleTypeTable
            Return "Table"
        Case $wdStyleTypeList
            Return "List"
        Case Else
            Return "Unknown (" & $iType & ")"
    EndSwitch
EndFunc

; So sanh 2 backups
Func _CompareBackups($sBackup1, $sBackup2)
    Local $sPath1 = $BACKUP_FOLDER & "\" & $sBackup1 & "\styles.txt"
    Local $sPath2 = $BACKUP_FOLDER & "\" & $sBackup2 & "\styles.txt"
    
    If Not FileExists($sPath1) Or Not FileExists($sPath2) Then
        ConsoleWrite("[ERROR] Khong tim thay file styles.txt!" & @CRLF)
        Return False
    EndIf
    
    ConsoleWrite(@CRLF & "=== SO SANH 2 BACKUPS ===" & @CRLF)
    ConsoleWrite("Backup 1: " & $sBackup1 & @CRLF)
    ConsoleWrite("Backup 2: " & $sBackup2 & @CRLF)
    ConsoleWrite(@CRLF)
    
    ; Doc file 1
    Local $aFile1
    _FileReadToArray($sPath1, $aFile1)
    
    ; Doc file 2
    Local $aFile2
    _FileReadToArray($sPath2, $aFile2)
    
    ; So sanh so dong
    ConsoleWrite("File 1 lines: " & $aFile1[0] & @CRLF)
    ConsoleWrite("File 2 lines: " & $aFile2[0] & @CRLF)
    
    If $aFile1[0] <> $aFile2[0] Then
        ConsoleWrite("[DIFF] So dong khac nhau!" & @CRLF)
    Else
        ConsoleWrite("[SAME] So dong giong nhau" & @CRLF)
    EndIf
    
    Return True
EndFunc

; Export danh sach styles ra CSV (de import vao Excel)
Func _ExportStylesToCSV($aStyles, $sFilePath)
    If Not IsArray($aStyles) Then Return SetError(1, 0, False)
    
    Local $hFile = FileOpen($sFilePath, 2)
    If $hFile = -1 Then Return SetError(2, 0, False)
    
    ; Header
    FileWriteLine($hFile, "Name,Type,BaseStyle,NextStyle,BuiltIn")
    
    ; Data
    For $i = 1 To $aStyles[0][0]
        FileWriteLine($hFile, $aStyles[$i][1] & "," & $aStyles[$i][2] & "," & _
            $aStyles[$i][3] & "," & $aStyles[$i][4] & "," & $aStyles[$i][5])
    Next
    
    FileClose($hFile)
    ConsoleWrite("[SUCCESS] Export CSV: " & $sFilePath & @CRLF)
    Return True
EndFunc

; Tim style theo ten
Func _FindStyleInBackup($sBackupName, $sStyleName)
    Local $sStylesFile = $BACKUP_FOLDER & "\" & $sBackupName & "\styles.txt"
    If Not FileExists($sStylesFile) Then Return SetError(1, 0, "")
    
    Local $hFile = FileOpen($sStylesFile, 0)
    If $hFile = -1 Then Return SetError(2, 0, "")
    
    Local $sResult = ""
    While True
        Local $sLine = FileReadLine($hFile)
        If @error Then ExitLoop
        
        If StringInStr($sLine, $sStyleName) Then
            $sResult = $sLine
            ExitLoop
        EndIf
    WEnd
    
    FileClose($hFile)
    Return $sResult
EndFunc

; ============================================
; QUICK ACTIONS
; ============================================

; Quick backup (1 click)
Func _QuickBackup($oWord)
    Local $sBackupPath = _BackupNormalDotm($oWord)
    If @error Then
        MsgBox($MB_ICONERROR, "Loi", "Backup that bai!")
        Return False
    EndIf
    
    MsgBox($MB_ICONINFORMATION, "Backup thanh cong", _
        "Normal.dotm da duoc backup!" & @CRLF & @CRLF & _
        "Location: " & $sBackupPath)
    
    Return True
EndFunc

; Quick scan (chi hien thi thong tin)
Func _QuickScan($oWord)
    Local $aStyles = _ScanNormalDotmStyles($oWord)
    If @error Then
        MsgBox($MB_ICONERROR, "Loi", "Khong the quet styles!")
        Return False
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
    
    Local $sMsg = "=== NORMAL.DOTM SCAN RESULT ===" & @CRLF & @CRLF
    $sMsg &= "Total Styles: " & $aStyles[0][0] & @CRLF & @CRLF
    $sMsg &= "By Type:" & @CRLF
    $sMsg &= "  - Paragraph: " & $iParagraph & @CRLF
    $sMsg &= "  - Character: " & $iCharacter & @CRLF
    $sMsg &= "  - Table: " & $iTable & @CRLF
    $sMsg &= "  - List: " & $iList & @CRLF & @CRLF
    $sMsg &= "By Origin:" & @CRLF
    $sMsg &= "  - Built-in: " & $iBuiltIn & @CRLF
    $sMsg &= "  - Custom: " & $iCustom
    
    MsgBox($MB_ICONINFORMATION, "Scan Result", $sMsg)
    Return True
EndFunc
