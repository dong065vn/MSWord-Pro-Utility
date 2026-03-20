; ============================================
; DIALOGS.AU3 - Popup Windows
; ============================================

#include-once

; Hien thi Help
Func _ShowHelp()
    Local $sHelp = "PDF TO WORD FIXER PRO v" & $APP_VERSION & @CRLF & @CRLF
    $sHelp &= "HUONG DAN SU DUNG:" & @CRLF & @CRLF
    $sHelp &= "1. Mo file Word can xu ly" & @CRLF
    $sHelp &= "2. Nhan 'Lam moi' de quet danh sach file" & @CRLF
    $sHelp &= "3. Chon file va nhan 'Ket noi'" & @CRLF
    $sHelp &= "4. Su dung cac tab de xu ly:" & @CRLF
    $sHelp &= "   - Tab 1: Sua loi PDF (xuong dong, khoang trang...)" & @CRLF
    $sHelp &= "   - Tab 2: Dinh dang (font, le trang, heading...)" & @CRLF
    $sHelp &= "   - Tab 3: Cong cu (tim thay the, hinh, bang...)" & @CRLF
    $sHelp &= "   - Tab 4: Muc luc va TLTK" & @CRLF
    $sHelp &= "   - Tab 5: Copy Style giua cac file" & @CRLF
    $sHelp &= "   - Tab 6: Tinh nang nang cao" & @CRLF & @CRLF
    $sHelp &= "LUU Y:" & @CRLF
    $sHelp &= "- Nen Backup truoc khi sua" & @CRLF
    $sHelp &= "- Dung Ctrl+Z de Undo neu can"
    
    MsgBox($MB_ICONINFORMATION, "Tro giup", $sHelp)
EndFunc

; Backup document
Func _BackupDocument()
    If Not _CheckConnection() Then Return
    
    Local $sPath = $g_oDoc.FullName
    If $sPath = "" Then
        MsgBox($MB_ICONWARNING, "Loi", "File chua duoc luu!")
        Return
    EndIf
    
    ; Luu file truoc khi copy (dam bao noi dung moi nhat)
    $g_oDoc.Save()
    
    Local $sDir = StringRegExpReplace($sPath, "\\[^\\]+$", "")
    Local $sName = StringRegExpReplace($sPath, "^.*\\", "")
    $sName = StringRegExpReplace($sName, "\.[^.]+$", "")
    
    Local $sBackup = $sDir & "\" & $sName & "_backup_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & ".docx"
    
    ; FIX: Dung FileCopy thay vi SaveAs2 de khong thay doi document dang ket noi
    If FileCopy($sPath, $sBackup) Then
        MsgBox($MB_ICONINFORMATION, "Backup", "Da tao backup:" & @CRLF & $sBackup)
    Else
        MsgBox($MB_ICONWARNING, "Loi", "Khong the tao backup!" & @CRLF & "Kiem tra quyen ghi file.")
    EndIf
EndFunc

; Save document
Func _SaveDocument()
    If Not _CheckConnection() Then Return
    $g_oDoc.Save()
    _UpdateProgress("Da luu file!")
EndFunc
