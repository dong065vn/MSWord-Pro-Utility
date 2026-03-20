; ============================================
; TOOLS.AU3 - Module Cong cu
; ============================================

#include-once

; Find & Replace
Func _DoFindReplace()
    If Not _CheckConnection() Then Return
    Local $sFind = GUICtrlRead($g_inputFind)
    Local $sReplace = GUICtrlRead($g_inputReplace)
    If $sFind = "" Then
        MsgBox($MB_ICONWARNING, "Loi", "Nhap tu can tim!")
        Return
    EndIf

    Local $oFind = $g_oDoc.Content.Find
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    
    Local $bCase = (GUICtrlRead($g_chkMatchCase) = $GUI_CHECKED)
    Local $bWord = (GUICtrlRead($g_chkWholeWord) = $GUI_CHECKED)
    
    $oFind.Execute($sFind, $bCase, $bWord, False, False, False, True, 1, False, $sReplace, $WD_REPLACE_ALL)
    _UpdateProgress("Da thay the xong!")
EndFunc

; Find Next
Func _DoFindNext()
    If Not _CheckConnection() Then Return
    Local $sFind = GUICtrlRead($g_inputFind)
    If $sFind = "" Then Return

    Local $oFind = $g_oWord.Selection.Find
    $oFind.ClearFormatting()
    $oFind.Execute($sFind)
EndFunc

; Resize Images
Func _ResizeImages()
    If Not _CheckConnection() Then Return
    Local $oShapes = $g_oDoc.InlineShapes
    If Not IsObj($oShapes) Or $oShapes.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co hinh!")
        Return
    EndIf

    Local $fMaxW = $g_oDoc.PageSetup.PageWidth - $g_oDoc.PageSetup.LeftMargin - $g_oDoc.PageSetup.RightMargin
    Local $n = 0

    _UpdateProgress("Dang resize hinh...")
    For $i = 1 To $oShapes.Count
        Local $oS = $oShapes.Item($i)
        If Not IsObj($oS) Then ContinueLoop
        
        If $oS.Width > $fMaxW Then
            Local $fRatio = $fMaxW / $oS.Width
            $oS.Width = $fMaxW
            $oS.Height = $oS.Height * $fRatio
            $n += 1
        EndIf
    Next
    _UpdateProgress("Da resize " & $n & "/" & $oShapes.Count & " hinh!")
EndFunc

; Center Images
Func _CenterImages()
    If Not _CheckConnection() Then Return
    Local $oShapes = $g_oDoc.InlineShapes
    If Not IsObj($oShapes) Or $oShapes.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co hinh!")
        Return
    EndIf
    
    _UpdateProgress("Dang can giua hinh...")
    For $i = 1 To $oShapes.Count
        Local $oShape = $oShapes.Item($i)
        If IsObj($oShape) And IsObj($oShape.Range) Then
            $oShape.Range.ParagraphFormat.Alignment = $WD_ALIGN_CENTER
        EndIf
    Next
    _UpdateProgress("Da can giua " & $oShapes.Count & " hinh!")
EndFunc

; Auto Caption Images
Func _AutoCaptionImg()
    _AutoNumberImages()
EndFunc

; Remove All Images
Func _RemoveAllImages()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", "Xoa tat ca hinh?") <> $IDYES Then Return
    
    Local $oShapes = $g_oDoc.InlineShapes
    Local $n = $oShapes.Count
    While $oShapes.Count > 0
        $oShapes.Item(1).Delete()
    WEnd
    _UpdateProgress("Da xoa " & $n & " hinh!")
EndFunc

; AutoFit Tables
Func _AutoFitTables($iMode)
    If Not _CheckConnection() Then Return
    Local $oTables = $g_oDoc.Tables
    If Not IsObj($oTables) Or $oTables.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co bang!")
        Return
    EndIf

    _UpdateProgress("Dang AutoFit bang...")
    Local $iSuccess = 0
    For $i = 1 To $oTables.Count
        Local $oTable = $oTables.Item($i)
        If IsObj($oTable) Then
            If $iMode = 1 Then
                $oTable.AutoFitBehavior(1) ; wdAutoFitContent
            Else
                $oTable.AutoFitBehavior(2) ; wdAutoFitWindow
            EndIf
            $iSuccess += 1
        EndIf
    Next
    _UpdateProgress("Da AutoFit " & $iSuccess & "/" & $oTables.Count & " bang!")
EndFunc

; Auto Caption Tables
Func _AutoCaptionTbl()
    _AutoNumberTables()
EndFunc

; Add Table Borders
Func _AddTableBorders()
    If Not _CheckConnection() Then Return
    Local $oTables = $g_oDoc.Tables
    If Not IsObj($oTables) Or $oTables.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co bang!")
        Return
    EndIf
    
    _UpdateProgress("Dang them vien bang...")
    For $i = 1 To $oTables.Count
        Local $oTable = $oTables.Item($i)
        If IsObj($oTable) Then
            $oTable.Borders.Enable = True
        EndIf
    Next
    _UpdateProgress("Da them vien " & $oTables.Count & " bang!")
EndFunc


; Word Count
Func _ShowWordCount()
    If Not _CheckConnection() Then Return
    Local $sMsg = "THONG KE TAI LIEU" & @CRLF & @CRLF
    $sMsg &= "So trang: " & $g_oDoc.ComputeStatistics(2) & @CRLF
    $sMsg &= "So tu: " & $g_oDoc.ComputeStatistics(0) & @CRLF
    $sMsg &= "So ky tu: " & $g_oDoc.ComputeStatistics(3) & @CRLF
    $sMsg &= "So doan van: " & $g_oDoc.ComputeStatistics(4) & @CRLF
    $sMsg &= "So bang: " & $g_oDoc.Tables.Count & @CRLF
    $sMsg &= "So hinh: " & $g_oDoc.InlineShapes.Count
    
    _LogPreview($sMsg)
    MsgBox($MB_ICONINFORMATION, "Thong ke", $sMsg)
EndFunc

; Check Spelling
Func _CheckSpelling()
    If Not _CheckConnection() Then Return
    $g_oDoc.CheckSpelling()
    _UpdateProgress("Da kiem tra chinh ta!")
EndFunc

; Check Format
Func _CheckFormat()
    _CheckThesisFormat()
EndFunc

; Show Detailed Stats
Func _ShowDetailedStats()
    _ShowWordCount()
EndFunc

; Export Stats
Func _ExportStats()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Luu bao cao", @DesktopDir, "Text (*.txt)", 16, "ThongKe.txt")
    If @error Then Return

    Local $sStats = "THONG KE TAI LIEU: " & $g_oDoc.Name & @CRLF
    $sStats &= "Ngay: " & @MDAY & "/" & @MON & "/" & @YEAR & @CRLF & @CRLF
    $sStats &= "So trang: " & $g_oDoc.ComputeStatistics(2) & @CRLF
    $sStats &= "So tu: " & $g_oDoc.ComputeStatistics(0) & @CRLF
    $sStats &= "So ky tu: " & $g_oDoc.ComputeStatistics(3) & @CRLF
    $sStats &= "So bang: " & $g_oDoc.Tables.Count & @CRLF
    $sStats &= "So hinh: " & $g_oDoc.InlineShapes.Count

    FileWrite($sPath, $sStats)
    _UpdateProgress("Da xuat bao cao!")
EndFunc
