; ============================================
; ADVANCED.AU3 - Module Nang cao
; ============================================

#include-once

; Auto Detect Heading
Func _AutoDetectHeading()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang tu dong gan Heading...")
    
    Local $iH1 = 0, $iH2 = 0, $iH3 = 0
    Local $oParas = $g_oDoc.Paragraphs
    
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        Local $oRange = $oPara.Range
        Local $sText = StringStripWS($oRange.Text, 3)
        
        If StringLen($sText) < 3 Or StringLen($sText) > 100 Then ContinueLoop
        
        ; Get font properties - handle mixed formatting (9999999 = wdUndefined)
        Local $bBold = $oRange.Font.Bold
        Local $bItalic = $oRange.Font.Italic
        Local $fSize = $oRange.Font.Size
        
        ; Skip if mixed formatting (undefined values)
        If $bBold = 9999999 Or $fSize = 9999999 Then ContinueLoop
        
        ; Detect H1: Bold, Size >= 14
        If $bBold = True And $fSize >= 14 Then
            $oRange.Style = "Heading 1"
            $iH1 += 1
            ContinueLoop
        EndIf
        
        ; Detect H2: Bold, Size >= 13
        If $bBold = True And $fSize >= 13 Then
            $oRange.Style = "Heading 2"
            $iH2 += 1
            ContinueLoop
        EndIf
        
        ; Detect H3: Bold + Italic
        If $bBold = True And $bItalic = True Then
            $oRange.Style = "Heading 3"
            $iH3 += 1
        EndIf
    Next
    
    _UpdateProgress("Da gan: H1=" & $iH1 & ", H2=" & $iH2 & ", H3=" & $iH3)
    MsgBox($MB_ICONINFORMATION, "Ket qua", _
        "Da tu dong gan Heading:" & @CRLF & _
        "Heading 1: " & $iH1 & @CRLF & _
        "Heading 2: " & $iH2 & @CRLF & _
        "Heading 3: " & $iH3)
EndFunc

; Reset All Headings
Func _ResetAllHeadings()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", "Reset tat ca Heading ve Normal?") <> $IDYES Then Return
    
    _UpdateProgress("Dang reset Headings...")
    Local $oParas = $g_oDoc.Paragraphs
    Local $iReset = 0
    
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        
        Local $oStyle = $oPara.Style
        If Not IsObj($oStyle) Then ContinueLoop
        
        Local $sStyle = $oStyle.NameLocal
        If StringInStr($sStyle, "Heading") Or StringInStr($sStyle, "Tieu de") Then
            $oPara.Style = "Normal"
            $iReset += 1
        EndIf
    Next
    
    _UpdateProgress("Da reset " & $iReset & " Heading!")
EndFunc

; Heading to TOC
Func _HeadingToTOC()
    _CreateTOC()
EndFunc

; List All Headings
Func _ListAllHeadings()
    If Not _CheckConnection() Then Return
    
    Local $sMsg = "DANH SACH HEADING:" & @CRLF & @CRLF
    Local $oParas = $g_oDoc.Paragraphs
    Local $iCount = 0
    
    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop
        
        Local $oStyle = $oPara.Style
        If Not IsObj($oStyle) Then ContinueLoop
        
        Local $sStyle = $oStyle.NameLocal
        If StringInStr($sStyle, "Heading") Or StringInStr($sStyle, "Tieu de") Then
            Local $sText = StringStripWS($oPara.Range.Text, 3)
            $sText = StringLeft($sText, 50)
            $sMsg &= $sStyle & ": " & $sText & @CRLF
            $iCount += 1
            If $iCount >= 50 Then
                $sMsg &= "... va con nua"
                ExitLoop
            EndIf
        EndIf
    Next
    
    If $iCount = 0 Then $sMsg &= "(Khong co Heading)"
    
    _LogPreview($sMsg)
    MsgBox($MB_ICONINFORMATION, "Danh sach Heading (" & $iCount & ")", $sMsg)
EndFunc

; Remove All Formatting
Func _RemoveAllFormatting()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", "Xoa tat ca dinh dang?") <> $IDYES Then Return
    $g_oDoc.Content.ClearFormatting()
    _UpdateProgress("Da xoa tat ca dinh dang!")
EndFunc

; Convert Text Case
Func _ConvertTextCase()
    If Not _CheckConnection() Then Return
    Local $sChoice = InputBox("Chuyen doi chu", _
        "1 - UPPER CASE" & @CRLF & _
        "2 - lower case" & @CRLF & _
        "3 - Title Case", "1")
    If @error Or $sChoice = "" Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Or $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon van ban truoc!")
        Return
    EndIf
    
    Switch $sChoice
        Case "1"
            $oSel.Range.Case = 1 ; wdUpperCase
        Case "2"
            $oSel.Range.Case = 2 ; wdLowerCase
        Case "3"
            $oSel.Range.Case = 4 ; wdTitleCase
        Case Else
            MsgBox($MB_ICONWARNING, "Loi", "Chon 1, 2 hoac 3!")
            Return
    EndSwitch
    _UpdateProgress("Da chuyen doi!")
EndFunc

; Remove All Hyperlinks
Func _RemoveAllHyperlinks()
    If Not _CheckConnection() Then Return
    Local $oLinks = $g_oDoc.Hyperlinks
    Local $n = $oLinks.Count
    While $oLinks.Count > 0
        $oLinks.Item(1).Delete()
    WEnd
    _UpdateProgress("Da xoa " & $n & " hyperlinks!")
EndFunc


; Remove All Comments
Func _RemoveAllComments()
    If Not _CheckConnection() Then Return
    Local $oComments = $g_oDoc.Comments
    Local $n = $oComments.Count
    While $oComments.Count > 0
        $oComments.Item(1).Delete()
    WEnd
    _UpdateProgress("Da xoa " & $n & " comments!")
EndFunc

; Accept All Changes
Func _AcceptAllChanges()
    If Not _CheckConnection() Then Return
    $g_oDoc.AcceptAllRevisions()
    _UpdateProgress("Da chap nhan tat ca thay doi!")
EndFunc

; Convert Numbering to Text
; FIX: Su dung while-loop de xu ly Live Collection, xoa list formatting sau khi convert
Func _ConvertNumberingToText()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Xac nhan", _
        "Chuyen Numbering/Bullet thanh text?" & @CRLF & @CRLF & _
        "Chuc nang nay se:" & @CRLF & _
        "1. Chuyen so thu tu (1. 2. 3.) thanh text" & @CRLF & _
        "2. Chuyen bullet thanh ky tu text" & @CRLF & _
        "3. Xoa dinh dang list (thut le, numbering)" & @CRLF & @CRLF & _
        "LUU Y: Nen Backup truoc!") <> $IDYES Then Return

    _UpdateProgress("Dang chuyen Numbering thanh text...")

    ; Buoc 1: Dem so luong lists ban dau
    Local $iTotal = $g_oDoc.Lists.Count
    If $iTotal = 0 Then
        MsgBox($MB_ICONINFORMATION, "Thong bao", "Khong tim thay Numbering/Bullet nao!")
        _UpdateProgress("")
        Return
    EndIf

    ; Buoc 2: Convert numbers thanh text
    ; QUAN TRONG: Lists la Live Collection - sau moi ConvertNumbersToText(),
    ; list bi xoa khoi collection. Dung While loop de xu ly an toan.
    Local $iConverted = 0
    Local $iFailed = 0

    While $g_oDoc.Lists.Count > 0
        _UpdateProgress("Dang chuyen list " & ($iConverted + 1) & "/" & $iTotal & "...")
        Local $oList = $g_oDoc.Lists.Item(1)
        If IsObj($oList) Then
            ; Luu range truoc khi convert de xoa formatting sau
            Local $oRange = $oList.Range
            $oList.ConvertNumbersToText()
            If @error Then
                $iFailed += 1
            Else
                ; Buoc 3: Xoa list formatting con sot lai (thut le, ListTemplate)
                ; Neu khong xoa, paragraph van giu indent va co the bi re-number
                If IsObj($oRange) Then
                    $oRange.ListFormat.RemoveNumbers()
                EndIf
                $iConverted += 1
            EndIf
        Else
            ExitLoop ; Tranh vong lap vo han
        EndIf
    WEnd

    ; Buoc 4: Bao cao ket qua
    Local $sMsg = "CHUYEN NUMBERING THANH TEXT" & @CRLF & @CRLF
    $sMsg &= "Tong cong: " & $iTotal & " lists" & @CRLF
    $sMsg &= "Thanh cong: " & $iConverted & @CRLF
    If $iFailed > 0 Then $sMsg &= "Loi: " & $iFailed & @CRLF

    _LogPreview($sMsg)
    _UpdateProgress("Da chuyen " & $iConverted & "/" & $iTotal & " lists!")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", $sMsg)
EndFunc

; Convert Bullet to Text
Func _ConvertBulletToText()
    _ConvertNumberingToText()
EndFunc

; Convert Numbering to Text (Selection)
; FIX: Them IsObj check, xu ly insertion point, xoa list formatting
Func _ConvertNumberingToTextSelection()
    If Not _CheckConnection() Then Return
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then
        MsgBox($MB_ICONWARNING, "Loi", "Khong lay duoc Selection!")
        Return
    EndIf

    ; Kiem tra selection co text khong
    If $oSel.Type = 1 Then ; wdSelectionIP = insertion point
        MsgBox($MB_ICONWARNING, "Loi", "Vui long boi den (select) vung van ban co numbering!")
        Return
    EndIf

    Local $oRange = $oSel.Range
    If Not IsObj($oRange) Then Return

    If $oRange.ListFormat.ListType <> 0 Then ; 0 = wdListNoNumbering
        _UpdateProgress("Dang chuyen vung chon...")
        $oRange.ListFormat.ConvertNumbersToText()
        ; Xoa list formatting con sot lai
        $oRange.ListFormat.RemoveNumbers()
        _UpdateProgress("Da chuyen vung chon thanh text!")
    Else
        MsgBox($MB_ICONWARNING, "Thong bao", _
            "Vung chon khong co Numbering/Bullet!" & @CRLF & @CRLF & _
            "Huong dan:" & @CRLF & _
            "1. Boi den cac dong co so thu tu hoac bullet" & @CRLF & _
            "2. Nhan lai nut nay")
    EndIf
EndFunc

; Export to PDF
Func _ExportToPDF()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Xuat PDF", @DesktopDir, "PDF (*.pdf)", 16, _
        StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "") & ".pdf")
    If @error Then Return
    
    _UpdateProgress("Dang xuat PDF...")
    _ExportCurrentDocumentToPath($sPath, $WD_EXPORT_PDF)
    _UpdateProgress("Da xuat PDF!")
    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da xuat: " & $sPath)
EndFunc

; Export to HTML
Func _ExportToHTML()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Xuat HTML", @DesktopDir, "HTML (*.html)", 16, _
        StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "") & ".html")
    If @error Then Return
    
    _UpdateProgress("Dang xuat HTML...")
    _ExportCurrentDocumentToPath($sPath, $WD_FORMAT_HTML)
    _UpdateProgress("Da xuat HTML!")
    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da xuat: " & $sPath)
EndFunc

; Export to TXT
Func _ExportToTXT()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Xuat TXT", @DesktopDir, "Text (*.txt)", 16, _
        StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "") & ".txt")
    If @error Then Return
    
    _UpdateProgress("Dang xuat TXT...")
    _ExportCurrentDocumentToPath($sPath, $WD_FORMAT_TEXT)
    _UpdateProgress("Da xuat TXT!")
    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da xuat: " & $sPath)
EndFunc

; Export to RTF
Func _ExportToRTF()
    If Not _CheckConnection() Then Return
    Local $sPath = FileSaveDialog("Xuat RTF", @DesktopDir, "RTF (*.rtf)", 16, _
        StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "") & ".rtf")
    If @error Then Return
    
    _UpdateProgress("Dang xuat RTF...")
    _ExportCurrentDocumentToPath($sPath, $WD_FORMAT_RTF)
    _UpdateProgress("Da xuat RTF!")
    MsgBox($MB_ICONINFORMATION, "Thanh cong", "Da xuat: " & $sPath)
EndFunc

Func _ExportCurrentDocumentToPath($sPath, $iFormat)
    If Not _CheckConnection() Then Return False
    If $sPath = "" Then Return False

    Local $sSourcePath = ""
    Local $bDeleteSource = False

    If $iFormat = $WD_EXPORT_PDF Then
        $g_oDoc.ExportAsFixedFormat($sPath, $iFormat)
        Return (Not @error And FileExists($sPath))
    EndIf

    If Not _PrepareDocumentPathForExport($g_oDoc, $sSourcePath, $bDeleteSource) Then Return False

    Local $oExportDoc = $g_oWord.Documents.Open($sSourcePath, False, False)
    If Not IsObj($oExportDoc) Then
        If $bDeleteSource And FileExists($sSourcePath) Then FileDelete($sSourcePath)
        Return False
    EndIf

    $oExportDoc.SaveAs2($sPath, $iFormat)
    Local $bOk = (Not @error And FileExists($sPath))
    $oExportDoc.Close(0)

    If IsObj($g_oDoc) Then $g_oDoc.Activate()

    If $bDeleteSource And FileExists($sSourcePath) Then FileDelete($sSourcePath)
    Return $bOk
EndFunc

Func _PrepareDocumentPathForExport($oDoc, ByRef $sSourcePath, ByRef $bDeleteSource)
    If Not IsObj($oDoc) Then Return False

    $sSourcePath = ""
    $bDeleteSource = False

    Local $sCurrentPath = $oDoc.FullName
    If $sCurrentPath <> "" And FileExists($sCurrentPath) Then
        $oDoc.Save()
        $sSourcePath = $sCurrentPath
        Return FileExists($sSourcePath)
    EndIf

    Local $sTempDir = @TempDir & "\PDFToWordFixer\Exports"
    DirCreate($sTempDir)
    $sSourcePath = $sTempDir & "\ExportSource_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & @SEC & "_" & @MSEC & ".docx"

    Local $oTempDoc = $g_oWord.Documents.Add()
    If Not IsObj($oTempDoc) Then Return False

    $oDoc.Content.Copy()
    $oTempDoc.Range(0, 0).Paste()
    $oTempDoc.PageSetup.LeftMargin = $oDoc.PageSetup.LeftMargin
    $oTempDoc.PageSetup.RightMargin = $oDoc.PageSetup.RightMargin
    $oTempDoc.PageSetup.TopMargin = $oDoc.PageSetup.TopMargin
    $oTempDoc.PageSetup.BottomMargin = $oDoc.PageSetup.BottomMargin
    $oTempDoc.SaveAs2($sSourcePath, 16)
    $oTempDoc.Close(0)

    If @error Or Not FileExists($sSourcePath) Then Return False

    $bDeleteSource = True
    Return True
EndFunc

; Show Print Preview
Func _ShowPrintPreview()
    If Not _CheckConnection() Then Return
    $g_oDoc.PrintPreview()
EndFunc

; Compare Documents
; FIX: Them error handling cho CompareDocuments COM call
Func _CompareDocuments()
    If Not _CheckConnection() Then Return
    Local $sPath = FileOpenDialog("Chon file de so sanh", @ScriptDir, "Word (*.docx;*.doc)", 1)
    If @error Then Return
    _UpdateProgress("Dang so sanh...")

    Local $oCompareDoc = _CompareDocumentsByPath($sPath)
    If Not IsObj($oCompareDoc) Then
        _UpdateProgress("Loi khi so sanh!")
        MsgBox($MB_ICONWARNING, "Loi", "Khong the so sanh 2 file!" & @CRLF & _
            "Co the do 2 file khong tuong thich hoac bi loi.")
        Return
    EndIf

    $oCompareDoc.Activate()
    _UpdateProgress("Da tao file so sanh!")
    MsgBox($MB_ICONINFORMATION, "So sanh", "Da tao file so sanh moi!" & @CRLF & @CRLF & _
        "Cac thay doi duoc danh dau:" & @CRLF & _
        "- Mau do: Xoa" & @CRLF & _
        "- Mau xanh: Them moi")
EndFunc

; Merge Documents
Func _MergeDocuments()
    If Not _CheckConnection() Then Return

    Local $sChoice = InputBox("Gop file", "Chon vi tri gop:" & @CRLF & _
        "1 - Chen vao cuoi file" & @CRLF & _
        "2 - Chen vao dau file" & @CRLF & _
        "3 - Chen tai vi tri con tro", "1")
    If @error Then Return

    Local $sPath = FileOpenDialog("Chon file de gop", @ScriptDir, "Word (*.docx;*.doc)", 1)
    If @error Then Return

    _UpdateProgress("Dang gop file...")
    If Not _MergeDocumentFromPath($sPath, $sChoice) Then
        MsgBox($MB_ICONWARNING, "Gop file", "Khong the gop file da chon!")
        Return
    EndIf
    _UpdateProgress("Da gop file!")
    MsgBox($MB_ICONINFORMATION, "Gop file", "Da gop file thanh cong!")
EndFunc

; Split Document
; FIX: Re-activate file goc sau khi tach, dung Selection moi thay vi $oSel cu
Func _SplitDocument()
    If Not _CheckConnection() Then Return
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Or $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Huong dan", "Chon phan van ban can tach truoc!" & @CRLF & @CRLF & _
            "Cach dung:" & @CRLF & _
            "1. Boi den phan van ban can tach" & @CRLF & _
            "2. Nhan nut 'Tach file'" & @CRLF & _
            "3. Chon noi luu file moi")
        Return
    EndIf

    Local $sBaseName = StringRegExpReplace($g_oDoc.Name, "\.[^.]+$", "")
    Local $sPath = FileSaveDialog("Luu phan tach", @ScriptDir, "Word (*.docx)", 16, $sBaseName & "_tach.docx")
    If @error Then Return

    _UpdateProgress("Dang tach file...")

    Local $bDeleteFromSource = (MsgBox($MB_YESNO, "Xoa phan da tach?", "Xoa phan van ban da tach khoi file goc?") = $IDYES)
    If Not _SplitRangeToPath($oSel.Range, $sPath, $bDeleteFromSource) Then
        MsgBox($MB_ICONWARNING, "Tach file", "Khong the tach vung da chon!")
        Return
    EndIf

    _UpdateProgress("Da tach file!")
    MsgBox($MB_ICONINFORMATION, "Tach file thanh cong", "Da luu phan tach tai:" & @CRLF & @CRLF & $sPath)
EndFunc

; Protect Document
; FIX: Them error handling cho Unprotect (sai mat khau throw COM error)
Func _ProtectDocument()
    If Not _CheckConnection() Then Return
    
    ; Kiem tra trang thai bao ve hien tai
    Local $bProtected = $g_oDoc.ProtectionType <> -1 ; -1 = wdNoProtection

    If $bProtected Then
        ; File dang duoc bao ve -> hoi bo bao ve
        Local $sPassword = InputBox("Bo bao ve", "Nhap mat khau de bo bao ve:" & @CRLF & _
            "(De trong neu khong co mat khau)", "", "*")
        If @error Then Return
        
        ; QUAN TRONG: Unprotect() throw COM error neu sai mat khau
        ; _ComErrorHandler se bat va set @error = 1
        If Not _RemoveDocumentProtection($sPassword) Then
            MsgBox($MB_ICONWARNING, "Loi", "Mat khau khong dung hoac khong the bo bao ve!")
            Return
        EndIf
        
        If $g_oDoc.ProtectionType = -1 Then
            _UpdateProgress("Da bo bao ve!")
            MsgBox($MB_ICONINFORMATION, "Bo bao ve", "Da bo bao ve file thanh cong!")
        Else
            MsgBox($MB_ICONWARNING, "Loi", "Mat khau khong dung!")
        EndIf
    Else
        ; File chua bao ve -> hoi bao ve
        Local $sChoice = InputBox("Bao ve file", "Chon loai bao ve:" & @CRLF & _
            "1 - Chi doc (Read Only)" & @CRLF & _
            "2 - Chi cho dien form" & @CRLF & _
            "3 - Chi cho comment" & @CRLF & _
            "4 - Chi cho track changes", "1")
        If @error Then Return

        Local $sPassword = InputBox("Mat khau", "Nhap mat khau bao ve (co the de trong):", "", "*")
        If @error Then Return

        Local $iProtectType = 0
        Switch $sChoice
            Case "1"
                $iProtectType = 3 ; wdAllowOnlyReading
            Case "2"
                $iProtectType = 2 ; wdAllowOnlyFormFields
            Case "3"
                $iProtectType = 1 ; wdAllowOnlyComments
            Case "4"
                $iProtectType = 0 ; wdAllowOnlyRevisions
            Case Else
                $iProtectType = 3
        EndSwitch

        If Not _SetDocumentProtection($iProtectType, $sPassword) Then
            MsgBox($MB_ICONWARNING, "Bao ve", "Khong the bao ve file!")
            Return
        EndIf
        _UpdateProgress("Da bao ve file!")
        MsgBox($MB_ICONINFORMATION, "Bao ve", "Da bao ve file thanh cong!")
    EndIf
EndFunc

Func _CompareDocumentsByPath($sPath)
    If Not _CheckConnection() Then Return 0
    If $sPath = "" Or Not FileExists($sPath) Then Return 0

    Local $oDoc2 = $g_oWord.Documents.Open($sPath, False, True)
    If Not IsObj($oDoc2) Then Return 0

    Local Const $wdCompareTargetNew = 2
    Local Const $wdGranularityWordLevel = 1
    Local $oCompareDoc = $g_oWord.CompareDocuments($g_oDoc, $oDoc2, $wdCompareTargetNew, $wdGranularityWordLevel, True)
    $oDoc2.Close(0)

    If @error Or Not IsObj($oCompareDoc) Then Return 0
    Return $oCompareDoc
EndFunc

Func _MergeDocumentFromPath($sPath, $sChoice = "1")
    If Not _CheckConnection() Then Return False
    If $sPath = "" Or Not FileExists($sPath) Then Return False

    Local $oRange = 0
    Switch $sChoice
        Case "1"
            $oRange = $g_oDoc.Content
            $oRange.Collapse($WD_COLLAPSE_END)
            $oRange.InsertParagraphAfter()
            $oRange.Collapse($WD_COLLAPSE_END)
        Case "2"
            $oRange = $g_oDoc.Range(0, 0)
        Case "3"
            $oRange = $g_oWord.Selection.Range
        Case Else
            $oRange = $g_oDoc.Content
            $oRange.Collapse($WD_COLLAPSE_END)
    EndSwitch

    If Not IsObj($oRange) Then Return False
    $oRange.InsertFile($sPath)
    Return (Not @error)
EndFunc

Func _SplitRangeToPath($oRange, $sPath, $bDeleteFromSource = False)
    If Not _CheckConnection() Then Return False
    If Not IsObj($oRange) Or $sPath = "" Then Return False

    Local $iStart = $oRange.Start
    Local $iEnd = $oRange.End

    $oRange.Copy()
    Local $oNewDoc = $g_oWord.Documents.Add()
    If Not IsObj($oNewDoc) Then Return False

    $oNewDoc.Content.Paste()
    $oNewDoc.SaveAs2($sPath)
    Local $bOk = (Not @error And FileExists($sPath))
    $oNewDoc.Close(0)

    If IsObj($g_oDoc) Then $g_oDoc.Activate()
    If Not $bOk Then Return False

    If $bDeleteFromSource Then
        Local $oNewRange = $g_oDoc.Range($iStart, $iEnd)
        $oNewRange.Delete()
    EndIf

    Return True
EndFunc

Func _SetDocumentProtection($iProtectType, $sPassword = "")
    If Not _CheckConnection() Then Return False
    $g_oDoc.Protect($iProtectType, False, $sPassword)
    Return (Not @error And $g_oDoc.ProtectionType <> -1)
EndFunc

Func _RemoveDocumentProtection($sPassword = "")
    If Not _CheckConnection() Then Return False
    $g_oDoc.Unprotect($sPassword)
    Return (Not @error And $g_oDoc.ProtectionType = -1)
EndFunc

; Show Document Properties
Func _ShowDocProperties()
    If Not _CheckConnection() Then Return
    
    Local $sMsg = "THUOC TINH FILE:" & @CRLF & @CRLF
    $sMsg &= "Ten: " & $g_oDoc.Name & @CRLF
    $sMsg &= "Duong dan: " & $g_oDoc.Path & @CRLF
    $sMsg &= "So trang: " & $g_oDoc.ComputeStatistics(2) & @CRLF
    $sMsg &= "So tu: " & $g_oDoc.ComputeStatistics(0) & @CRLF
    
    _LogPreview($sMsg)
    MsgBox($MB_ICONINFORMATION, "Thuoc tinh", $sMsg)
EndFunc

; Clean Document
Func _CleanDocument()
    If Not _CheckConnection() Then Return
    If MsgBox($MB_YESNO, "Don dep file", _
        "Se thuc hien:" & @CRLF & _
        "- Xoa comments" & @CRLF & _
        "- Xoa hyperlinks" & @CRLF & _
        "- Chap nhan thay doi" & @CRLF & _
        "Tiep tuc?") <> $IDYES Then Return
    
    _UpdateProgress("Dang don dep...")

    _CleanDocumentCore()
    _UpdateProgress("Da don dep file!")
EndFunc

Func _CleanDocumentCore()
    If Not _CheckConnection() Then Return False

    While $g_oDoc.Comments.Count > 0
        $g_oDoc.Comments.Item(1).Delete()
    WEnd

    While $g_oDoc.Hyperlinks.Count > 0
        $g_oDoc.Hyperlinks.Item(1).Delete()
    WEnd

    $g_oDoc.AcceptAllRevisions()
    Return (Not @error)
EndFunc
