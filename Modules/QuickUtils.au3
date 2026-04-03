; ============================================
; QUICKUTILS.AU3 - Module Tien ich Nhanh
; Cac chuc nang tien ich bo sung cho Word
; Version: 6.1
; ============================================

#include-once

; ============================================
; 1. SMART PASTE - Dan thong minh
; ============================================

; Dan van ban khong dinh dang (chi text)
Func _PasteAsPlainText()
    If Not _CheckConnection() Then Return
    
    _UpdateProgress("Dang dan van ban...")
    
    ; wdPasteText = 2
    $g_oWord.Selection.PasteSpecial(False, False, 0, False, 2)
    
    _UpdateProgress("Da dan van ban (khong dinh dang)!")
EndFunc

; Dan va giu dinh dang nguon
Func _PasteKeepSourceFormat()
    If Not _CheckConnection() Then Return
    
    _UpdateProgress("Dang dan van ban...")
    
    ; wdFormatOriginalFormatting = 16
    $g_oWord.Selection.PasteAndFormat(16)
    
    _UpdateProgress("Da dan van ban (giu dinh dang nguon)!")
EndFunc

; Dan va hop nhat dinh dang
Func _PasteMergeFormat()
    If Not _CheckConnection() Then Return
    
    _UpdateProgress("Dang dan van ban...")
    
    ; wdFormatSurroundingFormattingWithEmphasis = 20
    $g_oWord.Selection.PasteAndFormat(20)
    
    _UpdateProgress("Da dan van ban (hop nhat dinh dang)!")
EndFunc

; ============================================
; 2. QUICK SELECTION - Chon nhanh
; ============================================

; Chon toan bo doan van hien tai
Func _SelectCurrentParagraph()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    ; Mo rong selection den toan bo paragraph
    $oSel.Expand(4) ; wdParagraph = 4
    
    _UpdateProgress("Da chon doan van hien tai")
EndFunc

; Chon toan bo cau hien tai
Func _SelectCurrentSentence()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    ; Mo rong selection den toan bo sentence
    $oSel.Expand(3) ; wdSentence = 3
    
    _UpdateProgress("Da chon cau hien tai")
EndFunc

; Chon tu dau den vi tri hien tai
Func _SelectFromStart()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    Local $iEnd = $oSel.End
    $g_oDoc.Range(0, $iEnd).Select()
    
    _UpdateProgress("Da chon tu dau den vi tri hien tai")
EndFunc

; Chon tu vi tri hien tai den cuoi
Func _SelectToEnd()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    Local $iStart = $oSel.Start
    Local $iEnd = $g_oDoc.Content.End
    $g_oDoc.Range($iStart, $iEnd).Select()
    
    _UpdateProgress("Da chon tu vi tri hien tai den cuoi")
EndFunc

; ============================================
; 3. QUICK NAVIGATION - Di chuyen nhanh
; ============================================

; Nhay den trang cu the
Func _GoToPage()
    If Not _CheckConnection() Then Return
    
    Local $iPage = InputBox("Nhay den trang", "Nhap so trang:", "1")
    If @error Or $iPage = "" Then Return
    
    $iPage = Int($iPage)
    If $iPage < 1 Then Return
    
    ; wdGoToPage = 1
    $g_oWord.Selection.GoTo(1, 0, $iPage)
    
    _UpdateProgress("Da nhay den trang " & $iPage)
EndFunc

; Nhay den Heading tiep theo
Func _GoToNextHeading()
    If Not _CheckConnection() Then Return
    
    ; wdGoToHeading = 11
    $g_oWord.Selection.GoTo(11, 2) ; wdGoToNext = 2
    
    _UpdateProgress("Da nhay den Heading tiep theo")
EndFunc

; Nhay den Heading truoc do
Func _GoToPrevHeading()
    If Not _CheckConnection() Then Return
    
    ; wdGoToHeading = 11
    $g_oWord.Selection.GoTo(11, 3) ; wdGoToPrevious = 3
    
    _UpdateProgress("Da nhay den Heading truoc do")
EndFunc

; Nhay den bang tiep theo
Func _GoToNextTable()
    If Not _CheckConnection() Then Return
    
    ; wdGoToTable = 2
    $g_oWord.Selection.GoTo(2, 2) ; wdGoToNext = 2
    
    _UpdateProgress("Da nhay den bang tiep theo")
EndFunc

; Nhay den hinh tiep theo
Func _GoToNextImage()
    If Not _CheckConnection() Then Return
    
    ; wdGoToGraphic = 8
    $g_oWord.Selection.GoTo(8, 2) ; wdGoToNext = 2
    
    _UpdateProgress("Da nhay den hinh tiep theo")
EndFunc

; ============================================
; 4. QUICK INSERT - Chen nhanh
; ============================================

; Chen ngat trang
Func _InsertPageBreak()
    If Not _CheckConnection() Then Return
    
    ; wdPageBreak = 7
    $g_oWord.Selection.InsertBreak(7)
    
    _UpdateProgress("Da chen ngat trang")
EndFunc

; Chen ngat section
Func _InsertSectionBreak()
    If Not _CheckConnection() Then Return
    
    Local $sChoice = InputBox("Chen Section Break", _
        "Chon loai:" & @CRLF & _
        "1 - Trang moi (Next Page)" & @CRLF & _
        "2 - Lien tuc (Continuous)" & @CRLF & _
        "3 - Trang chan (Even Page)" & @CRLF & _
        "4 - Trang le (Odd Page)", "1")
    If @error Then Return
    
    Local $iType = 2 ; wdSectionBreakNextPage
    Switch $sChoice
        Case "1"
            $iType = 2 ; wdSectionBreakNextPage
        Case "2"
            $iType = 0 ; wdSectionBreakContinuous
        Case "3"
            $iType = 3 ; wdSectionBreakEvenPage
        Case "4"
            $iType = 4 ; wdSectionBreakOddPage
    EndSwitch
    
    $g_oWord.Selection.InsertBreak($iType)
    
    _UpdateProgress("Da chen Section Break")
EndFunc

; Chen ngay hien tai
Func _InsertCurrentDate()
    If Not _CheckConnection() Then Return
    
    Local $sDate = @MDAY & "/" & @MON & "/" & @YEAR
    $g_oWord.Selection.TypeText($sDate)
    
    _UpdateProgress("Da chen ngay: " & $sDate)
EndFunc

; Chen gio hien tai
Func _InsertCurrentTime()
    If Not _CheckConnection() Then Return
    
    Local $sTime = @HOUR & ":" & @MIN & ":" & @SEC
    $g_oWord.Selection.TypeText($sTime)
    
    _UpdateProgress("Da chen gio: " & $sTime)
EndFunc

; Chen duong ke ngang
Func _InsertHorizontalLine()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    $oSel.TypeParagraph()
    
    ; Tao duong ke bang border
    $oSel.Paragraphs(1).Borders(-3).LineStyle = 1 ; wdLineStyleSingle, wdBorderBottom = -3
    $oSel.Paragraphs(1).Borders(-3).LineWidth = 8 ; wdLineWidth050pt
    
    $oSel.TypeParagraph()
    
    _UpdateProgress("Da chen duong ke ngang")
EndFunc

; ============================================
; 5. QUICK FORMAT - Dinh dang nhanh
; ============================================

; Tang kich thuoc font
Func _IncreaseFontSize()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    Local $fSize = $oSel.Font.Size
    If $fSize <> 9999999 Then ; Khong phai mixed
        $oSel.Font.Size = $fSize + 1
        _UpdateProgress("Font size: " & ($fSize + 1) & "pt")
    EndIf
EndFunc

; Giam kich thuoc font
Func _DecreaseFontSize()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    Local $fSize = $oSel.Font.Size
    If $fSize <> 9999999 And $fSize > 1 Then
        $oSel.Font.Size = $fSize - 1
        _UpdateProgress("Font size: " & ($fSize - 1) & "pt")
    EndIf
EndFunc

; Toggle Bold
Func _ToggleBold()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    $oSel.Font.Bold = Not $oSel.Font.Bold
    _UpdateProgress($oSel.Font.Bold ? "Bold: ON" : "Bold: OFF")
EndFunc

; Toggle Italic
Func _ToggleItalic()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    $oSel.Font.Italic = Not $oSel.Font.Italic
    _UpdateProgress($oSel.Font.Italic ? "Italic: ON" : "Italic: OFF")
EndFunc

; Toggle Underline
Func _ToggleUnderline()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    ; wdUnderlineSingle = 1, wdUnderlineNone = 0
    $oSel.Font.Underline = ($oSel.Font.Underline = 0) ? 1 : 0
    _UpdateProgress($oSel.Font.Underline ? "Underline: ON" : "Underline: OFF")
EndFunc

; Toggle Subscript
Func _ToggleSubscript()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    $oSel.Font.Subscript = Not $oSel.Font.Subscript
    _UpdateProgress($oSel.Font.Subscript ? "Subscript: ON" : "Subscript: OFF")
EndFunc

; Toggle Superscript
Func _ToggleSuperscript()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    $oSel.Font.Superscript = Not $oSel.Font.Superscript
    _UpdateProgress($oSel.Font.Superscript ? "Superscript: ON" : "Superscript: OFF")
EndFunc

; ============================================
; 6. PARAGRAPH TOOLS - Cong cu doan van
; ============================================

; Tang thut le doan van
Func _IncreaseIndent()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    Local $fIndent = $oSel.ParagraphFormat.LeftIndent
    $oSel.ParagraphFormat.LeftIndent = $fIndent + 36 ; 0.5 inch = 36pt
    
    _UpdateProgress("Da tang thut le")
EndFunc

; Giam thut le doan van
Func _DecreaseIndent()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    Local $fIndent = $oSel.ParagraphFormat.LeftIndent
    If $fIndent >= 36 Then
        $oSel.ParagraphFormat.LeftIndent = $fIndent - 36
    Else
        $oSel.ParagraphFormat.LeftIndent = 0
    EndIf
    
    _UpdateProgress("Da giam thut le")
EndFunc

; Xoa thut dong dau tien
Func _RemoveFirstLineIndent()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    $oSel.ParagraphFormat.FirstLineIndent = 0
    
    _UpdateProgress("Da xoa thut dong dau tien")
EndFunc

; Dat thut dong dau tien 1.27cm
Func _SetFirstLineIndent()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return
    
    $oSel.ParagraphFormat.FirstLineIndent = 1.27 * $CM_TO_POINTS
    
    _UpdateProgress("Da dat thut dong dau tien 1.27cm")
EndFunc

; ============================================
; 7. SPECIAL CHARACTERS - Ky tu dac biet
; ============================================

; Chen ky tu dac biet
Func _InsertSpecialChar()
    If Not _CheckConnection() Then Return
    
    Local $sChoice = InputBox("Chen ky tu dac biet", _
        "Chon ky tu:" & @CRLF & _
        "1 - Em Dash (—)" & @CRLF & _
        "2 - En Dash (–)" & @CRLF & _
        "3 - Non-breaking Space" & @CRLF & _
        "4 - Ellipsis (…)" & @CRLF & _
        "5 - Copyright (©)" & @CRLF & _
        "6 - Registered (®)" & @CRLF & _
        "7 - Trademark (™)" & @CRLF & _
        "8 - Degree (°)" & @CRLF & _
        "9 - Plus-Minus (±)", "1")
    If @error Then Return
    
    Local $sChar = ""
    Switch $sChoice
        Case "1"
            $sChar = ChrW(8212) ; Em Dash
        Case "2"
            $sChar = ChrW(8211) ; En Dash
        Case "3"
            $sChar = ChrW(160)  ; Non-breaking Space
        Case "4"
            $sChar = ChrW(8230) ; Ellipsis
        Case "5"
            $sChar = ChrW(169)  ; Copyright
        Case "6"
            $sChar = ChrW(174)  ; Registered
        Case "7"
            $sChar = ChrW(8482) ; Trademark
        Case "8"
            $sChar = ChrW(176)  ; Degree
        Case "9"
            $sChar = ChrW(177)  ; Plus-Minus
        Case Else
            Return
    EndSwitch
    
    $g_oWord.Selection.TypeText($sChar)
    _UpdateProgress("Da chen ky tu dac biet")
EndFunc

; ============================================
; 8. BOOKMARK TOOLS - Cong cu Bookmark
; ============================================

; Them bookmark tai vi tri hien tai
Func _AddBookmark()
    If Not _CheckConnection() Then Return
    
    Local $sName = InputBox("Them Bookmark", "Nhap ten bookmark:", "Bookmark1")
    If @error Or $sName = "" Then Return
    
    ; Lam sach ten (chi cho phep chu cai, so va _)
    $sName = StringRegExpReplace($sName, "[^a-zA-Z0-9_]", "_")
    
    Local $oSel = $g_oWord.Selection
    $g_oDoc.Bookmarks.Add($sName, $oSel.Range)
    
    _UpdateProgress("Da them bookmark: " & $sName)
EndFunc

; Nhay den bookmark
Func _GoToBookmark()
    If Not _CheckConnection() Then Return
    
    Local $oBookmarks = $g_oDoc.Bookmarks
    If $oBookmarks.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co bookmark nao!")
        Return
    EndIf
    
    ; Tao danh sach bookmark
    Local $sList = ""
    For $i = 1 To $oBookmarks.Count
        $sList &= $i & ". " & $oBookmarks.Item($i).Name & @CRLF
    Next
    
    Local $sChoice = InputBox("Nhay den Bookmark", _
        "Danh sach bookmark:" & @CRLF & $sList & @CRLF & _
        "Nhap so thu tu:", "1")
    If @error Then Return
    
    Local $iIndex = Int($sChoice)
    If $iIndex < 1 Or $iIndex > $oBookmarks.Count Then Return
    
    $oBookmarks.Item($iIndex).Select()
    _UpdateProgress("Da nhay den bookmark: " & $oBookmarks.Item($iIndex).Name)
EndFunc

; Xoa bookmark
Func _DeleteBookmark()
    If Not _CheckConnection() Then Return
    
    Local $oBookmarks = $g_oDoc.Bookmarks
    If $oBookmarks.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co bookmark nao!")
        Return
    EndIf
    
    ; Tao danh sach bookmark
    Local $sList = ""
    For $i = 1 To $oBookmarks.Count
        $sList &= $i & ". " & $oBookmarks.Item($i).Name & @CRLF
    Next
    
    Local $sChoice = InputBox("Xoa Bookmark", _
        "Danh sach bookmark:" & @CRLF & $sList & @CRLF & _
        "Nhap so thu tu de xoa:", "1")
    If @error Then Return
    
    Local $iIndex = Int($sChoice)
    If $iIndex < 1 Or $iIndex > $oBookmarks.Count Then Return
    
    Local $sName = $oBookmarks.Item($iIndex).Name
    $oBookmarks.Item($iIndex).Delete()
    _UpdateProgress("Da xoa bookmark: " & $sName)
EndFunc

; ============================================
; 9. DOCUMENT INFO - Thong tin tai lieu
; ============================================

; Hien thi thong tin chi tiet
Func _ShowDetailedDocInfo()
    If Not _CheckConnection() Then Return
    
    Local $sMsg = "=== THONG TIN CHI TIET TAI LIEU ===" & @CRLF & @CRLF
    
    ; Thong tin co ban
    $sMsg &= "TEN FILE: " & $g_oDoc.Name & @CRLF
    $sMsg &= "DUONG DAN: " & $g_oDoc.Path & @CRLF & @CRLF
    
    ; Thong ke
    $sMsg &= "THONG KE:" & @CRLF
    $sMsg &= "  So trang: " & $g_oDoc.ComputeStatistics(2) & @CRLF
    $sMsg &= "  So tu: " & $g_oDoc.ComputeStatistics(0) & @CRLF
    $sMsg &= "  So ky tu (co dau cach): " & $g_oDoc.ComputeStatistics(3) & @CRLF
    $sMsg &= "  So ky tu (khong dau cach): " & $g_oDoc.ComputeStatistics(5) & @CRLF
    $sMsg &= "  So doan van: " & $g_oDoc.ComputeStatistics(4) & @CRLF
    $sMsg &= "  So dong: " & $g_oDoc.ComputeStatistics(1) & @CRLF & @CRLF
    
    ; Doi tuong
    $sMsg &= "DOI TUONG:" & @CRLF
    $sMsg &= "  So bang: " & $g_oDoc.Tables.Count & @CRLF
    $sMsg &= "  So hinh (InlineShapes): " & $g_oDoc.InlineShapes.Count & @CRLF
    $sMsg &= "  So hinh (Shapes): " & $g_oDoc.Shapes.Count & @CRLF
    $sMsg &= "  So hyperlink: " & $g_oDoc.Hyperlinks.Count & @CRLF
    $sMsg &= "  So bookmark: " & $g_oDoc.Bookmarks.Count & @CRLF
    $sMsg &= "  So comment: " & $g_oDoc.Comments.Count & @CRLF
    $sMsg &= "  So section: " & $g_oDoc.Sections.Count & @CRLF & @CRLF
    
    ; Page Setup
    Local $oPS = $g_oDoc.PageSetup
    $sMsg &= "THIET LAP TRANG:" & @CRLF
    $sMsg &= "  Le trai: " & Round($oPS.LeftMargin / $CM_TO_POINTS, 2) & " cm" & @CRLF
    $sMsg &= "  Le phai: " & Round($oPS.RightMargin / $CM_TO_POINTS, 2) & " cm" & @CRLF
    $sMsg &= "  Le tren: " & Round($oPS.TopMargin / $CM_TO_POINTS, 2) & " cm" & @CRLF
    $sMsg &= "  Le duoi: " & Round($oPS.BottomMargin / $CM_TO_POINTS, 2) & " cm" & @CRLF
    $sMsg &= "  Kich thuoc giay: " & Round($oPS.PageWidth / $CM_TO_POINTS, 2) & " x " & _
        Round($oPS.PageHeight / $CM_TO_POINTS, 2) & " cm" & @CRLF
    
    _LogPreview($sMsg)
    MsgBox($MB_ICONINFORMATION, "Thong tin tai lieu", $sMsg)
EndFunc

; ============================================
; 10. QUICK CLEANUP - Don dep nhanh
; ============================================

; Xoa tat ca highlight trong vung chon
Func _RemoveHighlightSelection()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Or $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon vung van ban truoc!")
        Return
    EndIf
    
    $oSel.Range.HighlightColorIndex = 0
    _UpdateProgress("Da xoa highlight trong vung chon")
EndFunc

; Xoa tat ca comment trong vung chon
Func _RemoveCommentsSelection()
    If Not _CheckConnection() Then Return
    
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Or $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon vung van ban truoc!")
        Return
    EndIf
    
    Local $oComments = $oSel.Comments
    Local $n = $oComments.Count
    While $oComments.Count > 0
        $oComments.Item(1).Delete()
    WEnd
    
    _UpdateProgress("Da xoa " & $n & " comment trong vung chon")
EndFunc

; Xoa tat ca field codes (chuyen thanh text)
Func _UnlinkAllFields()
    If Not _CheckConnection() Then Return
    
    If MsgBox($MB_YESNO, "Xac nhan", "Chuyen tat ca Field thanh text?" & @CRLF & _
        "(Khong the hoan tac)") <> $IDYES Then Return
    
    Local $oFields = $g_oDoc.Fields
    Local $n = $oFields.Count
    
    For $i = $n To 1 Step -1
        $oFields.Item($i).Unlink()
    Next
    
    _UpdateProgress("Da chuyen " & $n & " field thanh text")
EndFunc

; ============================================
; 11. CITATION CLEANUP - Xoa trich dan [4], (Nguyen, 2020)
; ============================================

Func _RemoveBracketCitationsSelection()
    If Not _CheckConnection() Then Return

    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Or $oSel.Type = 1 Then
        MsgBox($MB_ICONWARNING, "Loi", "Chon vung van ban truoc khi xoa trich dan!")
        Return
    EndIf

    _RemoveBracketCitationsInRange($oSel.Range, "vung chon")
EndFunc

Func _RemoveBracketCitationsDocument()
    If Not _CheckConnection() Then Return
    _RemoveBracketCitationsInRange($g_oDoc.Content, "toan bo tai lieu")
EndFunc

Func _PreviewBracketCitations()
    If Not _CheckConnection() Then Return

    Local $oRange = 0
    Local $sScopeLabel = "toan bo tai lieu"
    Local $oSel = $g_oWord.Selection
    If IsObj($oSel) And $oSel.Type <> 1 Then
        $oRange = $oSel.Range
        $sScopeLabel = "vung chon"
    Else
        $oRange = $g_oDoc.Content
    EndIf

    Local $sFilter = StringStripWS(GUICtrlRead($g_inputCitationFilter), 3)
    Local $aFilter = _ParseCitationFilter($sFilter)
    If $sFilter <> "" And @error Then
        MsgBox($MB_ICONWARNING, "Loi", _
            "Bo loc citation khong hop le." & @CRLF & @CRLF & _
            "Vi du hop le:" & @CRLF & _
            "- 2,5" & @CRLF & _
            "- 2;5;7" & @CRLF & _
            "- 2, 5-7")
        Return
    EndIf

    Local $iMode = _GetCitationMode()
    Local $aMatches = _CollectCitationMatches($oRange.Text, $aFilter, $iMode)
    If Not IsArray($aMatches) Or UBound($aMatches, 1) = 0 Then
        MsgBox($MB_ICONINFORMATION, "Xem truoc citation", _
            "Khong tim thay citation phu hop trong " & $sScopeLabel & "." & @CRLF & @CRLF & _
            _BuildCitationModeHint($iMode) & @CRLF & _
            _BuildCitationFilterHint($sFilter))
        Return
    EndIf

    Local $sPreview = "XEM TRUOC CITATION SE BI XOA - " & StringUpper($sScopeLabel) & @CRLF & @CRLF & _
        _BuildCitationModeHint($iMode) & @CRLF & _
        _BuildCitationFilterHint($sFilter) & @CRLF & _
        "Tong so: " & UBound($aMatches, 1) & @CRLF & @CRLF

    Local $iLimit = 40
    For $i = 0 To UBound($aMatches, 1) - 1
        If $i = $iLimit Then
            $sPreview &= "... va " & (UBound($aMatches, 1) - $iLimit) & " citation khac"
            ExitLoop
        EndIf
        $sPreview &= ($i + 1) & ". " & $aMatches[$i][2] & " | " & _CitationPreviewContext($oRange.Text, $aMatches[$i][0], $aMatches[$i][1]) & @CRLF
    Next

    _LogPreview($sPreview)
    MsgBox($MB_ICONINFORMATION, "Xem truoc citation", $sPreview)
EndFunc

Func _RemoveBracketCitationsInRange($oRange, $sScopeLabel)
    If Not IsObj($oRange) Then Return

    Local $sFilter = StringStripWS(GUICtrlRead($g_inputCitationFilter), 3)
    Local $aFilter = _ParseCitationFilter($sFilter)
    If $sFilter <> "" And @error Then
        MsgBox($MB_ICONWARNING, "Loi", _
            "Bo loc citation khong hop le." & @CRLF & @CRLF & _
            "Vi du hop le:" & @CRLF & _
            "- 2,5" & @CRLF & _
            "- 2;5;7" & @CRLF & _
            "- 2, 5-7")
        Return
    EndIf

    Local $iMode = _GetCitationMode()

    _UpdateProgress("Dang xoa trich dan [n] trong " & $sScopeLabel & "...")

    Local $sText = $oRange.Text
    Local $aMatches = _CollectCitationMatches($sText, $aFilter, $iMode)
    If Not IsArray($aMatches) Or UBound($aMatches, 1) = 0 Then
        _UpdateProgress("Khong tim thay trich dan [n] trong " & $sScopeLabel)
        MsgBox($MB_ICONINFORMATION, "Thong bao", _
            "Khong tim thay trich dan phu hop trong " & $sScopeLabel & "." & @CRLF & @CRLF & _
            _BuildCitationModeHint($iMode) & @CRLF & _
            _BuildCitationFilterHint($sFilter))
        Return
    EndIf

    Local $iRemoved = _DeleteCitationMatchesInWordRange($oRange, $aMatches)

    If $iRemoved = 0 Then
        _UpdateProgress("Khong xoa duoc citation trong " & $sScopeLabel)
        MsgBox($MB_ICONWARNING, "Thong bao", _
            "Da tim thay citation, nhung Word khong cho phep sua noi dung trong " & $sScopeLabel & ".")
        Return
    EndIf

    _CleanupCitationSpacing($oRange)
    _UpdateProgress("Da xoa " & $iRemoved & " trich dan trong " & $sScopeLabel)
    MsgBox($MB_ICONINFORMATION, "Hoan tat", _
        "Da xoa " & $iRemoved & " trich dan trong " & $sScopeLabel & "." & @CRLF & @CRLF & _
        _BuildCitationModeHint($iMode) & @CRLF & _
        _BuildCitationFilterHint($sFilter) & @CRLF & @CRLF & _
        "Ho tro cac dang:" & @CRLF & _
        "- [4]" & @CRLF & _
        "- [12, 15]" & @CRLF & _
        "- [3-5]" & @CRLF & _
        "- [4][5][6]" & @CRLF & _
        "- (Nguyen, 2020)" & @CRLF & _
        "- (Smith et al., 2021; Tran, 2022)")
EndFunc

Func _CollectCitationMatches($sText, $aFilter = 0, $iMode = 0)
    Local $aMatches[0][3]
    Local $iLen = StringLen($sText)
    Local $iPos = 1

    While $iPos <= $iLen
        Local $sOpen = StringMid($sText, $iPos, 1)
        Local $sClose = ""
        If $sOpen = "[" Then
            $sClose = "]"
        ElseIf $sOpen = "(" Then
            $sClose = ")"
        Else
            $iPos += 1
            ContinueLoop
        EndIf

        Local $iClose = StringInStr($sText, $sClose, 0, 1, $iPos + 1)
        If $iClose = 0 Then ExitLoop

        Local $sInner = StringMid($sText, $iPos + 1, $iClose - $iPos - 1)
        Local $iCitationType = _GetCitationType($sInner, $sOpen, $aFilter, $iMode)
        If $iCitationType <> 0 Then
            Local $iDeleteStart = $iPos
            Local $sPrev = ""
            Local $sNext = ""
            If $iPos > 1 Then $sPrev = StringMid($sText, $iPos - 1, 1)
            If $iClose < $iLen Then $sNext = StringMid($sText, $iClose + 1, 1)

            If $sPrev = " " Then
                If $sNext = "" Or $sNext = @CR Or $sNext = @LF Or $sNext = " " Or StringRegExp($sNext, "[\.,;:\!\?\)\]]") Then
                    $iDeleteStart -= 1
                EndIf
            EndIf

            Local $iCount = UBound($aMatches, 1)
            ReDim $aMatches[$iCount + 1][3]
            $aMatches[$iCount][0] = $iDeleteStart
            $aMatches[$iCount][1] = $iClose
            $aMatches[$iCount][2] = $sOpen & $sInner & $sClose
        EndIf

        $iPos = $iClose + 1
    WEnd

    Return $aMatches
EndFunc

Func _CitationPreviewContext($sText, $iStart1Based, $iEnd1Based)
    Local $iContext = 28
    Local $iStart = $iStart1Based - $iContext
    Local $iLen = ($iEnd1Based - $iStart1Based + 1) + ($iContext * 2)

    If $iStart < 1 Then
        $iLen -= (1 - $iStart)
        $iStart = 1
    EndIf
    If $iLen < 1 Then $iLen = 1

    Local $sSnippet = StringMid($sText, $iStart, $iLen)
    $sSnippet = StringReplace($sSnippet, @CR, " ")
    $sSnippet = StringReplace($sSnippet, @LF, " ")
    $sSnippet = StringRegExpReplace($sSnippet, "\s+", " ")
    $sSnippet = StringStripWS($sSnippet, 3)
    If $iStart > 1 Then $sSnippet = "..." & $sSnippet
    If ($iStart + $iLen - 1) < StringLen($sText) Then $sSnippet &= "..."
    Return $sSnippet
EndFunc

Func _DeleteCitationMatchesInWordRange($oScopeRange, $aMatches)
    If Not IsObj($oScopeRange) Or Not IsArray($aMatches) Then Return 0

    Local $iRemoved = 0
    For $i = 0 To UBound($aMatches, 1) - 1
        If _DeleteFirstCitationTextInScope($oScopeRange, $aMatches[$i][2]) Then
            $iRemoved += 1
        EndIf
    Next

    Return $iRemoved
EndFunc

Func _DeleteFirstCitationTextInScope($oScopeRange, $sCitationText)
    If Not IsObj($oScopeRange) Or $sCitationText = "" Then Return False

    Local $oSearch = $oScopeRange.Duplicate
    Local $oFind = $oSearch.Find
    If Not IsObj($oFind) Then Return False

    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()

    Local $bFound = $oFind.Execute($sCitationText, False, False, False, False, False, True, 1, False, "", 0)
    If Not $bFound Then Return False

    Local $iStart = $oSearch.Start
    Local $iEnd = $oSearch.End
    Local $sPrev = ""
    Local $sNext = ""

    If $iStart > 0 Then $sPrev = $g_oDoc.Range($iStart - 1, $iStart).Text
    If $iEnd < $g_oDoc.Content.End Then $sNext = $g_oDoc.Range($iEnd, $iEnd + 1).Text

    If $sPrev = " " Then
        If $sNext = "" Or $sNext = @CR Or $sNext = @LF Or $sNext = " " Or StringRegExp($sNext, "[\.,;:\!\?\)\]]") Then
            $iStart -= 1
        EndIf
    EndIf

    Local $oDeleteRange = $g_oDoc.Range($iStart, $iEnd)
    If Not IsObj($oDeleteRange) Then Return False

    $oDeleteRange.Text = ""
    Return True
EndFunc

Func _IsBracketCitationText($sInner)
    Local $sValue = StringStripWS($sInner, 3)
    If $sValue = "" Then Return False

    Return StringRegExp($sValue, "^\d+(?:\s*[-,;]\s*\d+)*$")
EndFunc

Func _GetCitationType($sInner, $sOpen, $aFilter = 0, $iMode = 0)
    Local $sValue = StringStripWS($sInner, 3)
    If $sValue = "" Then Return 0

    If $sOpen = "[" Then
        If Not _IsBracketCitationText($sValue) Then Return 0
        If $iMode = 2 Then Return 0
        If IsArray($aFilter) And Not _BracketCitationMatchesFilter($sValue, $aFilter) Then Return 0
        Return 1
    EndIf

    ; Khi co bo loc [n], chi xoa citation kieu [n], bo qua author-year.
    If IsArray($aFilter) Then Return 0
    If $iMode = 1 Then Return 0

    If StringLen($sValue) > 120 Then Return 0
    If Not StringRegExp($sValue, "(19|20)\d{2}[a-z]?") Then Return 0
    If Not StringRegExp($sValue, "[A-Za-z]") Then Return 0
    If Not StringRegExp($sValue, "[,;]") Then Return 0
    If StringRegExp($sValue, "^[\d\s,;.-]+$") Then Return 0

    ; Cac dang pho bien: Nguyen, 2020 | Smith et al., 2021; Tran, 2022
    If StringRegExp($sValue, "^[A-Za-zÀ-ỹĐđ][^()]{0,110}(19|20)\d{2}[a-z]?$") Then Return 2

    Return 0
EndFunc

Func _ParseCitationFilter($sFilter)
    If $sFilter = "" Then Return 0

    Local $sNormalized = StringRegExpReplace($sFilter, "[;\|]+", ",")
    $sNormalized = StringRegExpReplace($sNormalized, "\s+", "")
    If $sNormalized = "" Then Return SetError(1, 0, 0)

    Local $aParts = StringSplit($sNormalized, ",", 2)
    If Not IsArray($aParts) Or UBound($aParts) = 0 Then Return SetError(1, 0, 0)

    Local $aFilter[0][2]
    For $i = 0 To UBound($aParts) - 1
        Local $sPart = $aParts[$i]
        If $sPart = "" Then ContinueLoop

        Local $iStart = 0, $iEnd = 0
        If StringRegExp($sPart, "^\d+$") Then
            $iStart = Int($sPart)
            $iEnd = $iStart
        ElseIf StringRegExp($sPart, "^\d+-\d+$") Then
            Local $aRange = StringSplit($sPart, "-", 2)
            $iStart = Int($aRange[0])
            $iEnd = Int($aRange[1])
            If $iEnd < $iStart Then Return SetError(1, 0, 0)
        Else
            Return SetError(1, 0, 0)
        EndIf

        Local $iCount = UBound($aFilter, 1)
        ReDim $aFilter[$iCount + 1][2]
        $aFilter[$iCount][0] = $iStart
        $aFilter[$iCount][1] = $iEnd
    Next

    If UBound($aFilter, 1) = 0 Then Return SetError(1, 0, 0)
    Return $aFilter
EndFunc

Func _BracketCitationMatchesFilter($sValue, $aFilter)
    Local $sNormalized = StringRegExpReplace($sValue, "\s+", "")
    Local $aParts = StringSplit($sNormalized, ",", 2)
    If Not IsArray($aParts) Then Return False

    For $i = 0 To UBound($aParts) - 1
        If $aParts[$i] = "" Then ContinueLoop

        Local $iStart = 0, $iEnd = 0
        If StringRegExp($aParts[$i], "^\d+$") Then
            $iStart = Int($aParts[$i])
            $iEnd = $iStart
        ElseIf StringRegExp($aParts[$i], "^\d+-\d+$") Then
            Local $aRange = StringSplit($aParts[$i], "-", 2)
            $iStart = Int($aRange[0])
            $iEnd = Int($aRange[1])
        Else
            Return False
        EndIf

        For $j = 0 To UBound($aFilter, 1) - 1
            If $iStart <= $aFilter[$j][1] And $iEnd >= $aFilter[$j][0] Then Return True
        Next
    Next

    Return False
EndFunc

Func _BuildCitationFilterHint($sFilter)
    If $sFilter = "" Then Return "Bo loc [n]: tat ca."
    Return "Bo loc dang dung: [" & $sFilter & "]"
EndFunc

Func _GetCitationMode()
    Local $sMode = GUICtrlRead($g_cboCitationMode)
    Switch $sMode
        Case "Chi [n]"
            Return 1
        Case "Chi tac gia-nam"
            Return 2
    EndSwitch
    Return 0
EndFunc

Func _BuildCitationModeHint($iMode)
    Switch $iMode
        Case 1
            Return "Che do: chi xoa citation so [n]."
        Case 2
            Return "Che do: chi xoa citation tac gia-nam (Author, 2020)."
    EndSwitch
    Return "Che do: xoa tat ca citation ho tro."
EndFunc

Func _CleanupCitationSpacing($oRange)
    Local $oTarget = $g_oDoc.Content
    If IsObj($oRange) Then $oTarget = $oRange.Duplicate

    Local $oFind = $oTarget.Find
    If Not IsObj($oFind) Then Return

    For $i = 1 To 3
        $oFind.ClearFormatting()
        $oFind.Replacement.ClearFormatting()
        $oFind.Execute("  ", False, False, False, False, False, True, 1, False, " ", $WD_REPLACE_ALL)
    Next

    Local $aPairs[7][2] = [[" .", "."], [" ,", ","], [" ;", ";"], [" :", ":"], [" !", "!"], [" ?", "?"], [" )", ")"]]
    For $i = 0 To UBound($aPairs, 1) - 1
        $oFind.ClearFormatting()
        $oFind.Replacement.ClearFormatting()
        $oFind.Execute($aPairs[$i][0], False, False, False, False, False, True, 1, False, $aPairs[$i][1], $WD_REPLACE_ALL)
    Next
EndFunc

; ============================================
; 12. HEADING NUMBER CLEANUP - Sua dau cham de muc
; ============================================

; Chuan hoa tien to de muc:
; - "2.4. Tieu de" -> "2.4 Tieu de"
; - "2. 4 Tieu de" -> "2.4 Tieu de"
; - "2 . 4 . Tieu de" -> "2.4. Tieu de" neu separator = ". "
; Co the gioi han theo tien to, vi du "2." / "2" / "2.4" / "2.4."
Func _FixHeadingNumberDots()
    If Not _CheckConnection() Then Return

    Local $sPrefix = StringStripWS(GUICtrlRead($g_inputHeadingPrefixFix), 3)
    Local $sSeparator = GUICtrlRead($g_inputHeadingSeparatorFix)
    If $sSeparator = "" Then $sSeparator = " "

    Local $sPrefixNormalized = _NormalizeHeadingPrefixFilter($sPrefix)
    If $sPrefix <> "" And @error Then
        MsgBox($MB_ICONWARNING, "Loi", _
            "Tien to khong hop le." & @CRLF & @CRLF & _
            "Chap nhan cac dang:" & @CRLF & _
            "- 1" & @CRLF & _
            "- 1." & @CRLF & _
            "- 1.2" & @CRLF & _
            "- 1.2." & @CRLF & _
            "- 1.2.3")
        Return
    EndIf

    $sSeparator = _ResolveHeadingSeparatorAlias($sSeparator)
    _UpdateProgress("Dang sua de muc so...")

    Local $oParas = $g_oDoc.Paragraphs
    Local $iFixed = 0
    Local $iScanned = 0

    For $i = 1 To $oParas.Count
        Local $oPara = $oParas.Item($i)
        If Not IsObj($oPara) Then ContinueLoop

        Local $sRawText = $oPara.Range.Text
        If $sRawText = "" Then ContinueLoop

        $iScanned += 1
        Local $sNewText = _NormalizeHeadingParagraphPrefixText($sRawText, $sPrefixNormalized, $sSeparator)
        If @error Then ContinueLoop
        If $sNewText = $sRawText Then ContinueLoop

        $oPara.Range.Text = $sNewText
        $iFixed += 1
    Next

    _UpdateProgress("Da sua " & $iFixed & " de muc")
    MsgBox($MB_ICONINFORMATION, "Hoan tat", _
        "Da sua " & $iFixed & " de muc." & @CRLF & @CRLF & _
        "Da quet: " & $iScanned & " doan." & @CRLF & @CRLF & _
        "Ho tro cac dang:" & @CRLF & _
        "- ""2 Tieu de""" & @CRLF & _
        "- ""2. Tieu de""" & @CRLF & _
        "- ""2.4 Tieu de""" & @CRLF & _
        "- ""2. 4 Tieu de""" & @CRLF & _
        "- ""2.4. Tieu de""" & @CRLF & _
        "- ""2 . 4 . Tieu de""" & @CRLF & @CRLF & _
        "Meo dung nhanh:" & @CRLF & _
        "- Tien to ""2"" hoac ""2."" -> chi sua chuong 2" & @CRLF & _
        "- Tien to ""2.4"" hoac ""2.4."" -> chi sua nhanh muc 2.4.x" & @CRLF & _
        "- Ngat sau so = "" "" -> 2.4 Tieu de" & @CRLF & _
        "- Ngat sau so = "". "" -> 2.4. Tieu de" & @CRLF & _
        "- Ngat sau so = "" - "" -> 2.4 - Tieu de" & @CRLF & _
        "- Ngat sau so = "": "" -> 2.4: Tieu de")
EndFunc

Func _NormalizeHeadingPrefixFilter($sPrefix)
    Local $sClean = StringStripWS($sPrefix, 3)
    If $sClean = "" Then Return ""

    $sClean = StringRegExpReplace($sClean, "\s*\.\s*", ".")
    If Not StringRegExp($sClean, "^\d+(?:\.\d+)*\.?$") Then Return SetError(1, 0, "")
    If StringRight($sClean, 1) <> "." Then $sClean &= "."
    Return $sClean
EndFunc

Func _ResolveHeadingSeparatorAlias($sSeparator)
    Local $sValue = $sSeparator
    Switch StringLower($sValue)
        Case "\t", "{tab}", "<tab>", "tab"
            Return @TAB
        Case "\s", "{space}", "<space>", "space"
            Return " "
        Case "\none", "{none}", "<none>", "none"
            Return ""
    EndSwitch
    Return $sValue
EndFunc

Func _NormalizeHeadingParagraphPrefixText($sRawText, $sPrefixNormalized, $sSeparator)
    Local $sEol = ""
    Local $sBody = $sRawText
    If StringRight($sBody, 2) = @CRLF Then
        $sEol = @CRLF
        $sBody = StringTrimRight($sBody, 2)
    ElseIf StringRight($sBody, 1) = @CR Then
        $sEol = @CR
        $sBody = StringTrimRight($sBody, 1)
    ElseIf StringRight($sBody, 1) = @LF Then
        $sEol = @LF
        $sBody = StringTrimRight($sBody, 1)
    EndIf

    Local $aMatch = StringRegExp($sBody, "^(\s*)((?:\d+\s*(?:\.\s*\d+)*))(\s*(?:\.)?)\s+(.+?)\s*$", 1)
    If @error Or Not IsArray($aMatch) Or UBound($aMatch) < 4 Then Return SetError(1, 0, $sRawText)

    Local $sLeading = $aMatch[0]
    Local $sNumberBlock = $aMatch[1]
    Local $sTrailingMark = $aMatch[2]
    Local $sRest = $aMatch[3]

    Local $sNormalizedNumber = StringRegExpReplace(StringStripWS($sNumberBlock, 3), "\s*\.\s*", ".")
    Local $sFullPrefix = $sNormalizedNumber
    If StringStripWS($sTrailingMark, 3) = "." Then $sFullPrefix &= "."

    If $sPrefixNormalized <> "" Then
        Local $sPrefixNoDot = StringTrimRight($sPrefixNormalized, 1)
        Local $bMatchPrefix = (StringLeft($sFullPrefix, StringLen($sPrefixNormalized)) = $sPrefixNormalized)
        If Not $bMatchPrefix Then $bMatchPrefix = (StringLeft($sNormalizedNumber, StringLen($sPrefixNoDot)) = $sPrefixNoDot)
        If Not $bMatchPrefix Then Return SetError(2, 0, $sRawText)
    EndIf

    Return $sLeading & $sNormalizedNumber & $sSeparator & $sRest & $sEol
EndFunc

