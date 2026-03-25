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
; 11. HEADING NUMBER CLEANUP - Sua dau cham de muc
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

