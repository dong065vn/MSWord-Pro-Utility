; ============================================
; FORMAT.AU3 - Module Dinh dang
; ============================================

#include-once

; Ap dung dinh dang
Func _ApplyFormat($bSelOnly)
    If Not _CheckConnection() Then Return
    Local $oRange = $bSelOnly ? $g_oWord.Selection.Range : $g_oDoc.Content
    If Not IsObj($oRange) Then Return
    If Not $bSelOnly And MsgBox($MB_YESNO, "Xac nhan", "Ap dung toan bo?") <> $IDYES Then Return

    _UpdateProgress("Dang ap dung dinh dang...")

    $oRange.Font.Name = GUICtrlRead($g_cboFont)
    $oRange.Font.Size = Number(GUICtrlRead($g_cboFontSize))

    Local $fLS = Number(GUICtrlRead($g_cboLineSpacing))
    $oRange.ParagraphFormat.LineSpacingRule = $WD_LINE_SPACE_MULTIPLE
    $oRange.ParagraphFormat.LineSpacing = 12 * $fLS
    $oRange.ParagraphFormat.Alignment = _GetAlignmentConst(GUICtrlRead($g_cboAlignment))

    If GUICtrlRead($g_chkAutoFirstLine) = $GUI_CHECKED Then
        $oRange.ParagraphFormat.FirstLineIndent = 1.27 * $CM_TO_POINTS
    EndIf

    If Not $bSelOnly Then
        $g_oDoc.PageSetup.LeftMargin = Number(GUICtrlRead($g_inputLeftMargin)) * $CM_TO_POINTS
        $g_oDoc.PageSetup.RightMargin = Number(GUICtrlRead($g_inputRightMargin)) * $CM_TO_POINTS
        $g_oDoc.PageSetup.TopMargin = Number(GUICtrlRead($g_inputTopMargin)) * $CM_TO_POINTS
        $g_oDoc.PageSetup.BottomMargin = Number(GUICtrlRead($g_inputBottomMargin)) * $CM_TO_POINTS
    EndIf

    _UpdateProgress("Da ap dung dinh dang!")
EndFunc

; Preset VN
Func _LoadPresetVN()
    GUICtrlSetData($g_cboFont, "Times New Roman")
    GUICtrlSetData($g_cboFontSize, "13")
    GUICtrlSetData($g_cboLineSpacing, "1.5")
    GUICtrlSetData($g_cboAlignment, "Justify")
    GUICtrlSetData($g_inputLeftMargin, "3.5")
    GUICtrlSetData($g_inputRightMargin, "2")
    GUICtrlSetData($g_inputTopMargin, "2.5")
    GUICtrlSetData($g_inputBottomMargin, "2.5")
    GUICtrlSetState($g_chkAutoFirstLine, $GUI_CHECKED)
    _UpdateProgress("Da tai chuan VN (Times 13pt, 1.5, 3.5-2-2.5-2.5)")
EndFunc

; Preset US
Func _LoadPresetUS()
    GUICtrlSetData($g_cboFont, "Times New Roman")
    GUICtrlSetData($g_cboFontSize, "12")
    GUICtrlSetData($g_cboLineSpacing, "2.0")
    GUICtrlSetData($g_cboAlignment, "Left")
    GUICtrlSetData($g_inputLeftMargin, "2.54")
    GUICtrlSetData($g_inputRightMargin, "2.54")
    GUICtrlSetData($g_inputTopMargin, "2.54")
    GUICtrlSetData($g_inputBottomMargin, "2.54")
    GUICtrlSetState($g_chkAutoFirstLine, $GUI_CHECKED)
    _UpdateProgress("Da tai chuan US/APA (Times 12pt, 2.0, 1 inch)")
EndFunc

; Format Heading
Func _FormatHeading($iLevel)
    If Not _CheckConnection() Then Return
    Local $oSel = $g_oWord.Selection
    If Not IsObj($oSel) Then Return

    $oSel.Style = "Heading " & $iLevel
    $oSel.Font.Name = "Times New Roman"

    Switch $iLevel
        Case 1
            $oSel.Font.Size = 14
            $oSel.Font.Bold = True
            $oSel.ParagraphFormat.Alignment = $WD_ALIGN_CENTER
        Case 2
            $oSel.Font.Size = 13
            $oSel.Font.Bold = True
            $oSel.ParagraphFormat.Alignment = $WD_ALIGN_LEFT
        Case 3
            $oSel.Font.Size = 13
            $oSel.Font.Bold = True
            $oSel.Font.Italic = True
            $oSel.ParagraphFormat.Alignment = $WD_ALIGN_LEFT
    EndSwitch
    _UpdateProgress("Da ap dung Heading " & $iLevel)
EndFunc


; Format Caption
Func _FormatCaption()
    If Not _CheckConnection() Then Return
    Local $oSel = $g_oWord.Selection
    $oSel.Font.Name = "Times New Roman"
    $oSel.Font.Size = 12
    $oSel.Font.Italic = True
    $oSel.ParagraphFormat.Alignment = $WD_ALIGN_CENTER
    _UpdateProgress("Da ap dung Caption")
EndFunc

; Format Normal
Func _FormatNormal()
    If Not _CheckConnection() Then Return
    Local $oSel = $g_oWord.Selection
    $oSel.Style = "Normal"
    $oSel.Font.Name = "Times New Roman"
    $oSel.Font.Size = 13
    $oSel.Font.Bold = False
    $oSel.Font.Italic = False
    _UpdateProgress("Da ap dung Normal")
EndFunc

; Clear Format
Func _ClearFormat()
    If Not _CheckConnection() Then Return
    $g_oWord.Selection.ClearFormatting()
    _UpdateProgress("Da xoa dinh dang")
EndFunc

; Remove Highlight
Func _RemoveHighlight()
    If Not _CheckConnection() Then Return
    $g_oDoc.Content.HighlightColorIndex = 0
    _UpdateProgress("Da xoa highlight")
EndFunc

; Unify Font
Func _UnifyFont()
    If Not _CheckConnection() Then Return
    $g_oDoc.Content.Font.Name = "Times New Roman"
    $g_oDoc.Content.Font.Size = 13
    _UpdateProgress("Da thong nhat font")
EndFunc

; Fix All Spacing
Func _FixAllSpacing()
    If Not _CheckConnection() Then Return
    $g_oDoc.Content.ParagraphFormat.LineSpacingRule = $WD_LINE_SPACE_MULTIPLE
    $g_oDoc.Content.ParagraphFormat.LineSpacing = 18
    _UpdateProgress("Da sua gian dong 1.5")
EndFunc

; Add Page Numbers
Func _AddPageNumbers()
    If Not _CheckConnection() Then Return
    Local $sChoice = InputBox("Them so trang", _
        "Chon vi tri:" & @CRLF & _
        "1 - Giua duoi" & @CRLF & _
        "2 - Phai duoi" & @CRLF & _
        "3 - Giua tren" & @CRLF & _
        "4 - Phai tren", "1")
    If @error Or $sChoice = "" Then Return

    _UpdateProgress("Dang them so trang...")
    Local $oSections = $g_oDoc.Sections
    For $i = 1 To $oSections.Count
        Local $oSec = $oSections.Item($i)
        Switch $sChoice
            Case "1"
                $oSec.Footers(1).PageNumbers.Add(1)
            Case "2"
                $oSec.Footers(1).PageNumbers.Add(2)
            Case "3"
                $oSec.Headers(1).PageNumbers.Add(1)
            Case "4"
                $oSec.Headers(1).PageNumbers.Add(2)
        EndSwitch
    Next
    _UpdateProgress("Da them so trang!")
EndFunc

; Add Header
Func _AddHeader()
    If Not _CheckConnection() Then Return
    Local $sHeader = InputBox("Them Header", "Nhap noi dung:", "")
    If @error Then Return

    _UpdateProgress("Dang them header...")
    Local $oSections = $g_oDoc.Sections
    For $i = 1 To $oSections.Count
        Local $oH = $oSections.Item($i).Headers(1)
        $oH.Range.Text = $sHeader
        $oH.Range.Font.Name = "Times New Roman"
        $oH.Range.Font.Size = 12
        $oH.Range.ParagraphFormat.Alignment = $WD_ALIGN_CENTER
    Next
    _UpdateProgress("Da them header!")
EndFunc

; Remove Page Numbers
Func _RemovePageNumbers()
    If Not _CheckConnection() Then Return
    Local $oSections = $g_oDoc.Sections
    For $i = 1 To $oSections.Count
        Local $oSec = $oSections.Item($i)
        While $oSec.Footers(1).PageNumbers.Count > 0
            $oSec.Footers(1).PageNumbers.Item(1).Delete()
        WEnd
        While $oSec.Headers(1).PageNumbers.Count > 0
            $oSec.Headers(1).PageNumbers.Item(1).Delete()
        WEnd
    Next
    _UpdateProgress("Da xoa so trang")
EndFunc


; Auto Number Images
Func _AutoNumberImages()
    If Not _CheckConnection() Then Return
    Local $sPrefix = GUICtrlRead($g_inputCaptionPrefix)
    If $sPrefix = "" Then $sPrefix = "Hinh"

    Local $oShapes = $g_oDoc.InlineShapes
    If Not IsObj($oShapes) Or $oShapes.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co hinh anh!")
        Return
    EndIf

    Local $n = 0
    For $i = 1 To $oShapes.Count
        Local $oShape = $oShapes.Item($i)
        If Not IsObj($oShape) Then ContinueLoop
        
        ; Type: 3=wdInlineShapePicture, 5=wdInlineShapeLinkedPicture, 7=wdInlineShapeEmbeddedOLEObject
        Local $iType = $oShape.Type
        If $iType = 3 Or $iType = 5 Or $iType = 7 Then
            $n += 1
            Local $oR = $oShape.Range
            $oR.Collapse($WD_COLLAPSE_END)
            $oR.InsertParagraphAfter()
            $oR.Collapse($WD_COLLAPSE_END)
            $oR.Text = $sPrefix & " " & $n & ". "
            $oR.Font.Italic = True
            $oR.Font.Size = 12
            $oR.ParagraphFormat.Alignment = $WD_ALIGN_CENTER
        EndIf
    Next
    _UpdateProgress("Da danh so " & $n & " hinh!")
EndFunc

; Auto Number Tables
Func _AutoNumberTables()
    If Not _CheckConnection() Then Return
    Local $sPrefix = InputBox("Caption", "Tien to:", "Bang")
    If @error Or $sPrefix = "" Then Return

    Local $oTables = $g_oDoc.Tables
    If Not IsObj($oTables) Or $oTables.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co bang!")
        Return
    EndIf

    For $i = 1 To $oTables.Count
        Local $oR = $oTables.Item($i).Range
        $oR.Collapse($WD_COLLAPSE_START)
        $oR.InsertParagraphBefore()
        $oR.Collapse($WD_COLLAPSE_START)
        $oR.Text = $sPrefix & " " & $i & ". "
        $oR.Font.Bold = True
        $oR.Font.Size = 12
        $oR.ParagraphFormat.Alignment = $WD_ALIGN_CENTER
    Next
    _UpdateProgress("Da danh so " & $oTables.Count & " bang!")
EndFunc

; Number Equations
Func _NumberEquations()
    If Not _CheckConnection() Then Return
    Local $oOMaths = $g_oDoc.OMaths
    If Not IsObj($oOMaths) Or $oOMaths.Count = 0 Then
        MsgBox($MB_ICONWARNING, "Thong bao", "Khong co cong thuc!")
        Return
    EndIf

    Local $sChapter = GUICtrlRead($g_inputChapterNum)
    If $sChapter = "" Then $sChapter = "1"

    For $i = 1 To $oOMaths.Count
        Local $oR = $oOMaths.Item($i).Range
        $oR.Collapse($WD_COLLAPSE_END)
        $oR.InsertAfter(@TAB & "(" & $sChapter & "." & $i & ")")
    Next
    _UpdateProgress("Da danh so " & $oOMaths.Count & " cong thuc!")
EndFunc

; Remove Equation Numbers
Func _RemoveEquationNumbers()
    If Not _CheckConnection() Then Return
    Local $oFind = $g_oDoc.Content.Find
    $oFind.ClearFormatting()
    $oFind.Replacement.ClearFormatting()
    ; Pattern: TAB + (number.number) - escape special chars for wildcard
    $oFind.Text = "^t\([0-9].[0-9]*\)"
    $oFind.Replacement.Text = ""
    $oFind.MatchWildcards = True
    $oFind.Execute("", False, False, True, False, False, True, 1, False, "", $WD_REPLACE_ALL)
    _UpdateProgress("Da xoa so cong thuc!")
EndFunc

; Check Thesis Format
Func _CheckThesisFormat()
    If Not _CheckConnection() Then Return
    _UpdateProgress("Dang kiem tra dinh dang...")

    Local $sMsg = "=== KIEM TRA DINH DANG DO AN ===" & @CRLF & @CRLF
    Local $iIssues = 0, $iOK = 0

    ; Check Font - 9999999 means mixed fonts (wdUndefined)
    Local $sFont = $g_oDoc.Content.Font.Name
    If $sFont = "Times New Roman" Then
        $sMsg &= "[OK] Font: Times New Roman" & @CRLF
        $iOK += 1
    ElseIf $sFont = "" Or $sFont = 9999999 Then
        $sMsg &= "[!] Font: Hon hop (can thong nhat Times New Roman)" & @CRLF
        $iIssues += 1
    Else
        $sMsg &= "[!] Font: " & $sFont & " (can Times New Roman)" & @CRLF
        $iIssues += 1
    EndIf

    ; Check Margins
    Local $oPS = $g_oDoc.PageSetup
    Local $fLeft = Round($oPS.LeftMargin / $CM_TO_POINTS, 1)
    Local $fRight = Round($oPS.RightMargin / $CM_TO_POINTS, 1)

    If Abs($fLeft - 3.5) <= 0.2 Then
        $sMsg &= "[OK] Le trai: " & $fLeft & " cm" & @CRLF
        $iOK += 1
    Else
        $sMsg &= "[!] Le trai: " & $fLeft & " cm (can 3.5 cm)" & @CRLF
        $iIssues += 1
    EndIf

    If Abs($fRight - 2) <= 0.2 Then
        $sMsg &= "[OK] Le phai: " & $fRight & " cm" & @CRLF
        $iOK += 1
    Else
        $sMsg &= "[!] Le phai: " & $fRight & " cm (can 2 cm)" & @CRLF
        $iIssues += 1
    EndIf

    ; Stats
    $sMsg &= @CRLF & "THONG TIN:" & @CRLF
    $sMsg &= "  So trang: " & $g_oDoc.ComputeStatistics(2) & @CRLF
    $sMsg &= "  So tu: " & $g_oDoc.ComputeStatistics(0) & @CRLF
    $sMsg &= @CRLF & "TONG KET: " & $iOK & " OK, " & $iIssues & " can sua"

    _LogPreview($sMsg)
    _UpdateProgress("Kiem tra xong: " & $iIssues & " van de")
    MsgBox($iIssues > 0 ? $MB_ICONWARNING : $MB_ICONINFORMATION, "Ket qua", $sMsg)
EndFunc
